# -*- coding: utf-8 -*-
"""
Markdown 转纯文本小工具
- 图形界面：左侧输入 Markdown，右侧实时预览（标题/粗体/斜体/代码块带格式显示）
- 「复制（公式可编辑）」：Markdown → Pandoc → HTML(MathML) → OMML 条件注释 → 剪贴板，
  粘贴进 OneNote/Word 即为原生可编辑公式（免 Word，毫秒级；无 Pandoc 时自动降级内置转换器）
- 「复制（公式为图片）」：公式渲染为 PNG 图片的富文本（无需 Word），任何软件中公式均为图片
- 「复制源MD」：原样复制输入框中的 Markdown 源文本
- 列表自动编号：1. → (1) → a.，每个列表组独立计数
- 文件菜单：打开/保存 .md/.txt（Ctrl+O / Ctrl+S）；支持拖拽文件到输入框导入
- 内容与设置（窗口大小、开关状态）自动记忆
"""
import base64
import ctypes
import html
import io
import json
import os
import re
import shutil
import subprocess
import sys
import time
import unicodedata
from ctypes import wintypes
from html.entities import name2codepoint
import tkinter as tk
from tkinter import filedialog, ttk

try:
    import latex2mathml.converter
    import mathml2omml
    _HAS_LATEX = True
except ImportError:
    _HAS_LATEX = False

try:  # 拖拽支持（缺失时自动降级为不可拖拽）
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except ImportError:
    _HAS_DND = False

# ---------------- 占位符（私有区字符，避免与正文冲突） ----------------
PH_OPEN, PH_CLOSE = '\ue000', '\ue001'
PH_RE = re.compile(r'\ue000(\d+)\ue001')
# 链接 <a> 标签专用占位符（仅 HTML 富文本路径内部使用）
A_PH_OPEN, A_PH_CLOSE = '\ue002', '\ue003'
A_PH_RE = re.compile(r'\ue002(\d+)\ue003')

OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def _display_width(s):
    """计算显示宽度：全角字符算 2，半角算 1"""
    return sum(2 if unicodedata.east_asian_width(c) in 'FW' else 1 for c in s)


def _fix_math_delimiters(text):
    """公式分隔符规范化（修复 AI 输出的非标准写法，代码块已先被占位符保护）"""
    # $ x $ 内侧带空格 → $x$（否则下方 $ 正则不匹配，公式原样残留）
    text = re.sub(r'(?<!\$)\$(?!\$)[ \t]+([^\n$]+?)[ \t]+(?<!\$)\$(?!\$)',
                  r'$\1$', text)
    # 单独成行的成对 $ → $$（非标准块级公式分隔符）
    lines = text.split('\n')
    out = []
    in_block = False
    for line in lines:
        if re.fullmatch(r'\s*\$\s*', line):
            out.append(f"{line[:line.find('$')]}$$")
            in_block = not in_block
        else:
            out.append(line)
    return '\n'.join(out)


def _protect(text):
    """把公式、代码块、行内代码替换为占位符；存为 {kind, text, latex}"""
    store = []

    def add(kind, whole, latex=None, display=False):
        store.append({'kind': kind, 'text': whole, 'latex': latex,
                      'display': display})
        return f'{PH_OPEN}{len(store) - 1}{PH_CLOSE}'

    # 代码块 ```...```（内容原样保留）
    text = re.sub(r'```[^\n]*\n(.*?)(?:\n)?```',
                  lambda m: add('code', m.group(1)), text, flags=re.S)
    # 行内代码
    text = re.sub(r'`([^`\n]+)`', lambda m: add('code', m.group(1)), text)
    # 公式分隔符规范化
    text = _fix_math_delimiters(text)
    # 公式：$$...$$（可跨行，块级）
    text = re.sub(r'\$\$(.+?)\$\$',
                  lambda m: add('math', m.group(0), m.group(1), True), text, flags=re.S)
    # 公式：\[...\]（块级）
    text = re.sub(r'\\\[(.*?)\\\]',
                  lambda m: add('math', m.group(0), m.group(1), True), text, flags=re.S)
    # 公式：$...$（单行，内容不含 $、首尾不为空格）
    text = re.sub(r'(?<!\\)\$([^\s$](?:[^$\n]*[^\s$])?)\$',
                  lambda m: add('math', m.group(0), m.group(1)), text)
    # 公式：\(...\)
    text = re.sub(r'\\\((.*?)\\\)',
                  lambda m: add('math', m.group(0), m.group(1)), text)
    return text, store


def _restore(text, store):
    """纯文本模式：占位符还原为原文（公式带 $ 定界符）"""
    return PH_RE.sub(lambda m: store[int(m.group(1))]['text'], text)


def _plain_line(line, store):
    """纯文本视角的一行：先剥行内标记（粗体/斜体/链接等），再还原公式/代码占位符"""
    return _restore(_inline(line), store)


def _sanitize_latex(latex):
    """清洗 AI 生成 LaTeX 的常见噪声，返回可直接交给转换器的版本"""
    # \, \; \: \! \  等显式空格命令：latex2mathml 会把 \ 转成 m:nor 普通文本
    # run，Word 线性化为 " "，OneNote 粘贴后公式显示损坏；统一还原为普通空格
    latex = latex.replace('\\!', '').replace('\\ ', ' ')
    latex = re.sub(r'\\[,;:]', ' ', latex)
    # \kern1.5pt / {\kern1.5pt}：KaTeX 输出的间距（pandoc 生态常见），转成空格
    latex = re.sub(r'\{?\\kern\s*[-\d.]+(?:pt|em|cm|mm|ex|bp|mu)\}?', ' ', latex)
    # \operatorname{RNN} → \mathrm{RNN}：latex2mathml 不支持前者
    latex = latex.replace('\\operatorname', '\\mathrm')
    return latex


def latex_to_omml_xml(latex, display=False):
    """LaTeX → OMML XML 字符串（剪贴板条件注释用）；失败抛异常；
    display 块级公式外包 oMathPara（自带命名空间声明）"""
    mml = latex2mathml.converter.convert(_sanitize_latex(latex))
    omml = mathml2omml.convert(mml)
    omml = omml.replace('<m:box><m:e>', '').replace('</m:e></m:box>', '')
    m = re.match(r'^<m:oMath[^>]*>(.*)</m:oMath>$', omml, flags=re.S)
    body = m.group(1) if m else omml
    if display:
        return (f'<m:oMathPara xmlns:m="{OMML_NS}">'
                f'<m:oMath>{body}</m:oMath></m:oMathPara>')
    return f'<m:oMath xmlns:m="{OMML_NS}">{body}</m:oMath>'


def _math_html_img(seg):
    """公式 → base64 PNG 内嵌 img；渲染失败回退 $...$ 原文"""
    png = _render_math_png_bytes(seg['latex'])
    if png is None:
        return html.escape(seg['text'])
    b64 = base64.b64encode(png).decode('ascii')
    return (f'<img src="data:image/png;base64,{b64}" '
            f'alt="{html.escape(seg["latex"])}" style="vertical-align:middle;">')


_CODE_HTML_STYLE = ('style="background:#F2F2F2;padding:1px 3px;'
                    'border-radius:2px;font-family:Consolas,monospace;"')


def _html_frag(frag, store, math_fn):
    """片段内占位符处理：公式→math_fn 片段，行内代码→<code>；其余 HTML 转义"""
    parts = []
    pos = 0
    for m in PH_RE.finditer(frag):
        if m.start() > pos:
            parts.append(html.escape(frag[pos:m.start()]))
        seg = store[int(m.group(1))]
        if seg['kind'] == 'code':
            parts.append(f'<code {_CODE_HTML_STYLE}>{html.escape(seg["text"])}</code>')
        elif _HAS_LATEX:
            parts.append(math_fn(seg))
        else:
            parts.append(html.escape(seg['text']))
        pos = m.end()
    if pos < len(frag):
        parts.append(html.escape(frag[pos:]))
    return ''.join(parts)


def _restore_html(text, store, math_fn=None, keep_bold=True):
    """富文本模式：粗体/斜体/删除线转真标签，链接转 <a>，公式交给 math_fn；
    keep_bold=False 时 **加粗** 片段不输出 <b> 标签（全部不加粗）"""
    if math_fn is None:
        math_fn = _math_html_img
    t = _link_sub_html(text)
    # <a ...>文字</a> 原子化为占位符：避免 _parse_fmt 误入、_html_frag 转义破坏
    a_store = []

    def _stash_a(m):
        a_store.append(m.group(0))
        return f'{A_PH_OPEN}{len(a_store) - 1}{A_PH_CLOSE}'

    t = re.sub(r'<a href="[^"]*">.*?</a>', _stash_a, t)
    t = _RE_REFLINK.sub(r'\1', t)
    t = _RE_AUTOLINK.sub(r'\1', t)
    t = _RE_HTML.sub('', t)
    parts = []
    for frag, fmt in _parse_fmt(t):
        h = _html_frag(frag, store, math_fn)
        if 's' in fmt:
            h = f'<s>{h}</s>'
        if 'i' in fmt:
            h = f'<i>{h}</i>'
        if 'b' in fmt and keep_bold:
            h = f'<b>{h}</b>'
        parts.append(h)
    result = ''.join(parts)
    return A_PH_RE.sub(lambda m: a_store[int(m.group(1))], result)


_PNG_CACHE = {}


def _render_math_png_bytes(latex):
    """LaTeX → PNG 字节（matplotlib mathtext 渲染）；失败返回 None"""
    if latex in _PNG_CACHE:
        return _PNG_CACHE[latex]
    try:
        import matplotlib
        matplotlib.use('Agg')
        from matplotlib.figure import Figure
        fig = Figure()
        fig.text(0, 0, f'${latex}$', fontsize=13)
        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=110, transparent=True,
                    bbox_inches='tight', pad_inches=0.06)
        png = buf.getvalue()
    except Exception:
        png = None
    _PNG_CACHE[latex] = png
    return png


def _render_math_png(latex):
    """LaTeX → tk.PhotoImage（右侧预览用）；失败返回 None"""
    b = _render_math_png_bytes(latex)
    if b is None:
        return None
    try:
        return tk.PhotoImage(data=base64.b64encode(b))
    except Exception:
        return None


# ---------------- 行内元素转换 ----------------
_RE_IMG = re.compile(r'!\[[^\]]*\]\([^)]*\)')          # 图片：整体删除
_RE_LINK = re.compile(r'\[([^\]]*)\]\([^)]*\)')       # 链接：只留文字
_RE_REFLINK = re.compile(r'\[([^\]]*)\]\[[^\]]*\]')   # 引用式链接：只留文字
_RE_AUTOLINK = re.compile(r'<((?:https?|ftp)://[^>\s]+)>')
_RE_HTML = re.compile(r'</?[a-zA-Z][^<>]*>')           # HTML 标签：删除
_RE_STRIKE = re.compile(r'~~(.+?)~~')
_RE_BOLD3 = re.compile(r'\*\*\*(?!\s)(.+?)(?<!\s)\*\*\*')
_RE_BOLD = re.compile(r'\*\*(?!\s)(.+?)(?<!\s)\*\*')
_RE_BOLD2 = re.compile(r'(?<!\w)__(?!\s)(.+?)(?<!\s)__(?!\w)')
_RE_ITAL = re.compile(r'(?<![\w*\\])\*(?!\s)([^*\n]+?)(?<!\s)\*')
_RE_ITAL2 = re.compile(r'(?<![\w\\])_(?!\s)([^_\n]+?)(?<!\s)_(?!\w)')
_RE_ESCAPE = re.compile(r'\\([\\`*_{}\[\]()#+.!|>~-])')  # 转义符放最后处理


def _inline(t):
    t = _RE_IMG.sub('', t)
    t = _RE_LINK.sub(r'\1', t)
    t = _RE_REFLINK.sub(r'\1', t)
    t = _RE_AUTOLINK.sub(r'\1', t)
    t = _RE_HTML.sub('', t)
    t = _RE_STRIKE.sub(r'\1', t)
    t = _RE_BOLD3.sub(r'\1', t)
    t = _RE_BOLD.sub(r'\1', t)
    t = _RE_BOLD2.sub(r'\1', t)
    t = _RE_ITAL.sub(r'\1', t)
    t = _RE_ITAL2.sub(r'\1', t)
    t = _RE_ESCAPE.sub(r'\1', t)
    return t


# ---------------- 行内格式解析（docx/HTML 富文本用） ----------------
_RE_FMT = re.compile(
    r'\*\*\*(?!\s)(?P<b3>.+?)(?<!\s)\*\*\*'
    r'|\*\*(?!\s)(?P<b>.+?)(?<!\s)\*\*'
    r'|(?<![\w*\\])\*(?!\s)(?P<i>[^*\n]+?)(?<!\s)\*(?![\w*\\])'
    r'|~~(?!\s)(?P<s>.+?)(?<!\s)~~'
    r'|(?<!\w)__(?!\s)(?P<b2>.+?)(?<!\s)__(?!\w)'
    r'|(?<![\w\\])_(?!\s)(?P<i2>[^_\n]+?)(?<!\s)_(?!\w)'
)


def _parse_fmt(t, depth=0):
    """解析行内标记 → [(片段文本, fmt集合), …]，fmt ⊆ {'b','i','s'}；支持嵌套"""
    if depth > 4:
        return [(t, frozenset())]
    out = []
    pos = 0
    for m in _RE_FMT.finditer(t):
        pre = t[pos:m.start()]
        if pre:
            out.append((pre, frozenset()))
        g = m.groupdict()
        if g['b3'] is not None:
            inner, fmt = g['b3'], frozenset('bi')
        elif g['b'] is not None:
            inner, fmt = g['b'], frozenset('b')
        elif g['b2'] is not None:
            inner, fmt = g['b2'], frozenset('b')
        elif g['s'] is not None:
            inner, fmt = g['s'], frozenset('s')
        elif g['i'] is not None:
            inner, fmt = g['i'], frozenset('i')
        else:
            inner, fmt = g['i2'], frozenset('i')
        for seg, f in _parse_fmt(inner, depth + 1):
            out.append((seg, f | fmt))
        pos = m.end()
    if pos < len(t):
        out.append((t[pos:], frozenset()))
    return out


def _plain_pre(t):
    """富文本前的规整：删图片、链接只留文字、删裸 HTML 标签"""
    t = _RE_IMG.sub('', t)
    t = _RE_LINK.sub(r'\1', t)
    t = _RE_REFLINK.sub(r'\1', t)
    t = _RE_AUTOLINK.sub(r'\1', t)
    t = _RE_HTML.sub('', t)
    return t


_RE_LINK_H = re.compile(r'\[([^\]]*)\]\(([^)\s]+)[^)]*\)')


def _link_sub_html(t):
    """HTML 路径：[文字](url) → <a href>；图片删除"""
    t = _RE_IMG.sub('', t)
    return _RE_LINK_H.sub(
        lambda m: (f'<a href="{html.escape(m.group(2), quote=True)}">'
                   f'{_inline(m.group(1))}</a>'), t)


# ---------------- Markdown 解析 → 块 ----------------
_RE_LISTITEM = re.compile(r'^(\s*)(?:[-*+]|\d{1,3}[.)])\s+(.*)$')


def _to_letter(n):
    """1→a, 2→b, … 26→z, 27→aa"""
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(97 + r) + s
    return s


def _to_roman(n):
    """1→i, 2→ii …（小写罗马数字）"""
    out = ''
    for v, sym in ((1000, 'm'), (900, 'cm'), (500, 'd'), (400, 'cd'),
                   (100, 'c'), (90, 'xc'), (50, 'l'), (40, 'xl'),
                   (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')):
        while n >= v:
            out += sym
            n -= v
    return out


def _convert(md):
    """解析 Markdown → 块列表：
    ('line', str) | ('heading', 级别, str) | ('code', 原文) | ('table', cells)
    行内格式标记（**粗体** 等）保留在文本里，由各消费端自行解析"""
    text = md.replace('\r\n', '\n').replace('\r', '\n')
    text, store = _protect(text)

    lines = text.split('\n')
    blocks = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        s = line.strip()

        # 块级代码占位符（内容含换行）→ 独立 code 块
        m = re.fullmatch(f'{PH_OPEN}(\\d+){PH_CLOSE}', s)
        if m:
            seg = store[int(m.group(1))]
            if seg['kind'] == 'code' and '\n' in seg['text']:
                blocks.append(('code', seg['text']))
                i += 1
                continue

        if not s:  # 空行
            blocks.append(('line', ''))
            i += 1
            continue

        # 表格块（单元格保留占位符与行内标记，渲染时再处理）
        if s.startswith('|') and s.count('|') >= 2:
            rows = []
            while i < n and lines[i].strip().startswith('|'):
                raw = lines[i].strip()
                if raw.startswith('|'):
                    raw = raw[1:]
                if raw.endswith('|'):
                    raw = raw[:-1]
                rows.append([c.strip() for c in raw.split('|')])
                i += 1
            blocks.append(('table', rows))
            continue

        # 未闭合的代码块围栏：跳过
        if s.startswith('```'):
            i += 1
            continue

        # 链接/脚注定义行：删除
        if re.match(r'^\s*\[[^\]]+\]:\s*\S', line):
            i += 1
            continue

        # 分隔线 / Setext 标题下划线：丢弃
        if re.fullmatch(r'[-=*_]{3,}', s.replace(' ', '')):
            i += 1
            continue

        # 引用：去掉 > 标记
        line2 = re.sub(r'^\s*(?:>\s?)+', '', line)

        # 标题：去掉 #，记录层级（OneNote/Word 大纲用）
        m = re.match(r'^(#{1,6})\s+(.+?)(?:\s+#+)?\s*$', line2.strip())
        if m:
            blocks.append(('heading', len(m.group(1)), m.group(2)))
            i += 1
            continue

        # 列表块：连续列表项（中间允许空行）统一自动编号——
        # 第一层 1. 2. 3.，第二层 (1)(2)(3)，第三层 a. b. c.，第四层起 i. ii. iii.
        if _RE_LISTITEM.match(line2):
            items = []  # [(缩进, 内容), …]
            j = i
            while j < n:
                raw = re.sub(r'^\s*(?:>\s?)+', '', lines[j])  # 块内行同样去掉引用标记
                mj = _RE_LISTITEM.match(raw)
                if mj:
                    items.append((mj.group(1), mj.group(2)))
                    j += 1
                    continue
                if raw.strip() == '':  # 空行：后随仍是列表项则并入块内（松列表）
                    k = j + 1
                    while k < n and lines[k].strip() == '':
                        k += 1
                    if k < n and _RE_LISTITEM.match(
                            re.sub(r'^\s*(?:>\s?)+', '', lines[k])):
                        j += 1
                        continue
                break  # 非列表内容 → 列表块结束
            # 按缩进宽度分层：0=第一层，其余从小到大依次为更深层
            widths = sorted({len(it[0].expandtabs(4)) for it in items})
            counters = {}  # 层级 → 当前计数
            for indent, content in items:
                level = widths.index(len(indent.expandtabs(4))) + 1
                counters = {lv: c for lv, c in counters.items() if lv <= level}
                counters[level] = counters.get(level, 0) + 1
                num = counters[level]
                if level == 1:
                    label = f'{num}.'
                elif level == 2:
                    label = f'({num})'
                elif level == 3:
                    label = f'{_to_letter(num)}.'
                else:
                    label = f'{_to_roman(num)}.'
                blocks.append(('line', f'{indent}{label} {content}'))
            i = j
            continue

        blocks.append(('line', line2))
        i += 1
    return blocks, store


# ---------------- 表格渲染 ----------------
def _split_sep_row(rows):
    """识别对齐分隔行，返回 (对齐列表或 None, 数据行)"""
    sep_idx = None
    for idx, r in enumerate(rows):
        if r and all(re.fullmatch(r':?-+:?', c) for c in r):
            sep_idx = idx
            break
    if sep_idx is None:
        return None, rows
    aligns = []
    for c in rows[sep_idx]:
        a = 'left'
        if c.startswith(':') and c.endswith(':'):
            a = 'center'
        elif c.endswith(':'):
            a = 'right'
        aligns.append(a)
    data = [r for idx, r in enumerate(rows) if idx != sep_idx]
    return aligns, data


def _render_table_text(rows, store):
    """渲染成对齐的纯文本表格（右侧预览 / 纯文本复制用）"""
    _aligns, rows = _split_sep_row(rows)
    if not rows:
        return []
    ncols = max(len(r) for r in rows)
    rows = [r + [''] * (ncols - len(r)) for r in rows]
    rows = [[_plain_line(c, store) for c in r] for r in rows]
    widths = [max(_display_width(r[j]) for r in rows) for j in range(ncols)]

    def pad(cell, w):
        return cell + ' ' * (w - _display_width(cell))

    out = []
    for idx, r in enumerate(rows):
        out.append('| ' + ' | '.join(pad(r[j], widths[j])
                                     for j in range(ncols)) + ' |')
        if idx == 0:
            out.append('|' + '|'.join('-' * (w + 2) for w in widths) + '|')
    return out


def _render_table_html(rows, store, keep_bold=True, math_fn=None):
    """渲染成 HTML 真表格（富文本复制用，单元格内公式交给 math_fn）"""
    aligns, rows = _split_sep_row(rows)
    if not rows:
        return ''
    ncols = max(len(r) for r in rows)
    rows = [r + [''] * (ncols - len(r)) for r in rows]

    h = ['<table border="1" cellspacing="0" cellpadding="4" '
         'style="border-collapse:collapse;">']
    for idx, r in enumerate(rows):
        h.append('<tr>')
        for j, cell in enumerate(r):
            tag = 'th' if idx == 0 else 'td'
            style = ''
            if aligns and j < len(aligns) and aligns[j] != 'left':
                style = f' style="text-align:{aligns[j]};"'
            h.append(f'<{tag}{style}>{_restore_html(cell, store, math_fn=math_fn, keep_bold=keep_bold)}</{tag}>')
        h.append('</tr>')
    h.append('</table>')
    return ''.join(h)


# ---------------- 最终输出 ----------------
def md_to_text(md):
    """Markdown → 纯文本；公式保持 $...$ / $$...$$ 原文"""
    blocks, store = _convert(md)
    out = []
    for b in blocks:
        if b[0] == 'line':
            out.append(_plain_line(b[1], store))
        elif b[0] == 'heading':
            out.append(_plain_line(b[2], store))
        elif b[0] == 'code':
            out.extend(b[1].split('\n'))
        else:
            out.extend(_render_table_text(b[1], store))
    # 去行尾空白、压缩连续空行、去掉首尾空行
    out = [l.rstrip() for l in out]
    squeezed = []
    for l in out:
        if l == '' and squeezed and squeezed[-1] == '':
            continue
        squeezed.append(l)
    result = '\n'.join(squeezed).strip('\n')
    return result + '\n' if result else ''


def md_to_html(md, hsize=None, hfont=None, keep_bold=False, sp_before=0,
               sp_after=0, math_fn=None):
    """Markdown → HTML 片段：公式交给 math_fn（默认渲染为内嵌 PNG 图片）、
    表格转真表格（富文本复制用）
    标题转 h 标签、粗体/斜体/删除线转真标签、代码块转 pre；
    hsize/hfont：各级标题字号（磅）与字体，None 表示不设置
    （粘贴后由目标软件默认分级字号/正文字体渲染）
    keep_bold：Markdown **加粗** 片段是否保留 <b> 标签（False 则全部不加粗）
    sp_before/sp_after：段前/段后间距（磅，默认 0）"""
    blocks, store = _convert(md)
    parts = []
    margin = ('margin:0;' if sp_before == 0 and sp_after == 0
              else f'margin:{sp_before}pt 0 {sp_after}pt 0;')
    prev_blank = False
    for b in blocks:
        kind = b[0]
        payload = b[1]
        if kind == 'table':
            prev_blank = False
            t = _render_table_html(payload, store, keep_bold=keep_bold,
                                   math_fn=math_fn)
            if t:
                parts.append(t)
            continue
        if kind == 'code':
            prev_blank = False
            esc = html.escape(payload)
            parts.append(f'<pre style="{margin}background:#F2F2F2;'
                         f'padding:6px 8px;font-family:Consolas,monospace;'
                         f'font-size:12px;">{esc}</pre>')
            continue
        if kind == 'heading':
            lv = min(payload, 6)
            body = _restore_html(b[2], store, math_fn=math_fn,
                                 keep_bold=keep_bold)
            h_style = margin
            if hsize:
                h_style += f'font-size:{hsize}pt;'
            if hfont:
                h_style += f"font-family:'{hfont}';"
            parts.append(f'<h{lv} style="{h_style}">{body}</h{lv}>')
            prev_blank = False
            continue
        if _plain_line(payload, store).strip() == '':  # 空行
            if prev_blank:
                continue
            prev_blank = True
            parts.append('<p><br/></p>')
            continue
        prev_blank = False
        parts.append(f'<p style="{margin}">{_restore_html(payload, store, math_fn=math_fn, keep_bold=keep_bold)}</p>')
    result = ''.join(parts)
    # 去掉首尾的空段落
    blank = '<p><br/></p>'
    while result.startswith(blank):
        result = result[len(blank):]
    while result.endswith(blank):
        result = result[:-len(blank)]
    return result


# ---------------- Pandoc 集成 + Office HTML（OMML 条件注释） ----------------
# 原理（Word 官方剪贴板同款技术）：HTML 里写入
#   <!--[if gte msEquation 12]><m:oMath>…</m:oMath><![endif]-->
#   <![if !msEquation]>回退内容<![endif]-->
# OneNote/Word 会取原生 OMML 渲染为可编辑公式，其他软件取回退内容。

MS_OMML_HTML_NS = 'http://schemas.microsoft.com/office/2004/12/omml'

_RE_MATHML = re.compile(r'<math[^>]*>.*?</math>', re.S | re.I)

_PANDOC_PATH = None


def _find_pandoc():
    """定位 pandoc.exe：打包内置 → 程序目录 → PATH；找不到返回 None（结果缓存）"""
    global _PANDOC_PATH
    if _PANDOC_PATH is not None:
        return _PANDOC_PATH or None
    candidates = []
    if getattr(sys, 'frozen', False) and getattr(sys, '_MEIPASS', None):
        candidates.append(os.path.join(sys._MEIPASS, 'pandoc', 'pandoc.exe'))
    candidates.append(os.path.join(_app_dir(), 'pandoc', 'pandoc.exe'))
    which = shutil.which('pandoc')
    if which:
        candidates.append(which)
    for c in candidates:
        if c and os.path.isfile(c):
            _PANDOC_PATH = c
            return c
    _PANDOC_PATH = ''
    return None


def _pandoc_md_to_mathml_html(md):
    """Pandoc：Markdown → HTML 片段（公式为 MathML；语法覆盖远超内置转换器）"""
    pandoc = _find_pandoc()
    if not pandoc:
        raise RuntimeError('未找到 pandoc.exe')
    cmd = [os.path.abspath(pandoc),
           # hard_line_breaks：AI 笔记普遍用单换行分行（①②③条目等），
           # 标准 Markdown 会合并成一段；与内置转换器逐行处理的行为对齐
           '-f', 'markdown+hard_line_breaks'
                 '+tex_math_dollars+tex_math_single_backslash'
                 '+tex_math_double_backslash',
           '-t', 'html', '--math-method=mathml', '--wrap=none']
    kwargs = {}
    if os.name == 'nt':  # 窗口程序里不闪控制台黑框
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        kwargs = {'startupinfo': si,
                  'creationflags': subprocess.CREATE_NO_WINDOW}
    r = subprocess.run(cmd, input=md.encode('utf-8'), capture_output=True,
                       shell=False, **kwargs)
    if r.returncode != 0:
        raise RuntimeError('Pandoc 转换失败: ' +
                           (r.stderr or b'').decode('utf-8', 'ignore')[:300])
    return r.stdout.decode('utf-8', 'ignore')


def _mathml_to_omml(mathml, display=False):
    """MathML → OMML 字符串（m: 命名空间由剪贴板 HTML 的 <html> 声明）；
    display 块级公式外包 oMathPara；失败抛异常"""
    # 剥掉 pandoc 内嵌的原始 LaTeX 注释节点（mathml2omml 不认识且会污染回退文本）
    mathml = re.sub(r'<annotation[^>]*>.*?</annotation>', '', mathml, flags=re.S)
    omml = mathml2omml.convert(mathml, name2codepoint)
    omml = omml.replace('<m:box><m:e>', '').replace('</m:e></m:box>', '')
    m = re.match(r'^<m:oMath[^>]*>(.*)</m:oMath>$', omml, flags=re.S)
    body = m.group(1) if m else omml
    if display and m:
        # oMathPara 内必须嵌套完整 oMath（与 Word 自身剪贴板输出一致；
        # 无容器的裸 m:r/m:sSub 碎片会被 OneNote 当普通文本渲染）
        return f'<m:oMathPara><m:oMath>{body}</m:oMath></m:oMathPara>'
    if m:
        return omml  # mathml2omml 输出的完整 <m:oMath>…</m:oMath>
    return f'<m:oMath>{omml}</m:oMath>'


def _replace_mathml_with_omml(html_text):
    """把 HTML 里的 MathML 替换为 OMML 条件注释（回退内容为公式纯文本）"""
    def _sub(m):
        mathml = m.group(0)
        # 剥掉 pandoc 内嵌的原始 LaTeX 注释节点（会混入回退文本、mathml2omml 不认识）
        mathml = re.sub(r'<annotation[^>]*>.*?</annotation>', '', mathml,
                        flags=re.S)
        display = 'display="block"' in mathml
        fallback = html.escape(re.sub(r'<[^>]+>', '', mathml))
        try:
            omml = _mathml_to_omml(mathml, display)
        except Exception:
            return fallback
        return (f'<!--[if gte msEquation 12]>{omml}<![endif]-->'
                f'<![if !msEquation]>{fallback}<![endif]>')
    return _RE_MATHML.sub(_sub, html_text)


def _math_omml_conditional(seg):
    """无 Pandoc 的内置路径：公式 → OMML 条件注释（失败回退 $...$ 原文）"""
    try:
        omml = latex_to_omml_xml(seg['latex'], seg['display'])
        omml = omml.replace(f' xmlns:m="{OMML_NS}"', '')
    except Exception:
        return html.escape(seg['text'])
    return (f'<!--[if gte msEquation 12]>{omml}<![endif]-->'
            f'<![if !msEquation]>{html.escape(seg["text"])}<![endif]>')


def _strip_html_bold(body):
    """keep_bold=False：删除 <strong>/<b> 标签（保留内部内容）"""
    return re.sub(r'</?(?:strong|b)\b[^>]*>', '', body)


def _flatten_lists(md):
    """把 Markdown 列表压平为本应用的层级编号方案（1. → (1) → a. → i.）：
    输出转义后的普通段落（前后留空行、缩进用 &nbsp;），数字已转义，
    pandoc 不再解析为原生列表 → OneNote 中不出现圆点符号；
    编号方案与纯文本/预览路径保持一致"""
    out = []
    lines = md.split('\n')
    i = 0
    n = len(lines)
    in_code = False
    while i < n:
        line = lines[i]
        if line.lstrip().startswith('```'):  # 代码围栏内不处理
            in_code = not in_code
            out.append(line)
            i += 1
            continue
        if in_code:
            out.append(line)
            i += 1
            continue
        raw = re.sub(r'^\s*(?:>\s?)+', '', line)
        if not _RE_LISTITEM.match(raw):
            out.append(line)
            i += 1
            continue
        # 收集连续列表项（与 _convert 相同的松列表规则）
        items = []
        j = i
        while j < n:
            if lines[j].lstrip().startswith('```'):
                break
            raw_j = re.sub(r'^\s*(?:>\s?)+', '', lines[j])
            mj = _RE_LISTITEM.match(raw_j)
            if mj:
                items.append((mj.group(1), mj.group(2)))
                j += 1
                continue
            if raw_j.strip() == '':  # 空行：后随仍是列表项则并入块内
                k = j + 1
                while k < n and lines[k].strip() == '':
                    k += 1
                if k < n and _RE_LISTITEM.match(
                        re.sub(r'^\s*(?:>\s?)+', '', lines[k])):
                    j += 1
                    continue
            break
        # 按缩进宽度分层（与 _convert 一致）
        widths = sorted({len(it[0].expandtabs(4)) for it in items})
        counters = {}
        for indent, content in items:
            level = widths.index(len(indent.expandtabs(4))) + 1
            counters = {lv: c for lv, c in counters.items() if lv <= level}
            counters[level] = counters.get(level, 0) + 1
            num = counters[level]
            # 数字转义，防止 pandoc 重新解析为列表
            if level == 1:
                label = f'{num}\\.'
            elif level == 2:
                label = f'({num}\\)'
            elif level == 3:
                label = f'{_to_letter(num)}\\.'
            else:
                label = f'{_to_roman(num)}\\.'
            nbsp = '&nbsp;' * len(indent.expandtabs(4))
            if out and out[-1].strip():
                out.append('')  # 与前文空行分隔，避免被并进上一段落
            out.append(f'{nbsp}{label} {content}')
            out.append('')
        i = j
    return '\n'.join(out)


def _style_pandoc_html(body, *, keep, hsize, hfont, sp_before, sp_after):
    """给 pandoc HTML 注入内联样式（OneNote 对 <style> 块支持差）；
    复刻模式不动样式，使用目标软件默认渲染"""
    if keep:
        return body
    margin = ('' if sp_before == 0 and sp_after == 0
              else f'margin:{sp_before}pt 0 {sp_after}pt 0;')
    if margin:
        body = re.sub(r'<p>', f'<p style="{margin}">', body)
    # 标题统一字号/字体，且不加粗（对齐旧版「统一模式」行为：
    # OneNote 会按自身 Heading 样式加粗，须显式覆盖）
    h_style = 'font-weight:normal;'
    if hsize:
        h_style += f'font-size:{hsize}pt;'
    if hfont:
        h_style += f"font-family:'{hfont}';"
    body = re.sub(r'<h([1-6])(\s|>)',
                  lambda m: f'<h{m.group(1)} style="{h_style}"{m.group(2)}',
                  body)
    # pandoc 表格默认无边框，补边框样式（td/th 可能已带对齐 style，需合并）
    def _cell(m):
        tag = m.group(0)
        if 'style="' in tag:
            return tag.replace('style="',
                               'style="border:1px solid #999;padding:4px 8px;', 1)
        return tag[:-1] + ' style="border:1px solid #999;padding:4px 8px;">'
    body = re.sub(r'<table>', '<table style="border-collapse:collapse;">', body)
    body = re.sub(r'<t[dh][^>]*>', _cell, body)
    # 表头单元格同样不加粗（OneNote 默认 th 加粗）
    body = re.sub(r'<th style="border',
                  '<th style="font-weight:normal;border', body)
    return body


def _office_math_html(md, *, keep, hsize, hfont, keep_bold,
                      sp_before, sp_after):
    """「公式可编辑」：Markdown → 含 OMML 条件注释的 HTML 正文；
    优先 Pandoc（语法覆盖广、公式支持矩阵等复杂结构），无 Pandoc 时用内置转换器"""
    pandoc = _find_pandoc()
    if pandoc:
        # 压平列表（消除圆点、应用自动编号）并规范化公式分隔符；
        # 复刻模式不压平，保留 Markdown 原生列表形态
        md2 = md if keep else _flatten_lists(_fix_math_delimiters(md))
        body = _replace_mathml_with_omml(_pandoc_md_to_mathml_html(md2))
        if not (keep_bold or keep):  # 复刻模式强制保留加粗
            body = _strip_html_bold(body)
        body = _style_pandoc_html(body, keep=keep, hsize=hsize, hfont=hfont,
                                  sp_before=sp_before, sp_after=sp_after)
    else:
        body = md_to_html(md, hsize=None if keep else hsize,
                          hfont=None if keep else hfont,
                          keep_bold=keep_bold or keep,
                          sp_before=sp_before, sp_after=sp_after,
                          math_fn=_math_omml_conditional)
    return body


# ---------------- Windows 剪贴板（CF_HTML 富文本） ----------------
def _set_clipboard_html(html_fragment, plain_text, body_attr=''):
    """把 Office HTML（含 OMML 公式条件注释）和纯文本同时写入剪贴板；
    body_attr 写在 <body> 标签上（如字体样式；复刻模式传空）；
    <html> 上声明 Office OMML 命名空间，OneNote/Word 才能解析公式"""
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    user32.RegisterClipboardFormatW.restype = wintypes.UINT
    user32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    user32.SetClipboardData.restype = ctypes.c_void_p
    user32.SetClipboardData.argtypes = [wintypes.UINT, ctypes.c_void_p]

    cf_html = user32.RegisterClipboardFormatW('HTML Format')
    header = ('Version:0.9\r\nStartHTML:{:010d}\r\nEndHTML:{:010d}\r\n'
              'StartFragment:{:010d}\r\nEndFragment:{:010d}\r\n')
    pre = (f'<html xmlns:o="urn:schemas-microsoft-com:office:office"'
           f' xmlns:m="{MS_OMML_HTML_NS}"'
           f' xmlns="http://www.w3.org/TR/REC-html40">'
           f'<head><meta http-equiv="Content-Type" content="text/html;'
           f' charset=utf-8"></head>'
           f'<body{body_attr}>\r\n<!--StartFragment-->')
    post = '<!--EndFragment-->\r\n</body></html>'

    frag = html_fragment.encode('utf-8')
    pre_b = pre.encode('utf-8')
    post_b = post.encode('utf-8')
    header_len = len(header.format(0, 0, 0, 0).encode('utf-8'))
    start_frag = header_len + len(pre_b)
    end_frag = start_frag + len(frag)
    end_html = end_frag + len(post_b)
    data = (header.format(header_len, end_html, start_frag, end_frag)
            .encode('utf-8') + pre_b + frag + post_b)
    text_data = plain_text.encode('utf-16-le') + b'\x00\x00'

    opened = False
    for _ in range(5):  # 剪贴板可能被占用，重试
        if user32.OpenClipboard(0):
            opened = True
            break
        time.sleep(0.1)
    if not opened:
        raise RuntimeError('无法打开剪贴板')

    def put(handle_fmt, payload):
        h = kernel32.GlobalAlloc(0x0002, len(payload))
        p = kernel32.GlobalLock(h)
        if not p:
            raise RuntimeError('GlobalLock 失败')
        ctypes.memmove(p, payload, len(payload))
        kernel32.GlobalUnlock(h)
        user32.SetClipboardData(handle_fmt, h)

    try:
        user32.EmptyClipboard()
        put(cf_html, data)
        put(13, text_data)  # CF_UNICODETEXT
    finally:
        user32.CloseClipboard()


# ---------------- 图形界面 ----------------
DEMO = r'''# 示例文档

这是一段**粗体**、*斜体*、~~删除线~~和 `行内代码`，
以及行内公式 $E = mc^2$。

## 列表示例

- 第一项
- 第二项
  - 嵌套项

1. 步骤一
2. 步骤二

## 表格示例

| 名称 | 数值 | 备注 |
|------|-----:|------|
| 甲   | 1    | 无   |
| 乙   | 20   | 测试 |

## 代码块示例

```python
def hello(name):
    print(f'Hello, {name}!')
```

## 公式示例

块级公式：

$$\int_0^1 x^2\,dx = \frac{1}{3}$$

> 引用一行文字，去掉大于号。
[链接文字](https://example.com) 与 ![图片](x.png)。

使用方法：清空左侧，粘贴你自己的 Markdown，右侧实时显示结果。
'''

IN_FONT = ('Microsoft YaHei UI', 11)
OUT_FONT = ('Consolas', 11)
RICH_LABEL = '复制（公式可编辑）'
IMG_LABEL = '复制（公式为图片）'
PLAIN_LABEL = '复制源MD'

def _app_dir():
    """程序所在目录：打包成 exe 后 __file__ 指向临时解压目录，须用 exe 路径"""
    if getattr(sys, 'frozen', False):  # PyInstaller 打包
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(_app_dir(), 'config.json')


def _load_config():
    """读取配置（窗口内容、几何、开关状态）；不存在或损坏时返回 {}"""
    try:
        with open(CONFIG_PATH, encoding='utf-8') as f:
            cfg = json.load(f)
        return cfg if isinstance(cfg, dict) else {}
    except Exception:
        return {}


def _save_config(cfg):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False)
    except Exception:
        pass  # 配置只写失败不影响使用


class App:
    def __init__(self, root):
        self.root = root
        self._job = None
        self._img_cache = {}   # latex → PhotoImage/None
        self._imgs = []        # 防止图片被回收
        self.current_path = None  # 打开/保存的文件路径（标题栏显示）

        cfg = _load_config()
        root.title('MdClear')
        if isinstance(cfg.get('geometry'), str):
            root.geometry(cfg['geometry'])
        else:
            root.geometry('1150x680')
        # 上次保存的位置可能已跑出屏幕外（更换显示器/修改缩放后），此时居中显示
        m = re.match(r'(\d+)x(\d+)([+-]\d+)([+-]\d+)$', str(cfg.get('geometry') or ''))
        if m:
            w, h, x, y = (int(g) for g in m.groups())
            sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
            if x >= sw - 40 or y >= sh - 40 or x <= -(w - 40) or y <= -(h - 40):
                root.geometry(f'{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 2)}')
        # 转换后（粘贴产物）的默认字体设置：等线 / 12 磅（小四）/ 不加粗
        self.var_font = tk.StringVar(value=str(cfg.get('font') or '等线'))
        self.var_font_size = tk.StringVar(value=str(cfg.get('font_size') or 12))
        # 各级标题字号/字体（统一模式专用，默认 14 磅/等线；复刻模式不生效）
        self.var_head_size = tk.StringVar(value=str(cfg.get('head_size') or 14))
        self.var_head_font = tk.StringVar(value=str(cfg.get('head_font') or '等线'))
        # 正文加粗字段保留：勾选后 Markdown **加粗** 片段保留加粗；不勾选全部不加粗（默认不勾）
        self.var_keep_bold = tk.BooleanVar(value=bool(cfg.get('keep_bold')))
        # 段前/段后间距（磅，默认 0 紧凑）
        self.var_sp_before = tk.StringVar(value=str(cfg.get('sp_before') or 0))
        self.var_sp_after = tk.StringVar(value=str(cfg.get('sp_after') or 0))
        # 复刻 Markdown 排版：不设置字体字号加粗，标题/表头保持加粗（默认关闭）
        self.var_font_keep = tk.BooleanVar(value=bool(cfg.get('keep_markdown')))

        # 文件菜单：打开 / 保存输入 / 保存结果
        menubar = tk.Menu(root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label='打开文件… (Ctrl+O)', command=self.open_file)
        file_menu.add_separator()
        file_menu.add_command(label='保存输入为 Markdown… (Ctrl+S)',
                              command=self.save_input)
        file_menu.add_command(label='保存结果为纯文本…',
                              command=self.save_result)
        menubar.add_cascade(label='文件', menu=file_menu)
        # 设置菜单：粘贴产物的默认字体 + 关于
        set_menu = tk.Menu(menubar, tearoff=0)
        set_menu.add_command(label='粘贴设置…', command=self.open_settings)
        set_menu.add_separator()
        set_menu.add_command(label='关于 MdClear…', command=self.show_about)
        menubar.add_cascade(label='设置', menu=set_menu)
        root.config(menu=menubar)
        root.bind('<Control-o>', lambda e: self.open_file())
        root.bind('<Control-s>', lambda e: self.save_input())

        outer = ttk.Frame(root, padding=10)
        outer.pack(fill='both', expand=True)

        paned = ttk.PanedWindow(outer, orient='horizontal')
        paned.pack(fill='both', expand=True)

        # 左侧输入
        left = ttk.Frame(paned)
        paned.add(left, weight=1)
        ttk.Label(left, text='Markdown 输入').pack(anchor='w')
        self.in_tb = tk.Text(left, wrap='word', font=IN_FONT, undo=True)
        sb1 = ttk.Scrollbar(left, orient='vertical', command=self.in_tb.yview)
        self.in_tb.configure(yscrollcommand=sb1.set)
        sb1.pack(side='right', fill='y')
        self.in_tb.pack(fill='both', expand=True)

        # 右侧输出（只读，公式渲染显示）
        right = ttk.Frame(paned)
        paned.add(right, weight=1)
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)
        ttk.Label(right, text='转换结果（只读，公式渲染显示）').grid(
            row=0, column=0, columnspan=2, sticky='w')
        self.out_tb = tk.Text(right, wrap='none', state='disabled',
                              font=OUT_FONT, background='#f7f7f7')
        sy = ttk.Scrollbar(right, orient='vertical', command=self.out_tb.yview)
        sx = ttk.Scrollbar(right, orient='horizontal', command=self.out_tb.xview)
        self.out_tb.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.out_tb.grid(row=1, column=0, sticky='nsew')
        sy.grid(row=1, column=1, sticky='ns')
        sx.grid(row=2, column=0, sticky='ew')

        # 预览格式 tags：代码灰底 → 行内格式组合 → 标题层级（后配置的优先级高）
        self.out_tb.tag_configure('code', background='#F2F2F2')
        for key, attrs in (('b', 'bold'), ('i', 'italic'), ('bi', 'bold italic'),
                           ('s', 'overstrike'), ('bs', 'bold overstrike'),
                           ('is', 'italic overstrike'),
                           ('bis', 'bold italic overstrike')):
            self.out_tb.tag_configure(
                f'f_{key}', font=(OUT_FONT[0], OUT_FONT[1], *attrs.split()))
        for lv, size in zip(range(1, 7), (16, 13, 12, 11, 11, 11)):
            self.out_tb.tag_configure(f'h{lv}', font=(OUT_FONT[0], size, 'bold'))

        # 底部按钮栏
        bar = ttk.Frame(outer)
        bar.pack(fill='x', pady=(8, 0))
        self.btn_rich = ttk.Button(bar, text=RICH_LABEL, command=self.copy_rich)
        self.btn_rich.pack(side='left')
        self.btn_img = ttk.Button(bar, text=IMG_LABEL, command=self.copy_rich_img)
        self.btn_img.pack(side='left', padx=(8, 0))
        self.btn_copy = ttk.Button(bar, text=PLAIN_LABEL, command=self.copy_result)
        self.btn_copy.pack(side='left', padx=(8, 0))
        hint = ('支持拖入MD文件；粘贴时请使用保留源格式粘贴；'
                '内容与设置自动记忆')
        if not _HAS_DND:  # 缺少拖拽组件时不误导用户
            hint = '粘贴时请使用保留源格式粘贴；内容与设置自动记忆'
        ttk.Label(bar, text=hint).pack(side='right')

        # 实时转换：按键释放、粘贴、撤销等事件均触发（150ms 防抖）
        self.in_tb.bind('<KeyRelease>', self._schedule)
        for ev in ('<<Paste>>', '<<PasteSelected>>', '<<Cut>>',
                   '<<Undo>>', '<<Redo>>', '<<Clear>>'):
            self.in_tb.bind(ev, self._schedule)

        # 拖拽 md/txt 文件到输入框导入（tkinterdnd2 缺失时自动跳过）
        if _HAS_DND:
            self.in_tb.drop_target_register(DND_FILES)
            self.in_tb.dnd_bind('<<Drop>>', self._on_drop)

        # 内容记忆：上次退出时的输入直接恢复，否则用示例
        self.in_tb.insert('1.0', cfg.get('text') if cfg.get('text') else DEMO)
        self.convert_now()
        root.protocol('WM_DELETE_WINDOW', self._on_close)

    # ---------------- 文件操作 ----------------
    def _read_file(self, path):
        try:
            with open(path, encoding='utf-8-sig') as f:  # 兼容带 BOM 的记事本文件
                return f.read()
        except UnicodeDecodeError:
            with open(path, encoding='gbk', errors='replace') as f:
                return f.read()

    def _load_into_editor(self, content, path=None):
        self.in_tb.delete('1.0', 'end')
        self.in_tb.insert('1.0', content)
        self.current_path = path
        title = 'MdClear'
        if path:
            title += f' - {os.path.basename(path)}'
        self.root.title(title)
        self.convert_now()

    def open_file(self):
        path = filedialog.askopenfilename(
            title='打开 Markdown / 文本文件',
            filetypes=[('Markdown/文本', '*.md *.markdown *.txt'),
                       ('所有文件', '*.*')])
        if path:
            self._load_into_editor(self._read_file(path), path)

    def save_input(self):
        path = filedialog.asksaveasfilename(
            title='保存输入为 Markdown', defaultextension='.md',
            initialfile=os.path.basename(self.current_path) if self.current_path else '',
            filetypes=[('Markdown', '*.md'), ('文本', '*.txt')])
        if not path:
            return
        with open(path, 'w', encoding='utf-8') as f:
            f.write(self.in_tb.get('1.0', 'end-1c'))
        self.current_path = path
        self.root.title(f'MdClear - {os.path.basename(path)}')

    def save_result(self):
        path = filedialog.asksaveasfilename(
            title='保存结果为纯文本', defaultextension='.txt',
            filetypes=[('文本', '*.txt')])
        if not path:
            return
        with open(path, 'w', encoding='utf-8') as f:
            f.write(md_to_text(self.in_tb.get('1.0', 'end-1c')))

    def _on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        if not files:
            return
        path = files[0]  # 多个文件只取第一个
        if os.path.isfile(path):
            self._load_into_editor(self._read_file(path), path)

    def _on_close(self):
        """退出时：保存配置（内容/窗口/开关）"""
        _save_config({
            'text': self.in_tb.get('1.0', 'end-1c'),
            'geometry': self.root.geometry(),
            'font': self.var_font.get(),
            'font_size': self._font_size_val(),
            'head_size': self._head_size_val(),
            'head_font': self.var_head_font.get(),
            'keep_bold': self.var_keep_bold.get(),
            'sp_before': self._sp_val(self.var_sp_before),
            'sp_after': self._sp_val(self.var_sp_after),
            'keep_markdown': self.var_font_keep.get(),
        })
        self.root.destroy()

    def _schedule(self, _event=None):
        if self._job is not None:
            self.root.after_cancel(self._job)
        self._job = self.root.after(150, self.convert_now)

    def convert_now(self):
        self._job = None
        src = self.in_tb.get('1.0', 'end-1c')
        yview = self.out_tb.yview()
        self.out_tb.configure(state='normal')
        self.out_tb.delete('1.0', 'end')
        self._imgs = []
        try:
            blocks, store = _convert(src)
            prev_blank = False
            for b in blocks:
                kind = b[0]
                if kind == 'table':
                    prev_blank = False
                    for tl in _render_table_text(b[1], store):
                        self.out_tb.insert('end', tl + '\n')
                    continue
                if kind == 'code':
                    prev_blank = False
                    for cl in b[1].split('\n'):
                        self.out_tb.insert('end', cl + '\n', 'code')
                    continue
                if kind == 'heading':
                    prev_blank = False
                    self._insert_line(b[2], store, heading_lv=min(b[1], 6))
                    continue
                if _plain_line(b[1], store).strip() == '':  # 空行
                    if prev_blank:
                        continue
                    prev_blank = True
                    self.out_tb.insert('end', '\n')
                    continue
                prev_blank = False
                self._insert_line(b[1], store)
        except Exception as e:  # 转换出错时提示而不崩溃
            self.out_tb.insert('end', f'（转换出错：{e}）')
        self.out_tb.configure(state='disabled')
        self.out_tb.yview_moveto(yview[0])

    def _insert_line(self, line, store, heading_lv=0):
        """按行内格式与占位符分段插入：文本带格式 tag，公式渲染成图片插入"""
        t = _plain_pre(line)
        for frag, fmt in _parse_fmt(t):
            tags = [f'f_{"".join(sorted(fmt))}'] if fmt else []
            if heading_lv:
                tags.append(f'h{heading_lv}')
            last = 0
            for m in PH_RE.finditer(frag):
                if m.start() > last:
                    self.out_tb.insert('end', frag[last:m.start()], tags)
                seg = store[int(m.group(1))]
                if seg['kind'] == 'code':
                    self.out_tb.insert('end', seg['text'], tags + ['code'])
                else:
                    img = self._math_image(seg['latex'])
                    if img is not None:
                        self.out_tb.image_create('end-1c', image=img)
                        self._imgs.append(img)
                    else:
                        self.out_tb.insert('end', seg['text'], tags)  # 回退为 $...$ 原文
                last = m.end()
            if last < len(frag):
                self.out_tb.insert('end', frag[last:], tags)
        self.out_tb.insert('end', '\n')

    def _math_image(self, latex):
        """渲染公式图片并缓存（失败也缓存，避免重复渲染开销）"""
        if latex in self._img_cache:
            return self._img_cache[latex]
        img = _render_math_png(latex)
        self._img_cache[latex] = img
        return img

    def _flash(self, btn, text, restore_label):
        btn.config(text=text)
        self.root.after(1500, lambda: btn.config(text=restore_label))

    def copy_result(self):
        """复制输入框中的 Markdown 源文本（原样，不做任何转换）"""
        src = self.in_tb.get('1.0', 'end-1c')
        self.root.clipboard_clear()
        self.root.clipboard_append(src)
        self._flash(self.btn_copy, '已复制源MD ✓', PLAIN_LABEL)

    def _font_size_val(self):
        """字号设置的数值（非法输入回退 12 磅）"""
        try:
            return float(self.var_font_size.get())
        except ValueError:
            return 12

    def _head_size_val(self):
        """标题字号设置的数值（非法输入回退 14 磅）"""
        try:
            return float(self.var_head_size.get())
        except ValueError:
            return 14

    def _sp_val(self, var):
        """段前/段后间距数值（非法输入回退 0 磅）"""
        try:
            return float(var.get())
        except ValueError:
            return 0

    def show_about(self):
        """关于对话框（样式与粘贴设置一致）：作者 / 联系方式 / 开源地址 / 版权"""
        win = tk.Toplevel(self.root)
        win.title('关于 MdClear')
        win.resizable(False, False)
        win.transient(self.root)
        body = ttk.Frame(win, padding=12)
        body.pack(fill='both', expand=True)
        info = (
            'MdClear — Markdown 转纯文本小工具\n'
            '把 Markdown 转换成适合粘贴进 OneNote / Word / 聊天软件的格式\n\n'
            '作者：wangce\n'
            '联系方式：2253246@tongji.edu.cn\n'
            '开源地址：https://github.com/EthanWangHaven/md2txt_gui_tool\n\n'
            'Copyright © 2026 wangce. 保留所有权利。')
        ttk.Label(body, text=info, justify='left').pack(anchor='w')
        ttk.Button(body, text='确定', command=win.destroy).pack(
            anchor='e', pady=(10, 0))
        # 居中显示（与粘贴设置一致）
        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f'+{(sw - w) // 2}+{(sh - h) // 2}')
        win.grab_set()

    def open_settings(self):
        """粘贴设置：中文字体/字号/标题字体/标题字号/正文加粗字段保留/段前段后/复刻 Markdown 排版"""
        backup = (self.var_font.get(), self.var_font_size.get(),
                  self.var_head_font.get(), self.var_head_size.get(),
                  self.var_keep_bold.get(), self.var_sp_before.get(),
                  self.var_sp_after.get(), self.var_font_keep.get())
        win = tk.Toplevel(self.root)
        win.title('粘贴设置')
        win.resizable(False, False)
        win.transient(self.root)
        body = ttk.Frame(win, padding=12)
        body.pack(fill='both', expand=True)
        ttk.Label(body, text='中文字体：').grid(row=0, column=0, sticky='w', pady=4)
        cb_font = ttk.Combobox(body, textvariable=self.var_font, width=18,
                               values=('等线', '宋体', '微软雅黑', '黑体', '楷体', '仿宋',
                                       'Segoe UI', 'Calibri', 'Arial', 'Times New Roman'))
        cb_font.grid(row=0, column=1, sticky='w', pady=4)
        ttk.Label(body, text='字号（磅，12=小四）：').grid(
            row=1, column=0, sticky='w', pady=4)
        cb_size = ttk.Combobox(body, textvariable=self.var_font_size, width=8,
                               values=('9', '10.5', '11', '12', '14', '16', '18',
                                       '22', '24', '28'))
        cb_size.grid(row=1, column=1, sticky='w', pady=4)
        ttk.Label(body, text='标题字体：').grid(row=2, column=0, sticky='w', pady=4)
        cb_hfont = ttk.Combobox(body, textvariable=self.var_head_font, width=18,
                                values=('等线', '宋体', '微软雅黑', '黑体', '楷体', '仿宋',
                                        'Segoe UI', 'Calibri', 'Arial', 'Times New Roman'))
        cb_hfont.grid(row=2, column=1, sticky='w', pady=4)
        ttk.Label(body, text='标题字号（磅）：').grid(
            row=3, column=0, sticky='w', pady=4)
        cb_hsize = ttk.Combobox(body, textvariable=self.var_head_size, width=8,
                                values=('12', '14', '16', '18', '22', '24', '28'))
        cb_hsize.grid(row=3, column=1, sticky='w', pady=4)
        ttk.Label(body, text='正文加粗字段保留：').grid(
            row=4, column=0, sticky='w', pady=4)
        ck_bold = ttk.Checkbutton(
            body, variable=self.var_keep_bold,
            text='勾选后 Markdown 中 **加粗** 的字段保留加粗；不勾选则全部不加粗')
        ck_bold.grid(row=4, column=1, sticky='w', pady=4)
        ttk.Label(body, text='段前（磅）：').grid(
            row=5, column=0, sticky='w', pady=4)
        ttk.Entry(body, textvariable=self.var_sp_before, width=6).grid(
            row=5, column=1, sticky='w', pady=4)
        ttk.Label(body, text='段后（磅）：').grid(
            row=5, column=2, sticky='w', padx=(12, 0), pady=4)
        ttk.Entry(body, textvariable=self.var_sp_after, width=6).grid(
            row=5, column=3, sticky='w', pady=4)
        # 复刻 Markdown 排版开关（勾选后上面字体字号五项失效；段前/段后两种模式均生效）
        ck_keep = ttk.Checkbutton(
            body, variable=self.var_font_keep,
            text='复刻 Markdown 排版（不单独设置字体/字号/加粗）')
        ck_keep.grid(row=6, column=0, columnspan=4, sticky='w', pady=(8, 0))

        def _sync_keep_state(*_):
            """勾选复刻模式时，字体字号加粗五项变灰失效"""
            state = 'disabled' if self.var_font_keep.get() else '!disabled'
            for w in (cb_font, cb_size, cb_hfont, cb_hsize, ck_bold):
                w.state((state,))

        self.var_font_keep.trace_add('write', _sync_keep_state)
        _sync_keep_state()
        ttk.Label(body, foreground='#666',
                  text='对「公式可编辑」「公式为图片」两种复制的粘贴效果生效；'
                       '段前/段后与行距两种模式均生效；'
                       '复刻模式：保留 Word 默认分级字号'
                  ).grid(row=7, column=0, columnspan=4, sticky='w', pady=(6, 0))
        btns = ttk.Frame(body)
        btns.grid(row=8, column=0, columnspan=4, sticky='e', pady=(10, 0))
        ttk.Button(btns, text='确定',
                   command=win.destroy).pack(side='left', padx=4)
        ttk.Button(btns, text='取消',
                   command=lambda: self._restore_settings(backup, win)
                   ).pack(side='left')
        # 弹窗居中显示在屏幕中间
        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f'+{(sw - w) // 2}+{(sh - h) // 2}')
        win.grab_set()

    def _restore_settings(self, backup, win):
        """设置对话框取消：还原修改前的值"""
        self.var_font.set(backup[0])
        self.var_font_size.set(backup[1])
        self.var_head_font.set(backup[2])
        self.var_head_size.set(backup[3])
        self.var_keep_bold.set(backup[4])
        self.var_sp_before.set(backup[5])
        self.var_sp_after.set(backup[6])
        self.var_font_keep.set(backup[7])
        win.destroy()

    def copy_rich(self):
        """Markdown → 含 OMML 条件注释的 Office HTML 直写剪贴板（免 Word），
        粘贴进 OneNote/Word 即得原生可编辑公式"""
        src = self.in_tb.get('1.0', 'end-1c')
        if not _HAS_LATEX:
            self._flash(self.btn_rich, '缺少 latex2mathml/mathml2omml', RICH_LABEL)
            return
        self.btn_rich.config(text='正在转换…')
        self.root.update_idletasks()
        try:
            plain = md_to_text(src)
            keep = self.var_font_keep.get()
            body = _office_math_html(
                src, keep=keep,
                hsize=self._head_size_val(), hfont=self.var_head_font.get(),
                keep_bold=self.var_keep_bold.get(),
                sp_before=self._sp_val(self.var_sp_before),
                sp_after=self._sp_val(self.var_sp_after))
            body_attr = ('' if keep else
                         f' style="font-family:\'{self.var_font.get()}\';'
                         f'font-size:{self._font_size_val()}pt;"')
            _set_clipboard_html(body, plain, body_attr=body_attr)
        except Exception as e:
            self._flash(self.btn_rich, f'失败:{e}', RICH_LABEL)
            return
        self._flash(self.btn_rich, '已复制 ✓ 粘贴试试', RICH_LABEL)

    def copy_rich_img(self):
        """公式渲染为 PNG 图片 + 真表格，以富文本写入剪贴板（无需 Word）；
        OneNote/浏览器/聊天软件中公式均为图片"""
        src = self.in_tb.get('1.0', 'end-1c')
        try:
            plain = md_to_text(src)
            # 复刻模式不设标题字体字号；统一模式各级标题用设置的字体与字号；
            # 复刻模式强制保留 **加粗** 字段（Markdown 排版形态）
            keep = self.var_font_keep.get()
            hsize = None if keep else self._head_size_val()
            hfont = None if keep else self.var_head_font.get()
            body = md_to_html(src, hsize=hsize, hfont=hfont,
                              keep_bold=(keep or self.var_keep_bold.get()),
                              sp_before=self._sp_val(self.var_sp_before),
                              sp_after=self._sp_val(self.var_sp_after))
            body_attr = ('' if keep else
                         f' style="font-family:\'{self.var_font.get()}\';'
                         f'font-size:{self._font_size_val()}pt;"')
            _set_clipboard_html(body, plain, body_attr=body_attr)
        except Exception as e:
            self._flash(self.btn_img, f'失败:{e}', IMG_LABEL)
            return
        self._flash(self.btn_img, '已复制 ✓', IMG_LABEL)


# ---------------- 应用图标 ----------------
# 48px PNG base64：窗口标题栏/任务栏图标（exe 文件图标由 PyInstaller --icon 嵌入）
_APP_ICON_B64 = (
    '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'
    'lplk0/rUy1HNapmVpE5ICnW0SKt5k1cb1iaSnM84si9osv7pI5GLoWKpNfM0YXVptEsrIaNQrDdiWlgi'
    'D0AxoP2URlEScFLkZM6ZV2J4XgpcHvl0em22jrNquqsQJ7Fw8nCEy2PWJQVZT6wv6bct+Fwbcs1n4onz'
    '1j18i0xdSswvCwvoqXGLN/4SYt0dSfqcHLe4+JaBLia8FZECdPWCVOrkGrG4o1KqhFJAOM/4654qPvet'
    'HhfY508Z9NyQxlA7aWjm6Efbp/aVdN/ng2jGgKjtzlzTplN+LQkkqeQpKDc3V8Eb24lOhrIGeOSnvfjT'
    'M1UXW8Vy84NS803Z+IV+rxrPvqjZJrtyziVdNpLaSkqtmqVq/f4Cr88p+OKT885DzoPdGr467UmDOexG'
    'obxOTdPa5fj/5vrGPQIAc5OM9hzISS2dcijN3BjkzWAHD1zLev/r+UxcJpa2KNup6CauBhALdm0RH4tN'
    'KBoo4qV9McMLnmeVYBbs2szrg6FfGdeqNBTb5B0BlrEwM2tVQhhVxk2gD6rRQ4NTMHje93qILZsFP84t'
    '2WAjWAWzYFcMJqOCn80FM+965HtsrU0nLrMB6T6tYBSsglmwqx0PQo3++abLURx8xYNPsIrZGmkzsKyG'
    'NVawCUbBKpgFe/KiO31X/L17rjxW0gO7Y1tFzIF70d1chpaI9WCjydfC/Wo8tfP7h1Y9nWGu/9TgngN6'
    '9NB98Xe2TTxQoPJuX/cMBWYWLrBBS/ICisFKK598rw9BPH8p4srOHx6+cX+GtfXHHqlWX/vE28P95Rsf'
    'Z/BDBG+tp5Kn6cV6hUDuD8HYCAxzjkAvzFQmnvr5axvGM4xNc9spIZ+/sXVqsEyFbcbKz23sOiIuGhav'
    'Xa96YeGRJz8eCAlqzFP6eIWjwz95eSD5uU0OvBz7Dy+PkcKhLmFPAAAAAElFTkSuQmCC'
)


def _set_icon(root):
    """设置窗口与任务栏图标（数据内嵌，源码/打包运行均生效）"""
    try:
        root.iconphoto(True, tk.PhotoImage(data=_APP_ICON_B64))
    except Exception:
        pass


def main():
    if _HAS_DND:
        root = TkinterDnD.Tk()  # 支持文件拖拽
    else:
        root = tk.Tk()
    _set_icon(root)
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
