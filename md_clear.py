# -*- coding: utf-8 -*-
"""
Markdown 转纯文本小工具
- 图形界面：左侧输入 Markdown，右侧实时预览（标题/粗体/斜体/代码块带格式显示）
- 「复制纯文本」：所有标记转纯文本，公式保持 $...$ LaTeX 原文
- 「复制（公式可编辑）」：生成含 OMML 公式、真表格、Heading 标题、代码块灰底的 docx，
  Word 后台全选复制，粘贴进 OneNote 得可编辑公式与完整格式（需本机装有 Word）
- 「复制（公式为图片）」：公式渲染为 PNG 图片的富文本（无需 Word）
- 列表自动编号：1. → (1) → a.，每个列表组独立计数
- 文件菜单：打开/保存 .md/.txt（Ctrl+O / Ctrl+S）；支持拖拽文件到输入框导入
- 内容与设置（窗口大小、开关状态）自动记忆
- Word 常驻加速开关：开启后 Word 进程后台常驻，复制从秒级降到毫秒级
"""
import base64
import ctypes
import html
import io
import json
import os
import re
import sys
import tempfile
import time
import unicodedata
from ctypes import wintypes
import tkinter as tk
from tkinter import filedialog, ttk

try:
    import latex2mathml.converter
    import mathml2omml
    _HAS_LATEX = True
except ImportError:
    _HAS_LATEX = False

try:
    import win32com.client
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    from docx.shared import Pt, RGBColor
    _HAS_WORDCOM = True
except ImportError:
    _HAS_WORDCOM = False

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


def latex_to_omml_element(latex, display=False):
    """LaTeX → MathML → OMML；剥掉 mathml2omml 输出的非标准 m:box 包装"""
    mml = latex2mathml.converter.convert(latex)
    omml = mathml2omml.convert(mml)
    omml = omml.replace('<m:box><m:e>', '').replace('</m:e></m:box>', '')
    m = re.match(r'^<m:oMath[^>]*>(.*)</m:oMath>$', omml, flags=re.S)
    body = m.group(1) if m else omml
    if display:
        xml = (f'<m:oMathPara xmlns:m="{OMML_NS}">'
               f'<m:oMath>{body}</m:oMath></m:oMathPara>')
    else:
        xml = f'<m:oMath xmlns:m="{OMML_NS}">{body}</m:oMath>'
    return parse_xml(xml)


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


def _render_table_html(rows, store, keep_bold=True):
    """渲染成 HTML 真表格（富文本复制用，单元格内公式以图片嵌入）"""
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
            h.append(f'<{tag}{style}>{_restore_html(cell, store, keep_bold=keep_bold)}</{tag}>')
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
               sp_after=0):
    """Markdown → HTML 片段：公式渲染为内嵌 PNG 图片、表格转真表格（「公式为图片」复制用）
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
            t = _render_table_html(payload, store, keep_bold=keep_bold)
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
            body = _restore_html(b[2], store, keep_bold=keep_bold)
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
        parts.append(f'<p style="{margin}">{_restore_html(payload, store, keep_bold=keep_bold)}</p>')
    result = ''.join(parts)
    # 去掉首尾的空段落
    blank = '<p><br/></p>'
    while result.startswith(blank):
        result = result[len(blank):]
    while result.endswith(blank):
        result = result[:-len(blank)]
    return result


# ---------------- docx 生成 + Word 中转复制 ----------------
def _run_shd(r, fill='F2F2F2'):
    """run 加底纹（行内代码灰底）"""
    r._element.get_or_add_rPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def _para_shd(p, fill='F2F2F2'):
    """段落加底纹（代码块整行灰底）"""
    p._p.get_or_add_pPr().append(
        parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def _docx_fill_line(p, line, store, bold=False, keep_bold=True):
    """把一行写入 docx 段落：粗体/斜体/删除线转真格式 run，
    行内代码等宽灰底，公式转 OMML 公式对象（bold 用于表头行）；
    keep_bold=False 时 Markdown **加粗** 片段不产生加粗 run（全部不加粗）"""
    t = _plain_pre(line)

    def add_text(s, fmt=frozenset(), code=False):
        if not s:
            return
        r = p.add_run(s)
        if bold or (keep_bold and 'b' in fmt):
            r.bold = True
        if 'i' in fmt:
            r.italic = True
        if 's' in fmt:
            r.font.strike = True
        if code:
            r.font.name = 'Consolas'
            _run_shd(r)

    for frag, fmt in _parse_fmt(t):
        last = 0
        for m in PH_RE.finditer(frag):
            add_text(frag[last:m.start()], fmt)
            seg = store[int(m.group(1))]
            if seg['kind'] == 'code':
                add_text(seg['text'], fmt, code=True)
            elif _HAS_LATEX:
                try:
                    p._p.append(latex_to_omml_element(seg['latex'], seg['display']))
                except Exception:
                    add_text(seg['text'])  # 转换失败回退为 $...$ 原文
            else:
                add_text(seg['text'])
            last = m.end()
        add_text(frag[last:], fmt)


def _set_style_font(st, name, size_pt, bold):
    """样式设置字体：西文 + 中文字体（eastAsia）、字号、加粗；
    并清除模板自带的主题字体属性（asciiTheme 等）——按 OOXML 规范主题属性
    优先于显式字体，不清除会导致标题仍按 majorEastAsia 渲染成 MS Gothic/微软雅黑"""
    st.font.name = name          # ascii/hAnsi 西文字体
    st.font.size = Pt(size_pt)
    st.font.bold = bold
    rpr = st.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn('w:eastAsia'), name)  # 中文字体
    for attr in ('asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'):
        rfonts.attrib.pop(qn('w:' + attr), None)


def md_to_docx(md, font='等线', size=12, keep=False, hsize=14, hfont=None,
               keep_bold=False, sp_before=0, sp_after=0):
    """Markdown → docx 文档：公式为 OMML 公式对象、表格为真表格、标题用 Heading 样式、
    代码块等宽灰底；段落格式：单倍行距，段前/段后按设置（默认 0 磅紧凑，含表格内）
    font/size：正文字体、字号（磅）
    hsize/hfont：各级标题的统一字号（磅，默认 14）与字体（None 表示与正文同字体）
    keep_bold：正文加粗字段保留——勾选后 Markdown 中 **加粗** 的片段保留加粗，
    不勾选则全部字段不加粗（默认不勾选；复刻模式强制保留）
    sp_before/sp_after：段前/段后间距（磅，默认 0）
    keep：复刻 Markdown 排版——不设置字体/字号/加粗，标题与表头保持加粗，
    粘贴后使用目标软件的默认字体"""
    blocks, store = _convert(md)
    doc = Document()
    # 复刻模式强制保留 Markdown 加粗形态；统一模式由 keep_bold 选项控制
    kb = keep_bold or keep
    # Normal 样式：单倍行距、段前段后按设置（对所有段落生效，含表格内）；
    # 全局加粗已移除：Normal 恒不加粗，** 片段按 keep_bold 单独决定是否加粗
    pf = doc.styles['Normal'].paragraph_format
    pf.space_before = Pt(sp_before)
    pf.space_after = Pt(sp_after)
    pf.line_spacing = 1
    if not keep:
        _set_style_font(doc.styles['Normal'], font, size, False)
    # Heading 样式：段前段后按设置，颜色改黑（默认蓝色在笔记里太扎眼）；
    # 统一模式：各级标题统一用 hsize/hfont（设置里可配，默认 14 磅/与正文字体同），
    # 显式不加粗（覆盖模板默认加粗）
    # 复刻模式：字体字号全不动（保留 Word 默认的加粗与分级字号，即 Markdown 标题形态）
    for lv in range(1, 7):
        try:
            st = doc.styles[f'Heading {lv}']
        except KeyError:
            continue
        st.paragraph_format.space_before = Pt(sp_before)
        st.paragraph_format.space_after = Pt(sp_after)
        st.font.color.rgb = RGBColor(0, 0, 0)
        if not keep:
            _set_style_font(st, hfont or font, hsize, False)
    prev_table = False
    prev_blank = False
    for b in blocks:
        kind = b[0]
        payload = b[1]
        if kind == 'table':
            aligns, rows = _split_sep_row(payload)
            if not rows:
                continue
            if prev_table:  # 相邻两个表格间加空段，防止 Word 把它们合并
                doc.add_paragraph()
            ncols = max(len(r) for r in rows)
            rows = [r + [''] * (ncols - len(r)) for r in rows]
            table = doc.add_table(rows=len(rows), cols=ncols)
            table.style = 'Table Grid'
            for i, r in enumerate(rows):
                for j, cell in enumerate(r):
                    p = table.rows[i].cells[j].paragraphs[0]
                    # 表头：复刻模式加粗（Markdown 表格惯例）；统一模式不加粗
                    _docx_fill_line(p, cell, store, bold=(i == 0 and keep),
                                    keep_bold=kb)
                    if aligns and j < len(aligns) and aligns[j] != 'left':
                        p.alignment = (WD_ALIGN_PARAGRAPH.CENTER
                                       if aligns[j] == 'center'
                                       else WD_ALIGN_PARAGRAPH.RIGHT)
            prev_table = True
            prev_blank = False
            continue
        prev_table = False
        if kind == 'heading':
            p = doc.add_paragraph()
            p.style = doc.styles[f'Heading {min(payload, 6)}']
            _docx_fill_line(p, b[2], store, keep_bold=kb)
            prev_blank = False
            continue
        if kind == 'code':
            for cl in payload.split('\n'):
                p = doc.add_paragraph()
                _para_shd(p)
                r = p.add_run(cl)
                r.font.name = 'Consolas'
            prev_blank = False
            continue
        if _plain_line(payload, store).strip() == '':
            if prev_blank:  # 连续空行只保留一个
                continue
            prev_blank = True
            doc.add_paragraph()  # 空行 → 空段落（保留作者的分段意图）
            continue
        prev_blank = False
        _docx_fill_line(doc.add_paragraph(), payload, store, keep_bold=kb)
    return doc


_word_app = None  # 常驻的 Word 实例（开关开启后复用，避免每次启动 Word）


def _get_word():
    """获取/启动常驻 Word 实例（隐藏窗口）"""
    global _word_app
    if _word_app is None:
        _word_app = win32com.client.DispatchEx('Word.Application')
        _word_app.Visible = False
        _word_app.DisplayAlerts = 0
    return _word_app


def _release_word():
    """释放常驻 Word 实例（窗口退出时调用）"""
    global _word_app
    if _word_app is not None:
        try:
            _word_app.Quit()
        except Exception:
            pass
        _word_app = None


def _copy_via_word(docx_path, resident=False):
    """用 Word 打开 docx → 全选复制 → 关闭：剪贴板获得 Office 原生公式格式
    resident=True 复用常驻实例（首次约 2-3 秒，之后毫秒级）"""
    if not resident:
        # DispatchEx 起独立实例，用完即退，不会动用户正开着的 Word
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
            d = word.Documents.Open(os.path.abspath(docx_path), False, True, False)
            try:
                d.Content.Copy()
            finally:
                d.Close(0)  # 0 = 不保存
        finally:
            word.Quit()
        return
    for attempt in (1, 2):  # 实例可能被外部杀掉：失败则重建一次
        word = _get_word()
        try:
            d = word.Documents.Open(os.path.abspath(docx_path), False, True, False)
            try:
                d.Content.Copy()
            finally:
                d.Close(0)
            return
        except Exception:
            _release_word()
            if attempt == 2:
                raise


def copy_md_via_word(md, resident=False, font='等线', size=12, keep=False,
                     hsize=14, hfont=None, keep_bold=False, sp_before=0,
                     sp_after=0):
    """Markdown → docx（OMML 公式 + 真表格）→ Word 全选复制进剪贴板"""
    doc = md_to_docx(md, font=font, size=size, keep=keep, hsize=hsize,
                     hfont=hfont, keep_bold=keep_bold, sp_before=sp_before,
                     sp_after=sp_after)
    fd, path = tempfile.mkstemp(suffix='.docx')
    os.close(fd)
    try:
        doc.save(path)
        _copy_via_word(path, resident)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass  # Word 进程退出稍有延迟，删不掉就留在临时目录


# ---------------- Windows 剪贴板（CF_HTML 富文本） ----------------
def _set_clipboard_html(html_fragment, plain_text, font='等线', size=12,
                        keep=False):
    """把 HTML（内嵌 OMML 公式、真表格）和纯文本同时写入剪贴板；
    font/size 作为粘贴产物的全局默认字体设置；
    keep=True 时不写字体样式，粘贴后用目标软件默认字体（复刻 Markdown 排版）"""
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

    body_style = ('' if keep else
                  f'style="font-family:\'{font}\';font-size:{size}pt;"')
    cf_html = user32.RegisterClipboardFormatW('HTML Format')
    header = ('Version:0.9\r\nStartHTML:{:010d}\r\nEndHTML:{:010d}\r\n'
              'StartFragment:{:010d}\r\nEndFragment:{:010d}\r\n')
    pre = (f'<html><head><meta charset="utf-8"></head>'
           f'<body {body_style}>\r\n<!--StartFragment-->')
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
PLAIN_LABEL = '复制纯文本'

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
        self.var_word_keep = tk.BooleanVar(value=bool(cfg.get('word_keep')))
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
        # 设置菜单：粘贴产物的默认字体
        set_menu = tk.Menu(menubar, tearoff=0)
        set_menu.add_command(label='粘贴设置…', command=self.open_settings)
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
        self.chk_word = ttk.Checkbutton(
            bar, text='Word 常驻加速', variable=self.var_word_keep,
            command=self._toggle_word_keep)
        self.chk_word.pack(side='left', padx=(8, 0))
        hint = ('「公式可编辑」后台调用 Word；开启常驻后首次稍慢、之后毫秒级复制；'
                 '内容与设置自动记忆')
        if _HAS_DND:
            hint = '支持拖入 .md/.txt 文件。' + hint
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
        """退出时：保存配置（内容/窗口/开关）并释放常驻 Word"""
        _save_config({
            'text': self.in_tb.get('1.0', 'end-1c'),
            'geometry': self.root.geometry(),
            'word_keep': self.var_word_keep.get(),
            'font': self.var_font.get(),
            'font_size': self._font_size_val(),
            'head_size': self._head_size_val(),
            'head_font': self.var_head_font.get(),
            'keep_bold': self.var_keep_bold.get(),
            'sp_before': self._sp_val(self.var_sp_before),
            'sp_after': self._sp_val(self.var_sp_after),
            'keep_markdown': self.var_font_keep.get(),
        })
        _release_word()
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
        content = self.out_tb.get('1.0', 'end-1c')
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self._flash(self.btn_copy, '已复制 ✓', PLAIN_LABEL)

    def _toggle_word_keep(self):
        """开关切换：关闭时立即释放常驻实例"""
        if not self.var_word_keep.get():
            _release_word()

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
        """生成 docx（OMML 公式+真表格）→ Word 后台全选复制 → 粘贴进 OneNote 即可编辑公式"""
        src = self.in_tb.get('1.0', 'end-1c')
        if not _HAS_WORDCOM or not _HAS_LATEX:
            self._flash(self.btn_rich, '缺少 pywin32/python-docx', RICH_LABEL)
            return
        self.btn_rich.config(text='正在调用 Word…')
        self.root.update_idletasks()
        try:
            copy_md_via_word(src, resident=self.var_word_keep.get(),
                             font=self.var_font.get(),
                             size=self._font_size_val(),
                             keep=self.var_font_keep.get(),
                             hsize=self._head_size_val(),
                             hfont=self.var_head_font.get(),
                             keep_bold=self.var_keep_bold.get(),
                             sp_before=self._sp_val(self.var_sp_before),
                             sp_after=self._sp_val(self.var_sp_after))
        except Exception as e:
            self._flash(self.btn_rich, f'失败:{e}', RICH_LABEL)
            return
        self._flash(self.btn_rich, '已复制 ✓ 粘贴试试', RICH_LABEL)

    def copy_rich_img(self):
        """公式渲染为 PNG 图片 + 真表格，以富文本写入剪贴板（无需 Word）"""
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
            _set_clipboard_html(body, plain, font=self.var_font.get(),
                                size=self._font_size_val(), keep=keep)
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
