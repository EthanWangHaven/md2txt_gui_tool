# md2txt_gui — Markdown 转纯文本小工具

Windows 桌面小工具：把 Markdown 转换成适合粘贴进 **OneNote / Word / 聊天软件** 的格式。

左侧输入 Markdown，右侧实时预览，一键复制。公式可编辑、表格是真表格、列表自动编号。

## 功能特性

- **实时预览**：标题、粗体、斜体、删除线、代码块均带格式显示，公式渲染为图片
- **列表自动编号**：统一按层级编号 —— 第一层 `1. 2. 3.`，第二层 `(1)(2)(3)`，第三层 `a. b. c.`，每个列表组独立计数
- **标题层级**：`#` → Word/OneNote 大纲标题（Heading 样式，紧凑排版、黑色）
- **行内格式**：`**粗体**`、`*斜体*`、`***粗斜体***`、`~~删除线~~` → 真实格式，不再是星号
- **代码**：行内代码等宽灰底；代码块整块灰底
- **公式**：支持 `$...$`、`$$...$$`、`\(...\)`、`\[...\]`
- **表格**：转换为真表格（含对齐），单元格内格式与公式照常生效
- **文件操作**：菜单打开/保存 `.md`、`.txt`（Ctrl+O / Ctrl+S），支持拖拽文件到输入框
- **内容记忆**：输入内容、窗口大小、开关状态自动保存，下次启动直接恢复
- **Word 常驻加速**：开关开启后 Word 进程后台常驻，公式复制从秒级降到毫秒级

## 三种复制模式

| 按钮 | 原理 | 粘贴效果 |
|------|------|----------|
| 复制纯文本 | 直接转换 | 纯文本；公式保留 `$...$` LaTeX 原文 |
| 复制（公式可编辑） | 生成含 OMML 公式的 docx，Word 后台全选复制 | 公式为**可编辑的原生公式对象**，真表格、真标题（需本机装有 Word） |
| 复制（公式为图片） | 公式渲染为 PNG 的富文本 | 公式为图片，无需 Word |

## 环境要求

- Windows 10/11
- Python 3.10+（运行源码时需要；exe 版无需）
- 「公式可编辑」模式需要本机安装 Microsoft Word
- 「公式为图片」模式无额外要求

## 安装与运行（源码方式）

```bash
pip install latex2mathml mathml2omml python-docx pywin32 matplotlib tkinterdnd2
python md2txt_gui.py
```

依赖缺失时会自动降级（如未装 tkinterdnd2 则不可拖拽、未装 Word 组件则隐藏富文本复制），不会崩溃。

## 打包成 exe

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name md2txt_gui --collect-all tkinterdnd2 --collect-data latex2mathml md2txt_gui.py
```

生成 `dist/md2txt_gui.exe`，单文件约 45MB，无需 Python 环境即可运行。

## 配置文件

首次退出后自动在同目录生成 `config.json`，保存：

- 上次输入框内容
- 窗口大小位置
- 「Word 常驻加速」开关状态

删除该文件即可恢复初始状态。**该文件包含个人输入内容，请勿提交到仓库。**

## 目录说明

```
md2txt_gui/
├── md2txt_gui.py          # 主程序（单文件，含全部逻辑）
├── md2txt_gui.exe         # 打包后的可执行文件（构建产物）
├── Markdown转换工具.lnk    # 快捷方式
├── config.json            # 运行时自动生成（配置/内容记忆，勿提交）
└── README.md
```

建议配合 `.gitignore` 排除 `md2txt_gui.exe`、`config.json`、`__pycache__/`。

## 技术要点

- **OMML 公式链路**：LaTeX → MathML（latex2mathml）→ OMML（mathml2omml）→ 嵌入 docx，Word 打开后 `Content.Copy()` 使剪贴板获得 Office 原生公式格式
- **Word 中转**：`DispatchEx` 起独立 Word 实例，只读打开、复制、关闭，不影响用户正开着的 Word
- **富文本剪贴板**：CF_HTML + CF_UNICODETEXT 双格式写入（ctypes 直接调 Win32 API）
- **公式图片**：matplotlib mathtext 渲染 PNG，base64 内嵌 HTML
