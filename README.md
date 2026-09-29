# MdClear — Markdown 转纯文本小工具

<div align="center">

<img src="app.ico" width="128" alt="MdClear 图标">

</div>

Windows 桌面小工具：把 Markdown 转换成适合粘贴进 **OneNote / Word / 聊天软件** 的格式。

左侧输入 Markdown，右侧实时预览，一键复制。公式可编辑、表格是真表格、列表自动编号，全程**无需安装 Word**。

## 功能特性

- **实时预览**：标题、粗体、斜体、删除线、代码块均带格式显示；公式（含表格单元格内的公式）渲染为图片
- **列表自动编号**：统一按层级编号 —— 第一层 `1. 2. 3.`，第二层 `(1)(2)(3)`，第三层 `a. b. c.`，第四层起 `i. ii. iii.`，每个列表组独立计数
- **标题层级**：`#` → Word/OneNote 大纲标题（统一模式显式不加粗、黑色）
- **行内格式**：`**粗体**`、`*斜体*`、`***粗斜体***`、`~~删除线~~` → 真实格式，不再是星号
- **代码**：行内代码等宽灰底；代码块整块灰底
- **公式**：支持 `$...$`、`$$...$$`、`\(...\)`、`\[...\]`；自动修复 AI 输出的非标准写法（`$ x $` 内侧空格、单行 `$` 块、`\kern`、`\operatorname`、Unicode 连字符/减号等）
- **表格**：转换为真表格（含对齐），单元格内格式与公式照常生效
- **单换行保留**：AI 笔记常见的单换行分行（①②③条目等）粘贴后保持换行
- **文件操作**：菜单打开/保存 `.md`、`.txt`（Ctrl+O / Ctrl+S），支持拖拽文件到输入框
- **粘贴设置**：菜单「设置 → 粘贴设置」可配置粘贴产物的中文字体、字号、标题字号、是否加粗、段前段后间距，默认**等线 / 12 磅（小四）/ 不加粗**
- **关于**：菜单「设置 → 关于」查看作者、联系方式与开源地址
- **内容记忆**：输入内容、窗口大小位置、粘贴设置自动保存，下次启动直接恢复
- **内置 Pandoc**：打包内置 pandoc.exe，「公式可编辑」语法覆盖远超内置转换器（矩阵、对齐环境、任务列表、脚注等）

## 三种复制模式

| 按钮 | 原理 | 粘贴效果 |
|------|------|----------|
| 复制源MD | 原样复制输入框内容 | Markdown 源文本，无任何转换 |
| 复制（公式可编辑） | Markdown → Pandoc → HTML(MathML) → OMML 条件注释直写剪贴板（免 Word） | OneNote/Word 中公式为**可编辑的原生公式对象**，真表格、真标题；其他软件显示公式文本回退 |
| 复制（公式为图片） | 公式渲染为 PNG 的富文本 | 任何软件（OneNote/浏览器/聊天软件）中公式均为图片 |

## 环境要求

- Windows 10/11
- Python 3.10+（运行源码时需要；exe 版无需）
- pandoc.exe（源码方式可选，放到 `pandoc/pandoc.exe` 增强「公式可编辑」转换质量；exe 版已内置）

## 安装与运行（源码方式）

```bash
pip install latex2mathml mathml2omml matplotlib tkinterdnd2
python md_clear.py
```

依赖缺失时会自动降级（如未装 tkinterdnd2 则不可拖拽、未装 latex2mathml/mathml2omml 则公式以原文输出），不会崩溃。

## 打包成 exe

```bash
pip install pyinstaller latex2mathml mathml2omml matplotlib tkinterdnd2
# pandoc.exe 放到 pandoc/pandoc.exe 后打包内置
pyinstaller --noconfirm --onefile --windowed --icon=app.ico --name MdClear --add-binary "pandoc\pandoc.exe;pandoc" --collect-all tkinterdnd2 --collect-data latex2mathml md_clear.py
```

生成 `dist/MdClear.exe`（约 79MB），无需 Python 环境即可运行。`app.ico` 为应用图标（窗口/任务栏图标已内嵌在源码中，替换图标时同步更新 `md_clear.py` 里的 `_APP_ICON_B64`）。

## 配置文件

首次退出后自动在同目录生成 `config.json`，保存：

- 上次输入框内容
- 窗口大小位置（越界时自动居中）
- 粘贴设置（字体/字号/标题字号/加粗/段前段后间距/复刻模式）

删除该文件即可恢复初始状态。**该文件包含个人输入内容，请勿提交到仓库。**

## 目录说明

```
md2txt_gui/
├── md_clear.py            # 主程序（单文件，含全部逻辑）
├── MdClear.exe            # 打包后的可执行文件（构建产物）
├── pandoc/pandoc.exe      # Pandoc 引擎（打包内置，不入库）
├── config.json            # 运行时自动生成（配置/内容记忆，勿提交）
└── README.md
```

## 使用提示

- 粘贴到 OneNote 时请使用「**保留源格式粘贴**」，公式与表格才能保持完整格式
- 左侧输入框支持直接拖入 `.md` / `.txt` 文件
- 公式从聊天窗口复制时若混入不可见字符（如非断行连字符），程序会自动还原

## 技术要点

- **OMML 公式剪贴板链路（免 Word）**：Markdown → Pandoc（`--math-method=mathml`）→ MathML → mathml2omml 转 OMML → 包进 `<!--[if gte msEquation 12]>…<![endif]-->` 条件注释 → CF_HTML 剪贴板；OneNote/Word 读取原生 OMML 渲染为可编辑公式，其他软件取回退内容
- **公式图片**：matplotlib mathtext 渲染 PNG，base64 内嵌 HTML；预览与「公式为图片」模式共用
- **LaTeX 兼容清洗**：还原 AI 输出的 `\ `、`\,`、`\kern`、`\operatorname` 等噪声命令与 Unicode 变体字符，修复 `$ x $`、单行 `$` 块等非标准分隔符；清洗覆盖 OMML、预览、图片三条链路
- **CF_HTML**：ctypes 直接调 Win32 API 写入 HTML Format + CF_UNICODETEXT 双格式
- **实时预览**：内置轻量 Markdown 解析器逐块转换，Tk Text 渲染，公式 matplotlib 出图，表格内公式同样渲染
- **列表压平**：交给 Pandoc 前把列表压平为带转义编号的段落，OneNote 中不出现圆点符号

## 作者

- **作者**：wangce
- **联系方式**：2253246@tongji.edu.cn
- **开源地址**：https://github.com/EthanWangHaven/md2txt_gui_tool

Copyright © 2026 wangce. 保留所有权利。