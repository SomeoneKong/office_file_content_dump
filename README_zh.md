# office-dump

> 把 Microsoft Office OOXML 文件中嵌入的所有内容完整导出到一个统一目录。

[English](README.md)

`office-dump` 是一个 [Claude Code](https://claude.com/claude-code) skill（也可以作为普通 CLI 使用），用于打开 `.docx` / `.xlsx` / `.pptx`（以及启用宏的 `.docm` / `.xlsm` / `.pptm`），并把以下内容导出到一个目录：

- 主体内容（**Markdown**）
- 所有**图片、音频、视频**
- 所有**嵌入对象**（OLE 容器解一层）
- 所有**图表**——既导出**底层 CSV**，又用 matplotlib 渲染**PNG**
- **批注**、**脚注**、**页眉页脚**
- 核心 / 扩展 / 自定义**元数据**
- 提取的 **VBA 源码**（`.bas`）+ `olevba` 静态分析报告
- 对 PPTX：每张幻灯片一张**渲染 PNG**，以及可选的完整 **PDF**

跨平台支持 Windows / macOS / Linux。唯一依赖平台的部分是幻灯片渲染后端。

## 快速开始

```bash
git clone https://github.com/SomeoneKong/office_file_content_dump.git
cd office_file_content_dump
python -m venv venv
# Windows
venv\Scripts\python.exe -m pip install -r requirements.txt
venv\Scripts\python.exe dump.py path\to\file.docx
# macOS / Linux
venv/bin/python -m pip install -r requirements.txt
venv/bin/python dump.py path/to/file.docx
```

产物输出到 `./dump_output/<basename>/`。

## 输出结构

```
<OUTDIR>/<basename>/
├── content.md          # 主体 Markdown
├── metadata.json       # core / app / custom 属性
├── media/              # 图片、音视频（原样从 OOXML 拷贝）
├── embedded/           # 嵌入对象；OLE 容器解一层
├── charts/             # chart*.csv + chart*.png
├── comments.json       # 批注（若有）
├── macros/             # VBA .bas + olevba 报告（若有）
├── slides/             # 仅 PPTX：slide_01.png, slide_02.png, ...
├── full.pdf            # PPTX + LibreOffice 后端的副产品
└── manifest.json       # 总清单 + warnings[]
```

## 命令行

```
python dump.py <INPUT_FILE> [-o OUTDIR] [选项]
```

| 选项 | 默认 | 说明 |
|---|---|---|
| `-o`, `--output` | `./dump_output` | 输出根目录 |
| `--slide-renderer` | `auto` | `auto` / `powerpoint` / `libreoffice` / `none` |
| `--dpi` | `150` | 幻灯片渲染 DPI |
| `--no-charts` | 关闭 | 跳过 chart 提取 |
| `--no-macros` | 关闭 | 跳过 VBA 宏提取 |

## 幻灯片渲染后端

| 后端 | 适用环境 | 说明 |
|---|---|---|
| `powerpoint` | Windows + 已装 Office | 通过 COM 自动化，保真度最高 |
| `libreoffice` | 任意 OS | 要求 `soffice` 在 PATH 或常见位置；约 300 MB |
| `none` | 任意 OS | 仍输出文本/媒体/元数据，但不生成 `slides/*.png` |

`auto`：Windows 上优先 `powerpoint`，否则 `libreoffice`，再否则 `none`。

## 支持的格式

| 扩展名 | 支持 | 说明 |
|---|---|---|
| `.docx`, `.docm` | ✅ | macro 变体会额外提取 VBA |
| `.xlsx`, `.xlsm` | ✅ | 公式与缓存值并排显示 |
| `.pptx`, `.pptm` | ✅ | 额外生成每页 PNG 及可选 PDF |
| `.doc`, `.xls`, `.ppt` | ❌ | 请先用 LibreOffice 等转换为 OOXML |
| 加密文档 | ❌ | CLI 会以明确错误退出 |

## 作为 Claude Code skill 安装

安装到 `~/.claude/skills/office-dump/` 后，Claude Code 会自动发现并在用户表达"导出这个 pptx""提取 Word 里的图片"等意图时触发本 skill（其他语言的等价表达同样有效）。

最稳妥的安装方式是按 [`INSTALL.md`](INSTALL.md) 执行——该文档专门为 AI agent 自动化设计：求出目标路径、拷贝源码、建 venv、装依赖、（可选）装 LibreOffice、跑端到端测试。

## 设计要点

- 结构化文本走 `python-docx` / `openpyxl` / `python-pptx`；这些库够不着的部分（chart XML、嵌入对象、自定义 part）直接从 OOXML ZIP 读。
- 嵌入对象**不**递归 dump；只把顶层 OLE 容器解一层。
- 顺序保真：Word 遍历 body 子元素（不仅是 `paragraphs`）；PPT 按 slide 顺序 + 形状 z-order。
- 单个 part 失败仅记 warning 到 `manifest.json`，不中断整体导出。

## License

MIT.
