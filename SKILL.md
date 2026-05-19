---
name: office-dump
description: Dump every embedded piece of a Microsoft Office OOXML file (.docx / .xlsx / .pptx and the .docm / .xlsm / .pptm variants) into a single output directory — body as Markdown, plus media, embedded objects, chart data and renderings, comments, metadata, VBA macros; PPTX also gets one rendered PNG per slide. Trigger on intents like "export Office content", "extract images / charts / macros from docx/xlsx/pptx", "convert PPT pages to images", "pull the embedded Excel out of a Word doc", as well as the Chinese equivalents "导出 Office 文件内容", "把 PPT 每页转成图片", "提取 docx 里的图片和嵌入", "把 Word 里的 Excel 拿出来".
---

# Office Dump Skill

Export every embedded piece of a Microsoft Office OOXML file (.docx / .xlsx / .pptx and .docm / .xlsm / .pptm) into a single output directory, in a layout that is easy for humans or LLMs to consume.

## When to use

Trigger this skill when the user says any of the following (or the equivalents in any language):

- "dump / extract / unpack this docx / xlsx / pptx"
- "pull the images / embedded objects / comments / macros out of an Office file"
- "convert each slide of this PPT to a PNG"
- "extract the Excel that is embedded inside this Word doc"

Do **not** use it for:

- Legacy binary formats `.doc` / `.xls` / `.ppt` (first convert them to OOXML with LibreOffice etc., then run this skill).
- Encrypted documents (the CLI exits with an error).

## How to invoke

Run the CLI from the project's own venv:

```
venv/Scripts/python.exe dump.py <INPUT_FILE> [-o OUTDIR] [options]
```

Main options:

- `-o, --output DIR` – output root, defaults to `./dump_output`
- `--slide-renderer {auto,powerpoint,libreoffice,none}` – PPTX rendering backend, default `auto`
- `--dpi N` – slide rendering DPI, default 150
- `--no-charts` – skip chart CSV/PNG extraction
- `--no-macros` – skip VBA macro extraction

## Output layout

```
<OUTDIR>/<basename>/
├── content.md          # body as Markdown
├── metadata.json       # core / app / custom properties
├── media/              # images, audio, video
├── embedded/           # embedded objects, copied verbatim; OLE containers peeled one layer
├── charts/             # chart*.csv + chart*.png
├── comments.json       # comments (if any)
├── macros/             # VBA .bas sources + olevba report (if any)
├── slides/             # PPTX only: slide_01.png, slide_02.png, ...
├── full.pdf            # by-product when PPTX is rendered via LibreOffice
└── manifest.json       # full inventory + warnings[]
```

## Cross-platform prerequisites

- **Required**: Python 3.10+ and the packages listed in `requirements.txt`
- **Optional** (only needed for PPTX slide rendering — pick either):
  - Windows: installed Microsoft PowerPoint (used via COM automation)
  - Any OS: installed LibreOffice
- Without either, PPTX still gets text/media/metadata; only `slides/*.png` and `full.pdf` are missing.

## Design notes

- Structured text is read via python-docx / openpyxl / python-pptx; whatever they cannot reach (media, embeddings, raw chart XML, custom parts) is read directly from the OOXML ZIP.
- Embedded objects are **not** recursively dumped (only their top-level OLE container is peeled).
- Order preservation: Word walks body child elements (not just `paragraphs`); PPT walks slides in deck order and shapes in z-order.
- A failure on a single part is recorded as a warning in the manifest; the run does not abort.
