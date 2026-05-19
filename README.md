# office-dump

> Dump every embedded piece of a Microsoft Office OOXML file into a single, predictable output directory.

[简体中文](README_zh.md)

`office-dump` is a [Claude Code](https://claude.com/claude-code) skill (also usable as a plain CLI) that opens `.docx` / `.xlsx` / `.pptx` (and the macro-enabled `.docm` / `.xlsm` / `.pptm`) and writes out:

- the body as **Markdown**,
- every **image, audio, and video** asset,
- every **embedded object** (with one-layer OLE peeling),
- every **chart** as both data **CSV** and a rendered **PNG**,
- **comments**, **footnotes**, **headers & footers**,
- core / extended / custom **metadata**,
- extracted **VBA source** (`.bas`) plus an `olevba` static-analysis report,
- and for PPTX: one **rendered PNG per slide**, optionally a full **PDF**.

It works on Windows, macOS, and Linux. The only platform-specific part is the slide-rendering backend.

## Quick start

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

The output lands in `./dump_output/<basename>/`.

## Output layout

```
<OUTDIR>/<basename>/
├── content.md          # body as Markdown
├── metadata.json       # core / app / custom properties
├── media/              # images, audio, video (verbatim from the OOXML)
├── embedded/           # embedded objects; OLE containers peeled one level
├── charts/             # chart*.csv + chart*.png
├── comments.json       # comments, when present
├── macros/             # VBA .bas + olevba report, when present
├── slides/             # PPTX only: slide_01.png, slide_02.png, ...
├── full.pdf            # PPTX + LibreOffice backend by-product
└── manifest.json       # full inventory + warnings[]
```

## CLI

```
python dump.py <INPUT_FILE> [-o OUTDIR] [options]
```

| Option | Default | Description |
|---|---|---|
| `-o`, `--output` | `./dump_output` | output root directory |
| `--slide-renderer` | `auto` | `auto` / `powerpoint` / `libreoffice` / `none` |
| `--dpi` | `150` | slide rendering DPI |
| `--no-charts` | off | skip chart extraction |
| `--no-macros` | off | skip VBA macro extraction |

## Slide rendering backends

| Backend | Where | Notes |
|---|---|---|
| `powerpoint` | Windows + Office installed | uses COM automation, highest fidelity |
| `libreoffice` | any OS | needs `soffice` on PATH or in a standard location; ~300 MB to install |
| `none` | any OS | text / media / metadata still exported, but no `slides/*.png` |

`auto` picks `powerpoint` first on Windows, then `libreoffice`, then degrades to `none`.

## Supported formats

| Extension | Supported | Notes |
|---|---|---|
| `.docx`, `.docm` | ✅ | macro variant also extracts VBA |
| `.xlsx`, `.xlsm` | ✅ | formula + cached value rendered side-by-side |
| `.pptx`, `.pptm` | ✅ | adds per-slide PNGs and (optionally) full PDF |
| `.doc`, `.xls`, `.ppt` | ❌ | convert to OOXML first (e.g. via LibreOffice) |
| encrypted documents | ❌ | the CLI exits with a clear error |

## Install as a Claude Code skill

Once installed under `~/.claude/skills/office-dump/`, Claude Code will discover and trigger it when the user says things like "dump this pptx", "extract images from this Word file", or the equivalents in any language.

The most reliable path is to follow [`INSTALL.md`](INSTALL.md), which is written so an AI agent can execute it end-to-end (computes destination, copies sources, builds venv, installs deps, optionally installs LibreOffice, and runs a smoke test).

## Design

- Structured text comes from `python-docx` / `openpyxl` / `python-pptx`; everything those libraries can't reach — raw chart XML, embeddings, custom parts — is read directly from the OOXML ZIP.
- Embedded objects are **not** dumped recursively. Only the top-level OLE container is peeled one layer.
- Order preservation: Word walks body child elements (not just `paragraphs`); PowerPoint walks slides in deck order and shapes in z-order.
- A failure on a single part is recorded as a warning in `manifest.json`; the run does not abort.

## License

MIT.
