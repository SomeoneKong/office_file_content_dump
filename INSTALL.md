# Office Dump Skill — installation guide (for agents)

This document is written for an AI agent to follow autonomously. Each step lists the **command**, the **expected result**, and a **troubleshooting** section. Execute the steps in order. Never skip a failing step; fix it per the troubleshooting notes and retry.

`<SRC>` is the directory that contains the source tree. If the agent is reading this file from inside the source tree, `<SRC>` is the current working directory.

---

## 0. Prerequisite check

Run:

```
python --version
```

or (on macOS / Linux, where `python` may not exist):

```
python3 --version
```

**Expected**: version ≥ 3.10.

**Troubleshooting**:
- `command not found` → ask the user to install Python 3.10+ (https://www.python.org/downloads/) and retry. Do not attempt to auto-install Python.
- version < 3.10 → stop; tell the user to upgrade Python.

Record the working command name (`python` or `python3`) as `<PY>` for use below.

---

## 1. Choose an install location and copy the source

Office Dump is a Claude Code skill, so it must live under a path Claude Code searches.

### 1.1 Compute the destination

Target: `<DEST> = <HOME>/.claude/skills/office-dump`. Resolution of `<HOME>` per OS:

| OS | `<HOME>` |
|---|---|
| Windows | `%USERPROFILE%` or `$env:USERPROFILE` (PowerShell) |
| macOS / Linux | `$HOME` or `~` |

Cross-platform robust form — use Python to print the absolute path:

```
<PY> -c "from pathlib import Path; print(Path.home() / '.claude' / 'skills' / 'office-dump')"
```

Record the printed path as `<DEST>`.

### 1.2 Check whether the destination already exists

```
<PY> -c "from pathlib import Path; p=Path(r'<DEST>'); print('EXISTS' if p.exists() else 'MISSING')"
```

- `MISSING` → continue with 1.3.
- `EXISTS` → ask the user whether to overwrite. If they decline, abort. If they accept, remove `<DEST>` first (PowerShell `Remove-Item -Recurse -Force`; bash `rm -rf`) and then continue.

### 1.3 Copy the source to the destination

Cross-platform copy via Python:

```
<PY> -c "import shutil, pathlib; shutil.copytree(r'<SRC>', r'<DEST>', ignore=shutil.ignore_patterns('venv','__pycache__','*.pyc','tests/out','dump_output'))"
```

**Expected**: no output; exit code 0.

**Troubleshooting**:
- `PermissionError` → check that the parent of `<DEST>` is writable; you may need to `mkdir -p ~/.claude/skills/` (or PowerShell `New-Item -ItemType Directory -Force`) first.
- `FileExistsError` → go back to 1.2.

### 1.4 Verify

```
<PY> -c "from pathlib import Path; p=Path(r'<DEST>'); print('OK' if (p/'SKILL.md').exists() and (p/'dump.py').exists() and (p/'lib').is_dir() else 'BROKEN')"
```

Expected: `OK`.

From here, treat `<DEST>` as the working directory.

---

## 2. Create the in-project venv

In `<DEST>`:

```
<PY> -m venv venv
```

**Expected**: no output; exit code 0; the directory `<DEST>/venv/` now exists.

**Verify**:

```
<PY> -c "from pathlib import Path; import platform; p=Path(r'<DEST>')/'venv'/('Scripts' if platform.system()=='Windows' else 'bin')/('python.exe' if platform.system()=='Windows' else 'python'); print('OK' if p.exists() else 'MISSING:', p)"
```

Expected: `OK`. Record the venv's python as `<VPY>`:

- Windows: `<DEST>\venv\Scripts\python.exe`
- macOS / Linux: `<DEST>/venv/bin/python`

**Troubleshooting**:
- `ensurepip is not available` → on Debian/Ubuntu run `apt-get install python3-venv`. Tell the user and abort.

---

## 3. Install Python dependencies

```
<VPY> -m pip install --upgrade pip
<VPY> -m pip install -r requirements.txt
```

**Expected**: the last lines contain `Successfully installed ...`; exit code 0.

**Troubleshooting**:
- network timeout → ask the user; consider retrying with a mirror, e.g. `-i https://pypi.tuna.tsinghua.edu.cn/simple`.
- `pywin32` errors on non-Windows → should not happen since the requirement is gated by environment marker. If it does, ensure pip ≥ 20.3.
- `Microsoft Visual C++ 14.0 or greater is required` (rare; only when a wheel is missing) → ask the user to install Build Tools or to retry on a newer Python.

**Verify**:

```
<VPY> -c "import docx, openpyxl, pptx, oletools, matplotlib, fitz; print('OK')"
```

On Windows additionally:

```
<VPY> -c "import win32com.client; print('OK')"
```

Both should print `OK`.

---

## 4. (Optional) PPTX slide rendering backend

Slide rendering requires either PowerPoint COM (Windows only, Office installed locally) or LibreOffice (any OS). **Without either, the skill still exports text / media / metadata; it just won't produce `slides/*.png`.**

### 4.1 Detect what is already available

In `<DEST>`:

```
<VPY> -c "from lib.pptx_renderer import pick_renderer; print(pick_renderer('auto'))"
```

Output:
- `powerpoint` → Windows already has Office; nothing to do.
- `libreoffice` → LibreOffice is present; nothing to do.
- `None` → no backend; decide whether to install one (see 4.2).

### 4.2 Decide whether to install LibreOffice

If 4.1 printed `None`, **ask the user** whether slide rendering is needed:

- "no" → skip 4.3 and continue at step 5.
- "yes" → continue with 4.3.

### 4.3 Install LibreOffice

By OS:

| OS | Command |
|---|---|
| Windows | `winget install TheDocumentFoundation.LibreOffice`, or ask the user to install from https://www.libreoffice.org/download/ |
| macOS | `brew install --cask libreoffice` (requires Homebrew) |
| Debian/Ubuntu | `sudo apt-get install -y libreoffice` |
| Fedora/RHEL | `sudo dnf install -y libreoffice` |
| Arch | `sudo pacman -S libreoffice-fresh` |

**Caveat**: the package is ~300 MB and the install may take minutes. Commands that need `sudo` must be confirmed by the user first; do not call `sudo` unauthorized.

After installation re-run 4.1 to confirm the backend now resolves to `libreoffice`.

---

## 5. End-to-end smoke test

In `<DEST>`:

```
<VPY> tests/create_samples.py
<VPY> dump.py tests/samples/sample.docx -o tests/out/docx_out
<VPY> dump.py tests/samples/sample.xlsx -o tests/out/xlsx_out
<VPY> dump.py tests/samples/sample.pptx -o tests/out/pptx_out
```

**Expected**: each command prints `Done: ...` with no error.

**Verify outputs**:

```
<VPY> -c "
import json
from pathlib import Path
root = Path(r'<DEST>') / 'tests' / 'out'
for sub in ('docx_out', 'xlsx_out', 'pptx_out'):
    m = json.load(open(root / sub / 'sample' / 'manifest.json', encoding='utf-8'))
    print(sub, 'warnings:', len(m['warnings']), '| keys:', sorted(m.keys()))
"
```

Expected: each line shows `warnings: 0` (PPTX may have 1 warning when no rendering backend is present — acceptable).

**Troubleshooting**:
- module import errors → return to step 3 and check dependencies.
- path-related errors → make sure you are in `<DEST>` and the path contains no characters that break Python import.
- PPTX rendering fails while other fields look fine → backend issue; return to step 4.

---

## 6. Register with Claude Code

If `<DEST>` is `~/.claude/skills/office-dump`, the next Claude Code session should auto-discover this skill under the name `office-dump`.

**How to verify**: ask the user, in a fresh Claude Code session, to run `/help` or to say "export this docx" and see whether the skill triggers.

If it is not discovered:
- confirm the directory name is exactly `office-dump` (matching the `name` field in SKILL.md's frontmatter).
- confirm `SKILL.md` exists and the frontmatter is valid YAML.
- ask the user to restart Claude Code.

---

## 7. Cleanup (optional)

After install, you may delete the smoke-test artifacts under `<DEST>/tests/out/`:

```
<PY> -c "import shutil, pathlib; p=pathlib.Path(r'<DEST>')/'tests'/'out'; shutil.rmtree(p, ignore_errors=True)"
```

---

## Uninstall

```
<PY> -c "import shutil, pathlib; shutil.rmtree(pathlib.Path.home()/'.claude'/'skills'/'office-dump', ignore_errors=True)"
```

---

## Final report

After all mandatory steps, report back to the user:

1. install path `<DEST>`
2. Python version and venv path
3. PPTX rendering backend: `powerpoint` / `libreoffice` / `none` (with a short note)
4. whether the three end-to-end tests passed
5. summary of any warnings seen
