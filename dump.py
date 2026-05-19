"""Office OOXML content export CLI.

Usage:
    python dump.py <input.docx|xlsx|pptx> [-o OUTDIR] [options]
"""
from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path

from lib import docx_dumper, xlsx_dumper, pptx_dumper
from lib.macros import extract_macros


DOCTYPE_MAP = {
    ".docx": "docx", ".docm": "docx",
    ".xlsx": "xlsx", ".xlsm": "xlsx",
    ".pptx": "pptx", ".pptm": "pptx",
}


def _check_encrypted(path: Path) -> bool:
    """An encrypted OOXML is an OLE container (not a ZIP). Returns True if encrypted."""
    if not zipfile.is_zipfile(path):
        return True
    return False


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Dump Office OOXML files to a directory.")
    ap.add_argument("input", help="input file (.docx/.xlsx/.pptx and the .docm/.xlsm/.pptm variants)")
    ap.add_argument("-o", "--output", default="./dump_output", help="output root directory")
    ap.add_argument("--slide-renderer", choices=["auto", "powerpoint", "libreoffice", "none"],
                    default="auto", help="PPTX slide rendering backend")
    ap.add_argument("--dpi", type=int, default=150, help="slide rendering DPI")
    ap.add_argument("--no-charts", action="store_true", help="skip chart extraction")
    ap.add_argument("--no-macros", action="store_true", help="skip VBA macro extraction")
    args = ap.parse_args(argv)

    input_path = Path(args.input).resolve()
    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 2

    ext = input_path.suffix.lower()
    doctype = DOCTYPE_MAP.get(ext)
    if not doctype:
        print(f"Unsupported extension: {ext} (supported: {', '.join(DOCTYPE_MAP)})", file=sys.stderr)
        return 2

    if _check_encrypted(input_path):
        print(f"File is encrypted or not a ZIP container; cannot process: {input_path}", file=sys.stderr)
        return 3

    out_root = Path(args.output).resolve()
    out_dir = out_root / input_path.stem
    if out_dir.exists():
        import shutil
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    options = {
        "slide_renderer": args.slide_renderer,
        "dpi": args.dpi,
        "no_charts": args.no_charts,
        "no_macros": args.no_macros,
    }

    dumper = {"docx": docx_dumper, "xlsx": xlsx_dumper, "pptx": pptx_dumper}[doctype]
    manifest = dumper.dump(input_path, out_dir, options)

    # Macros (independent flow, only for *m variants).
    if not args.no_macros and ext in (".docm", ".xlsm", ".pptm"):
        macro_info = extract_macros(input_path, out_dir / "macros")
        if macro_info:
            manifest["macros"] = macro_info

    manifest_path = out_dir / "manifest.json"
    manifest_path.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"Done: {out_dir}")
    if manifest.get("warnings"):
        print(f"  {len(manifest['warnings'])} warning(s); see {manifest_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
