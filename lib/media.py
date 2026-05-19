"""Extract media files (images, audio, video) from an OOXML package."""
from __future__ import annotations

from pathlib import Path
from typing import Any

from .ooxml_pack import OOXMLPackage


# Where media is stored under each document type. `embeddings/` is handled
# separately by embedded.py — here we only look at `media/`.
MEDIA_PREFIXES = {
    "docx": ["word/media/", "word/embeddings/"],
    "xlsx": ["xl/media/"],
    "pptx": ["ppt/media/"],
}


def extract_media(pkg: OOXMLPackage, out_dir: Path, doctype: str) -> list[dict[str, Any]]:
    """Copy every file under */media/ to out_dir, keeping its base name. Returns a manifest list."""
    items: list[dict[str, Any]] = []
    prefixes = [p for p in MEDIA_PREFIXES.get(doctype, []) if "/media/" in p]
    if not prefixes:
        return items
    out_dir.mkdir(parents=True, exist_ok=True)
    for prefix in prefixes:
        for name in pkg.iter_prefix(prefix):
            if name.endswith("/"):
                continue
            basename = name.split("/")[-1]
            dest = out_dir / basename
            # de-duplicate by suffixing _N
            i = 1
            while dest.exists():
                stem, dot, ext = basename.rpartition(".")
                dest = out_dir / (f"{stem}_{i}.{ext}" if dot else f"{basename}_{i}")
                i += 1
            data = pkg.read(name)
            dest.write_bytes(data)
            items.append({
                "src": name,
                "dest": dest.relative_to(out_dir.parent).as_posix(),
                "size": len(data),
            })
    return items
