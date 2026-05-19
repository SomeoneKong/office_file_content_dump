"""Extract embedded objects: copy files under embeddings/ verbatim; peel one
layer off OLE containers (.bin)."""
from __future__ import annotations

import io
from pathlib import Path
from typing import Any

from .ooxml_pack import OOXMLPackage


EMBED_PREFIXES = {
    "docx": ["word/embeddings/"],
    "xlsx": ["xl/embeddings/"],
    "pptx": ["ppt/embeddings/"],
}


def _try_unpack_ole(data: bytes, out_dir: Path, basename: str) -> list[dict[str, Any]]:
    """For an OLE container (.bin), pull out the inner payload(s).

    Typical streams of interest:
      - 'Package' / 'CONTENTS' / 'Ole10Native' streams that wrap a whole file
      - The container itself may also be a binary Office document
    """
    extracted: list[dict[str, Any]] = []
    try:
        import olefile  # type: ignore
    except Exception:
        return extracted
    if not olefile.isOleFile(io.BytesIO(data)):
        return extracted
    try:
        ole = olefile.OleFileIO(io.BytesIO(data))
    except Exception:
        return extracted
    try:
        streams = ole.listdir()
        # Prefer Package / Ole10Native (an embedded whole file).
        for stream_path in streams:
            joined = "/".join(stream_path)
            lowered = joined.lower()
            if lowered.endswith("package") or "ole10native" in lowered:
                try:
                    stream = ole.openstream(stream_path)
                    raw = stream.read()
                except Exception:
                    continue
                inner_name, inner_bytes = _parse_ole10native(raw) if "ole10native" in lowered else (None, raw)
                fname = inner_name or f"{basename}__{joined.replace('/', '_')}"
                dest = out_dir / fname
                i = 1
                while dest.exists():
                    stem, dot, ext = fname.rpartition(".")
                    dest = out_dir / (f"{stem}_{i}.{ext}" if dot else f"{fname}_{i}")
                    i += 1
                dest.write_bytes(inner_bytes)
                extracted.append({
                    "dest": dest.relative_to(out_dir.parent).as_posix(),
                    "size": len(inner_bytes),
                    "from_ole_stream": joined,
                })
    finally:
        ole.close()
    return extracted


def _parse_ole10native(raw: bytes) -> tuple[str | None, bytes]:
    """Best-effort parse of Ole10Native: skip 4-byte total length + 2-byte flag,
    then read C-string filename, C-string path, 8 unused bytes, C-string temp
    path, 4-byte payload size, then payload bytes. On any failure return raw."""
    try:
        if len(raw) < 8:
            return None, raw
        pos = 4  # skip total_size
        if pos + 2 > len(raw):
            return None, raw
        pos += 2  # flag
        end = raw.find(b"\x00", pos)
        if end < 0:
            return None, raw
        filename = raw[pos:end].decode("latin-1", errors="replace")
        pos = end + 1
        end = raw.find(b"\x00", pos)
        if end < 0:
            return None, raw
        pos = end + 1  # skip path
        pos += 8  # reserved
        end = raw.find(b"\x00", pos)
        if end < 0:
            return None, raw
        pos = end + 1  # skip temp path
        if pos + 4 > len(raw):
            return None, raw
        size = int.from_bytes(raw[pos:pos + 4], "little")
        pos += 4
        if pos + size > len(raw):
            return None, raw
        return filename or None, raw[pos:pos + size]
    except Exception:
        return None, raw


def extract_embedded(pkg: OOXMLPackage, out_dir: Path, doctype: str) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    prefixes = EMBED_PREFIXES.get(doctype, [])
    for prefix in prefixes:
        for name in pkg.iter_prefix(prefix):
            if name.endswith("/"):
                continue
            out_dir.mkdir(parents=True, exist_ok=True)
            basename = name.split("/")[-1]
            dest = out_dir / basename
            i = 1
            while dest.exists():
                stem, dot, ext = basename.rpartition(".")
                dest = out_dir / (f"{stem}_{i}.{ext}" if dot else f"{basename}_{i}")
                i += 1
            data = pkg.read(name)
            dest.write_bytes(data)
            entry: dict[str, Any] = {
                "src": name,
                "dest": dest.relative_to(out_dir.parent).as_posix(),
                "size": len(data),
            }
            # Peel one layer off OLE containers.
            if basename.lower().endswith(".bin"):
                unpacked = _try_unpack_ole(data, out_dir, basename.rsplit(".", 1)[0])
                if unpacked:
                    entry["ole_unpacked"] = unpacked
            items.append(entry)
    return items
