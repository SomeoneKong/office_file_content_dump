"""Word (.docx / .docm) content export as Markdown, plus media / embeddings / comments / headers-footers / footnotes."""
from __future__ import annotations

import json
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

from docx import Document  # type: ignore
from docx.oxml.ns import qn  # type: ignore
from docx.table import Table  # type: ignore
from docx.text.paragraph import Paragraph  # type: ignore

from .ooxml_pack import OOXMLPackage
from .media import extract_media
from .embedded import extract_embedded
from .metadata import extract_metadata
from .charts import extract_charts


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _heading_level(para: Paragraph) -> int | None:
    style = (para.style.name or "") if para.style else ""
    if style.startswith("Heading "):
        try:
            return int(style.split(" ", 1)[1])
        except ValueError:
            return None
    if style == "Title":
        return 1
    return None


def _list_prefix(para: Paragraph) -> str | None:
    """Return '-' or '1.' or None."""
    pPr = para._p.find(qn("w:pPr"))
    if pPr is not None and pPr.find(qn("w:numPr")) is not None:
        return "-"
    style = (para.style.name or "") if para.style else ""
    if style.startswith("List Number"):
        return "1."
    if style.startswith("List Bullet") or style.startswith("List Paragraph"):
        return "-"
    return None


def _para_md(para: Paragraph, rels: dict[str, str]) -> str:
    """Render a paragraph as one Markdown line (no trailing newline)."""
    parts: list[str] = []
    # Walk paragraph children: handle hyperlinks specially.
    for child in para._p.iter():
        tag = child.tag
        if tag == qn("w:hyperlink"):
            text = "".join(t.text or "" for t in child.iter(qn("w:t")))
            rid = child.get(qn("r:id"))
            url = rels.get(rid, "") if rid else ""
            if text:
                if url:
                    parts.append(f"[{text}]({url})")
                else:
                    parts.append(text)
            # Clear inner t's so the next pass does not double-count them.
            for t in child.iter(qn("w:t")):
                t.text = None  # type: ignore
    # Pick up any remaining (non-hyperlink) text runs.
    for t in para._p.iter(qn("w:t")):
        if t.text:
            parts.append(t.text)

    text = "".join(parts).strip()
    if not text:
        return ""

    level = _heading_level(para)
    if level:
        level = max(1, min(6, level))
        return f"{'#' * level} {text}"
    prefix = _list_prefix(para)
    if prefix:
        return f"{prefix} {text}"
    return text


def _table_md(tbl: Table) -> str:
    rows: list[list[str]] = []
    for row in tbl.rows:
        cells = []
        for cell in row.cells:
            # Join multi-paragraph cells with <br>.
            t = " <br> ".join((p.text or "").strip() for p in cell.paragraphs if (p.text or "").strip())
            t = t.replace("|", "\\|")
            cells.append(t)
        rows.append(cells)
    if not rows:
        return ""
    col_count = max(len(r) for r in rows)
    rows = [r + [""] * (col_count - len(r)) for r in rows]
    lines = ["| " + " | ".join(rows[0]) + " |",
             "| " + " | ".join(["---"] * col_count) + " |"]
    for r in rows[1:]:
        lines.append("| " + " | ".join(r) + " |")
    return "\n".join(lines)


def _read_rels(pkg: OOXMLPackage) -> dict[str, str]:
    rels: dict[str, str] = {}
    rels_path = "word/_rels/document.xml.rels"
    if not pkg.has(rels_path):
        return rels
    try:
        root = ET.fromstring(pkg.read(rels_path))
    except Exception:
        return rels
    for r in root:
        rid = r.attrib.get("Id")
        target = r.attrib.get("Target", "")
        if rid:
            rels[rid] = target
    return rels


def _extract_xml_text(pkg: OOXMLPackage, name: str) -> str:
    """Concatenate all w:t text in a w:* XML part, paragraph by paragraph."""
    if not pkg.has(name):
        return ""
    try:
        root = ET.fromstring(pkg.read(name))
    except Exception:
        return ""
    lines: list[str] = []
    for p in root.iter(f"{W}p"):
        parts = [t.text or "" for t in p.iter(f"{W}t")]
        text = "".join(parts).strip()
        if text:
            lines.append(text)
    return "\n\n".join(lines)


def _extract_comments(pkg: OOXMLPackage) -> list[dict[str, Any]]:
    name = "word/comments.xml"
    if not pkg.has(name):
        return []
    try:
        root = ET.fromstring(pkg.read(name))
    except Exception:
        return []
    items = []
    for c in root.iter(f"{W}comment"):
        text_parts = [t.text or "" for t in c.iter(f"{W}t")]
        items.append({
            "id": c.attrib.get(f"{W}id"),
            "author": c.attrib.get(f"{W}author", ""),
            "date": c.attrib.get(f"{W}date", ""),
            "text": "".join(text_parts).strip(),
        })
    return items


def _extract_footnotes(pkg: OOXMLPackage) -> list[dict[str, Any]]:
    name = "word/footnotes.xml"
    if not pkg.has(name):
        return []
    try:
        root = ET.fromstring(pkg.read(name))
    except Exception:
        return []
    items = []
    for fn in root.iter(f"{W}footnote"):
        fn_type = fn.attrib.get(f"{W}type", "")
        if fn_type in ("separator", "continuationSeparator"):
            continue
        text_parts = [t.text or "" for t in fn.iter(f"{W}t")]
        text = "".join(text_parts).strip()
        if not text:
            continue
        items.append({"id": fn.attrib.get(f"{W}id"), "text": text})
    return items


def _headers_footers(pkg: OOXMLPackage) -> dict[str, list[str]]:
    headers: list[str] = []
    footers: list[str] = []
    for name in sorted(pkg.names()):
        if name.startswith("word/header") and name.endswith(".xml"):
            txt = _extract_xml_text(pkg, name)
            if txt:
                headers.append(txt)
        elif name.startswith("word/footer") and name.endswith(".xml"):
            txt = _extract_xml_text(pkg, name)
            if txt:
                footers.append(txt)
    return {"headers": headers, "footers": footers}


def dump(input_path: Path, out_dir: Path, options: dict[str, Any]) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest: dict[str, Any] = {
        "doctype": "docx",
        "source": input_path.name,
        "warnings": [],
    }

    # Structured body via python-docx.
    try:
        doc = Document(str(input_path))
    except Exception as e:
        manifest["warnings"].append(f"python-docx open failed: {e}")
        doc = None

    with OOXMLPackage(input_path) as pkg:
        manifest["metadata"] = extract_metadata(pkg)

        media_items = extract_media(pkg, out_dir / "media", "docx")
        manifest["media"] = media_items

        embedded_items = extract_embedded(pkg, out_dir / "embedded", "docx")
        manifest["embedded"] = embedded_items

        if not options.get("no_charts"):
            chart_items, chart_warns = extract_charts(pkg, out_dir / "charts", "docx")
            manifest["charts"] = chart_items
            manifest["warnings"].extend(chart_warns)

        rels = _read_rels(pkg)
        comments = _extract_comments(pkg)
        footnotes = _extract_footnotes(pkg)
        hf = _headers_footers(pkg)

    # Markdown body.
    md_lines: list[str] = [f"# {input_path.stem}", ""]
    if doc is not None:
        # python-docx 1.0+ returns paragraphs and tables in source order.
        try:
            iterator = doc.iter_inner_content()
        except AttributeError:
            iterator = doc.paragraphs  # fallback; order may differ slightly
        for item in iterator:
            if isinstance(item, Paragraph):
                line = _para_md(item, rels)
                if line:
                    md_lines.append(line)
                    md_lines.append("")
            elif isinstance(item, Table):
                md_lines.append(_table_md(item))
                md_lines.append("")
    else:
        md_lines.append("_(python-docx could not open the file; body not extracted)_")

    if hf["headers"]:
        md_lines.append("\n## Headers")
        for h in hf["headers"]:
            md_lines.append(h)
            md_lines.append("")
    if hf["footers"]:
        md_lines.append("\n## Footers")
        for f in hf["footers"]:
            md_lines.append(f)
            md_lines.append("")
    if footnotes:
        md_lines.append("\n## Footnotes")
        for fn in footnotes:
            md_lines.append(f"- [{fn['id']}] {fn['text']}")
        md_lines.append("")

    (out_dir / "content.md").write_text("\n".join(md_lines), encoding="utf-8")

    if comments:
        (out_dir / "comments.json").write_text(
            json.dumps(comments, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        manifest["comments"] = {"count": len(comments), "file": "comments.json"}

    return manifest
