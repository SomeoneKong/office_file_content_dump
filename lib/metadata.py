"""Read docProps/core.xml, app.xml, custom.xml."""
from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Any

from .ooxml_pack import OOXMLPackage


NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
}


def _local(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _parse_core(xml_bytes: bytes) -> dict[str, Any]:
    root = ET.fromstring(xml_bytes)
    out: dict[str, Any] = {}
    for child in root:
        out[_local(child.tag)] = (child.text or "").strip()
    return out


def _parse_app(xml_bytes: bytes) -> dict[str, Any]:
    root = ET.fromstring(xml_bytes)
    out: dict[str, Any] = {}
    for child in root:
        key = _local(child.tag)
        # Take text directly for simple leaves; skip complex structures
        # like HeadingPairs / TitlesOfParts.
        if len(child) == 0:
            out[key] = (child.text or "").strip()
    return out


def _parse_custom(xml_bytes: bytes) -> list[dict[str, Any]]:
    root = ET.fromstring(xml_bytes)
    items = []
    for prop in root:
        name = prop.attrib.get("name", "")
        # Child is a vt:* typed element.
        value = None
        vtype = None
        for v in prop:
            vtype = _local(v.tag)
            value = (v.text or "").strip()
        items.append({"name": name, "type": vtype, "value": value})
    return items


def extract_metadata(pkg: OOXMLPackage) -> dict[str, Any]:
    meta: dict[str, Any] = {"core": {}, "app": {}, "custom": []}
    if pkg.has("docProps/core.xml"):
        try:
            meta["core"] = _parse_core(pkg.read("docProps/core.xml"))
        except Exception as e:
            meta["core_error"] = str(e)
    if pkg.has("docProps/app.xml"):
        try:
            meta["app"] = _parse_app(pkg.read("docProps/app.xml"))
        except Exception as e:
            meta["app_error"] = str(e)
    if pkg.has("docProps/custom.xml"):
        try:
            meta["custom"] = _parse_custom(pkg.read("docProps/custom.xml"))
        except Exception as e:
            meta["custom_error"] = str(e)
    return meta
