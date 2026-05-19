"""Parse DrawingML chart XML; export series data to CSV and render PNG via matplotlib.

Supported chart types: bar / line / pie / scatter / area / doughnut. Others emit a warning.
"""
from __future__ import annotations

import csv
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

from .ooxml_pack import OOXMLPackage


C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
CHART_TYPES = {
    f"{C}barChart": "bar",
    f"{C}bar3DChart": "bar",
    f"{C}lineChart": "line",
    f"{C}line3DChart": "line",
    f"{C}pieChart": "pie",
    f"{C}pie3DChart": "pie",
    f"{C}doughnutChart": "doughnut",
    f"{C}scatterChart": "scatter",
    f"{C}areaChart": "area",
    f"{C}area3DChart": "area",
}


def _extract_pts(ref_elem: ET.Element | None, resolver=None) -> list[str]:
    """Pull idx-sorted strings out of c:strRef/c:numRef -> c:*Cache/c:pt[@idx]/c:v.

    If cache is missing and a resolver callback is provided, fall back to the
    c:f formula reference. resolver: (formula_str) -> list[str].
    """
    if ref_elem is None:
        return []
    pts_by_idx: dict[int, str] = {}
    for pt in ref_elem.iter(f"{C}pt"):
        try:
            idx = int(pt.attrib.get("idx", "0"))
        except ValueError:
            continue
        v = pt.find(f"{C}v")
        pts_by_idx[idx] = (v.text or "") if v is not None else ""
    if pts_by_idx:
        max_idx = max(pts_by_idx)
        return [pts_by_idx.get(i, "") for i in range(max_idx + 1)]
    # fallback: resolve c:f reference
    if resolver is not None:
        f_elem = ref_elem.find(f".//{C}f")
        if f_elem is not None and f_elem.text:
            try:
                return resolver(f_elem.text)
            except Exception:
                return []
    return []


def _parse_series(ser: ET.Element, resolver=None) -> dict[str, Any]:
    """Extract a series: name / categories / values (or xVal / yVal)."""
    name = ""
    tx = ser.find(f"{C}tx")
    if tx is not None:
        for v in tx.iter(f"{C}v"):
            if v.text:
                name = v.text
                break
        if not name and resolver is not None:
            f_elem = tx.find(f".//{C}f")
            if f_elem is not None and f_elem.text:
                try:
                    resolved = resolver(f_elem.text)
                    if resolved:
                        name = resolved[0]
                except Exception:
                    pass

    cats: list[str] = []
    vals: list[float | None] = []
    x_vals: list[float | None] = []
    y_vals: list[float | None] = []

    for child in ser:
        tag = child.tag
        if tag == f"{C}cat":
            cats = _extract_pts(child, resolver)
        elif tag == f"{C}val":
            for s in _extract_pts(child, resolver):
                vals.append(_to_float(s))
        elif tag == f"{C}xVal":
            for s in _extract_pts(child, resolver):
                x_vals.append(_to_float(s))
        elif tag == f"{C}yVal":
            for s in _extract_pts(child, resolver):
                y_vals.append(_to_float(s))

    return {
        "name": name,
        "categories": cats,
        "values": vals,
        "x_values": x_vals,
        "y_values": y_vals,
    }


def _to_float(s: str) -> float | None:
    if s == "" or s is None:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _title_of(root: ET.Element) -> str:
    title = root.find(f".//{C}title")
    if title is None:
        return ""
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    # Prefer DrawingML text (a:t); fall back to c:v.
    parts = [t.text or "" for t in title.iter(f"{A}t")]
    text = "".join(p for p in parts if p).strip()
    if text:
        return text
    parts = [v.text or "" for v in title.iter(f"{C}v")]
    return "".join(p for p in parts if p).strip()


def _render(chart_type: str, title: str, series: list[dict[str, Any]], out_path: Path) -> str | None:
    """Render to PNG. Returns None on success, otherwise an error string."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception as e:
        return f"matplotlib unavailable: {e}"

    fig, ax = plt.subplots(figsize=(8, 5))
    try:
        if chart_type in ("bar",):
            cats = series[0]["categories"] if series else []
            x = list(range(len(cats)))
            width = 0.8 / max(len(series), 1)
            for i, ser in enumerate(series):
                vals = [v if v is not None else 0 for v in ser["values"]]
                ax.bar([xi + i * width for xi in x], vals, width=width, label=ser["name"] or f"Series{i + 1}")
            ax.set_xticks([xi + width * (len(series) - 1) / 2 for xi in x])
            ax.set_xticklabels(cats, rotation=30, ha="right")
            if any(s["name"] for s in series):
                ax.legend()
        elif chart_type in ("line", "area"):
            cats = series[0]["categories"] if series else []
            for ser in series:
                vals = ser["values"]
                ax.plot(cats or list(range(len(vals))), vals, marker="o", label=ser["name"] or "")
            if any(s["name"] for s in series):
                ax.legend()
            for label in ax.get_xticklabels():
                label.set_rotation(30)
                label.set_ha("right")
        elif chart_type in ("pie", "doughnut"):
            if not series:
                return "no series"
            s = series[0]
            vals = [v if v is not None else 0 for v in s["values"]]
            ax.pie(vals, labels=s["categories"], autopct="%1.1f%%", startangle=90)
            ax.axis("equal")
        elif chart_type == "scatter":
            for ser in series:
                xs = ser["x_values"]
                ys = ser["y_values"]
                ax.scatter(xs, ys, label=ser["name"] or "")
            if any(s["name"] for s in series):
                ax.legend()
        else:
            plt.close(fig)
            return f"unsupported chart type: {chart_type}"

        if title:
            ax.set_title(title)
        fig.tight_layout()
        fig.savefig(out_path, dpi=120)
    finally:
        plt.close(fig)
    return None


def _write_csv(series: list[dict[str, Any]], chart_type: str, csv_path: Path) -> None:
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if chart_type == "scatter":
            # Two columns per series: x and y.
            header = []
            for ser in series:
                name = ser["name"] or "Series"
                header += [f"{name}_x", f"{name}_y"]
            w.writerow(header)
            max_len = max((len(s["x_values"]) for s in series), default=0)
            for i in range(max_len):
                row = []
                for ser in series:
                    xs, ys = ser["x_values"], ser["y_values"]
                    row.append(xs[i] if i < len(xs) else "")
                    row.append(ys[i] if i < len(ys) else "")
                w.writerow(row)
        else:
            cats = series[0]["categories"] if series else []
            w.writerow(["Category"] + [(s["name"] or f"Series{i + 1}") for i, s in enumerate(series)])
            max_len = max((len(s["values"]) for s in series), default=len(cats))
            for i in range(max(max_len, len(cats))):
                row = [cats[i] if i < len(cats) else ""]
                for ser in series:
                    vals = ser["values"]
                    row.append(vals[i] if i < len(vals) else "")
                w.writerow(row)


def extract_charts(pkg: OOXMLPackage, out_dir: Path, doctype: str, resolver=None) -> tuple[list[dict[str, Any]], list[str]]:
    items: list[dict[str, Any]] = []
    warnings: list[str] = []
    pattern = re.compile(r"^(word|xl|ppt)/charts/chart\d+\.xml$")
    chart_names = [n for n in pkg.names() if pattern.match(n)]
    if not chart_names:
        return items, warnings
    out_dir.mkdir(parents=True, exist_ok=True)

    for name in sorted(chart_names):
        chart_id = Path(name).stem  # chart1, chart2 ...
        try:
            root = ET.fromstring(pkg.read(name))
        except Exception as e:
            warnings.append(f"chart {name} parse failed: {e}")
            continue
        # Identify type by the first *Chart element.
        chart_type = None
        chart_elem = None
        for child in root.iter():
            if child.tag in CHART_TYPES:
                chart_type = CHART_TYPES[child.tag]
                chart_elem = child
                break
        if chart_elem is None:
            warnings.append(f"chart {name} type not recognized")
            continue
        series = [_parse_series(ser, resolver) for ser in chart_elem.findall(f"{C}ser")]
        title = _title_of(root)

        csv_path = out_dir / f"{chart_id}.csv"
        png_path = out_dir / f"{chart_id}.png"
        try:
            _write_csv(series, chart_type, csv_path)
        except Exception as e:
            warnings.append(f"chart {name} CSV write failed: {e}")
            csv_path = None  # type: ignore

        png_err = _render(chart_type, title, series, png_path)
        if png_err:
            warnings.append(f"chart {name} render failed: {png_err}")
            png_path = None  # type: ignore

        items.append({
            "src": name,
            "type": chart_type,
            "title": title,
            "series": [s["name"] for s in series],
            "csv": csv_path.relative_to(out_dir.parent).as_posix() if csv_path else None,
            "png": png_path.relative_to(out_dir.parent).as_posix() if png_path else None,
        })
    return items, warnings
