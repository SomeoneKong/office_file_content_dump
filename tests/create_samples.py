"""Generate test docx / xlsx / pptx samples covering text, tables, images, charts, links.

A small amount of non-ASCII text is kept on purpose as a Unicode smoke test.
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw
from docx import Document
from docx.shared import Inches

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, LineChart, PieChart

from pptx import Presentation
from pptx.util import Inches as PInches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE


HERE = Path(__file__).parent
SAMPLES = HERE / "samples"
SAMPLES.mkdir(parents=True, exist_ok=True)


def _make_png(path: Path, text: str, color=(70, 130, 180)) -> None:
    img = Image.new("RGB", (320, 200), color=color)
    d = ImageDraw.Draw(img)
    d.text((20, 80), text, fill=(255, 255, 255))
    img.save(path)


def make_docx() -> Path:
    img_path = SAMPLES / "_tmp.png"
    _make_png(img_path, "Hello DOCX")

    doc = Document()
    doc.core_properties.title = "Sample Word Document"
    doc.core_properties.author = "office-dump test"
    doc.core_properties.subject = "test docx_dumper"

    doc.add_heading("Heading 1", level=1)
    # Mix of ASCII and CJK to verify Unicode handling end-to-end.
    doc.add_paragraph("This is an ordinary paragraph with English and 中文.")
    doc.add_heading("Heading 2", level=2)
    p = doc.add_paragraph("Visit ")
    p.add_run("Anthropic homepage ").bold = True
    doc.add_paragraph("Bullet 1", style="List Bullet")
    doc.add_paragraph("Bullet 2", style="List Bullet")
    doc.add_paragraph("Bullet 3", style="List Bullet")

    doc.add_heading("Table sample", level=2)
    table = doc.add_table(rows=3, cols=3)
    table.style = "Table Grid"
    header = table.rows[0].cells
    header[0].text = "Name"
    header[1].text = "Team"
    header[2].text = "ID"
    table.rows[1].cells[0].text = "Alice"
    table.rows[1].cells[1].text = "R&D"
    table.rows[1].cells[2].text = "1001"
    table.rows[2].cells[0].text = "Bob"
    table.rows[2].cells[1].text = "QA"
    table.rows[2].cells[2].text = "1002"

    doc.add_heading("Image sample", level=2)
    doc.add_picture(str(img_path), width=Inches(3))

    out = SAMPLES / "sample.docx"
    doc.save(str(out))
    img_path.unlink(missing_ok=True)
    return out


def make_xlsx() -> Path:
    wb = Workbook()
    wb.properties.title = "Sample Excel"
    wb.properties.creator = "office-dump test"

    ws = wb.active
    ws.title = "Sales"
    ws["A1"], ws["B1"], ws["C1"] = "Month", "Region", "Amount"
    rows = [
        ("Jan", "North", 100),
        ("Feb", "North", 130),
        ("Mar", "North", 160),
        ("Jan", "South", 80),
        ("Feb", "South", 110),
        ("Mar", "South", 140),
    ]
    for r in rows:
        ws.append(r)
    ws["E1"] = "Total"
    ws["E2"] = "=SUM(C2:C7)"

    # Bar chart.
    chart = BarChart()
    chart.title = "Sales by Month"
    data = Reference(ws, min_col=3, min_row=1, max_row=7, max_col=3)
    cats = Reference(ws, min_col=1, min_row=2, max_row=7)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "G2")

    # Second sheet + line chart.
    ws2 = wb.create_sheet("Trend")
    ws2["A1"], ws2["B1"] = "Day", "Value"
    for i in range(1, 8):
        ws2.append([i, i * i])
    line = LineChart()
    line.title = "Trend"
    line.add_data(Reference(ws2, min_col=2, min_row=1, max_row=8, max_col=2), titles_from_data=True)
    line.set_categories(Reference(ws2, min_col=1, min_row=2, max_row=8))
    ws2.add_chart(line, "D2")

    # Third sheet + pie chart.
    ws3 = wb.create_sheet("Share")
    ws3["A1"], ws3["B1"] = "Group", "Share"
    for k, v in [("A", 30), ("B", 50), ("C", 20)]:
        ws3.append([k, v])
    pie = PieChart()
    pie.title = "Group Share"
    pie.add_data(Reference(ws3, min_col=2, min_row=1, max_row=4, max_col=2), titles_from_data=True)
    pie.set_categories(Reference(ws3, min_col=1, min_row=2, max_row=4))
    ws3.add_chart(pie, "D2")

    out = SAMPLES / "sample.xlsx"
    wb.save(str(out))
    return out


def make_pptx() -> Path:
    img_path = SAMPLES / "_tmp.png"
    _make_png(img_path, "Hello PPTX", color=(220, 90, 60))

    pres = Presentation()
    pres.core_properties.title = "Sample PowerPoint"
    pres.core_properties.author = "office-dump test"

    # Slide 1: title.
    slide1 = pres.slides.add_slide(pres.slide_layouts[0])
    slide1.shapes.title.text = "Sample PPT Title"
    if len(slide1.placeholders) > 1:
        slide1.placeholders[1].text = "Subtitle: testing pptx_dumper"

    # Slide 2: bullets + image.
    slide2 = pres.slides.add_slide(pres.slide_layouts[1])
    slide2.shapes.title.text = "Bullets + Image"
    body = slide2.placeholders[1].text_frame
    body.text = "Bullet 1"
    p2 = body.add_paragraph(); p2.text = "Bullet 2"
    p3 = body.add_paragraph(); p3.text = "Bullet 3"
    slide2.shapes.add_picture(str(img_path), PInches(5.5), PInches(1.5), width=PInches(3))

    # Slide 3: table.
    slide3 = pres.slides.add_slide(pres.slide_layouts[5])
    slide3.shapes.title.text = "Table sample"
    rows, cols = 3, 3
    left, top, width, height = PInches(0.5), PInches(1.5), PInches(9), PInches(3)
    table = slide3.shapes.add_table(rows, cols, left, top, width, height).table
    for c, h in enumerate(["A", "B", "C"]):
        table.cell(0, c).text = h
    for r in range(1, 3):
        for c in range(3):
            table.cell(r, c).text = f"r{r}c{c}"

    # Slide 4: chart.
    slide4 = pres.slides.add_slide(pres.slide_layouts[5])
    slide4.shapes.title.text = "Chart sample"
    chart_data = CategoryChartData()
    chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
    chart_data.add_series("North", (10, 20, 30, 40))
    chart_data.add_series("South", (15, 25, 20, 35))
    slide4.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        PInches(1), PInches(1.5), PInches(8), PInches(5),
        chart_data,
    )

    # Speaker notes.
    slide2.notes_slide.notes_text_frame.text = "Speaker notes for slide 2"

    out = SAMPLES / "sample.pptx"
    pres.save(str(out))
    img_path.unlink(missing_ok=True)
    return out


if __name__ == "__main__":
    print("docx:", make_docx())
    print("xlsx:", make_xlsx())
    print("pptx:", make_pptx())
