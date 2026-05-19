"""PPTX slide rendering: PowerPoint COM (Windows) and LibreOffice headless backends."""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any


# Candidate LibreOffice paths across platforms.
SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "/snap/bin/libreoffice",
]


def _find_soffice() -> str | None:
    for cmd in ("soffice", "libreoffice"):
        p = shutil.which(cmd)
        if p:
            return p
    for p in SOFFICE_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


def _powerpoint_available() -> bool:
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
    except Exception:
        return False
    try:
        pythoncom.CoInitialize()
        app = win32com.client.Dispatch("PowerPoint.Application")
        # Don't Quit; Dispatch may attach to an existing instance.
        return app is not None
    except Exception:
        return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def pick_renderer(prefer: str = "auto") -> str | None:
    """Return 'powerpoint' / 'libreoffice' / None."""
    if prefer == "none":
        return None
    if prefer == "powerpoint":
        return "powerpoint" if _powerpoint_available() else None
    if prefer == "libreoffice":
        return "libreoffice" if _find_soffice() else None
    # auto
    if _powerpoint_available():
        return "powerpoint"
    if _find_soffice():
        return "libreoffice"
    return None


def _slide_pixel_size(input_path: Path, dpi: int) -> tuple[int, int]:
    """Read slide size (EMU) from pptx, convert to pixels. Fall back to 16:9 default."""
    try:
        from pptx import Presentation  # type: ignore
        pres = Presentation(str(input_path))
        w_px = int(pres.slide_width / 914400 * dpi)
        h_px = int(pres.slide_height / 914400 * dpi)
        return w_px, h_px
    except Exception:
        # 13.333 x 7.5 inches at the requested DPI.
        return int(13.333 * dpi), int(7.5 * dpi)


def render_powerpoint_com(input_path: Path, slides_dir: Path, dpi: int) -> tuple[list[str], list[str]]:
    """One PNG per slide. Returns (relative paths, warnings)."""
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    warnings: list[str] = []
    rendered: list[str] = []
    slides_dir.mkdir(parents=True, exist_ok=True)
    width_px, height_px = _slide_pixel_size(input_path, dpi)

    pythoncom.CoInitialize()
    app = None
    pres = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        # PowerPoint refuses Visible=False; minimize the window instead.
        try:
            app.WindowState = 2  # ppWindowMinimized
        except Exception:
            pass
        # Absolute path required.
        pres = app.Presentations.Open(
            str(input_path.resolve()),
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )
        for i, slide in enumerate(pres.Slides, start=1):
            out_file = slides_dir / f"slide_{i:02d}.png"
            try:
                slide.Export(str(out_file.resolve()), "PNG", width_px, height_px)
                rendered.append(out_file.relative_to(slides_dir.parent).as_posix())
            except Exception as e:
                warnings.append(f"slide {i} export failed: {e}")
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return rendered, warnings


def render_libreoffice(input_path: Path, slides_dir: Path, dpi: int, pdf_dest: Path | None) -> tuple[list[str], list[str]]:
    """Convert to PDF via soffice headless, then split pages to PNG via PyMuPDF."""
    warnings: list[str] = []
    rendered: list[str] = []
    slides_dir.mkdir(parents=True, exist_ok=True)
    soffice = _find_soffice()
    if not soffice:
        return rendered, ["soffice not found"]

    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        cmd = [soffice, "--headless", "--norestore", "--nofirststartwizard",
               "--convert-to", "pdf", "--outdir", str(td_path), str(input_path)]
        try:
            r = subprocess.run(cmd, capture_output=True, timeout=300)
            if r.returncode != 0:
                warnings.append(f"soffice PDF conversion failed: rc={r.returncode}, stderr={r.stderr.decode('utf-8', errors='replace')[:500]}")
                return rendered, warnings
        except subprocess.TimeoutExpired:
            warnings.append("soffice PDF conversion timed out")
            return rendered, warnings
        except Exception as e:
            warnings.append(f"soffice invocation error: {e}")
            return rendered, warnings

        pdfs = list(td_path.glob("*.pdf"))
        if not pdfs:
            warnings.append("soffice did not produce a PDF")
            return rendered, warnings
        pdf_path = pdfs[0]

        # Copy the PDF to the requested destination.
        if pdf_dest is not None:
            pdf_dest.parent.mkdir(parents=True, exist_ok=True)
            pdf_dest.write_bytes(pdf_path.read_bytes())

        try:
            import fitz  # type: ignore  # PyMuPDF
        except Exception as e:
            warnings.append(f"PyMuPDF unavailable: {e}")
            return rendered, warnings

        try:
            doc = fitz.open(str(pdf_path))
        except Exception as e:
            warnings.append(f"PyMuPDF failed to open PDF: {e}")
            return rendered, warnings

        try:
            scale = dpi / 72.0
            mat = fitz.Matrix(scale, scale)
            for i, page in enumerate(doc, start=1):
                try:
                    pix = page.get_pixmap(matrix=mat, alpha=False)
                    out_file = slides_dir / f"slide_{i:02d}.png"
                    pix.save(str(out_file))
                    rendered.append(out_file.relative_to(slides_dir.parent).as_posix())
                except Exception as e:
                    warnings.append(f"slide {i} render failed: {e}")
        finally:
            doc.close()

    return rendered, warnings


def render_slides(input_path: Path, out_dir: Path, dpi: int, prefer: str = "auto") -> dict[str, Any]:
    """Returns {renderer, slides[], pdf, warnings[]}."""
    backend = pick_renderer(prefer)
    if backend is None:
        return {"renderer": None, "slides": [], "pdf": None,
                "warnings": ["no usable PPTX renderer (PowerPoint / LibreOffice not available)"]}

    slides_dir = out_dir / "slides"
    if backend == "powerpoint":
        slides, warns = render_powerpoint_com(input_path, slides_dir, dpi)
        return {"renderer": "powerpoint", "slides": slides, "pdf": None, "warnings": warns}
    else:
        pdf_dest = out_dir / "full.pdf"
        slides, warns = render_libreoffice(input_path, slides_dir, dpi, pdf_dest)
        return {
            "renderer": "libreoffice",
            "slides": slides,
            "pdf": pdf_dest.relative_to(out_dir.parent).as_posix() if pdf_dest.exists() else None,
            "warnings": warns,
        }
