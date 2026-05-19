"""Extract VBA macros: write each module as a .bas file with an olevba report."""
from __future__ import annotations

from pathlib import Path
from typing import Any


def extract_macros(input_path: str | Path, out_dir: Path) -> dict[str, Any] | None:
    """Return a dict when macros are found; None otherwise."""
    try:
        from oletools.olevba import VBA_Parser, FileOpenError  # type: ignore
    except Exception as e:
        return {"error": f"oletools unavailable: {e}"}

    try:
        vba = VBA_Parser(str(input_path))
    except FileOpenError:
        return None
    except Exception as e:
        return {"error": f"olevba open failed: {e}"}

    try:
        if not vba.detect_vba_macros():
            return None
        out_dir.mkdir(parents=True, exist_ok=True)
        modules: list[dict[str, Any]] = []
        for (filename, stream_path, vba_filename, vba_code) in vba.extract_macros():
            # Use the module name as the file name.
            safe = (vba_filename or "module").replace("/", "_").replace("\\", "_")
            if not safe.lower().endswith(".bas"):
                safe = f"{safe}.bas"
            dest = out_dir / safe
            i = 1
            while dest.exists():
                dest = out_dir / f"{safe.rsplit('.', 1)[0]}_{i}.bas"
                i += 1
            dest.write_text(vba_code or "", encoding="utf-8", errors="replace")
            modules.append({
                "module": vba_filename,
                "stream": stream_path,
                "dest": dest.relative_to(out_dir.parent).as_posix(),
                "size": len(vba_code or ""),
            })

        # Static analysis report (keywords / suspicious APIs).
        try:
            analysis = vba.analyze_macros()
            report = [{"type": kind, "keyword": kw, "description": desc} for kind, kw, desc in analysis]
        except Exception as e:
            report = [{"error": f"analyze_macros failed: {e}"}]

        return {
            "modules": modules,
            "analysis": report,
        }
    finally:
        try:
            vba.close()
        except Exception:
            pass
