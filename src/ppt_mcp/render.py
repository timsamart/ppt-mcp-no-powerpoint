"""Rendering service (DESIGN.md §8): headless LibreOffice + pypdfium2.

soffice converts the deck to PDF; pypdfium2 rasterizes pages to PNG — both
fully local. Renders are validation evidence, not pixel truth: where
LibreOffice and PowerPoint disagree, PowerPoint is the manual arbiter (§15).
Everything is cached by (file hash, dpi); a missing LibreOffice degrades the
render tools with a clear error while the rest of the server keeps working.
"""

from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from .errors import PptMcpError
from .store import Store

ENV_SOFFICE = "PPT_MCP_SOFFICE"

_WINDOWS_CANDIDATES = (
    r"C:\Program Files\LibreOffice\program\soffice.com",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)

CONVERT_TIMEOUT_S = 120


def find_soffice() -> str | None:
    explicit = os.environ.get(ENV_SOFFICE)
    if explicit and Path(explicit).is_file():
        return explicit
    found = shutil.which("soffice")
    if found:
        return found
    for candidate in _WINDOWS_CANDIDATES:
        if Path(candidate).is_file():
            return candidate
    return None


class RenderService:
    def __init__(self, store: Store):
        self.store = store
        self.soffice = find_soffice()

    @property
    def available(self) -> bool:
        return self.soffice is not None

    def _require_soffice(self) -> str:
        if self.soffice is None:
            raise PptMcpError(
                "Rendering unavailable: LibreOffice not found. Install it "
                "(e.g. 'winget install TheDocumentFoundation.LibreOffice') or set "
                f"the {ENV_SOFFICE} environment variable to soffice's path. "
                "All non-render tools keep working without it."
            )
        return self.soffice

    def _cache_dir(self, pptx_path: Path) -> Path:
        digest = hashlib.sha256(pptx_path.read_bytes()).hexdigest()[:16]
        return self.store.renders_dir / digest

    def to_pdf(self, pptx_path: Path) -> Path:
        """Convert a deck to PDF, cached by content hash."""
        soffice = self._require_soffice()
        cache_dir = self._cache_dir(pptx_path)
        pdf_path = cache_dir / "deck.pdf"
        if pdf_path.is_file():
            return pdf_path
        cache_dir.mkdir(parents=True, exist_ok=True)
        # fresh profile dir: parallel-safe, no first-run dialogs (§13)
        with tempfile.TemporaryDirectory(prefix="ppt-mcp-lo-") as profile:
            profile_uri = Path(profile).as_uri()
            result = subprocess.run(
                [
                    soffice,
                    f"-env:UserInstallation={profile_uri}",
                    "--headless",
                    "--norestore",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(cache_dir),
                    str(pptx_path),
                ],
                capture_output=True,
                text=True,
                timeout=CONVERT_TIMEOUT_S,
            )
        produced = cache_dir / (pptx_path.stem + ".pdf")
        if result.returncode != 0 or not produced.is_file():
            detail = (result.stderr or result.stdout or "").strip()[:500]
            raise PptMcpError(f"LibreOffice PDF conversion failed: {detail or 'no output'}")
        produced.replace(pdf_path)
        return pdf_path

    def render_slides(
        self, pptx_path: Path, slide_indices: list[int] | None = None, dpi: int = 96
    ) -> dict[int, Path]:
        """Render slides (1-based indices; None = all) to PNG. Returns
        {slide_index: png_path}, cached."""
        import pypdfium2 as pdfium

        pdf_path = self.to_pdf(pptx_path)
        cache_dir = pdf_path.parent
        pdf = pdfium.PdfDocument(str(pdf_path))
        try:
            page_count = len(pdf)
            wanted = slide_indices or list(range(1, page_count + 1))
            out: dict[int, Path] = {}
            for index in wanted:
                if not 1 <= index <= page_count:
                    raise PptMcpError(
                        f"slide_index {index} out of range; the rendered deck has "
                        f"{page_count} page(s)."
                    )
                png_path = cache_dir / f"slide-{index}-{dpi}dpi.png"
                if not png_path.is_file():
                    page = pdf[index - 1]
                    bitmap = page.render(scale=dpi / 72)
                    bitmap.to_pil().save(png_path)
                out[index] = png_path
            return out
        finally:
            pdf.close()


def diff_images(path_a: Path, path_b: Path) -> dict:
    """Pixel diff between two renders. Returns changed fraction and bbox."""
    from PIL import Image, ImageChops

    with Image.open(path_a) as im_a, Image.open(path_b) as im_b:
        if im_a.size != im_b.size:
            return {"changed": True, "changed_fraction": 1.0, "bbox": None,
                    "note": "image sizes differ"}
        diff = ImageChops.difference(im_a.convert("RGB"), im_b.convert("RGB"))
        bbox = diff.getbbox()
        if bbox is None:
            return {"changed": False, "changed_fraction": 0.0, "bbox": None}
        histogram = diff.convert("L").histogram()
        changed_px = sum(histogram[8:])  # tolerate tiny anti-aliasing noise
        total = im_a.size[0] * im_a.size[1]
        return {
            "changed": changed_px / total > 0.0005,
            "changed_fraction": round(changed_px / total, 4),
            "bbox": list(bbox),
        }
