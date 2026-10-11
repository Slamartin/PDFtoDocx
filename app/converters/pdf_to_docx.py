"""PDF -> DOCX using pdf2docx (layout analysis on top of PyMuPDF)."""

from __future__ import annotations

import logging
from pathlib import Path

import pymupdf
from pdf2docx import Converter
from pdf2docx.shape.Path import Path as _Pdf2DocxPath

from app.report import ConversionReport, CorruptFileError, PasswordError

log = logging.getLogger(__name__)


def _blend_over_white(components, opacity):
    """pdf2docx ignores opacity, so a 3% black fill becomes solid black shading.
    Pre-blend the colour against white; CMYK white is all zeros, Gray/RGB white is all ones."""
    if len(components) == 4:
        return [opacity * c for c in components]
    return [1 - opacity * (1 - c) for c in components]


def _patch_pdf2docx_opacity() -> None:
    if getattr(_Pdf2DocxPath, "_opacity_patched", False):
        return
    original_init = _Pdf2DocxPath.__init__

    def patched_init(self, raw):
        raw = dict(raw)
        for colour_key, opacity_key in (("fill", "fill_opacity"), ("color", "stroke_opacity")):
            colour, opacity = raw.get(colour_key), raw.get(opacity_key)
            if colour is not None and opacity is not None and 0 <= opacity < 1:
                raw[colour_key] = _blend_over_white(list(colour), opacity)
                raw[opacity_key] = 1
        original_init(self, raw)

    _Pdf2DocxPath.__init__ = patched_init
    _Pdf2DocxPath._opacity_patched = True


_patch_pdf2docx_opacity()


def _summarise(nums: list[int]) -> str:
    """[1, 2, 3, 7] -> '1-3, 7'."""
    out: list[str] = []
    start = prev = nums[0]
    for n in nums[1:] + [None]:
        if n is not None and n == prev + 1:
            prev = n
            continue
        out.append(str(start) if start == prev else f"{start}-{prev}")
        if n is not None:
            start = prev = n
    return ", ".join(out)


def inspect_pdf(src: Path, password: str | None) -> ConversionReport:
    """Open the PDF once to check the password and gather counts for the report."""
    try:
        doc = pymupdf.open(str(src))
    except Exception as exc:  # pymupdf raises a plain Exception family for bad files
        raise CorruptFileError(f"Not a readable PDF: {exc}") from exc
    try:
        if doc.needs_pass:
            if not password or not doc.authenticate(password):
                raise PasswordError("wrong or missing password")
        report = ConversionReport(pages=doc.page_count)
        textless: list[int] = []
        for page in doc:
            if not page.get_text("text").strip():
                textless.append(page.number + 1)
            report.images += len(page.get_images(full=True))
            try:
                report.tables += len(page.find_tables().tables)
            except Exception:  # table finder is best-effort only
                pass
        if textless:
            if len(textless) == doc.page_count:
                report.warn("No text layer found. Every page went in as a picture; the text is not editable (no OCR).")
            else:
                if len(textless) == 1:
                    report.warn(f"Page {textless[0]} has no text layer and went in as a picture (no OCR).")
                else:
                    report.warn(f"Pages {_summarise(textless)} have no text layer and went in as pictures (no OCR).")
        return report
    finally:
        doc.close()


def convert_pdf(src: Path, dst: Path, password: str | None = None) -> ConversionReport:
    report = inspect_pdf(src, password)
    cv = Converter(str(src), password=password or None)
    try:
        cv.convert(str(dst))
    except Exception as exc:
        raise CorruptFileError(f"pdf2docx could not convert this file: {exc}") from exc
    finally:
        cv.close()
    if not dst.exists():
        raise CorruptFileError("pdf2docx produced no output")
    report.note("Fonts are mapped to the closest installed font. Hyperlinks and form fields become plain text.")
    return report
