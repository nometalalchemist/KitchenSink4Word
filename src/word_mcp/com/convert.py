"""File-format conversions through a dedicated invisible Word instance.

Same discipline as com/bridge.py, guaranteed by reusing its `_word` context
manager: every call gets its OWN invisible DispatchEx instance (never
attaches to the user's visible Word), DisplayAlerts is off, and Quit runs in
a finally block. Nothing here touches a running Word or an open document.
"""

from __future__ import annotations

from pathlib import Path

from ..core.errors import DocumentNotFound, WordMcpError
from .bridge import _WD_DO_NOT_SAVE, _word

_WD_FORMAT_DOCX_DEFAULT = 16  # wdFormatDocumentDefault (.docx)
_WD_STAT_WORDS = 0  # wdStatisticWords
_WD_STAT_PAGES = 2  # wdStatisticPages

# Below this many words the conversion almost certainly found no text layer.
_SCANNED_PDF_WORD_FLOOR = 5


def import_pdf(pdf_path: str, output_path: str | None = None) -> dict:
    """Convert a PDF to .docx via Word's built-in PDF reflow.

    Word opens the PDF (ConfirmConversions off; the one-time "Word will now
    convert your PDF" alert is suppressed by DisplayAlerts=0) and the
    reflowed document is saved as .docx. Output defaults to the PDF's path
    with a .docx extension; an existing output file is refused, never
    overwritten. The produced .docx is validated by a DocxPackage
    round-trip (full ZIP + XML well-formedness check) before returning.

    Honesty about quality: reflow fidelity depends entirely on the PDF's
    internal structure. Text-based PDFs convert well; complex layouts
    (multi-column, heavy tables) may reflow imperfectly; scanned image PDFs
    have no text layer and produce little or no text — Word does not OCR.
    A near-zero word count in the result triggers a warning saying exactly
    that.
    """
    src = Path(pdf_path)
    if not src.exists():
        raise DocumentNotFound(f"no file at {pdf_path}")
    if src.suffix.lower() != ".pdf":
        raise WordMcpError(
            f"import_pdf expects a .pdf, got {src.suffix!r} — for other "
            "formats open the file in Word manually"
        )
    # magic-byte check: Word "successfully" text-fallback-imports a renamed
    # text file after a ~30s stall; a non-PDF deserves a typed refusal
    with open(src, "rb") as fh:
        if fh.read(5) != b"%PDF-":
            raise WordMcpError(
                f"{src.name} is not a PDF (missing %PDF- header) — rename "
                "tricks do not survive the magic-byte check"
            )
    out = Path(output_path) if output_path else src.with_suffix(".docx")
    if out.exists():
        raise WordMcpError(
            f"output already exists: {out} — refusing to overwrite; pass a "
            "different output_path or remove the file first"
        )
    if not out.parent.exists():
        raise WordMcpError(
            f"output directory does not exist: {out.parent} — create it "
            "first (Word's converter cannot)"
        )

    try:
        with _word() as app:
            doc = app.Documents.Open(
                str(src.resolve()),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
            )
            try:
                doc.SaveAs2(
                    str(out.resolve()), FileFormat=_WD_FORMAT_DOCX_DEFAULT
                )
                words = int(doc.ComputeStatistics(_WD_STAT_WORDS))
                pages = int(doc.ComputeStatistics(_WD_STAT_PAGES))
            finally:
                doc.Close(SaveChanges=_WD_DO_NOT_SAVE)
    except WordMcpError:
        raise
    except Exception as exc:
        raise WordMcpError(
            f"Word could not convert {src.name}: {exc} — the PDF may be "
            "encrypted, damaged, or the output location unwritable"
        ) from exc

    if not out.exists():
        raise WordMcpError("Word reported success but produced no output")

    # Validate: the produced .docx must survive our own load/save round-trip
    # (ZIP integrity, document.xml present, every XML part well-formed).
    from ..core.package import DocxPackage

    pkg = DocxPackage(out)
    pkg.save(do_backup=False)
    from ..ops.read import body_items

    paragraphs = sum(
        1 for kind, _, _ in body_items(pkg) if kind == "paragraph"
    )

    result = {
        "docx": str(out),
        "bytes": out.stat().st_size,
        "pages": pages,
        "words": words,
        "paragraphs": paragraphs,
        "note": (
            "Word's PDF reflow: fidelity depends on the PDF's internal "
            "structure — text-based PDFs convert well; complex layouts may "
            "reflow imperfectly and need review"
        ),
    }
    if words <= _SCANNED_PDF_WORD_FLOOR:
        result["warning"] = (
            f"result contains little or no text (word count {words}) — the "
            "PDF is likely a scanned image; Word's reflow does not OCR, so "
            "the page content was not recovered as text"
        )
    return result
