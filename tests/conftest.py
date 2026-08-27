"""The test suite runs against a corpus of REAL documents in tests/corpus/
(the author's own dissertation drafts and manuscripts — private, not shipped).

Without the corpus, corpus-dependent tests are skipped with an explanation.
To run the full suite, drop your own .docx files into tests/corpus/ using the
expected names below — documents with genuine footnotes, tracked changes, and
comments give the most meaningful coverage.
"""

from pathlib import Path

import pytest

CORPUS = Path(__file__).parent / "corpus"

# name -> features the tests expect it to contain
EXPECTED = {
    "ch4.docx": "ordinary prose document with a couple of tables",
    "ch5.docx": "ordinary prose document",
    "ch1-3.docx": "long document with headings, sections, and a real TOC",
    "codebook.docx": "document with many tables",
    "unitar.docx": "document containing images",
    "ch4_chair.docx": "any second prose document",
    "niu.docx": "document with many real footnotes (tests expect 171)",
    "ejir_rw.docx": "document with tracked changes (tests expect 126) and comments",
    "outline.docx": "document with threaded comments",
}


def pytest_collection_modifyitems(config, items):
    present = {p.name for p in CORPUS.glob("*.docx")}
    missing = set(EXPECTED) - present
    if not missing:
        return
    skip = pytest.mark.skip(
        reason=(
            f"test corpus incomplete (missing: {sorted(missing)}). "
            "The corpus is private real-world documents — see tests/conftest.py "
            "for what to supply to run these tests yourself."
        )
    )
    uses_corpus: dict[str, bool] = {}
    for item in items:
        fname = str(item.fspath)
        if fname not in uses_corpus:
            uses_corpus[fname] = "corpus" in Path(fname).read_text(
                encoding="utf-8", errors="ignore"
            )
        # Only corpus-dependent tests are skipped; corpus-free tests still run.
        if uses_corpus[fname]:
            item.add_marker(skip)
