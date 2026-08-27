"""The test suite runs against a corpus of REAL documents in tests/corpus/
(the author's own long-form manuscripts — private, not shipped).

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


def pytest_configure(config):
    """Missing corpus files are GENERATED as structural stand-ins (see
    tests/make_corpus.py), so the full suite runs anywhere — CI included.
    Real local documents, when present, always take precedence."""
    present = {p.name for p in CORPUS.glob("*.docx")}
    if set(EXPECTED) - present:
        import make_corpus  # noqa: F401  (lives beside this file)

        made = make_corpus.generate_missing(verbose=True)
        if made:
            print(f"conftest: generated {len(made)} synthetic corpus file(s)")


def pytest_collection_modifyitems(config, items):
    present = {p.name for p in CORPUS.glob("*.docx")}
    missing = set(EXPECTED) - present
    if not missing:
        return
    skip = pytest.mark.skip(
        reason=(
            f"test corpus incomplete (missing: {sorted(missing)}) and "
            "generation failed — run python tests/make_corpus.py for details."
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
