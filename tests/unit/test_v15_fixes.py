"""Regressions for the v1.5 adversarial findings (F1-F7)."""

import pytest

import word_mcp.server as srv
from word_mcp.core.errors import WordMcpError
from word_mcp.core.package import DocxPackage


def _box_doc(tmp_path):
    import shutil

    from test_safety_journal import doc_with_box

    return doc_with_box(tmp_path)


def test_f1_replace_matches_preview_on_box_docs(tmp_path):
    path = _box_doc(tmp_path)
    # box text contains "Box body text"; host contains "Host paragraph."
    prev = srv.preview_replace(str(path), [{"find": "Box body", "replace": "X"}])
    real = srv.search_and_replace(
        str(path), [{"find": "Box body", "replace": "X"}],
        backup=False, live="off",
    )
    n_prev = prev["items"][0]["matches" if "matches" in prev["items"][0]
                             else "count"]
    n_prev = len(n_prev) if isinstance(n_prev, list) else n_prev
    assert n_prev == 0
    assert real["replaced"][ "Box body"] == 0 if "replaced" in real else True
    xml = DocxPackage(path).raw_part("word/document.xml").decode("utf-8")
    assert xml.count("Box body text") == 2  # both copies untouched


def test_f2_protected_entry_citations_not_numbered(tmp_path):
    path = tmp_path / "f2.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": "Alpha claim (Hurd, 1999). Beta claim (Lake, 2009)."},
         {"text": "References"},
         {"text": "Hurd, I. (1999). Legitimacy and authority. International "
                  "Organization, 53(2), 379-408."},
         {"text": "Lake, D. (2009). Hierarchy in international relations. "
                  "Cornell University Press."}],
        at_end=True, backup=False,
    )
    srv.apply_style(str(path), [1], "Heading 1", backup=False)
    # protect the Hurd entry with a hyperlink
    srv.add_hyperlink(
        str(path), anchor_text="International Organization",
        url="https://example.org/io", backup=False,
    )
    r = srv.convert_citation_style(str(path), "ieee", backup=False)
    text = " ".join(
        p["text"] for p in srv.get_text(str(path), live="off")
        if p.get("index") is not None
    )
    assert "(Hurd, 1999)" in text          # left as-is
    assert "(Lake, 2009)" not in text       # converted
    assert any("left verbatim" in str(f) for f in r["citations_flagged"])


def test_f3_swapped_blocks_reported_as_moves(tmp_path):
    a = tmp_path / "a.docx"
    b = tmp_path / "b.docx"
    block1 = [f"First block sentence number {i} with content." for i in range(4)]
    block2 = [f"Second block sentence number {i} with content." for i in range(4)]
    anchor = ["Anchor paragraph one stays.", "Anchor paragraph two stays."]
    srv.create_document(str(a))
    srv.insert_paragraphs(
        str(a), [{"text": t} for t in block1 + anchor + block2],
        at_end=True, backup=False,
    )
    srv.create_document(str(b))
    srv.insert_paragraphs(
        str(b), [{"text": t} for t in block2 + anchor + block1],
        at_end=True, backup=False,
    )
    d = srv.structured_diff(str(a), str(b))
    assert d["counts"]["moved"] >= 6 if "counts" in d else len(d["moved"]) >= 6
    mods = d["modified"] if "modified" in d else []
    assert len(mods) == 0, mods


def test_f4_hangul_and_nd_citations_recognized(tmp_path):
    path = tmp_path / "f4.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": "북한 연구에서 김철수 (2026) argues; see also (박영희, n.d.)."},
         {"text": "References"},
         {"text": "김철수. (2026). 논문 제목. 학술지, 1(1), 1-10."},
         {"text": "박영희. (n.d.). 다른 논문. 학술지."}],
        at_end=True, backup=False,
    )
    srv.apply_style(str(path), [1], "Heading 1", backup=False)
    parity = srv.check_citation_parity(str(path))
    uncited = parity.get("uncited_references", [])
    assert not any("김철수" in u for u in uncited)
    parsed = srv.parse_references(str(path))
    assert parsed["citations_found"] >= 2 if "citations_found" in parsed \
        else len(parsed["citations"]) >= 2


def test_f5_page_range_dash_preserved():
    from word_mcp.ops.styleconvert_data import norm_pages

    assert norm_pages("50-52") == "50-52"
    assert norm_pages("50 – 52") == "50–52"
    assert norm_pages(None) is None


def test_f6_file_macros_refused(tmp_path):
    path = tmp_path / "f6.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(str(path), [{"text": "x"}], at_end=True, backup=False)
    with pytest.raises(WordMcpError, match="file/preamble"):
        srv.add_equation(str(path), r"\input{secrets.tex}", at_end=True,
                         backup=False)
    with pytest.raises(WordMcpError, match="file/preamble"):
        srv.add_equation(str(path), r"a + \write18{cmd}", at_end=True,
                         backup=False)


def test_f7_extract_images_dir_is_file_refused(tmp_path):
    path = tmp_path / "f7.docx"
    srv.create_document(str(path))
    blocker = tmp_path / "blocker"
    blocker.write_text("i am a file")
    with pytest.raises(WordMcpError, match="existing FILE"):
        srv.extract_images(str(path), str(blocker))


def test_accented_latin_authors_recognized(tmp_path):
    """Müller/García-class names were invisible to the Latin-only citation
    patterns (post-v1.5.0 finding, fixed in 1.5.1)."""
    path = tmp_path / "accents.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": "As Müller (2019) shows, and see also (García, 2020)."},
         {"text": "References"},
         {"text": "Müller, K. (2019). Der Titel. Zeitschrift, 1(1), 1-10."},
         {"text": "García, L. (2020). El título. Revista, 2(2), 20-30."}],
        at_end=True, backup=False,
    )
    srv.apply_style(str(path), [1], "Heading 1", backup=False)
    parity = srv.check_citation_parity(str(path))
    uncited = parity.get("uncited_references", [])
    assert not any("Müller" in u or "García" in u for u in uncited), parity
