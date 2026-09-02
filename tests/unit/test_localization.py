"""International language support gate: localized style-name aliases (Korean
first, plus the other majors), CJK-aware word counting, and the localized
detection sites (journalcount reference/abstract headings, TOC-heading
matching used by the live layer). Synthetic documents built with the
server-layer functions; no COM, no Word."""

import pytest

import word_mcp.server as srv
from word_mcp.core.package import DocxPackage
from word_mcp.ops import journalcount as jc
from word_mcp.ops import stats
from word_mcp.ops.citecheck import _REF_HEADINGS
from word_mcp.ops.localization import (
    STYLE_ALIASES,
    canonical_for_name,
    cjk_aware_word_count,
    localized_names,
    style_name_matches,
)


def new_doc(tmp_path, name="doc.docx"):
    path = str(tmp_path / name)
    srv.create_document(path)
    return path


# ------------------------------------------------------------ alias coverage


def test_heading1_aliases_all_languages():
    for name in (
        "Heading 1",        # en
        "제목 1",           # ko
        "Überschrift 1",    # de
        "Titre 1",          # fr
        "Título 1",         # es/pt
        "見出し 1",         # ja
        "标题 1",           # zh-CN
        "Titolo 1",         # it
    ):
        assert style_name_matches(name, "heading1"), name
        assert canonical_for_name(name) == "heading1", name


def test_korean_style_names():
    assert style_name_matches("제목 3", "heading3")
    assert style_name_matches("목차 제목", "toc_heading")
    assert style_name_matches("캡션", "caption")
    assert style_name_matches("인용", "quote")
    assert canonical_for_name("제목") == "title"  # Title, NOT a heading


def test_matching_is_case_and_whitespace_insensitive():
    assert style_name_matches("toc heading", "toc_heading")
    assert style_name_matches("TOCHeading", "toc_heading")
    assert style_name_matches("  TOC   Heading  ", "toc_heading")
    assert style_name_matches("제목1", "heading1")  # no-space Korean
    assert style_name_matches("HEADING 2", "heading2")


def test_english_names_still_match_and_negatives():
    assert style_name_matches("Heading 1", "heading1")
    assert not style_name_matches("Heading 1", "heading2")
    assert not style_name_matches("Body Text", "quote")
    assert canonical_for_name("Some Custom Style") is None
    assert canonical_for_name(None) is None
    assert not style_name_matches(None, "heading1")


def test_alias_table_shape_and_helpers():
    expected_keys = (
        [f"heading{n}" for n in range(1, 10)]
        + ["toc_heading", "caption", "quote", "block_text", "footnote_text",
           "title", "subtitle", "normal", "list_paragraph"]
        + [f"toc{n}" for n in range(1, 4)]
    )
    assert set(STYLE_ALIASES) == set(expected_keys)
    assert "목차 제목" in localized_names("toc_heading")
    with pytest.raises(KeyError):
        localized_names("bogus_key")
    with pytest.raises(KeyError):
        style_name_matches("Heading 1", "bogus_key")


def test_no_alias_maps_to_two_keys():
    seen = {}
    for key, names in STYLE_ALIASES.items():
        for name in names:
            norm = "".join(name.split()).casefold()
            assert seen.setdefault(norm, key) == key, name


# ------------------------------------------------------ CJK-aware word count


def test_pure_korean_counts_by_space_tokens():
    r = cjk_aware_word_count("한국어 문서의 단어를 공백으로 셉니다")
    assert r == {"words": 5, "cjk_chars": 0, "counting": "spaces"}


def test_pure_japanese_counts_by_character():
    # 10 kana/kanji characters; the fullwidth period is punctuation, not a word.
    r = cjk_aware_word_count("日本語のテキストです。")
    assert r == {"words": 10, "cjk_chars": 10, "counting": "cjk"}


def test_pure_chinese_counts_by_character():
    r = cjk_aware_word_count("中文文本没有空格")
    assert r == {"words": 8, "cjk_chars": 8, "counting": "cjk"}


def test_mixed_english_korean():
    # Hangul uses spaces, so mixed English+Korean is still pure token counting.
    r = cjk_aware_word_count("Delta Model 연구를 진행한다")
    assert r == {"words": 4, "cjk_chars": 0, "counting": "spaces"}


def test_mixed_english_japanese():
    r = cjk_aware_word_count("The word 単語 appears here.")
    assert r["cjk_chars"] == 2
    assert r["words"] == 4 + 2  # four space tokens + two CJK characters
    assert r["counting"] == "mixed"


def test_english_only_invariance_against_stats_tokenizer():
    text = "Space-delimited text must count exactly as before, 1 2 3."
    import re

    assert cjk_aware_word_count(text)["words"] == len(
        re.findall(r"\S+", text)
    )
    assert cjk_aware_word_count("")["words"] == 0
    assert cjk_aware_word_count("")["counting"] == "spaces"


def test_stats_word_count_invariance_and_new_fields(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "one two three four five"}], backup=False
    )
    r = stats.word_count(DocxPackage(path))
    assert r["totals"]["words"] == 5  # unchanged semantics
    assert r["totals"]["cjk_chars"] == 0
    assert r["counting"] == "spaces"


def test_stats_word_count_reports_cjk(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "日本語のテキスト"}], backup=False
    )
    r = stats.word_count(DocxPackage(path))
    assert r["totals"]["cjk_chars"] == 8
    assert r["counting"] == "cjk"


# --------------------------------------------- localized reference headings


def test_ref_headings_regex_all_languages():
    for heading in (
        "References", "Bibliography", "Works Cited", "Reference List",
        "참고문헌", "참고 문헌",       # ko
        "Literaturverzeichnis",        # de
        "Bibliographie",               # fr
        "Bibliografía", "Bibliografia",  # es / it+pt
        "参考文献",                    # ja + zh
        "Referências", "Referencias",  # pt / es
    ):
        assert _REF_HEADINGS.match(heading), heading
    assert not _REF_HEADINGS.match("Reference materials and other things")
    assert not _REF_HEADINGS.match("참고문헌과 부록")


def test_journalcount_korean_reference_heading(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "본문 내용이 여기에 있다."}], backup=False
    )                                                                # 4 words
    srv.insert_paragraphs(path, [{"text": "참고문헌", "heading_level": 1}], backup=False)  # 1 word
    srv.insert_paragraphs(
        path,
        [{"text": "Hurd, I. (1999). Legitimacy and authority."}], backup=False)                                   # 6 words
    out = jc.word_count_with_exclusions(
        DocxPackage(path), exclude=("references",)
    )
    assert out["zones_detected"]["references_section"]["heading"] == "참고문헌"
    assert out["excluded"] == {"references": 7}
    assert out["included"] == 4
    assert out["total"] == out["included"] + out["excluded_total"]
    assert out["counting"] == "spaces"  # Korean text is space-delimited
    assert out["cjk_chars"] == 0


def test_journalcount_korean_abstract_heading(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(path, [{"text": "초록", "heading_level": 1}], backup=False)      # 1 word
    srv.insert_paragraphs(
        path, [{"text": "이 논문은 무언가를 주장한다."}], backup=False)                                   # 4 words
    srv.insert_paragraphs(path, [{"text": "서론", "heading_level": 1}], backup=False)      # 1 word
    srv.insert_paragraphs(
        path, [{"text": "본문이 이어진다."}], backup=False
    )                                                                # 2 words
    out = jc.word_count_with_exclusions(
        DocxPackage(path), exclude=("abstract",)
    )
    assert out["zones_detected"]["abstract_section"]["heading"] == "초록"
    assert out["excluded"] == {"abstract": 5}  # heading + 4 body words
    assert out["total"] == out["included"] + out["excluded_total"]


def test_journalcount_japanese_doc_char_counting(tmp_path):
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "日本語の本文です。"}], backup=False
    )                                                       # 8 CJK characters
    out = jc.word_count_with_exclusions(DocxPackage(path), exclude=())
    assert out["total"] == 8
    assert out["cjk_chars"] == 8
    assert out["counting"] == "cjk"
    assert out["total"] == out["included"] + out["excluded_total"]


def test_journalcount_english_numbers_unchanged(tmp_path):
    """English docs must produce identical numbers to the pre-localization
    behavior (the invariance the full suite also enforces)."""
    path = new_doc(tmp_path)
    srv.insert_paragraphs(
        path, [{"text": "Plain English body text here."}], backup=False)
    srv.insert_paragraphs(path, [{"text": "References", "heading_level": 1}], backup=False)
    srv.insert_paragraphs(
        path, [{"text": "Author, A. (2020). Title."}], backup=False)
    out = jc.word_count_with_exclusions(
        DocxPackage(path), exclude=("references",)
    )
    assert out["total"] == 5 + 1 + 4
    assert out["excluded"] == {"references": 5}
    assert out["included"] == 5
    assert out["cjk_chars"] == 0
    assert out["counting"] == "spaces"


# --------------------------------------------------- TOC-heading recognition
# _sdt_regions (com/live_ops.py) matches Style.NameLocal against the
# "toc_heading" aliases; the match itself is exercised unit-level here (a live
# Word round-trip with a Korean UI language is not reproducible in CI).


def test_toc_heading_recognition_unit_level():
    for name_local in (
        "TOC Heading", "toc heading", "TOCHeading",   # en + legacy forms
        "목차 제목", "목차제목",                       # ko
        "Inhaltsverzeichnisüberschrift",              # de
        "En-tête de table des matières",              # fr
        "目次の見出し",                               # ja
    ):
        assert style_name_matches(name_local, "toc_heading"), name_local
    assert not style_name_matches("Heading 1", "toc_heading")
    assert not style_name_matches("목차 1", "toc_heading")  # TOC 1 != heading
