"""Publication-style conversion bundle: parse / convert / manuscript format.

v2-staged copy (Phase 2 Wave B): server-surface calls renamed to the v2
grammar (manage_source, manage_note, insert_citation with a location
object); ops-level calls, fixtures, and assertions unchanged from the v1
suite. Fixtures are synthetic manuscripts built with the server's own
tools: a 10-entry APA reference list with 2 deliberately malformed entries
and a body mixing parenthetical, multi-work, and narrative citations. Gate
cases per the v1.5 spec: APA -> Chicago author-date, APA -> IEEE
(first-appearance numbering, narrative names kept), APA -> Chicago notes
(REAL footnotes, long-then-short), IEEE -> APA round trip, dry_run
byte-identity, malformed entries untouched + flagged, native-field
routing, and the APA manuscript format verified via sectPr/pPr reads.
"""

import hashlib
import shutil

import pytest
from lxml import etree

import word_mcp.server as srv
from word_mcp.core.errors import TargetNotFound, WordMcpError
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import notes, read, styleconvert as sc

BODY = [
    "Introduction",
    "Legitimacy shapes compliance in international institutions (Hurd, 1999, p. 384).",
    "Lake (2009) argues that hierarchy persists even among formal equals.",
    "Alliance behavior tracks perceived threat rather than raw power (Walt, 1987; Snyder, 1997).",
    "Mearsheimer (2001) rejects this claim entirely, and the point recurs elsewhere (Hurd, 1999).",
    "References",
]
ENTRIES = [
    "Hurd, I. (1999). Legitimacy and authority in international politics. International Organization, 53(2), 379-408.",
    "Jervis, R. (1976). Perception and misperception in international politics. Princeton University Press.",
    "Keohane, R. O. (1984). After hegemony: Cooperation and discord in the world political economy. Princeton University Press.",
    "Lake, D. A. (2009). Hierarchy in international relations. Cornell University Press.",
    "Mearsheimer, J. J. (2001). The tragedy of great power politics. W. W. Norton.",
    "Snyder, G. H. (1997). Alliance politics. Cornell University Press.",
    "Walt, S. M. (1987). The origins of alliances. Cornell University Press.",
    "Wendt, A. (1999). Social theory of international politics. Cambridge University Press.",
    # deliberately malformed:
    "The legitimacy problem, various authors, undated conference manuscript.",
    "assorted notes without any recognizable structure or date",
]
MALFORMED = ENTRIES[8:]

# Clean 4-entry manuscript (every entry parses) for reorder/notes round trips.
CLEAN_BODY = [
    "Legitimacy shapes compliance in institutions (Hurd, 1999, p. 384).",
    "Lake (2009) argues that hierarchy persists.",
    "Alliance behavior tracks threat (Walt, 1987; Snyder, 1997).",
    "References",
]
CLEAN_ENTRIES = [
    "Walt, S. M. (1987). The origins of alliances. Cornell University Press.",
    "Hurd, I. (1999). Legitimacy and authority in international politics. International Organization, 53(2), 379-408.",
    "Snyder, G. H. (1997). Alliance politics. Cornell University Press.",
    "Lake, D. A. (2009). Hierarchy in international relations. Cornell University Press.",
]


def _build(tmp_path, name, body, entries, heading_index):
    path = str(tmp_path / name)
    srv.create_document(path)
    srv.insert_paragraphs(
        path, [{"text": t} for t in body + entries],
        backup=False, live="off",
    )
    srv.apply_style(path, style="Heading1", range={"start": heading_index, "end": heading_index}, backup=False)
    return path


@pytest.fixture
def manuscript(tmp_path):
    return _build(tmp_path, "m.docx", BODY, ENTRIES, 5)


@pytest.fixture
def clean_manuscript(tmp_path):
    return _build(tmp_path, "clean.docx", CLEAN_BODY, CLEAN_ENTRIES, 3)


def _convert(path, target, **kw):
    pkg = DocxPackage(path)
    result = sc.convert_citation_style(pkg, target, **kw)
    if result.get("converted"):
        pkg.save(do_backup=False)
    return result


def _texts(path):
    return [p["text"] for p in read.get_paragraphs(DocxPackage(path))]


def _md5(path):
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


# ------------------------------------------------------------------- parsing


def test_parse_references_model(manuscript):
    rep = sc.parse_references(DocxPackage(manuscript))
    assert rep["detected_source_system"] == "author-date"
    assert rep["confidence_counts"] == {"full": 8, "partial": 0, "failed": 2}
    hurd = next(e for e in rep["entries"] if e["raw"].startswith("Hurd"))
    assert hurd["type"] == "article"
    assert hurd["authors"] == [{"family": "Hurd", "given": "I."}]
    assert hurd["year"] == "1999"
    assert hurd["container"] == "International Organization"
    assert (hurd["volume"], hurd["issue"], hurd["pages"]) == ("53", "2", "379-408")
    lake = next(e for e in rep["entries"] if e["raw"].startswith("Lake"))
    assert lake["type"] == "book"
    assert lake["publisher"] == "Cornell University Press"
    kinds = sorted(c["kind"] for c in rep["citations"])
    assert kinds == ["narrative", "narrative", "parenthetical",
                     "parenthetical", "parenthetical"]
    multi = next(c for c in rep["citations"] if "Walt" in c["raw"])
    assert len(multi["items"]) == 2


def test_parse_failed_entries_verbatim(manuscript):
    rep = sc.parse_references(DocxPackage(manuscript))
    failed = [e for e in rep["entries"] if e["parse_confidence"] == "failed"]
    assert {e["raw"] for e in failed} == set(MALFORMED)


def test_no_reference_heading_raises(tmp_path):
    path = str(tmp_path / "bare.docx")
    srv.create_document(path)
    srv.insert_paragraphs(
        path, [{"text": "Just prose, no list."}],
        backup=False, live="off",
    )
    with pytest.raises(TargetNotFound):
        sc.parse_references(DocxPackage(path))


def test_unknown_styles_rejected(manuscript):
    with pytest.raises(WordMcpError):
        sc.parse_references(DocxPackage(manuscript), style_hint="bluebook")
    with pytest.raises(WordMcpError):
        sc.convert_citation_style(DocxPackage(manuscript), "chicago")  # ambiguous


# ----------------------------------------------------- gate: APA -> Chicago AD


def test_apa_to_chicago_author_date(manuscript):
    result = _convert(manuscript, "chicago17-author-date")
    assert result["converted"] and result["entries_converted"] == 8
    texts = _texts(manuscript)
    # in-text: comma dropped, page without p.
    assert any("(Hurd 1999, 384)" in t for t in texts)
    assert any("(Walt 1987; Snyder 1997)" in t for t in texts)
    # narrative form is identical in both styles, so it must survive
    assert any(t.startswith("Lake (2009) argues") for t in texts)
    # reference entry: inverted name, year after period, quoted title-case
    # title, journal + vol (iss): pages
    hurd = next(t for t in texts if t.startswith("Hurd"))
    assert hurd.startswith("Hurd, I. 1999. “Legitimacy and Authority in International Politics.”")
    assert "International Organization 53 (2): 379-408." in hurd
    # honesty: Chicago wants full given names; source had initials
    assert any(
        "initials only" in f["problem"] for f in result["entries_flagged"]
    )


# ---------------------------------------------------------- gate: APA -> IEEE


def test_apa_to_ieee_numbering_and_narrative(manuscript):
    result = _convert(manuscript, "ieee")
    texts = _texts(manuscript)
    # numbers by first appearance: Hurd 1, Lake 2, Walt 3, Snyder 4, Mearsheimer 5
    assert any("institutions [1, p. 384]." in t for t in texts)
    assert any("power [3], [4]." in t for t in texts)
    assert any("elsewhere [1]." in t for t in texts)
    # narrative keeps the name in prose
    assert any(t.startswith("Lake [2] argues") for t in texts)
    assert any(t.startswith("Mearsheimer [5] rejects") for t in texts)
    # entries carry their number prefix and IEEE shape
    hurd = next(t for t in texts if "Legitimacy and authority" in t)
    assert hurd.startswith("[1] I. Hurd, “Legitimacy and authority in international politics,”")
    assert "vol. 53, no. 2, pp. 379-408, 1999." in hurd
    lake = next(t for t in texts if "Hierarchy in international relations" in t)
    assert lake.startswith("[2] D. A. Lake,")
    # mixed doc: list order could not be normalized, which must be flagged
    assert any("could not be normalized" in f for f in result["flags"])


def test_malformed_entries_untouched_and_flagged(manuscript):
    result = _convert(manuscript, "ieee")
    texts = _texts(manuscript)
    for raw in MALFORMED:
        assert raw in texts  # byte-for-byte identical paragraph text
    flagged = {f["entry"] for f in result["entries_flagged"]
               if f["confidence"] == "failed"}
    assert len(flagged) == 2
    assert result["review_required"]


# ------------------------------------------------- gate: APA -> Chicago notes


def test_apa_to_chicago_notes_real_footnotes(manuscript):
    result = _convert(manuscript, "chicago17-notes")
    assert result["footnotes_created"] == 5
    pkg = DocxPackage(manuscript)
    fns = read.list_footnotes(pkg)
    assert len(fns) == 5
    # notes part integrity: separator/continuation present, refs paired
    v = notes.validate_notes(pkg)
    assert v["footnotes"]["ok"] is True
    # parentheticals removed from the body
    texts = _texts(manuscript)
    assert not any("(Hurd, 1999" in t for t in texts)
    assert not any("(Walt, 1987" in t for t in texts)
    # narrative citations keep the name, lose the year paren
    assert any(t.startswith("Lake argues") for t in texts)
    # first use = long note (journal, year, cited page); later = short note
    hurd_notes = [f["text"] for f in fns if "Hurd" in f["text"]]
    assert len(hurd_notes) == 2
    long_note = next(t for t in hurd_notes if "International Organization" in t)
    assert "I. Hurd, “Legitimacy and Authority in International Politics,”" in long_note
    assert "(1999): 384." in long_note
    short_note = next(t for t in hurd_notes if "International Organization" not in t)
    assert "Hurd, “Legitimacy and Authority.”" in short_note
    # multi-work citation becomes ONE note with both works
    multi = next(f["text"] for f in fns if "Walt" in f["text"])
    assert "Snyder" in multi
    # heading renamed for notes-bibliography style
    assert result["heading_renamed"] == "Bibliography"
    assert any("Bibliography" == t for t in texts)


# ------------------------------------------------ gate: IEEE -> APA round trip


def test_ieee_to_apa_round_trip(manuscript):
    original_body = _texts(manuscript)[:5]
    _convert(manuscript, "ieee")
    result = _convert(manuscript, "apa7")
    assert result["converted"]
    texts = _texts(manuscript)
    # body citations return to their exact original APA form
    assert texts[:5] == original_body
    # entries survive structurally
    rep = sc.parse_references(DocxPackage(manuscript))
    assert rep["confidence_counts"]["full"] == 8
    hurd = next(e for e in rep["entries"] if "Legitimacy" in (e.get("title") or ""))
    assert hurd["type"] == "article"
    assert (hurd["volume"], hurd["issue"]) == ("53", "2")
    assert hurd["year"] == "1999"
    lake = next(e for e in rep["entries"] if "Hierarchy" in (e.get("title") or ""))
    assert lake["type"] == "book"
    assert lake["publisher"] == "Cornell University Press"
    # malformed entries still byte-identical after TWO conversions
    for raw in MALFORMED:
        assert raw in texts


# ------------------------------------------------------- gate: dry_run inert


def test_dry_run_changes_nothing(manuscript):
    before = _md5(manuscript)
    plan = sc.convert_citation_style(
        DocxPackage(manuscript), "ieee", dry_run=True
    )
    assert _md5(manuscript) == before  # byte-identical
    assert plan["dry_run"] is True and plan["converted"] is False
    assert plan["counts"]["citations_to_convert"] == 5
    assert plan["counts"]["entries_to_convert"] == 8
    assert plan["counts"]["entries_flagged"] == 2
    # the plan carries per-item before/after
    ops = plan["plan"]["citations"]
    assert all("before" in o and "after" in o for o in ops)
    narrative = next(o for o in ops if o["before"] == "Lake (2009)")
    assert narrative["after"] == "[2]"
    entry_ops = plan["plan"]["entries"]
    assert all("before" in o and "after" in o for o in entry_ops)


# -------------------------------------------------- gate: native-field routing


def test_native_field_doc_routed_not_rewritten(tmp_path):
    path = str(tmp_path / "native.docx")
    srv.create_document(path)
    srv.insert_paragraphs(
        path,
        [{"text": "Legitimacy matters here."}, {"text": "References"}], backup=False, live="off",
    )
    srv.manage_source(
        path, action="add", tag="Hurd1999", source_type="JournalArticle",
        title="Legitimacy and authority",
        year="1999", authors=[{"last": "Hurd", "first": "Ian"}],
        journal_name="International Organization",
        backup=False,
    )
    srv.insert_citation(
        path, "Hurd1999", location={"search": {"text": "matters here"}},
        backup=False,
    )
    before = _md5(path)
    result = sc.convert_citation_style(DocxPackage(path), "apa7")
    assert result["routed"] is True and result["converted"] is False
    assert result["citation_fields"]["native"] >= 1
    assert any("set_bibliography_style" in a for a in result["action_required"])
    assert _md5(path) == before  # nothing touched, nothing saved


# ------------------------------------------------------ Vancouver superscript


def test_apa_to_vancouver_superscript(clean_manuscript):
    result = _convert(clean_manuscript, "vancouver")
    texts = _texts(clean_manuscript)
    assert any(t == "Legitimacy shapes compliance in institutions.1" for t in texts)
    assert any(t.startswith("Lake2 argues") for t in texts)
    assert any(t.endswith("threat.3,4") for t in texts)
    # the number is a REAL superscript run
    pkg = DocxPackage(clean_manuscript)
    supers = [
        r for r in pkg.root().iter(qn("w:r"))
        if r.find(qn("w:rPr")) is not None
        and r.find(qn("w:rPr")).find(qn("w:vertAlign")) is not None
        and "".join(t.text or "" for t in r.findall(qn("w:t"))).strip().isdigit()
    ]
    assert supers
    # citation-order reordering happened (all entries parsed)
    assert result["list_reordered"] is True
    ref_pos = texts.index("References")
    assert texts[ref_pos + 1].startswith("1. Hurd I.")
    # locator loss is flagged, not silent
    assert any(
        "locator dropped" in f["problem"] for f in result["citations_flagged"]
    )


def test_vancouver_back_to_apa(clean_manuscript):
    _convert(clean_manuscript, "vancouver")
    result = _convert(clean_manuscript, "apa7")
    texts = _texts(clean_manuscript)
    assert any("institutions (Hurd, 1999)." in t for t in texts)
    assert any(t.startswith("Lake (2009) argues") for t in texts)
    assert any("(Walt, 1987; Snyder, 1997)" in t for t in texts)
    assert result["list_reordered"] is True  # back to alphabetical


# --------------------------------------------------------- notes -> APA (harvest)


def test_notes_round_trip_harvest(clean_manuscript):
    _convert(clean_manuscript, "chicago17-notes")
    result = _convert(
        clean_manuscript, "apa7", source_style="chicago17-notes"
    )
    assert result["footnotes_harvested"] == 3
    assert read.list_footnotes(DocxPackage(clean_manuscript)) == []
    texts = _texts(clean_manuscript)
    assert any("(Hurd, 1999, p. 384)" in t for t in texts)  # locator restored
    assert any("(Walt, 1987; Snyder, 1997)" in t for t in texts)


def test_mixed_content_footnote_left_alone(clean_manuscript):
    _convert(clean_manuscript, "chicago17-notes")
    # add a commentary footnote that must NOT be harvested
    srv.manage_note(
        clean_manuscript, action="insert", note_type="footnote",
        text="This claim is contested; see the discussion in chapter 3.",
        location={"search": {"text": "hierarchy persists"}},
        backup=False,
    )
    result = _convert(
        clean_manuscript, "apa7", source_style="chicago17-notes"
    )
    assert result["footnotes_harvested"] == 3
    remaining = read.list_footnotes(DocxPackage(clean_manuscript))
    assert len(remaining) == 1
    assert "contested" in remaining[0]["text"]
    assert any(
        "not recognizably a pure citation" in f["problem"]
        for f in result["citations_flagged"]
    )


# ------------------------------------------------------ gate: manuscript format


def test_manuscript_format_apa(manuscript):
    pkg = DocxPackage(manuscript)
    result = sc.apply_manuscript_format(pkg, "apa7")
    pkg.save(do_backup=False)
    assert result["not_applied"]  # honesty list is never silently empty here

    pkg = DocxPackage(manuscript)
    # sectPr: 1-inch margins
    from word_mcp.ops.furniture import list_sections

    margins = list_sections(pkg)[0]["margins_pt"]
    assert margins == {"top": 72.0, "bottom": 72.0, "left": 72.0, "right": 72.0}
    # pPr: double spacing on a body paragraph
    body_p = next(
        el for kind, idx, el in read.body_items(pkg)
        if kind == "paragraph" and "Legitimacy shapes" in read.paragraph_text(el)
    )
    spacing = body_p.find(qn("w:pPr")).find(qn("w:spacing"))
    assert spacing.get(qn("w:line")) == "480"
    assert spacing.get(qn("w:lineRule")) == "auto"
    # header carries a PAGE field (student paper: page number only)
    header_parts = [
        p for p in pkg.part_names() if p.startswith("word/header")
    ]
    assert header_parts
    assert any(
        b"PAGE" in pkg.read_part(p) if hasattr(pkg, "read_part")
        else b"PAGE" in etree.tostring(pkg.root(p))
        for p in header_parts
    )
    # hanging indent on reference entries
    hurd_p = next(
        el for kind, idx, el in read.body_items(pkg)
        if kind == "paragraph" and read.paragraph_text(el).startswith("Hurd, I.")
    )
    ind = hurd_p.find(qn("w:pPr")).find(qn("w:ind"))
    assert ind.get(qn("w:hanging")) == "720"
    assert ind.get(qn("w:left")) == "720"


def test_manuscript_format_journal_styles_refused(manuscript):
    with pytest.raises(WordMcpError):
        sc.apply_manuscript_format(DocxPackage(manuscript), "ieee")


def test_manuscript_format_mla_header(manuscript):
    pkg = DocxPackage(manuscript)
    result = sc.apply_manuscript_format(
        pkg, "mla9", author_last_name="Smith"
    )
    pkg.save(do_backup=False)
    assert any("Smith" in a for a in result["applied"])
    pkg = DocxPackage(manuscript)
    header_parts = [p for p in pkg.part_names() if p.startswith("word/header")]
    assert any(
        b"Smith" in etree.tostring(pkg.root(p)) for p in header_parts
    )


# ----------------------------------------------------------- error atomicity


def test_error_leaves_file_untouched(manuscript):
    before = _md5(manuscript)
    with pytest.raises(WordMcpError):
        _convert(manuscript, "not-a-style")
    assert _md5(manuscript) == before
