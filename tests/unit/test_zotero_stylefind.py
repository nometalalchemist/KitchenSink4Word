"""ZOTERO TIER-2 + STYLE-AWARE FIND bundle tests.

The Zotero tests run against a SYNTHETIC zotero.sqlite built here with the
real Zotero 7 schema (verified against a live database, userdata v125) and
entirely fictional items. The user's real Zotero database is never read by
any test. Style-find tests run on documents built with the server's own
formatting tools, covering explicit-rPr hits, style-inherited hits, mixed
criteria, and unknown-key rejection.
"""

import json
import sqlite3

import pytest

import word_mcp.server as srv
from word_mcp.core.errors import TargetNotFound, WordMcpError
from word_mcp.core.package import DocxPackage
from word_mcp.ops import stylefind, zoterolib
from word_mcp.ops.reffields import (
    check_reference_field_integrity,
    list_reference_fields,
    scan_complex_fields,
)

LOCAL_USER_KEY = "zZ8kQp2A"

# Column definitions mirror the live Zotero 7 database (only the tables the
# module touches). fieldIDs are arbitrary on purpose: the code must resolve
# fields by name, never by hardcoded id.
_SCHEMA = """
CREATE TABLE libraries (libraryID INTEGER PRIMARY KEY, type TEXT,
    editable INT, filesEditable INT, version INT DEFAULT 0,
    storageVersion INT DEFAULT 0, lastSync INT DEFAULT 0,
    archived INT DEFAULT 0, isAdmin INT);
CREATE TABLE items (itemID INTEGER PRIMARY KEY, itemTypeID INT,
    dateAdded TIMESTAMP, dateModified TIMESTAMP,
    clientDateModified TIMESTAMP, libraryID INT, key TEXT,
    version INT DEFAULT 0, synced INT DEFAULT 0);
CREATE TABLE itemTypes (itemTypeID INTEGER PRIMARY KEY, typeName TEXT,
    templateItemTypeID INT, display INT);
CREATE TABLE fields (fieldID INTEGER PRIMARY KEY, fieldName TEXT,
    fieldFormatID INT);
CREATE TABLE itemData (itemID INT, fieldID INT, valueID);
CREATE TABLE itemDataValues (valueID INTEGER PRIMARY KEY, value);
CREATE TABLE creators (creatorID INTEGER PRIMARY KEY, firstName TEXT,
    lastName TEXT, fieldMode INT);
CREATE TABLE itemCreators (itemID INT, creatorID INT, creatorTypeID INT,
    orderIndex INT);
CREATE TABLE creatorTypes (creatorTypeID INTEGER PRIMARY KEY,
    creatorType TEXT);
CREATE TABLE deletedItems (itemID INTEGER PRIMARY KEY, dateDeleted);
CREATE TABLE itemAttachments (itemID INTEGER PRIMARY KEY, parentItemID INT,
    linkMode INT, contentType TEXT, charsetID INT, path TEXT, syncState INT,
    storageModTime INT, storageHash TEXT,
    lastProcessedModificationTime INT, lastRead INT);
CREATE TABLE itemNotes (itemID INTEGER PRIMARY KEY, parentItemID INT,
    note TEXT, title TEXT);
CREATE TABLE settings (setting TEXT, key TEXT, value);
CREATE TABLE users (userID INTEGER PRIMARY KEY, name TEXT);
CREATE TABLE groups (groupID INTEGER PRIMARY KEY, libraryID INT, name TEXT,
    description TEXT, version INT);
"""


def _make_zotero_db(path) -> str:
    con = sqlite3.connect(path)
    con.executescript(_SCHEMA)
    con.executemany(
        "INSERT INTO itemTypes (itemTypeID, typeName) VALUES (?, ?)",
        [(1, "annotation"), (3, "attachment"), (7, "book"),
         (22, "journalArticle"), (28, "note"), (34, "report")],
    )
    con.executemany(
        "INSERT INTO fields (fieldID, fieldName) VALUES (?, ?)",
        [(101, "title"), (102, "date"), (103, "publicationTitle"),
         (104, "volume"), (105, "pages"), (106, "publisher"),
         (107, "institution")],
    )
    con.executemany(
        "INSERT INTO creatorTypes (creatorTypeID, creatorType) VALUES (?, ?)",
        [(10, "author"), (12, "editor")],
    )
    con.execute(
        "INSERT INTO libraries (libraryID, type, editable, filesEditable) "
        "VALUES (1, 'user', 1, 1)"
    )
    con.execute(
        "INSERT INTO settings (setting, key, value) "
        "VALUES ('account', 'localUserKey', ?)",
        (LOCAL_USER_KEY,),
    )
    items = [
        (1, 22, 1, "AAAA1111"),  # journal article
        (2, 7, 1, "BBBB2222"),   # book
        (3, 22, 1, "CCCC3333"),  # trashed journal article
        (4, 3, 1, "DDDD4444"),   # attachment
        (5, 34, 1, "EEEE5555"),  # report, institutional author
    ]
    con.executemany(
        "INSERT INTO items (itemID, itemTypeID, libraryID, key) "
        "VALUES (?, ?, ?, ?)",
        items,
    )
    values = {
        1: [("title", "Signals and Standing in Alliance Politics"),
            ("date", "1999-00-00 1999"),
            ("publicationTitle", "Journal of Synthetic Examples"),
            ("volume", "41"), ("pages", "379-408")],
        2: [("title", "The Fabricated Order"),
            ("date", "2007-03-00 March 2007"),
            ("publisher", "Nonesuch Press")],
        3: [("title", "Trashed Signals Study"), ("date", "2001-00-00 2001")],
        5: [("title", "Annual Overview of Invented Metrics"),
            ("date", "2021-06-15 June 15, 2021"),
            ("institution", "Institute for Invented Studies")],
    }
    field_ids = dict(
        con.execute("SELECT fieldName, fieldID FROM fields").fetchall()
    )
    vid = 0
    for item_id, pairs in values.items():
        for fname, value in pairs:
            vid += 1
            con.execute(
                "INSERT INTO itemDataValues (valueID, value) VALUES (?, ?)",
                (vid, value),
            )
            con.execute(
                "INSERT INTO itemData (itemID, fieldID, valueID) "
                "VALUES (?, ?, ?)",
                (item_id, field_ids[fname], vid),
            )
    con.executemany(
        "INSERT INTO creators (creatorID, firstName, lastName, fieldMode) "
        "VALUES (?, ?, ?, ?)",
        [(1, "Dorian", "Vexley", 0),
         (2, "Marta", "Quenneville", 0),
         (3, "", "Institute for Invented Studies", 1)],
    )
    con.executemany(
        "INSERT INTO itemCreators (itemID, creatorID, creatorTypeID, "
        "orderIndex) VALUES (?, ?, ?, ?)",
        [(1, 1, 10, 0), (2, 2, 10, 0), (3, 1, 10, 0), (5, 3, 10, 0)],
    )
    con.execute("INSERT INTO deletedItems (itemID) VALUES (3)")
    con.execute(
        "INSERT INTO itemAttachments (itemID, parentItemID, contentType) "
        "VALUES (4, 1, 'application/pdf')"
    )
    con.commit()
    con.close()
    return str(path)


@pytest.fixture()
def zdb(tmp_path):
    return _make_zotero_db(tmp_path / "zotero.sqlite")


@pytest.fixture()
def cited_doc(tmp_path):
    """A doc with an anchor sentence, built through the server tools."""
    path = tmp_path / "cited.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": "Opening paragraph before the citation."},
         {"text": "The alliance argument rests on strong evidence."},
         {"text": "Closing paragraph after the citation."}],
        at_end=True,
        backup=False,
    )
    return str(path)


def _field_json(pkg: DocxPackage) -> dict:
    """Parse the CSL_CITATION JSON out of the (single) Zotero field."""
    fields, orphans = scan_complex_fields(pkg.root())
    assert not orphans
    zotero = [f for f in fields if "ZOTERO_ITEM" in f["instr"]]
    assert len(zotero) == 1
    instr = zotero[0]["instr"]
    prefix = "ADDIN ZOTERO_ITEM CSL_CITATION "
    assert instr.startswith(prefix)
    return json.loads(instr[len(prefix):])


# ------------------------------------------------------------ Zotero: search


def test_search_by_title(zdb):
    res = zoterolib.search_zotero_library("signals standing", db_path=zdb)
    assert res["total_matches"] == 1
    item = res["items"][0]
    assert item["key"] == "AAAA1111"
    assert item["itemType"] == "journalArticle"
    assert item["year"] == "1999"
    assert item["publication"] == "Journal of Synthetic Examples"
    assert item["creators"][0]["lastName"] == "Vexley"


def test_search_by_creator_and_year(zdb):
    res = zoterolib.search_zotero_library("vexley 1999", db_path=zdb)
    assert [i["key"] for i in res["items"]] == ["AAAA1111"]


def test_search_excludes_deleted_and_attachments(zdb):
    res = zoterolib.search_zotero_library("trashed signals", db_path=zdb)
    assert res["total_matches"] == 0
    res = zoterolib.search_zotero_library("signals", db_path=zdb)
    keys = {i["key"] for i in res["items"]}
    assert "CCCC3333" not in keys and "DDDD4444" not in keys


def test_search_limit(zdb):
    res = zoterolib.search_zotero_library("e", db_path=zdb, limit=1)
    assert res["returned"] == 1
    assert res["total_matches"] >= 2


def test_search_missing_db_refused(tmp_path):
    missing = tmp_path / "nowhere" / "zotero.sqlite"
    with pytest.raises(WordMcpError) as exc:
        zoterolib.search_zotero_library("anything", db_path=str(missing))
    assert str(missing) in str(exc.value)


def test_search_non_database_file_refused(tmp_path):
    junk = tmp_path / "junk.sqlite"
    junk.write_bytes(b"this is not a sqlite database at all")
    with pytest.raises(WordMcpError) as exc:
        zoterolib.search_zotero_library("anything", db_path=str(junk))
    assert "not a readable Zotero database" in str(exc.value)


def test_search_while_db_locked_by_writer(zdb):
    """The snapshot-copy path must keep working while another connection
    holds a write transaction on the database (Zotero running)."""
    holder = sqlite3.connect(zdb)
    holder.execute("BEGIN IMMEDIATE")
    holder.execute(
        "INSERT INTO itemDataValues (valueID, value) VALUES (999, 'x')"
    )
    try:
        res = zoterolib.search_zotero_library("vexley", db_path=zdb)
        assert res["total_matches"] == 1
    finally:
        holder.rollback()
        holder.close()


# ------------------------------------------------------------ Zotero: insert


def test_insert_reports_intact_via_reffields_oracle(cited_doc, zdb):
    pkg = DocxPackage(cited_doc)
    res = zoterolib.insert_zotero_citation(
        pkg, ["AAAA1111"], anchor_text="rests on", db_path=zdb
    )
    pkg.save(do_backup=False)
    assert res["field_intact"] is True
    assert res["placeholder_text"] == "(Vexley, 1999)"

    reopened = DocxPackage(cited_doc)
    inv = list_reference_fields(reopened)
    assert inv["total"] == 1
    field = inv["fields"][0]
    assert field["manager"] == "zotero"
    assert field["kind"] == "citation"
    assert field["intact"] is True
    assert field["has_cached_result"] is True
    assert field["cached_text"] == "(Vexley, 1999)"

    health = check_reference_field_integrity(reopened)
    assert health["ok"] is True
    assert health["citations"] == 1
    assert health["by_manager"] == {"zotero": 1}


def test_insert_field_json_shape(cited_doc, zdb):
    pkg = DocxPackage(cited_doc)
    res = zoterolib.insert_zotero_citation(
        pkg, ["AAAA1111"], anchor_text="rests on", db_path=zdb
    )
    pkg.save(do_backup=False)
    payload = _field_json(DocxPackage(cited_doc))

    assert payload["citationID"] == res["citation_id"]
    props = payload["properties"]
    assert props["formattedCitation"] == props["plainCitation"]
    assert props["noteIndex"] == 0
    assert payload["schema"].endswith("csl-citation.json")

    (ci,) = payload["citationItems"]
    assert ci["uris"] == [
        f"http://zotero.org/users/local/{LOCAL_USER_KEY}/items/AAAA1111"
    ]
    data = ci["itemData"]
    assert data["type"] == "article-journal"
    assert data["title"] == "Signals and Standing in Alliance Politics"
    assert data["container-title"] == "Journal of Synthetic Examples"
    assert data["volume"] == "41"
    assert data["page"] == "379-408"
    assert data["author"] == [{"family": "Vexley", "given": "Dorian"}]
    assert data["issued"] == {"date-parts": [[1999]]}


def test_insert_with_page_prefix_suffix(cited_doc, zdb):
    pkg = DocxPackage(cited_doc)
    res = zoterolib.insert_zotero_citation(
        pkg, ["AAAA1111"], anchor_text="rests on", db_path=zdb,
        page="45", prefix="see", suffix="; but compare",
    )
    pkg.save(do_backup=False)
    payload = _field_json(DocxPackage(cited_doc))
    (ci,) = payload["citationItems"]
    assert ci["locator"] == "45"
    assert ci["label"] == "page"
    assert ci["prefix"] == "see"
    assert ci["suffix"] == "; but compare"
    assert res["placeholder_text"] == "(see Vexley, 1999, p. 45; but compare)"


def test_insert_multiple_items(cited_doc, zdb):
    pkg = DocxPackage(cited_doc)
    res = zoterolib.insert_zotero_citation(
        pkg, ["AAAA1111", "BBBB2222"], anchor_text="rests on", db_path=zdb
    )
    pkg.save(do_backup=False)
    assert res["placeholder_text"] == "(Vexley, 1999; Quenneville, 2007)"
    payload = _field_json(DocxPackage(cited_doc))
    assert len(payload["citationItems"]) == 2
    assert payload["citationItems"][1]["itemData"]["type"] == "book"
    assert check_reference_field_integrity(DocxPackage(cited_doc))["ok"]


def test_insert_multiple_items_with_page_refused(cited_doc, zdb):
    pkg = DocxPackage(cited_doc)
    with pytest.raises(WordMcpError, match="single cited item"):
        zoterolib.insert_zotero_citation(
            pkg, ["AAAA1111", "BBBB2222"], anchor_text="rests on",
            db_path=zdb, page="45",
        )


def test_insert_unknown_or_trashed_key_refused(cited_doc, zdb):
    pkg = DocxPackage(cited_doc)
    with pytest.raises(TargetNotFound, match="ZZZZ9999"):
        zoterolib.insert_zotero_citation(
            pkg, ["ZZZZ9999"], anchor_text="rests on", db_path=zdb
        )
    with pytest.raises(TargetNotFound, match="CCCC3333"):
        zoterolib.insert_zotero_citation(
            pkg, ["CCCC3333"], anchor_text="rests on", db_path=zdb
        )


def test_insert_institutional_author(cited_doc, zdb):
    pkg = DocxPackage(cited_doc)
    res = zoterolib.insert_zotero_citation(
        pkg, ["EEEE5555"], anchor_text="rests on", db_path=zdb
    )
    pkg.save(do_backup=False)
    assert res["placeholder_text"] == (
        "(Institute for Invented Studies, 2021)"
    )
    payload = _field_json(DocxPackage(cited_doc))
    data = payload["citationItems"][0]["itemData"]
    assert data["author"] == [{"literal": "Institute for Invented Studies"}]
    assert data["type"] == "report"
    assert data["issued"] == {"date-parts": [[2021, 6, 15]]}


# --------------------------------------------------------------- style find


@pytest.fixture()
def styled_doc(tmp_path):
    path = tmp_path / "styled.docx"
    srv.create_document(str(path))
    srv.insert_paragraphs(
        str(path),
        [{"text": "The delta model requires sustained attention."},
         {"text": "Alpha beta gamma and beta again in one sentence."},
         {"text": "A completely plain closing paragraph."}],
        at_end=True,
        backup=False,
    )
    srv.add_heading(str(path), "Findings Overview", level=1, at_end=True,
                    backup=False)
    srv.format_text(str(path), {"bold": True}, find="delta model",
                    backup=False)
    srv.format_text(str(path), {"bold": True, "size_pt": 13},
                    find="sustained attention", backup=False)
    idx = next(
        p["index"] for p in srv.get_text(str(path))
        if "Alpha beta" in p["text"]
    )
    srv.format_text(str(path), {"italic": True}, paragraph_index=idx,
                    backup=False)
    return str(path)


def test_explicit_bold_hit(styled_doc):
    res = stylefind.find_formatted(
        DocxPackage(styled_doc), formatting={"bold": True}
    )
    texts = {m["text"] for m in res["matches"]}
    assert "delta model" in texts
    hit = next(m for m in res["matches"] if m["text"] == "delta model")
    assert hit["matched_via"]["bold"] == "explicit"
    assert hit["paragraph_index"] is not None
    assert hit["part"] == "word/document.xml"


def test_style_inherited_bold_hit(styled_doc):
    # The blank-document template ships Heading1 as bold, 14pt (sz 28).
    res = stylefind.find_formatted(
        DocxPackage(styled_doc), formatting={"bold": True, "size_pt": 14}
    )
    assert [m["text"] for m in res["matches"]] == ["Findings Overview"]
    via = res["matches"][0]["matched_via"]
    assert via["bold"] == "paragraph_style"
    assert via["size_pt"] == "paragraph_style"


def test_mixed_criteria_narrow_correctly(styled_doc):
    res = stylefind.find_formatted(
        DocxPackage(styled_doc), formatting={"bold": True, "size_pt": 13}
    )
    assert [m["text"] for m in res["matches"]] == ["sustained attention"]
    via = res["matches"][0]["matched_via"]
    assert via["bold"] == "explicit"
    assert via["size_pt"] == "explicit"


def test_query_within_formatting(styled_doc):
    res = stylefind.find_formatted(
        DocxPackage(styled_doc), "beta", formatting={"italic": True}
    )
    assert res["total"] == 2
    assert all(m["text"] == "beta" for m in res["matches"])
    assert all(
        m["matched_via"]["italic"] == "explicit" for m in res["matches"]
    )
    res = stylefind.find_formatted(
        DocxPackage(styled_doc), "delta", formatting={"italic": True}
    )
    assert res["total"] == 0


def test_style_criterion_by_id_and_name(styled_doc):
    for name in ("Heading1", "HEADING 1"):
        res = stylefind.find_formatted(
            DocxPackage(styled_doc), formatting={"style": name}
        )
        assert [m["text"] for m in res["matches"]] == ["Findings Overview"]
        assert res["matches"][0]["matched_via"]["style"] == "paragraph_style"


def test_highlight_absent_then_found(styled_doc):
    res = stylefind.find_formatted(
        DocxPackage(styled_doc), formatting={"highlight": "yellow"}
    )
    assert res["total"] == 0
    srv.format_text(styled_doc, {"highlight": "yellow"},
                    find="plain closing", backup=False)
    res = stylefind.find_formatted(
        DocxPackage(styled_doc), formatting={"highlight": "yellow"}
    )
    assert [m["text"] for m in res["matches"]] == ["plain closing"]
    assert res["matches"][0]["matched_via"]["highlight"] == "explicit"


def test_unknown_key_and_bad_inputs_rejected(styled_doc):
    pkg = DocxPackage(styled_doc)
    with pytest.raises(WordMcpError) as exc:
        stylefind.find_formatted(pkg, formatting={"boldface": True})
    assert "allowed" in str(exc.value) and "bold" in str(exc.value)
    with pytest.raises(WordMcpError):
        stylefind.find_formatted(pkg, formatting={})
    with pytest.raises(WordMcpError, match="scope"):
        stylefind.find_formatted(
            pkg, formatting={"bold": True}, scope="everything"
        )
    with pytest.raises(WordMcpError):
        stylefind.find_formatted(pkg, "", formatting={"bold": True})
    with pytest.raises(WordMcpError, match="true or false"):
        stylefind.find_formatted(pkg, formatting={"bold": "yes"})


def test_bold_false_matches_unbold_text(styled_doc):
    res = stylefind.find_formatted(
        DocxPackage(styled_doc), "plain closing", formatting={"bold": False}
    )
    assert res["total"] == 1
    assert res["matches"][0]["matched_via"]["bold"] == "absent_default_off"
