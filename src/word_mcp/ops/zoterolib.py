"""Read the user's LOCAL Zotero library and write plugin-compatible citations.

Two capabilities:

- search_zotero_library: query the local zotero.sqlite (title, creator last
  names, year, publication) and return matching items with their Zotero keys.
- insert_zotero_citation: build a real `ADDIN ZOTERO_ITEM CSL_CITATION {json}`
  complex field — the exact structure the Zotero Word plugin writes — so the
  citation is recognized, refreshable, and re-styleable by Zotero itself.

Database safety contract (non-negotiable):
- The Zotero database is NEVER written to and NEVER opened read-write. Every
  read goes through a temporary copy of the file opened with
  `mode=ro&immutable=1`, so even SQLite's side files (-wal/-shm) are never
  created next to the real database. If the copy cannot be made, a direct
  `mode=ro` connection is the fallback.
- Copying while Zotero is running means the snapshot can be a few moments
  stale (changes still in Zotero's WAL or made after the copy are not seen).
  That is the accepted tradeoff for never touching the live file; each call
  takes a fresh snapshot, so staleness is bounded to the single call.

Honesty notes:
- The inserted citation's visible text is a locally generated (Author, Year)
  PLACEHOLDER. It becomes a properly styled citation the next time the user
  clicks Refresh in Zotero's Word plugin; until then it may not match the
  document's citation style.
- itemData embedded in the field is a best-effort CSL-JSON conversion of the
  item's metadata (used by Zotero only when the item cannot be resolved from
  the library). Resolution normally happens through the item URI, which is
  always exact.

Verified against a live Zotero 7 database (userdata schema version 125):
tables items / itemData / itemDataValues / fields / creators / itemCreators /
creatorTypes / itemTypes / deletedItems / itemAttachments / itemNotes /
settings / users / libraries / groups, and the localUserKey stored at
settings(setting='account', key='localUserKey').
"""

from __future__ import annotations

import json
import re
import secrets
import shutil
import sqlite3
import tempfile
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import quote

from lxml import etree

from ..core.errors import TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from ..core.sandbox import check_path
from . import _runmap
from .fields import _find_anchor_span

# ------------------------------------------------------------------ DB access

_EXCLUDED_TYPE_NAMES = ("attachment", "note", "annotation")


def _default_db_path() -> Path:
    """Zotero's default data directory is <home>/Zotero on all platforms."""
    return Path.home() / "Zotero" / "zotero.sqlite"


def _connect_uri(path: Path, *, immutable: bool) -> sqlite3.Connection:
    uri = "file:" + quote(path.as_posix(), safe="/:")
    uri += "?mode=ro" + ("&immutable=1" if immutable else "")
    return sqlite3.connect(uri, uri=True)


@contextmanager
def _open_db(db_path: str | None):
    """Yield a read-only connection to a snapshot of the Zotero database.

    The real database file is never opened for writing and never has SQLite
    side files created next to it: we copy it to a temp directory and open
    the copy immutable. Direct read-only open is the fallback if the copy
    fails. Missing/invalid databases raise WordMcpError with the searched
    path so the caller can point at a nonstandard location via db_path.
    """
    # An explicitly supplied db_path is caller-controlled and sandbox-gated;
    # the auto-discovered default is a fixed well-known location and exempt
    # (otherwise enabling KS4W_ALLOWED_ROOTS would break Zotero search).
    if db_path:
        check_path(db_path, "read Zotero database")
    path = Path(db_path) if db_path else _default_db_path()
    if not path.exists():
        raise WordMcpError(
            f"Zotero database not found at {path}. Zotero stores it at "
            "<home>/Zotero/zotero.sqlite by default; if your data directory "
            "is elsewhere, pass db_path explicitly."
        )
    tmpdir = Path(tempfile.mkdtemp(prefix="word-mcp-zotero-"))
    con: sqlite3.Connection | None = None
    try:
        snapshot = tmpdir / "zotero-snapshot.sqlite"
        try:
            shutil.copy2(path, snapshot)
            con = _connect_uri(snapshot, immutable=True)
        except OSError:
            # Copy failed (permissions, disk); read the real file, still ro.
            con = _connect_uri(path, immutable=False)
        try:
            con.execute("SELECT itemID FROM items LIMIT 1")
        except sqlite3.Error as exc:
            raise WordMcpError(
                f"{path} is not a readable Zotero database "
                f"(no items table): {exc}"
            ) from exc
        yield con
    finally:
        if con is not None:
            con.close()
        shutil.rmtree(tmpdir, ignore_errors=True)


# ----------------------------------------------------------- item extraction

_CONTAINER_FIELDS = (
    "publicationTitle",
    "bookTitle",
    "proceedingsTitle",
    "encyclopediaTitle",
    "dictionaryTitle",
    "websiteTitle",
    "blogTitle",
)


def _field_values(con: sqlite3.Connection, item_id: int) -> dict[str, str]:
    rows = con.execute(
        "SELECT f.fieldName, v.value FROM itemData d "
        "JOIN fields f ON f.fieldID = d.fieldID "
        "JOIN itemDataValues v ON v.valueID = d.valueID "
        "WHERE d.itemID = ?",
        (item_id,),
    ).fetchall()
    return {name: str(value) for name, value in rows}


def _creators(con: sqlite3.Connection, item_id: int) -> list[dict]:
    rows = con.execute(
        "SELECT c.lastName, c.firstName, c.fieldMode, ct.creatorType "
        "FROM itemCreators ic "
        "JOIN creators c ON c.creatorID = ic.creatorID "
        "JOIN creatorTypes ct ON ct.creatorTypeID = ic.creatorTypeID "
        "WHERE ic.itemID = ? ORDER BY ic.orderIndex",
        (item_id,),
    ).fetchall()
    return [
        {
            "lastName": last or "",
            "firstName": first or "",
            "fieldMode": mode or 0,
            "creatorType": ctype,
        }
        for last, first, mode, ctype in rows
    ]


def _year_of(date_value: str | None) -> str | None:
    if not date_value:
        return None
    m = re.search(r"\b(\d{4})\b", date_value)
    return m.group(1) if m else None


def _load_items(con: sqlite3.Connection, *, keys: list[str] | None = None):
    """Regular (non-deleted, non-attachment/note/annotation) items with
    metadata. keys=None loads the whole library."""
    sql = (
        "SELECT i.itemID, i.key, t.typeName, i.libraryID FROM items i "
        "JOIN itemTypes t ON t.itemTypeID = i.itemTypeID "
        "WHERE t.typeName NOT IN ({placeholders}) "
        "AND i.itemID NOT IN (SELECT itemID FROM deletedItems)"
    ).format(placeholders=",".join("?" * len(_EXCLUDED_TYPE_NAMES)))
    params: list = list(_EXCLUDED_TYPE_NAMES)
    if keys is not None:
        sql += " AND i.key IN ({})".format(",".join("?" * len(keys)))
        params += keys
    items = []
    for item_id, key, type_name, library_id in con.execute(sql, params):
        fields = _field_values(con, item_id)
        creators = _creators(con, item_id)
        publication = next(
            (fields[f] for f in _CONTAINER_FIELDS if fields.get(f)), None
        )
        items.append(
            {
                "itemID": item_id,
                "key": key,
                "itemType": type_name,
                "libraryID": library_id,
                "title": fields.get("title", ""),
                "creators": creators,
                "year": _year_of(fields.get("date")),
                "publication": publication,
                "fields": fields,
            }
        )
    return items


# ------------------------------------------------------------------- search


def search_zotero_library(
    query: str, *, db_path: str | None = None, limit: int = 20
) -> dict:
    """Search the user's local Zotero library (read-only).

    Every whitespace-separated token in `query` must appear (case-insensitive
    substring) in the item's title, a creator's last name, the year, or the
    publication title. Attachments, notes, annotations, and trashed items are
    excluded. Returns the Zotero item key per match — the handle
    insert_zotero_citation needs."""
    if not query or not query.strip():
        raise WordMcpError("query must be a non-empty string")
    if limit < 1:
        raise WordMcpError("limit must be >= 1")
    tokens = [t.lower() for t in query.split()]
    with _open_db(db_path) as con:
        items = _load_items(con)
    matches = []
    for it in items:
        hay = " ".join(
            [it["title"]]
            + [c["lastName"] for c in it["creators"]]
            + [it["year"] or "", it["publication"] or ""]
        ).lower()
        if all(t in hay for t in tokens):
            matches.append(it)
    matches.sort(
        key=lambda it: (query.lower() not in it["title"].lower(), it["title"].lower())
    )
    out = [
        {
            "key": it["key"],
            "itemType": it["itemType"],
            "title": it["title"],
            "creators": [
                {
                    "lastName": c["lastName"],
                    "firstName": c["firstName"],
                    "creatorType": c["creatorType"],
                }
                for c in it["creators"]
            ],
            "year": it["year"],
            "publication": it["publication"],
        }
        for it in matches[:limit]
    ]
    return {
        "query": query,
        "total_matches": len(matches),
        "returned": len(out),
        "items": out,
    }


# ------------------------------------------- CSL-JSON conversion for itemData

# Zotero itemType -> CSL type (subset of Zotero's own mapping; unlisted types
# fall back to "document", which citeproc accepts for any item).
_CSL_TYPES = {
    "journalArticle": "article-journal",
    "magazineArticle": "article-magazine",
    "newspaperArticle": "article-newspaper",
    "book": "book",
    "bookSection": "chapter",
    "conferencePaper": "paper-conference",
    "thesis": "thesis",
    "report": "report",
    "webpage": "webpage",
    "blogPost": "post-weblog",
    "forumPost": "post",
    "manuscript": "manuscript",
    "interview": "interview",
    "letter": "personal_communication",
    "email": "personal_communication",
    "presentation": "speech",
    "patent": "patent",
    "statute": "legislation",
    "case": "legal_case",
    "bill": "bill",
    "map": "map",
    "computerProgram": "software",
    "encyclopediaArticle": "entry-encyclopedia",
    "dictionaryEntry": "entry-dictionary",
    "dataset": "dataset",
    "standard": "standard",
    "preprint": "article",
    "document": "document",
}

# Zotero field -> CSL variable (subset; unmapped fields are omitted from
# itemData — Zotero resolves the item through its URI, so nothing is lost
# for the library owner).
_CSL_FIELDS = {
    "title": "title",
    "publicationTitle": "container-title",
    "bookTitle": "container-title",
    "proceedingsTitle": "container-title",
    "encyclopediaTitle": "container-title",
    "dictionaryTitle": "container-title",
    "websiteTitle": "container-title",
    "blogTitle": "container-title",
    "seriesTitle": "collection-title",
    "series": "collection-title",
    "publisher": "publisher",
    "institution": "publisher",
    "university": "publisher",
    "place": "publisher-place",
    "volume": "volume",
    "issue": "issue",
    "pages": "page",
    "DOI": "DOI",
    "url": "URL",
    "ISBN": "ISBN",
    "ISSN": "ISSN",
    "edition": "edition",
    "shortTitle": "shortTitle",
    "abstractNote": "abstract",
    "language": "language",
    "numPages": "number-of-pages",
    "reportNumber": "number",
}

_CSL_ROLES = {
    "author": "author",
    "editor": "editor",
    "translator": "translator",
    "bookAuthor": "container-author",
    "seriesEditor": "collection-editor",
    "reviewedAuthor": "reviewed-author",
    "director": "director",
    "interviewer": "interviewer",
    "recipient": "recipient",
    "composer": "composer",
}


def _issued_of(date_value: str | None) -> dict | None:
    if not date_value:
        return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", date_value)
    if m:
        parts = [int(m.group(1))]
        if m.group(2) != "00":
            parts.append(int(m.group(2)))
            if m.group(3) != "00":
                parts.append(int(m.group(3)))
        if parts[0] != 0:
            return {"date-parts": [parts]}
    year = _year_of(date_value)
    return {"date-parts": [[int(year)]]} if year else None


def _csl_item_data(item: dict) -> dict:
    data: dict = {
        "id": item["itemID"],
        "type": _CSL_TYPES.get(item["itemType"], "document"),
    }
    for zfield, cslvar in _CSL_FIELDS.items():
        val = item["fields"].get(zfield)
        if val and cslvar not in data:
            data[cslvar] = val
    for c in item["creators"]:
        role = _CSL_ROLES.get(c["creatorType"])
        if role is None:
            continue  # role has no CSL equivalent; URI resolution covers it
        if c["fieldMode"] == 1:  # single-field (institutional) name
            name = {"literal": c["lastName"]}
        else:
            name = {"family": c["lastName"], "given": c["firstName"]}
        data.setdefault(role, []).append(name)
    issued = _issued_of(item["fields"].get("date"))
    if issued:
        data["issued"] = issued
    return data


def _item_uri(con: sqlite3.Connection, item: dict) -> str:
    """The URI Zotero uses to resolve the citation back to the library item.

    Synced accounts use the numeric userID; unsynced installs use the
    localUserKey (settings table, setting='account'). Group-library items use
    the groupID. Verified against a live Zotero 7 database and real plugin
    field codes."""
    lib_type = con.execute(
        "SELECT type FROM libraries WHERE libraryID = ?",
        (item["libraryID"],),
    ).fetchone()
    if lib_type and lib_type[0] == "group":
        row = con.execute(
            "SELECT groupID FROM groups WHERE libraryID = ?",
            (item["libraryID"],),
        ).fetchone()
        if row is None:
            raise WordMcpError(
                f"item {item['key']} is in a group library with no groups "
                "entry; cannot build its citation URI"
            )
        return f"http://zotero.org/groups/{row[0]}/items/{item['key']}"
    row = con.execute("SELECT userID FROM users LIMIT 1").fetchone()
    if row is not None:
        return f"http://zotero.org/users/{row[0]}/items/{item['key']}"
    row = con.execute(
        "SELECT value FROM settings WHERE setting='account' "
        "AND key='localUserKey'"
    ).fetchone()
    if row is None:
        raise WordMcpError(
            "Zotero database has neither a synced userID nor a localUserKey; "
            "cannot build citation URIs this Zotero install would recognize"
        )
    return f"http://zotero.org/users/local/{row[0]}/items/{item['key']}"


# ---------------------------------------------------------------- insertion

_CITATION_SCHEMA = (
    "https://github.com/citation-style-language/schema/raw/master/"
    "csl-citation.json"
)


def _placeholder_segment(item: dict, page: str | None) -> str:
    if item["creators"]:
        first = item["creators"][0]
        name = first["lastName"] or first["firstName"] or "Anon."
        if len(item["creators"]) > 1:
            name += " et al."
    else:
        name = item["title"].split(":")[0][:40] or "Untitled"
    seg = f"{name}, {item['year'] or 'n.d.'}"
    if page:
        seg += f", p. {page}"
    return seg


def _citation_field_runs(instr: str, cached_text: str) -> list[etree._Element]:
    """begin / instrText / separate / cached result / end — the run shape the
    Zotero plugin writes (no w:dirty: ADDIN fields are owned by the add-in,
    not recomputed by Word)."""
    els = []
    r1 = etree.Element(qn("w:r"))
    etree.SubElement(r1, qn("w:fldChar")).set(qn("w:fldCharType"), "begin")
    els.append(r1)
    r2 = etree.Element(qn("w:r"))
    it = etree.SubElement(r2, qn("w:instrText"))
    it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    it.text = f" {instr} "
    els.append(r2)
    r3 = etree.Element(qn("w:r"))
    etree.SubElement(r3, qn("w:fldChar")).set(qn("w:fldCharType"), "separate")
    els.append(r3)
    r4 = etree.Element(qn("w:r"))
    t = etree.SubElement(r4, qn("w:t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = cached_text
    els.append(r4)
    r5 = etree.Element(qn("w:r"))
    etree.SubElement(r5, qn("w:fldChar")).set(qn("w:fldCharType"), "end")
    els.append(r5)
    return els


def insert_zotero_citation(
    pkg: DocxPackage,
    item_keys: list[str],
    *,
    anchor_text: str,
    occurrence: int = 1,
    page: str | None = None,
    prefix: str | None = None,
    suffix: str | None = None,
    db_path: str | None = None,
) -> dict:
    """Insert a Zotero-plugin-compatible citation field after `anchor_text`.

    `item_keys` are Zotero item keys (from search_zotero_library). The field
    is a genuine `ADDIN ZOTERO_ITEM CSL_CITATION {json}` complex field with
    the item URIs and CSL-JSON metadata the plugin expects, so Zotero's Word
    plugin recognizes it, restyles it, and includes it in the bibliography.

    The visible text is a locally generated (Author, Year) placeholder: it is
    NOT rendered in the document's citation style until the user clicks
    Refresh in Zotero's Word plugin. page becomes the citation's locator;
    prefix/suffix are the Zotero prefix/suffix of the cited item. With more
    than one item key, page/prefix/suffix are refused (Zotero applies them
    per item; insert separate citations instead of letting this tool guess).

    The Zotero database itself is only ever read (see module docstring)."""
    if not item_keys:
        raise WordMcpError("item_keys must contain at least one Zotero key")
    if len(item_keys) > 1 and (page or prefix or suffix):
        raise WordMcpError(
            "page/prefix/suffix apply to a single cited item; with multiple "
            "item_keys Zotero attaches them per item — insert one citation "
            "per item instead"
        )
    page = str(page) if page is not None else None

    with _open_db(db_path) as con:
        found = {it["key"]: it for it in _load_items(con, keys=item_keys)}
        missing = [k for k in item_keys if k not in found]
        if missing:
            raise TargetNotFound(
                f"Zotero item key(s) not found in the library (or the item "
                f"is trashed / an attachment): {', '.join(missing)}"
            )
        items = [found[k] for k in item_keys]  # caller's order
        uris = {it["key"]: _item_uri(con, it) for it in items}

    citation_items = []
    for it in items:
        entry: dict = {
            "id": it["itemID"],
            "uris": [uris[it["key"]]],
            "itemData": _csl_item_data(it),
        }
        if page:
            entry["locator"] = page
            entry["label"] = "page"
        if prefix:
            entry["prefix"] = prefix
        if suffix:
            entry["suffix"] = suffix
        citation_items.append(entry)

    inner = "; ".join(
        _placeholder_segment(it, page if len(items) == 1 else None)
        for it in items
    )
    if prefix:
        inner = f"{prefix.strip()} {inner}"
    if suffix:
        inner = f"{inner}{suffix}"
    placeholder = f"({inner})"

    citation_id = "".join(
        secrets.choice("abcdefghijklmnopqrstuvwxyz0123456789")
        for _ in range(10)
    )
    payload = {
        "citationID": citation_id,
        "properties": {
            "formattedCitation": placeholder,
            "plainCitation": placeholder,
            "noteIndex": 0,
        },
        "citationItems": citation_items,
        "schema": _CITATION_SCHEMA,
    }
    instr = "ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(
        payload, ensure_ascii=False
    )

    p, _, end = _find_anchor_span(pkg, anchor_text, occurrence)
    covered = _runmap.split_for_range(p, end - 1, end)
    ref = covered[-1]
    for el in reversed(_citation_field_runs(instr, placeholder)):
        ref.addnext(el)
    pkg.mark_dirty()

    # Self-check against the reference-field scanner: the field we just wrote
    # must inventory as an intact Zotero citation, or we refuse to pretend.
    from .reffields import list_reference_fields

    inv = list_reference_fields(pkg)
    ours = [
        f
        for f in inv["fields"]
        if f["manager"] == "zotero" and f["kind"] == "citation"
        and f["cached_text"] == placeholder and f["intact"]
    ]
    if not ours:
        raise WordMcpError(
            "internal check failed: the inserted field did not inventory as "
            "an intact Zotero citation; the document was not saved"
        )

    return {
        "inserted": True,
        "citation_id": citation_id,
        "items": [
            {"key": it["key"], "title": it["title"], "year": it["year"]}
            for it in items
        ],
        "placeholder_text": placeholder,
        "field_intact": True,
        "note": (
            "renders as a plain (Author, Year) placeholder until Refresh is "
            "clicked in Zotero's Word plugin"
        ),
    }
