"""Publication-style conversion for plain-text citations (v1.5 headliner).

Converts academic manuscripts between the major publication styles:
APA 7, Chicago 17 (author-date and notes-bibliography, the latter being
Turabian for practical purposes), MLA 9, Harvard (generic), IEEE, Vancouver,
and ASA (Sage). Three stages:

1. parse_references()      — locate and parse the reference list and in-text
                             citations into a structured model (read-only).
2. convert_citation_style() — re-emit the reference list in the target style
                             and convert in-text citations between systems
                             (author-date <-> numbered <-> notes).
3. apply_manuscript_format() — page-level conventions (margins, spacing,
                             running heads) for the styles that define them.

HONESTY CONTRACT (this is heuristic TEXT conversion, not a citation
processor):
- Every reference entry gets a parse_confidence (full/partial/failed).
  Only FULL entries are ever rewritten. Partial and failed entries are
  returned/left VERBATIM and flagged — never guessed at, never half-
  converted.
- In-text citations are converted only when they resolve unambiguously to a
  fully-parsed entry; everything else is left in place and flagged.
- Documents using Word-native CITATION fields or reference-manager (Zotero/
  Mendeley/EndNote) fields are ROUTED, not rewritten: field-based citations
  must be restyled by their owning tool, or the field data is destroyed.
- The flagged list in every result is part of the output, not decoration:
  a conversion is not done until a human has reviewed it.

Builds on ops/citecheck.py's proven in-text patterns; all text surgery goes
through ops/_runmap.py (runmap-safe); footnotes are real footnotes via the
ops/notes.py machinery.
"""

from __future__ import annotations

import re

from lxml import etree

from ..core.errors import TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from . import _runmap, notes as _notes
from .citecheck import _REF_HEADINGS, _YEAR
from .read import body_items, get_outline, paragraph_text
from .reffields import _classify_manager, scan_complex_fields
from .styleconvert_data import (
    STYLE_INFO,
    canonical_style,
    emit_entry,
    emit_intext_item,
    emit_intext_year_only,
    emit_note,
    norm_pages,
)

# ------------------------------------------------------------------ patterns
# Derived from ops/citecheck.py's tuned patterns, extended with locator
# (page) capture and multi-style year forms.

# Latin capitalized (accents included) OR hangul author tokens; Korean uses
# 와/과 connectors. (_YEAR imported from citecheck already accepts "n.d.".)
_AUTHOR_TOKEN = r"(?:[A-ZÀ-ÞĀ-Žƀ-Ƀ][\w'’\-]+|[가-힣]+)"
_AUTHOR_SEQ = (
    rf"{_AUTHOR_TOKEN}(?:\s+(?:and|&|와|과)\s+{_AUTHOR_TOKEN})*"
)

_PAREN = re.compile(r"\(([^()]{2,300})\)")

_ITEM_AD = re.compile(  # author-date item: APA / Chicago AD / Harvard / ASA
    rf"^(?P<prefix>(?:e\.g\.,?|cf\.,?|see(?:\s+also)?)\s+)?"
    rf"(?P<names>{_AUTHOR_SEQ})"
    rf"(?P<etal>,?\s+et al\.?)?"
    rf",?\s+(?P<year>{_YEAR})"
    rf"(?:\s*[,:]\s*(?:pp?\.\s*)?(?P<loc>\d[\d\s,\-–—]*))?\s*$"
)
_ITEM_AP = re.compile(  # author-page item (MLA): Hurd 24 / Hurd and Lake 24-28
    rf"^(?P<names>{_AUTHOR_SEQ})(?P<etal>\s+et al\.?)?"
    rf"\s+(?P<loc>\d{{1,4}}(?:[\-–]\d{{1,4}})?)\s*$"
)
_NARRATIVE = re.compile(  # Hurd (1999) / Hurd et al. (1999, p. 24)
    rf"(?<![\w가-힣])(?P<names>{_AUTHOR_SEQ})(?P<etal>\s+et al\.?)?\s*"
    rf"(?P<paren>\((?P<year>{_YEAR})"
    rf"(?:,\s*(?:pp?\.\s*)?(?P<loc>\d[\d\s,\-–—]*))?\))"
)
_BRACKET = re.compile(
    r"\[(?P<nums>\d+(?:\s*[,;–\-]\s*\d+)*)"
    r"(?:,\s*pp?\.\s*(?P<loc>[\d\-–]+))?\]"
)
_SUPERSCRIPT_TEXT = re.compile(r"^[\d,\s\-–]+$")

_ENTRY_NUM = re.compile(r"^\s*(?:\[(\d+)\]|(\d+)\.)\s+")
_DOI_URL = re.compile(r"(https?://[^\s]+|doi:\s*[^\s]+)\.?\s*$", re.I)
_SENT_END = re.compile(r"[.!?]")


# =========================================================== reference list


def _locate_reference_list(pkg: DocxPackage) -> tuple[int, int | None]:
    """(heading paragraph index, exclusive end index or None) — mirrors
    ops/citecheck.py's location logic."""
    outline = get_outline(pkg)
    ref_heading = next(
        (h for h in outline if _REF_HEADINGS.match(h["text"])), None
    )
    if ref_heading is None:
        candidates = [
            idx
            for kind, idx, el in body_items(pkg)
            if kind == "paragraph" and _REF_HEADINGS.match(paragraph_text(el) or "")
        ]
        if not candidates:
            raise TargetNotFound(
                "no References/Bibliography/Works Cited heading found; "
                "cannot locate the reference list"
            )
        ref_start = candidates[-1]
    else:
        ref_start = ref_heading["paragraph_index"]
    next_headings = [
        h["paragraph_index"] for h in outline if h["paragraph_index"] > ref_start
    ]
    return ref_start, (min(next_headings) if next_headings else None)


def _body_paragraph_elements(pkg: DocxPackage) -> list:
    """[(index, element)] for top-level body paragraphs (keep the list alive
    while using the elements)."""
    return [
        (idx, el) for kind, idx, el in body_items(pkg) if kind == "paragraph"
    ]


def _italic_spans(p: etree._Element) -> list[str]:
    """Consecutive-italic-run text chunks of a paragraph (parser hints)."""
    spans: list[str] = []
    current = ""
    for r in p.iter(qn("w:r")):
        if _runmap._in_deleted(r):
            continue
        rpr = r.find(qn("w:rPr"))
        italic = False
        if rpr is not None:
            i = rpr.find(qn("w:i"))
            if i is not None and i.get(qn("w:val")) not in ("0", "false", "none"):
                italic = True
        text = "".join(t.text or "" for t in r.findall(qn("w:t")))
        if italic and text:
            current += text
        else:
            if current.strip():
                spans.append(current.strip())
            current = ""
    if current.strip():
        spans.append(current.strip())
    return spans


# ------------------------------------------------------------ author parsing

_GIVEN_INITIALS = re.compile(r"^(?:[A-Z]\.(?:[\s\-]?[A-Z]\.)*)$")
_NAT_INITIALS = re.compile(
    r"^(?P<ini>(?:[A-Z]\.[\s\-]?)+)(?P<fam>[A-Z][\w'’\-]+)$"
)
_VANC_UNIT = re.compile(r"^(?P<fam>[A-Z][\w'’\-]+)\s+(?P<ini>[A-Z]{1,3})$")
_FAMILY = r"[A-Z][\w'’\-]+(?:\s+[A-Z][\w'’\-]+)?"
_INV_PAIR = re.compile(
    rf"(?P<fam>{_FAMILY}),\s*"
    r"(?P<giv>(?:[A-Z]\.(?:[\s\-]?[A-Z]\.)*)|[A-Z][\w'’\-]+(?:\s+[A-Z]\.(?:\s?[A-Z]\.)*)?)"
)


def _parse_single_author(seg: str) -> dict:
    seg = seg.strip().rstrip(",")
    m = _NAT_INITIALS.match(seg)
    if m:  # IEEE: I. Hurd
        return {"family": m.group("fam"), "given": m.group("ini").strip()}
    m = _VANC_UNIT.match(seg)
    if m:  # Vancouver: Hurd IA
        ini = " ".join(f"{c}." for c in m.group("ini"))
        return {"family": m.group("fam"), "given": ini}
    if "," in seg:
        fam, giv = seg.split(",", 1)
        giv = giv.strip()
        # Normalize bare/partial initials to dotted form ("I" -> "I.",
        # "D. A" -> "D. A.") so downstream emitters render them correctly.
        if re.fullmatch(r"[A-Z](?:[.\s\-]*[A-Z])*\.?", giv):
            letters = re.findall(r"[A-Z]", giv)
            giv = " ".join(f"{c}." for c in letters)
        return {"family": fam.strip(), "given": giv}
    words = seg.split()
    if len(words) >= 2 and all(w[0].isupper() for w in words if w):
        # natural full: Ian Hurd / Jean-Paul B. Sartre
        return {"family": words[-1], "given": " ".join(words[:-1])}
    return {"literal": seg}


def _parse_authors(s: str) -> tuple[list[dict], bool, bool]:
    """(authors, et_al, clean_parse). Never invents names: on failure the
    whole string becomes a literal author with clean_parse False."""
    s = s.strip().rstrip(".").rstrip(",").strip()
    if not s:
        return [], False, False
    et_al = False
    m = re.search(r",?\s+et al\.?$", s)
    if m:
        et_al = True
        s = s[: m.start()].rstrip(",. ")

    # Vancouver whole-string form: Hurd I, Lake DA, Walt SM
    if re.fullmatch(
        r"[A-Z][\w'’\-]+\s+[A-Z]{1,3}(,\s*[A-Z][\w'’\-]+\s+[A-Z]{1,3})*",
        s,
    ):
        return [
            _parse_single_author(seg) for seg in s.split(",")
        ], et_al, True

    # Normalize explicit separators to |
    work = re.sub(r",?\s+(?:and|&)\s+", "|", s)
    work = work.replace("; ", "|")
    segs = [seg.strip() for seg in work.split("|") if seg.strip()]
    authors: list[dict] = []
    ok = True
    for seg in segs:
        # A segment may hold several inverted-initials authors (APA):
        # "Hurd, I., Lake, D. A."
        pairs = list(_INV_PAIR.finditer(seg))
        covered = sum(m.end() - m.start() for m in pairs)
        if len(pairs) >= 2 and covered >= len(seg) * 0.8:
            for m in pairs:
                authors.append(
                    {"family": m.group("fam"), "given": m.group("giv")}
                )
            continue
        a = _parse_single_author(seg)
        if "literal" in a and (" " in a["literal"] and "," not in seg):
            # Might be an organization — keep literal but mark unclean.
            ok = ok and a["literal"][0].isupper()
        authors.append(a)
    if not authors:
        return [{"literal": s}], et_al, False
    return authors, et_al, ok


# ------------------------------------------------------------- entry parsing


def _split_place_publisher(s: str) -> tuple[str | None, str | None]:
    s = s.strip().rstrip(".")
    if not s:
        return None, None
    if ": " in s:
        place, pub = s.split(": ", 1)
        return place.strip(), pub.strip()
    return None, s


def _finish(entry: dict) -> dict:
    """Assign parse_confidence."""
    has_authors = bool(entry.get("authors"))
    literal_only = has_authors and all(
        "literal" in a for a in entry["authors"]
    )
    has_core = has_authors and entry.get("year") and entry.get("title")
    has_venue = (
        entry.get("container")
        or entry.get("publisher")
        or entry.get("url")
        or entry.get("doi")
    )
    typed = entry.get("type") in ("article", "book", "chapter", "report", "web")
    if has_core and has_venue and typed and not literal_only:
        entry["parse_confidence"] = "full"
    elif has_core:
        entry["parse_confidence"] = "partial"
        if literal_only:
            entry.setdefault("flags", []).append(
                "author parsed as a literal string (organization or "
                "unrecognized name format)"
            )
    else:
        entry["parse_confidence"] = "failed"
    return entry


_APA_HEAD = re.compile(
    rf"^(?P<auth>[^()]{{1,250}}?)\s*\((?P<year>{_YEAR})(?P<yextra>[^)]*)\)"
    rf"\s*(?P<punct>[.,]?)\s*(?P<rest>.+)$",
    re.S,
)
_APA_ARTICLE_TAIL = re.compile(
    r"^(?P<cont>.+?),\s*(?P<vol>\d+)\s*(?:\((?P<iss>[^)]+)\))?"
    r"(?:,\s*(?P<pp>[\dexvi\-–—,\s]+?))?\.?\s*$"
)
_QUOTED = re.compile(r"[“\"](?P<t>[^”\"]{3,300}?)[,.]?[”\"]")
_CHI_YEAR_TAIL = re.compile(rf"[.,]?\s*(?P<year>{_YEAR})\.\s*$")
_NB_TAIL = re.compile(
    rf"^(?P<cont>.+?)\s+(?P<vol>\d+)(?:,\s*no\.\s*(?P<iss>\w+))?\s*"
    rf"\((?P<year>{_YEAR})\)(?::\s*(?P<pp>[\d\-–]+))?\.?\s*$"
)
_MLA_TAIL = re.compile(
    rf"^(?P<cont>.+?),\s*vol\.\s*(?P<vol>\d+)(?:,\s*no\.\s*(?P<iss>\w+))?"
    rf",\s*(?P<year>{_YEAR})(?:,\s*pp?\.\s*(?P<pp>[\d\-–]+))?\.?\s*$"
)
_IEEE_TAIL = re.compile(
    rf"^(?P<cont>.+?),\s*vol\.\s*(?P<vol>\d+)(?:,\s*no\.\s*(?P<iss>\w+))?"
    rf"(?:,\s*pp?\.\s*(?P<pp>[\d\-–]+))?,\s*(?P<year>{_YEAR})\.?\s*$"
)
_CHIAD_TAIL = re.compile(
    rf"^(?P<cont>.+?)\s+(?P<vol>\d+)\s*(?:\(\s*(?P<iss>\d+)\s*\))?\s*:"
    rf"\s*(?P<pp>[\d\-–]+)\.?\s*$"
)
_VANC_ENTRY = re.compile(
    rf"^(?P<auth>[^.]+?)\.\s+(?P<title>[^.]+?)\.\s+(?P<cont>[^.]+?)\.\s+"
    rf"(?P<year>{_YEAR})\s*;\s*(?P<vol>\d+)(?:\((?P<iss>\w+)\))?"
    rf":(?P<pp>[\d\-–]+)\.?\s*$"
)
_VANC_BOOK = re.compile(
    rf"^(?P<auth>[^.]+?)\.\s+(?P<title>.+?)\.\s+"
    rf"(?:(?P<place>[^:;.]+):\s*)?(?P<pub>[^;]+?);\s*(?P<year>{_YEAR})\.?\s*$"
)
_VANC_AUTHORS = re.compile(
    r"[A-Z][\w'’\-]+\s+[A-Z]{1,3}(,\s*[A-Z][\w'’\-]+\s+[A-Z]{1,3})*(,\s*et al)?"
)
_CHI_BOOK = re.compile(
    rf"^(?P<auth>.+?)\.\s+(?P<year>{_YEAR})\.\s+(?P<rest>.+)$"
)


def parse_entry_text(
    raw: str, italics: list[str] | None = None, style_hint: str | None = None
) -> dict:
    """Parse one reference-list paragraph into the entry model. Heuristic:
    unparseable structure yields parse_confidence 'failed' with the raw text
    intact; nothing is guessed."""
    italics = italics or []
    e: dict = {"raw": raw, "flags": [], "type": "unknown", "authors": []}
    work = raw.strip()
    if not work:
        e["parse_confidence"] = "failed"
        return e

    m = _ENTRY_NUM.match(work)
    if m:
        e["number"] = int(m.group(1) or m.group(2))
        work = work[m.end():]

    m = _DOI_URL.search(work)
    if m:
        ref = m.group(1).rstrip(".")
        if ref.lower().startswith("http") and "doi.org" in ref:
            e["doi"] = ref
        elif ref.lower().startswith("doi"):
            e["doi"] = ref
        else:
            e["url"] = ref
        work = work[: m.start()].rstrip()

    # ---- Vancouver shapes (before the generic paths; very rigid author form)
    m = _VANC_ENTRY.match(work)
    if m and _VANC_AUTHORS.fullmatch(m.group("auth").strip()):
        auths, et_al, _ = _parse_authors(m.group("auth"))
        e.update(
            type="article", authors=auths, et_al=et_al,
            year=m.group("year"), title=m.group("title").strip(),
            container=m.group("cont").strip(), volume=m.group("vol"),
            issue=m.group("iss"), pages=m.group("pp"),
        )
        return _finish(e)
    m = _VANC_BOOK.match(work)
    if m and _VANC_AUTHORS.fullmatch(m.group("auth").strip()):
        auths, et_al, _ = _parse_authors(m.group("auth"))
        e.update(
            type="book", authors=auths, et_al=et_al,
            year=m.group("year"), title=m.group("title").strip(),
            place=(m.group("place") or "").strip() or None,
            publisher=m.group("pub").strip(),
        )
        return _finish(e)

    # ---- APA / Harvard: Authors (Year). Rest
    # (a quoted title BEFORE the year paren means Chicago NB — not this path)
    m = _APA_HEAD.match(work)
    if m and re.search(r"[“”\"]", m.group("auth")):
        m = None
    if m and m.group("auth").strip():
        auths, et_al, ok = _parse_authors(m.group("auth"))
        e["authors"], e["et_al"] = auths, et_al
        if not ok:
            e["flags"].append("author string parsed loosely")
        e["year"] = m.group("year")
        yextra = m.group("yextra").strip(" ,")
        if yextra:
            e["year_extra"] = yextra
        rest = m.group("rest").strip()

        hm = re.match(r"^[‘'](?P<t>[^’']{3,300})[’']\s*,\s*(?P<after>.+)$", rest)
        if hm:  # Harvard quoted article title
            e["title"] = hm.group("t")
            after = hm.group("after").strip()
            am = re.match(
                rf"^(?P<cont>.+?),\s*(?P<vol>\d+)\s*(?:\((?P<iss>\w+)\))?"
                rf"(?:,\s*pp?\.\s*(?P<pp>[\d\-–]+))?\.?\s*$",
                after,
            )
            if am:
                e.update(
                    type="article", container=am.group("cont").strip(),
                    volume=am.group("vol"), issue=am.group("iss"),
                    pages=am.group("pp"),
                )
            else:
                e["container"] = after.rstrip(".")
                e["type"] = "article"
            return _finish(e)

        # split title from the remainder at the first sentence boundary
        parts = re.split(r"(?<!\bIn)\.\s+(?=[A-Z“])", rest, maxsplit=1)
        if len(parts) == 2:
            title, remainder = parts[0].strip(), parts[1].strip()
            if remainder.startswith("In "):
                e["title"] = title
                e["type"] = "chapter"
                cm = re.match(
                    r"^In\s+(?:(?P<eds>.+?)\s+\(Eds?\.\),\s*)?(?P<book>.+?)"
                    r"(?:\s*\(pp\.\s*(?P<pp>[\d\-–]+)\))?\.\s*(?P<pub>.+?)\.?\s*$",
                    remainder,
                )
                if cm:
                    e["container"] = cm.group("book").strip()
                    e["pages"] = cm.group("pp")
                    place, pub = _split_place_publisher(cm.group("pub"))
                    e["place"], e["publisher"] = place, pub
                    if cm.group("eds"):
                        eds, _, _ = _parse_authors(cm.group("eds"))
                        # editors arrive as "E. Editor" natural order
                        e["editors"] = eds
                else:
                    e["container"] = remainder[3:].rstrip(".")
                return _finish(e)
            am = _APA_ARTICLE_TAIL.match(remainder)
            if am and not remainder.rstrip(".").endswith(("Press", "Books")):
                e.update(
                    type="article", title=title,
                    container=am.group("cont").strip(),
                    volume=am.group("vol"), issue=am.group("iss"),
                    pages=(am.group("pp") or "").strip() or None,
                )
                return _finish(e)
            # book / report: title then publisher
            e["title"] = title
            if re.search(r"\d{2,}", remainder) and "(" not in remainder:
                e["flags"].append(
                    "publisher segment contains digits; check parse"
                )
            place, pub = _split_place_publisher(remainder)
            e["place"], e["publisher"] = place, pub
            e["type"] = "book"
            return _finish(e)
        # single segment after year: title only (web/report without venue)
        e["title"] = rest.rstrip(".")
        if e.get("url") or e.get("doi"):
            e["type"] = "web"
        return _finish(e)

    # ---- quoted-title styles: Chicago (both), MLA, IEEE, ASA
    qm = _QUOTED.search(work)
    if qm:
        e["title"] = qm.group("t")
        pre = work[: qm.start()].strip()
        post = work[qm.end():].strip().lstrip(",.").strip()
        cm = re.search(rf"(?:^|\.\s+)(?P<year>{_YEAR})\.\s*$", pre)
        if cm:  # Chicago AD / ASA: Authors. Year. "Title." Journal ...
            e["year"] = cm.group("year")
            pre = pre[: cm.start()].rstrip(".")
        auths, et_al, ok = _parse_authors(pre)
        e["authors"], e["et_al"] = auths, et_al
        if not ok:
            e["flags"].append("author string parsed loosely")
        for pat in (_NB_TAIL, _MLA_TAIL, _IEEE_TAIL, _CHIAD_TAIL):
            am = pat.match(post)
            if am:
                g = am.groupdict()
                e.update(
                    type="article",
                    container=g.get("cont", "").strip() or None,
                    volume=g.get("vol"), issue=g.get("iss"),
                    pages=g.get("pp"),
                )
                if g.get("year"):
                    e["year"] = g["year"]
                return _finish(e)
        if post.lower().startswith("in "):
            e["type"] = "chapter"
            e["container"] = re.split(r",\s*edited by", post[3:])[0].rstrip(".,")
            ym = re.search(rf"\b({_YEAR})\b", post)
            if ym and not e.get("year"):
                e["year"] = ym.group(1)
            return _finish(e)
        # journal without volume, or unknown tail
        ym = re.search(rf"[,(]\s*({_YEAR})\s*[).]", post)
        if ym and not e.get("year"):
            e["year"] = ym.group(1)
        cont = re.split(rf"[,(]\s*{_YEAR}", post)[0].strip().rstrip(",.")
        if cont:
            e["container"] = cont
            e["type"] = "article"
        return _finish(e)

    # ---- Chicago/ASA/MLA/IEEE books (no quotes; italics are the best hint)
    m = _CHI_BOOK.match(work)
    if m and not m.group("auth").count(". ") > 6:
        auths, et_al, ok = _parse_authors(m.group("auth"))
        if auths and "literal" not in auths[0]:
            e["authors"], e["et_al"] = auths, et_al
            e["year"] = m.group("year")
            rest = m.group("rest").strip()
            title = None
            for it in italics:
                if it in rest:
                    title = it
                    tail = rest.split(it, 1)[1].strip().lstrip(".").strip()
                    place, pub = _split_place_publisher(tail)
                    e["place"], e["publisher"] = place, pub
                    break
            if title is None:
                segs = rest.split(". ")
                title = segs[0].rstrip(".")
                if len(segs) > 1:
                    place, pub = _split_place_publisher(
                        ". ".join(segs[1:])
                    )
                    e["place"], e["publisher"] = place, pub
            e["title"] = title
            e["type"] = "book"
            return _finish(e)

    # ---- MLA/IEEE book: Authors. Title. Publisher, Year. (italic title hint)
    ym = _CHI_YEAR_TAIL.search(work)
    if ym and italics:
        for it in italics:
            if it in work:
                pre = work.split(it, 1)[0].strip().rstrip(",.")
                tail = work.split(it, 1)[1].strip().lstrip(",.").strip()
                tail = tail[: ym.start() - work.index(tail)] if tail else ""
                auths, et_al, ok = _parse_authors(pre)
                if auths and "literal" not in auths[0]:
                    e["authors"], e["et_al"] = auths, et_al
                    e["title"] = it
                    e["year"] = ym.group("year")
                    place, pub = _split_place_publisher(
                        tail.rstrip(",. ")
                    )
                    e["place"], e["publisher"] = place, pub
                    e["type"] = "book"
                    return _finish(e)

    # ---- give up: keep whatever partial signal exists
    lead = re.match(r"^\s*([A-Z][\w'’\-]+)", work)
    ymatch = re.search(rf"\(({_YEAR})\b[^)]*\)", work) or re.search(
        rf"\b({_YEAR})\b", work
    )
    if lead:
        e["authors"] = [{"literal": lead.group(1)}]
    if ymatch:
        e["year"] = ymatch.group(1)
    e["flags"].append("structure not recognized; entry left verbatim")
    e["parse_confidence"] = "failed"
    return e


# ========================================================= citation scanning


def _split_names(names: str) -> list[str]:
    return re.split(r"\s+(?:and|&)\s+", names)


def _scan_paragraph_citations(text: str, superscript_spans: list[tuple[int, int]]):
    """All in-text citations in one paragraph's visible text.

    Returns a list of dicts: kind (narrative|parenthetical|bracket|
    superscript), start, end, raw, items. Item: {"surname","year","loc"} or
    {"surname","loc"} (author-page) or {"number","loc"}."""
    found = []
    consumed: list[tuple[int, int]] = []

    for m in _NARRATIVE.finditer(text):
        pstart = m.start("paren")
        found.append({
            "kind": "narrative",
            "start": pstart,
            "end": m.end("paren"),
            "name_start": m.start("names"),
            "raw": m.group(0),
            "items": [{
                "surname": _split_names(m.group("names"))[0],
                "all_names": m.group("names"),
                "et_al": bool(m.group("etal")),
                "year": m.group("year"),
                "loc": (m.group("loc") or "").strip() or None,
            }],
        })
        consumed.append((pstart, m.end("paren")))

    for m in _PAREN.finditer(text):
        if any(s < m.end() and e > m.start() for s, e in consumed):
            continue
        inner = m.group(1)
        if not re.search(_YEAR, inner) and not _ITEM_AP.match(inner):
            continue
        items = []
        clean = True
        for chunk in inner.split(";"):
            chunk = chunk.strip()
            im = _ITEM_AD.match(chunk)
            if im:
                if im.group("prefix"):
                    clean = False  # "see also X" — do not rewrite prose
                items.append({
                    "surname": _split_names(im.group("names"))[0],
                    "all_names": im.group("names"),
                    "et_al": bool(im.group("etal")),
                    "year": im.group("year"),
                    "loc": (im.group("loc") or "").strip() or None,
                })
                continue
            am = _ITEM_AP.match(chunk)
            if am:
                items.append({
                    "surname": _split_names(am.group("names"))[0],
                    "all_names": am.group("names"),
                    "et_al": bool(am.group("etal")),
                    "year": None,
                    "loc": am.group("loc"),
                })
                continue
            clean = False
        if items:
            found.append({
                "kind": "parenthetical",
                "start": m.start(),
                "end": m.end(),
                "raw": m.group(0),
                "items": items,
                "clean": clean,
            })
            consumed.append((m.start(), m.end()))

    bracket_cites: list[dict] = []
    for m in _BRACKET.finditer(text):
        nums = []
        for tok in re.split(r"[,;]", m.group("nums")):
            tok = tok.strip()
            if re.match(r"^\d+\s*[\-–]\s*\d+$", tok):
                a, b = re.split(r"[\-–]", tok)
                nums.extend(range(int(a), int(b) + 1))
            elif tok:
                nums.append(int(tok))
        cite = {
            "kind": "bracket",
            "start": m.start(),
            "end": m.end(),
            "raw": m.group(0),
            "items": [
                {"number": n, "loc": m.group("loc") if len(nums) == 1 else None}
                for n in nums
            ],
        }
        # merge "[3], [4]" (adjacent brackets, IEEE multi-cite) into one
        if bracket_cites and re.fullmatch(
            r"[,;]?\s*", text[bracket_cites[-1]["end"]:m.start()]
        ):
            prev = bracket_cites[-1]
            prev["end"] = m.end()
            prev["raw"] = text[prev["start"]:prev["end"]]
            prev["items"].extend(cite["items"])
        else:
            bracket_cites.append(cite)
    for cite in bracket_cites:
        wm = re.search(r"([A-Za-z][\w'’\-]*)\s*$", text[: cite["start"]])
        cite["preceding_word"] = wm.group(1) if wm else None
        found.append(cite)

    for s, e in superscript_spans:
        raw = text[s:e]
        nums = [int(t) for t in re.split(r"[,\-–]", raw) if t.strip().isdigit()]
        if nums:
            wm = re.search(r"([A-Za-z][\w'’\-]*)\s*$", text[:s].rstrip(".!?"))
            found.append({
                "kind": "superscript",
                "start": s,
                "end": e,
                "raw": raw,
                "items": [{"number": n, "loc": None} for n in nums],
                "preceding_word": wm.group(1) if wm else None,
            })

    found.sort(key=lambda c: c["start"])
    return found


def _superscript_spans(p: etree._Element) -> list[tuple[int, int]]:
    """Character spans of superscripted digit runs (Vancouver citations)."""
    _, segments = _runmap.build_map(p)
    spans: list[tuple[int, int]] = []
    for seg in segments:
        if seg.atomic:
            continue
        rpr = seg.run.find(qn("w:rPr"))
        if rpr is None:
            continue
        va = rpr.find(qn("w:vertAlign"))
        if va is None or va.get(qn("w:val")) != "superscript":
            continue
        # exclude real note references (they carry a footnoteReference, not w:t)
        text = seg.el.text or ""
        if not text or not _SUPERSCRIPT_TEXT.match(text):
            continue
        if spans and spans[-1][1] == seg.start:
            spans[-1] = (spans[-1][0], seg.end)
        else:
            spans.append((seg.start, seg.end))
    return spans


# ====================================================== native-field routing


def detect_citation_fields(pkg: DocxPackage) -> dict:
    """Word-native CITATION/BIBLIOGRAPHY fields and reference-manager fields
    present in the document. Non-empty means text conversion must not run."""
    counts = {"native": 0, "zotero": 0, "mendeley": 0, "endnote": 0}
    for part in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
        if not pkg.has_part(part):
            continue
        fields, _ = scan_complex_fields(pkg.root(part))
        for f in fields:
            instr = " ".join(f["instr"].split())
            if instr.startswith("CITATION") or instr.startswith("BIBLIOGRAPHY"):
                counts["native"] += 1
                continue
            cls = _classify_manager(instr)
            if cls:
                counts[cls[0]] += 1
    return counts


# =========================================================== parse (stage 1)


def parse_references(pkg: DocxPackage, *, style_hint: str | None = None) -> dict:
    """Stage 1: parse the reference list and in-text citations into a
    structured model. Read-only.

    Heuristic text parsing: every entry carries parse_confidence
    (full/partial/failed); failed entries are returned verbatim. Citations
    include positions and whether they resolve to a parsed entry."""
    hint = canonical_style(style_hint) if style_hint else None
    if style_hint and hint is None:
        raise WordMcpError(
            f"unknown style_hint {style_hint!r}; known: "
            + ", ".join(sorted(STYLE_INFO))
        )
    fields = detect_citation_fields(pkg)
    ref_start, ref_end = _locate_reference_list(pkg)
    paras = _body_paragraph_elements(pkg)
    by_index = dict(paras)

    entries = []
    for idx, el in paras:
        if idx <= ref_start or (ref_end is not None and idx >= ref_end):
            continue
        raw = paragraph_text(el)
        if not raw.strip():
            continue
        entry = parse_entry_text(raw, _italic_spans(el), hint)
        entry["paragraph_index"] = idx
        entries.append(entry)

    citations = []
    kind_counts: dict[str, int] = {}
    for idx, el in paras:
        if idx >= ref_start:
            continue
        text, _ = _runmap.build_map(el)
        for c in _scan_paragraph_citations(text, _superscript_spans(el)):
            c["paragraph_index"] = idx
            kind_counts[c["kind"]] = kind_counts.get(c["kind"], 0) + 1
            citations.append(c)

    n_foot = 0
    if pkg.has_part("word/footnotes.xml"):
        n_foot = sum(
            1
            for n in pkg.root("word/footnotes.xml").findall(qn("w:footnote"))
            if n.get(qn("w:type")) is None
        )

    if hint:
        system = STYLE_INFO[hint]["system"]
    elif kind_counts.get("bracket", 0) > (
        kind_counts.get("parenthetical", 0) + kind_counts.get("narrative", 0)
    ):
        system = "numbered-bracket"
    elif kind_counts.get("superscript", 0) > (
        kind_counts.get("parenthetical", 0) + kind_counts.get("narrative", 0)
    ):
        system = "numbered-superscript"
    elif (
        n_foot > 0
        and not kind_counts.get("parenthetical")
        and not kind_counts.get("narrative")
        and not kind_counts.get("bracket")
    ):
        system = "notes"
    else:
        system = "author-date"

    heading_el = by_index.get(ref_start)
    conf = {"full": 0, "partial": 0, "failed": 0}
    for e in entries:
        conf[e["parse_confidence"]] += 1
    return {
        "reference_list": {
            "heading_index": ref_start,
            "heading_text": paragraph_text(heading_el) if heading_el is not None else "",
            "end_index": ref_end,
            "entry_count": len(entries),
        },
        "entries": entries,
        "citations": citations,
        "citation_fields": fields,
        "detected_source_system": system,
        "confidence_counts": conf,
        "note": (
            "Heuristic text parsing. parse_confidence 'failed' entries are "
            "verbatim and will never be converted; 'partial' entries are "
            "flagged and left untouched by conversion."
        ),
    }


# ========================================================= convert (stage 2)


def _entry_key_maps(entries: list[dict]):
    by_key: dict[tuple[str, str], list[int]] = {}
    by_fam: dict[str, list[int]] = {}
    by_num: dict[int, int] = {}
    for i, e in enumerate(entries):
        auths = e.get("authors") or []
        fam = ""
        if auths:
            fam = (auths[0].get("family") or auths[0].get("literal") or "").lower()
        year = (e.get("year") or "").lower()
        if fam and year:
            by_key.setdefault((fam, year), []).append(i)
        if fam:
            by_fam.setdefault(fam, []).append(i)
        if e.get("number"):
            by_num[e["number"]] = i
    return by_key, by_fam, by_num


def _resolve_item(item, entries, by_key, by_fam, by_num) -> tuple[int | None, str | None]:
    """(entry index, problem)."""
    if "number" in item:
        idx = by_num.get(item["number"])
        if idx is None and 1 <= item["number"] <= len(entries):
            idx = item["number"] - 1  # positional fallback (unnumbered list)
        if idx is None:
            return None, f"no reference entry numbered {item['number']}"
        return idx, None
    fam = item["surname"].lower()
    year = (item.get("year") or "").lower()
    if year:
        hits = by_key.get((fam, year), [])
        if len(hits) == 1:
            return hits[0], None
        if len(hits) > 1:
            return None, (
                f"{item['surname']} {item['year']}: multiple entries match "
                "(needs year-letters)"
            )
        # MLA edge: "(Hurd 1999)" could be author-page with a year-like page
        hits = by_fam.get(fam, [])
        if len(hits) == 1 and not by_key.get((fam, year)):
            e = entries[hits[0]]
            if (e.get("year") or "").lower() == year:
                return hits[0], None
        return None, f"{item['surname']} ({item['year']}) not in reference list"
    hits = by_fam.get(fam, [])
    if len(hits) == 1:
        return hits[0], None
    if not hits:
        return None, f"{item['surname']} not in reference list"
    return None, (
        f"{item['surname']}: multiple works by this author; author-page "
        "citation is ambiguous without a short title"
    )


def _rebuildable(p: etree._Element) -> bool:
    """True when the paragraph holds only plain runs (safe to rebuild)."""
    complex_tags = {
        "fldChar", "instrText", "fldSimple", "footnoteReference",
        "endnoteReference", "drawing", "pict", "object", "hyperlink",
        "commentReference", "ins", "del", "sdt",
    }
    for el in p.iter():
        if etree.QName(el).localname in complex_tags:
            return False
    return True


def _rebuild_paragraph(p: etree._Element, segments: list[tuple[str, bool]]) -> None:
    for r in list(p.findall(qn("w:r"))):
        p.remove(r)
    for text, italic in segments:
        r = etree.SubElement(p, qn("w:r"))
        if italic:
            rpr = etree.SubElement(r, qn("w:rPr"))
            etree.SubElement(rpr, qn("w:i"))
            etree.SubElement(rpr, qn("w:iCs"))
        t = etree.SubElement(r, qn("w:t"))
        t.text = text
        _runmap._preserve_space(t)


def _insert_runs_at(p: etree._Element, pos: int, runs: list) -> None:
    if pos <= 0:
        anchor = p.find(qn("w:pPr"))
        if anchor is not None:
            for r in reversed(runs):
                anchor.addnext(r)
        else:
            for r in reversed(runs):
                p.insert(0, r)
        return
    covered = _runmap.split_for_range(p, pos - 1, pos)
    last = covered[-1]
    for r in reversed(runs):
        last.addnext(r)


def _make_text_run(text: str, *, superscript: bool = False) -> etree._Element:
    r = etree.Element(qn("w:r"))
    if superscript:
        rpr = etree.SubElement(r, qn("w:rPr"))
        etree.SubElement(rpr, qn("w:vertAlign")).set(qn("w:val"), "superscript")
    t = etree.SubElement(r, qn("w:t"))
    t.text = text
    _runmap._preserve_space(t)
    return r


def _add_footnote_with_segments(
    pkg: DocxPackage, p: etree._Element, anchor_pos: int,
    segments: list[tuple[str, bool]],
) -> int:
    """Real footnote via the ops/notes machinery, with italic-capable
    content, anchored at character position anchor_pos."""
    _notes._ensure_part(pkg, "footnote")
    _notes._ensure_styles(pkg, "footnote")
    cfg = _notes._KINDS["footnote"]
    note_id = _notes._next_id(pkg, "footnote")

    notes_root = pkg.root(cfg["part"])
    note = etree.SubElement(notes_root, qn(cfg["note"]))
    note.set(qn("w:id"), str(note_id))
    np = etree.SubElement(note, qn("w:p"))
    ppr = etree.SubElement(np, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), cfg["text_style"][0])
    ref_run = etree.SubElement(np, qn("w:r"))
    rpr = etree.SubElement(ref_run, qn("w:rPr"))
    etree.SubElement(rpr, qn("w:rStyle")).set(qn("w:val"), cfg["ref_style"][0])
    etree.SubElement(ref_run, qn(cfg["self_ref"]))
    sp = etree.SubElement(np, qn("w:r"))
    t = etree.SubElement(sp, qn("w:t"))
    t.text = " "
    _runmap._preserve_space(t)
    for text, italic in segments:
        r = etree.SubElement(np, qn("w:r"))
        if italic:
            rp = etree.SubElement(r, qn("w:rPr"))
            etree.SubElement(rp, qn("w:i"))
            etree.SubElement(rp, qn("w:iCs"))
        rt = etree.SubElement(r, qn("w:t"))
        rt.text = text
        _runmap._preserve_space(rt)
    pkg.mark_dirty(cfg["part"])

    body_ref = etree.Element(qn("w:r"))
    brpr = etree.SubElement(body_ref, qn("w:rPr"))
    etree.SubElement(brpr, qn("w:rStyle")).set(qn("w:val"), cfg["ref_style"][0])
    ref = etree.SubElement(body_ref, qn(cfg["body_ref"]))
    ref.set(qn("w:id"), str(note_id))
    _insert_runs_at(p, anchor_pos, [body_ref])
    pkg.mark_dirty()
    return note_id


def _extend_left_over_space(text: str, start: int) -> int:
    while start > 0 and text[start - 1] == " ":
        start -= 1
    return start


def _sentence_end_after(text: str, pos: int) -> int:
    m = _SENT_END.search(text, pos)
    return m.start() if m else len(text)


def convert_citation_style(
    pkg: DocxPackage,
    target_style: str,
    *,
    source_style: str = "auto",
    dry_run: bool = False,
) -> dict:
    """Stage 2: convert plain-text citations and the reference list to
    target_style.

    HEURISTIC TEXT CONVERSION — the flagged list in the result must be
    reviewed by a human. Only fully-parsed entries are rewritten; only
    citations that resolve unambiguously are converted; everything else is
    left verbatim and flagged. Word-native/Zotero/Mendeley/EndNote CITATION
    fields cause the document to be ROUTED (report, no text rewriting).
    dry_run=True returns the full change plan without touching the file."""
    target = canonical_style(target_style)
    if target is None:
        raise WordMcpError(
            f"unknown target_style {target_style!r}; known styles: "
            + ", ".join(sorted(STYLE_INFO))
            + " (aliases: turabian, chicago-ad, chicago-nb, asa-sage, ...)"
        )
    if source_style != "auto" and canonical_style(source_style) is None:
        raise WordMcpError(
            f"unknown source_style {source_style!r}; use 'auto' or one of "
            + ", ".join(sorted(STYLE_INFO))
        )

    fields = detect_citation_fields(pkg)
    if any(fields.values()):
        actions = []
        if fields["native"]:
            actions.append(
                f"{fields['native']} Word-native CITATION/BIBLIOGRAPHY "
                "field(s): use set_bibliography_style (Word's built-in "
                "styles) instead of text conversion"
            )
        for mgr in ("zotero", "mendeley", "endnote"):
            if fields[mgr]:
                actions.append(
                    f"{fields[mgr]} {mgr.capitalize()} field(s): switch the "
                    f"citation style inside {mgr.capitalize()} and refresh; "
                    "rewriting them as text would disconnect the library"
                )
        return {
            "routed": True,
            "converted": False,
            "citation_fields": fields,
            "action_required": actions,
            "note": (
                "This document's citations are live fields, not plain text. "
                "Text conversion was NOT performed and the file is unchanged."
            ),
        }

    parsed = parse_references(
        pkg, style_hint=None if source_style == "auto" else source_style
    )
    entries = parsed["entries"]
    citations = parsed["citations"]
    source_system = parsed["detected_source_system"]
    target_info = STYLE_INFO[target]
    target_system = target_info["system"]
    by_key, by_fam, by_num = _entry_key_maps(entries)

    # Entries whose paragraphs cannot be rebuilt (fields/hyperlinks/tracked
    # changes) stay verbatim at write time — so their dependent citations
    # must not be converted or numbered either, or the body would carry [n]
    # references no visible entry announces (v1.5 adversarial finding F2).
    _by_index = dict(_body_paragraph_elements(pkg))
    protected_entries = set()
    for _i, _e in enumerate(entries):
        _el = _by_index.get(_e["paragraph_index"])
        if (
            _e["parse_confidence"] == "full"
            and _el is not None
            and not _rebuildable(_el)
        ):
            protected_entries.add(_i)

    flags: list[str] = []
    citation_flags: list[dict] = []
    entry_flags: list[dict] = []

    # ---- resolve citations
    for c in citations:
        c["resolved"] = []
        c["problems"] = []
        for item in c["items"]:
            idx, problem = _resolve_item(item, entries, by_key, by_fam, by_num)
            c["resolved"].append(idx)
            if problem:
                c["problems"].append(problem)
            elif idx is not None and idx in protected_entries:
                c["problems"].append(
                    "cites an entry left verbatim (its paragraph holds "
                    "fields/hyperlinks/tracked changes); citation kept as-is"
                )
            elif idx is not None and entries[idx]["parse_confidence"] != "full":
                c["problems"].append(
                    f"cites an entry that did not fully parse "
                    f"(entry at paragraph {entries[idx]['paragraph_index']})"
                )
        if c.get("clean") is False:
            c["problems"].append(
                "parenthetical contains prose (e.g./cf./see); left untouched"
            )

    convertible = [c for c in citations if not c["problems"]]
    doc_order = sorted(
        convertible, key=lambda c: (c["paragraph_index"], c["start"])
    )

    # ---- numbering (numbered targets): first appearance, full entries only
    numbers: dict[int, int] = {}
    if target_system in ("numbered-bracket", "numbered-superscript"):
        n = 0
        for c in doc_order:
            for idx in c["resolved"]:
                if (
                    idx is not None
                    and idx not in numbers
                    and idx not in protected_entries
                ):
                    n += 1
                    numbers[idx] = n
        for i, e in enumerate(entries):  # uncited full entries get numbers too
            if (
                e["parse_confidence"] == "full"
                and i not in numbers
                and i not in protected_entries
            ):
                n += 1
                numbers[i] = n

    # ---- first-use tracking (notes target)
    first_use_seen: set[int] = set()

    # ---- build citation ops
    cite_ops: list[dict] = []
    footnote_plan_count = 0
    for c in doc_order:
        op = {
            "paragraph_index": c["paragraph_index"],
            "start": c["start"],
            "end": c["end"],
            "kind": c["kind"],
            "before": c["raw"],
        }
        ents = [entries[i] for i in c["resolved"]]

        if target_system == "notes":
            segs: list[tuple[str, bool]] = []
            for pos_i, (i, e) in enumerate(zip(c["resolved"], ents)):
                fu = i not in first_use_seen
                first_use_seen.add(i)
                loc = c["items"][pos_i].get("loc")
                part = emit_note(e, loc, first_use=fu)
                if segs:
                    # join multi-work citations into one note: "A.; B." -> "A; B."
                    last_text, last_it = segs[-1]
                    segs[-1] = (last_text.rstrip("."), last_it)
                    segs.append(("; ", False))
                segs.extend(part)
            op["action"] = "footnote"
            op["note_segments"] = segs
            op["after"] = "[footnote] " + "".join(t for t, _ in segs)
            footnote_plan_count += 1

        elif target_system == "numbered-superscript":
            nums = sorted(numbers[i] for i in c["resolved"])
            op["action"] = "superscript"
            op["number_text"] = ",".join(str(n) for n in nums)
            op["after"] = f"^{op['number_text']}"
            if any(item.get("loc") for item in c["items"]):
                citation_flags.append({
                    "citation": c["raw"],
                    "paragraph_index": c["paragraph_index"],
                    "problem": "page locator dropped (Vancouver superscript "
                               "citations carry no locators)",
                })

        elif target_system == "numbered-bracket":
            if c["kind"] == "narrative":
                # keep the name: Hurd (1999) -> Hurd [12]
                op["action"] = "replace"
                op["after"] = emit_intext_item(
                    "ieee", ents[0], c["items"][0].get("loc"),
                    numbers[c["resolved"][0]],
                )
            else:
                parts = [
                    emit_intext_item(
                        "ieee", e, item.get("loc"), numbers[i]
                    )
                    for i, e, item in zip(c["resolved"], ents, c["items"])
                ]
                op["action"] = "replace"
                op["after"] = ", ".join(parts)

        else:  # author-date / author-page targets
            if c["kind"] == "narrative" and target_system == "author-date":
                op["action"] = "replace"
                op["after"] = emit_intext_year_only(
                    target, ents[0], c["items"][0].get("loc")
                )
            elif c["kind"] == "narrative":  # MLA target
                loc = c["items"][0].get("loc")
                op["action"] = "replace_or_remove"
                op["after"] = f"({norm_pages(loc)})" if loc else ""
            else:
                # numbered -> author-date narrative: "Lake [2] argues" must
                # become "Lake (2009) argues", keeping the name in prose.
                fam0 = ""
                if ents and (ents[0].get("authors") or []):
                    a0 = ents[0]["authors"][0]
                    fam0 = (a0.get("family") or a0.get("literal") or "").lower()
                numbered_narrative = (
                    c["kind"] in ("bracket", "superscript")
                    and len(c["resolved"]) == 1
                    and fam0
                    and (c.get("preceding_word") or "").lower() == fam0
                )
                if numbered_narrative and target_system == "author-date":
                    op["action"] = "replace"
                    op["after"] = emit_intext_year_only(
                        target, ents[0], c["items"][0].get("loc")
                    )
                elif numbered_narrative:  # MLA target
                    loc = c["items"][0].get("loc")
                    op["action"] = "replace_or_remove"
                    op["after"] = f"({norm_pages(loc)})" if loc else ""
                else:
                    parts = [
                        emit_intext_item(target, e, item.get("loc"), None)
                        for e, item in zip(ents, c["items"])
                    ]
                    op["action"] = "replace"
                    op["after"] = "(" + "; ".join(parts) + ")"
                if c["kind"] in ("bracket", "superscript"):
                    op["from_numbered"] = True
        cite_ops.append(op)

    for c in citations:
        if c["problems"]:
            citation_flags.append({
                "citation": c["raw"],
                "paragraph_index": c["paragraph_index"],
                "problem": "; ".join(c["problems"]),
            })

    # ---- build entry ops
    entry_ops: list[dict] = []
    for i, e in enumerate(entries):
        if i in protected_entries:
            entry_flags.append({
                "paragraph_index": e["paragraph_index"],
                "entry": e["raw"][:120],
                "confidence": "full",
                "problem": (
                    "entry paragraph holds fields/hyperlinks/tracked "
                    "changes; rebuilding would destroy them — left verbatim "
                    "(its citations were also left unconverted)"
                ),
            })
            continue
        if e["parse_confidence"] != "full":
            entry_flags.append({
                "paragraph_index": e["paragraph_index"],
                "entry": e["raw"][:120],
                "confidence": e["parse_confidence"],
                "problem": (
                    "entry left verbatim (not fully parsed); convert by hand "
                    "or fix the source formatting"
                ),
            })
            continue
        segs, eflags = emit_entry(target, e, numbers.get(i))
        entry_ops.append({
            "paragraph_index": e["paragraph_index"],
            "before": e["raw"],
            "after": "".join(t for t, _ in segs),
            "segments": segs,
            "number": numbers.get(i),
        })
        for f in eflags:
            entry_flags.append({
                "paragraph_index": e["paragraph_index"],
                "entry": e["raw"][:80],
                "confidence": "full",
                "problem": f,
            })

    # ---- reference list ordering
    all_full = all(e["parse_confidence"] == "full" for e in entries)
    reorder: list[int] | None = None
    if all_full and entries:
        if target_info["ordering"] == "citation-order" and numbers:
            desired = sorted(
                range(len(entries)), key=lambda i: numbers.get(i, 10 ** 6)
            )
        else:
            def _alpha(i):
                e = entries[i]
                a = (e.get("authors") or [{}])[0]
                return (
                    (a.get("family") or a.get("literal") or "").lower(),
                    (e.get("year") or "").lower(),
                    (e.get("title") or "").lower(),
                )
            desired = sorted(range(len(entries)), key=_alpha)
        if desired != list(range(len(entries))):
            reorder = desired
    elif not all_full and entries:
        flags.append(
            "reference list NOT reordered: some entries failed to parse and "
            "will not be moved"
        )
        if numbers:
            flags.append(
                "numbered citations were assigned by first appearance, but "
                "the list order could not be normalized (unparsed entries "
                "present); entry numbers are carried in each entry's [n] "
                "prefix"
            )

    # ---- heading
    heading_op = None
    ref_start = parsed["reference_list"]["heading_index"]
    cur_heading = parsed["reference_list"]["heading_text"].strip()
    if cur_heading.lower() != target_info["ref_heading"].lower():
        heading_op = {
            "paragraph_index": ref_start,
            "before": cur_heading,
            "after": target_info["ref_heading"],
        }

    review = []
    if citation_flags:
        review.append(
            f"{len(citation_flags)} citation(s) not converted or degraded — review each"
        )
    if entry_flags:
        review.append(
            f"{len(entry_flags)} reference-entry issue(s) — review each"
        )
    if target_system == "notes" and pkg.has_part("word/footnotes.xml"):
        existing = sum(
            1
            for n in pkg.root("word/footnotes.xml").findall(qn("w:footnote"))
            if n.get(qn("w:type")) is None
        )
        if existing:
            review.append(
                f"document already has {existing} footnote(s); new citation "
                "footnotes are interleaved with them — check numbering"
            )
    if source_system == "notes" and target_system != "notes":
        review.append(
            "source uses notes: only footnotes that are RECOGNIZABLY pure "
            "citations are harvested; mixed-content footnotes are left alone "
            "and flagged"
        )

    plan = {
        "target_style": target_info["label"],
        "source_system": source_system,
        "target_system": target_system,
        "citations": cite_ops,
        "citations_flagged": citation_flags,
        "entries": entry_ops,
        "entries_flagged": entry_flags,
        "entry_reorder": reorder,
        "heading": heading_op,
        "flags": flags,
        "review_required": review,
    }

    # ---- notes-source harvesting is planned separately (needs XML walk)
    harvest_ops = []
    if source_system == "notes" and target_system != "notes":
        harvest_ops, harvest_flags = _plan_note_harvest(
            pkg, entries, by_key, by_fam, target, numbers
        )
        plan["note_harvest"] = [
            {k: v for k, v in op.items() if k not in ("ref_el",)}
            for op in harvest_ops
        ]
        citation_flags.extend(harvest_flags)

    if dry_run:
        return {
            "dry_run": True,
            "converted": False,
            "plan": plan,
            "counts": {
                "citations_to_convert": len(cite_ops),
                "citations_flagged": len(citation_flags),
                "entries_to_convert": len(entry_ops),
                "entries_flagged": len(entry_flags),
                "footnotes_to_create": footnote_plan_count,
            },
            "review_required": review,
            "note": (
                "Dry run: the file was NOT modified. This is heuristic text "
                "conversion — review the plan and flags before applying."
            ),
        }

    # =================================================================
    # APPLY. Any exception leaves the caller's file untouched (nothing is
    # written to disk here; the tool layer only saves on success).
    # =================================================================
    paras = _body_paragraph_elements(pkg)
    by_index = dict(paras)

    if heading_op:
        p = by_index[heading_op["paragraph_index"]]
        text, segments = _runmap.build_map(p)
        if text:
            _runmap.replace_range(p, segments, 0, len(text), heading_op["after"])
            pkg.mark_dirty()

    entries_converted = 0
    for op in entry_ops:
        p = by_index[op["paragraph_index"]]
        if not _rebuildable(p):
            entry_flags.append({
                "paragraph_index": op["paragraph_index"],
                "entry": op["before"][:120],
                "confidence": "full",
                "problem": (
                    "entry paragraph holds fields/hyperlinks/tracked "
                    "changes; rebuilding would destroy them — left verbatim"
                ),
            })
            continue
        _rebuild_paragraph(p, op["segments"])
        entries_converted += 1
    if entries_converted:
        pkg.mark_dirty()

    if reorder:
        ordered_els = [
            by_index[entries[i]["paragraph_index"]] for i in reorder
        ]
        anchor = by_index[ref_start]
        prev = anchor
        for el in ordered_els:
            prev.addnext(el)
            prev = el
        pkg.mark_dirty()

    citations_converted = 0
    footnotes_created = 0
    by_par: dict[int, list[dict]] = {}
    for op in cite_ops:
        by_par.setdefault(op["paragraph_index"], []).append(op)
    for pidx, ops in by_par.items():
        p = by_index[pidx]
        for op in sorted(ops, key=lambda o: -o["start"]):
            text, segments = _runmap.build_map(p)
            start, end = op["start"], op["end"]
            if text[start:end] != (
                op["before"] if op["kind"] != "narrative" else text[start:end]
            ) and op["kind"] != "narrative":
                # paragraph text drifted (should not happen); skip safely
                citation_flags.append({
                    "citation": op["before"],
                    "paragraph_index": pidx,
                    "problem": "text changed during conversion; skipped",
                })
                continue
            action = op["action"]
            if action in ("replace", "replace_or_remove") and op["kind"] == "superscript":
                # Source span sits AFTER punctuation ("…claim.¹"): delete it,
                # then place the replacement BEFORE that punctuation (or after
                # the preceding name for narrative superscripts).
                _runmap.replace_range(p, segments, start, end, "")
                if op["after"]:
                    text2, segments2 = _runmap.build_map(p)
                    s = min(start, len(text2))
                    if s > 0 and text2[s - 1] in ".!?":
                        _runmap.replace_range(
                            p, segments2, s - 1, s,
                            " " + op["after"] + text2[s - 1],
                        )
                    elif s > 0:
                        _runmap.replace_range(
                            p, segments2, s - 1, s,
                            text2[s - 1] + " " + op["after"],
                        )
                    else:
                        _insert_runs_at(p, 0, [_make_text_run(op["after"])])
                citations_converted += 1
            elif action == "replace":
                _runmap.replace_range(p, segments, start, end, op["after"])
                citations_converted += 1
            elif action == "replace_or_remove":
                s = start if op["after"] else _extend_left_over_space(text, start)
                _runmap.replace_range(p, segments, s, end, op["after"])
                citations_converted += 1
            elif action == "superscript":
                s = _extend_left_over_space(text, start)
                _runmap.replace_range(p, segments, s, end, "")
                text2, _ = _runmap.build_map(p)
                if op["kind"] in ("narrative", "superscript"):
                    ins = min(s, len(text2))  # already at its display spot
                else:
                    pe = _sentence_end_after(text2, s)
                    ins = pe + 1 if pe < len(text2) else len(text2)
                _insert_runs_at(
                    p, ins, [_make_text_run(op["number_text"], superscript=True)]
                )
                citations_converted += 1
            elif action == "footnote":
                s = _extend_left_over_space(text, start)
                _runmap.replace_range(p, segments, s, end, "")
                text2, _ = _runmap.build_map(p)
                if op["kind"] == "superscript":
                    anchor_pos = min(s, len(text2))
                else:
                    pe = _sentence_end_after(text2, s)
                    anchor_pos = pe + 1 if pe < len(text2) else len(text2)
                _add_footnote_with_segments(
                    pkg, p, anchor_pos, op["note_segments"]
                )
                footnotes_created += 1
                citations_converted += 1
            pkg.mark_dirty()

    harvested = 0
    for op in harvest_ops:
        if _apply_note_harvest(pkg, op):
            harvested += 1
            citations_converted += 1

    return {
        "converted": True,
        "target_style": target_info["label"],
        "source_system": source_system,
        "entries_converted": entries_converted,
        "citations_converted": citations_converted,
        "footnotes_created": footnotes_created,
        "footnotes_harvested": harvested,
        "entries_flagged": entry_flags,
        "citations_flagged": citation_flags,
        "list_reordered": bool(reorder),
        "heading_renamed": heading_op["after"] if heading_op else None,
        "flags": flags,
        "review_required": review or ["clean conversion — still worth a read-through"],
        "note": (
            "Heuristic text conversion. Flagged entries/citations were left "
            "verbatim; review every flag before treating the manuscript as "
            "converted."
        ),
    }


# ------------------------------------------------- notes-source harvesting

_NOTE_ARTICLE = re.compile(
    rf"^(?P<auth>.+?),\s*[“\"](?P<title>.+?)[,.]?[”\"]\s*"
    rf"(?P<cont>.+?)\s+(?P<vol>\d+)(?:,\s*no\.\s*(?P<iss>\w+))?\s*"
    rf"\((?P<year>{_YEAR})\)(?::\s*(?P<pp>[\d\-–]+))?\.?\s*$"
)
_NOTE_BOOK = re.compile(
    rf"^(?P<auth>.+?),\s*(?P<title>[^,(]+?)\s*"
    rf"\((?:(?P<place>[^:()]+):\s*)?(?P<pub>[^,()]+),\s*(?P<year>{_YEAR})\)"
    rf"(?:,\s*(?P<pp>[\d\-–]+))?\.?\s*$"
)
_NOTE_SHORT = re.compile(
    r"^(?P<fam>[A-Z][\w'’\-]+),\s*[“\"](?P<st>.+?)[,.]?[”\"]"
    r"(?:,?\s*(?P<pp>[\d\-–]+))?\.?\s*$"
)


def _match_note_part(part, by_key, by_fam):
    """(entry index, locator) for ONE citation inside a footnote, or
    (None, None) when it is not recognizably a citation."""
    part = part.strip()
    for pat in (_NOTE_ARTICLE, _NOTE_BOOK):
        m = pat.match(part)
        if m:
            fam_guess = m.group("auth").split()[-1].lower().strip(",.")
            hits = by_key.get((fam_guess, m.group("year").lower()), [])
            if len(hits) == 1:
                return hits[0], m.groupdict().get("pp")
            return None, None
    m = _NOTE_SHORT.match(part)
    if m:
        hits = by_fam.get(m.group("fam").lower(), [])
        if len(hits) == 1:
            return hits[0], m.group("pp")
    return None, None


def _plan_note_harvest(pkg, entries, by_key, by_fam, target, numbers):
    """Footnotes that are recognizably pure citations -> conversion ops.
    Multi-work notes ("A; B.") convert only when EVERY part resolves;
    mixed-content or unrecognized notes are flagged, never touched."""
    ops, flags = [], []
    if not pkg.has_part("word/footnotes.xml"):
        return ops, flags
    root = pkg.root("word/footnotes.xml")
    for note in root.findall(qn("w:footnote")):
        if note.get(qn("w:type")) is not None:
            continue
        nid = note.get(qn("w:id"))
        text = "\n".join(
            paragraph_text(p) for p in note.findall(qn("w:p"))
        ).strip()
        parts = re.split(r";\s+(?=[A-Z])", text) if "; " in text else [text]
        items = []
        for part in parts:
            idx, loc = _match_note_part(part, by_key, by_fam)
            if idx is None:
                items = None
                break
            items.append((idx, loc))
        if not items:
            flags.append({
                "citation": f"footnote {nid}: {text[:80]}",
                "paragraph_index": None,
                "problem": (
                    "footnote is not recognizably a pure citation (or a part "
                    "of it is ambiguous/mixed content); left alone"
                ),
            })
            continue
        if any(entries[idx]["parse_confidence"] != "full" for idx, _ in items):
            flags.append({
                "citation": f"footnote {nid}: {text[:80]}",
                "paragraph_index": None,
                "problem": "matches an entry that did not fully parse; left alone",
            })
            continue
        if target == "ieee":
            replacement = ", ".join(
                emit_intext_item("ieee", entries[idx], loc, numbers.get(idx))
                for idx, loc in items
            )
        elif target == "vancouver":
            replacement = None  # handled as a superscript number run
        else:
            replacement = "(" + "; ".join(
                emit_intext_item(target, entries[idx], loc, None)
                for idx, loc in items
            ) + ")"
        number_text = ",".join(
            str(numbers.get(idx)) for idx, _ in items if numbers.get(idx)
        )
        ops.append({
            "note_id": nid,
            "before": text[:120],
            "after": replacement or f"^{number_text}",
            "entry_indexes": [idx for idx, _ in items],
            "replacement": replacement,
            "number": number_text,
        })
    return ops, flags


def _apply_note_harvest(pkg, op) -> bool:
    """Replace one citation footnote with an in-text citation."""
    body = pkg.root()
    cfg = _notes._KINDS["footnote"]
    target_ref = None
    for ref in body.iter(qn(cfg["body_ref"])):
        if ref.get(qn("w:id")) == op["note_id"]:
            target_ref = ref
            break
    if target_ref is None:
        return False
    run = target_ref.getparent()
    p = run.getparent()
    while p is not None and etree.QName(p).localname != "p":
        p = p.getparent()
    if p is None:
        return False

    # Character offset of the reference run = text length of preceding runs.
    _, segments = _runmap.build_map(p)
    pos = 0
    for seg in segments:
        if seg.run is run or seg.run.getparent() is run:
            break
        # element order: stop when the segment's run comes after our run
        if run.getparent() is p and seg.run.getparent() is p:
            if p.index(seg.run) >= p.index(run):
                break
        pos = seg.end
    p.remove(run)
    pkg.mark_dirty()

    text, segments = _runmap.build_map(p)
    if op["replacement"] is None:  # Vancouver: superscript number stays put
        _insert_runs_at(p, min(pos, len(text)), [
            _make_text_run(str(op["number"]), superscript=True)
        ])
    else:
        # insert " (…)" before the sentence punctuation preceding the mark
        ins = min(pos, len(text))
        punct = None
        for m in _SENT_END.finditer(text[:ins]):
            punct = m.start()
        if punct is not None and ins - punct <= 2:
            _runmap.replace_range(
                p, segments, punct, punct + 1,
                " " + op["replacement"] + text[punct],
            )
        else:
            _runmap.replace_range(
                p, segments, max(ins - 1, 0), ins,
                (text[ins - 1] if ins else "") + " " + op["replacement"],
            )
    pkg.mark_dirty()

    root = pkg.root(cfg["part"])
    for note in root.findall(qn(cfg["note"])):
        if note.get(qn("w:id")) == op["note_id"]:
            root.remove(note)
            pkg.mark_dirty(cfg["part"])
            break
    return True


# =============================================== manuscript format (stage 3)

_MANUSCRIPT_STYLES = {
    "apa7": "apa-student",
    "apa7-student": "apa-student",
    "apa-student": "apa-student",
    "apa7-professional": "apa-professional",
    "apa-professional": "apa-professional",
    "mla9": "mla",
    "mla": "mla",
    "chicago17": "chicago",
    "chicago": "chicago",
    "turabian": "chicago",
    "chicago17-notes": "chicago",
}


def _set_heading_style_fmt(
    pkg: DocxPackage, level: int, *, align: str | None, bold: bool,
    italic: bool, indent_first_pt: float | None,
) -> None:
    from .text import ensure_heading_style

    style_id = ensure_heading_style(pkg, level)
    root = pkg.root("word/styles.xml")
    style = next(
        s for s in root.findall(qn("w:style"))
        if s.get(qn("w:styleId")) == style_id
    )
    ppr = style.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.SubElement(style, qn("w:pPr"))
    if align:
        jc = ppr.find(qn("w:jc"))
        if jc is None:
            jc = etree.SubElement(ppr, qn("w:jc"))
        jc.set(qn("w:val"), align)
    if indent_first_pt is not None:
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = etree.SubElement(ppr, qn("w:ind"))
        ind.set(qn("w:firstLine"), str(int(indent_first_pt * 20)))
    rpr = style.find(qn("w:rPr"))
    if rpr is None:
        rpr = etree.SubElement(style, qn("w:rPr"))
    for tag, want in (("w:b", bold), ("w:i", italic)):
        el = rpr.find(qn(tag))
        if want and el is None:
            etree.SubElement(rpr, qn(tag))
        elif not want and el is not None:
            rpr.remove(el)
    pkg.mark_dirty("word/styles.xml")


def _set_normal_size(pkg: DocxPackage, size_pt: int) -> None:
    root = pkg.root("word/styles.xml")
    normal = next(
        (
            s for s in root.findall(qn("w:style"))
            if s.get(qn("w:styleId")) == "Normal"
        ),
        None,
    )
    if normal is None:
        normal = etree.SubElement(root, qn("w:style"))
        normal.set(qn("w:type"), "paragraph")
        normal.set(qn("w:styleId"), "Normal")
        etree.SubElement(normal, qn("w:name")).set(qn("w:val"), "Normal")
    rpr = normal.find(qn("w:rPr"))
    if rpr is None:
        rpr = etree.SubElement(normal, qn("w:rPr"))
    for tag in ("w:sz", "w:szCs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rpr, qn(tag))
        el.set(qn("w:val"), str(size_pt * 2))
    pkg.mark_dirty("word/styles.xml")


def _patch_running_head(pkg: DocxPackage, section: int) -> bool:
    """Turn the just-written default header into 'TEXT<tab>PAGE' with a
    right tab stop at the text width."""
    from .furniture import _sect_prs, list_sections

    sects = _sect_prs(pkg)
    sp = sects[section]
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ref = next(
        (
            r for r in sp.findall(qn("w:headerReference"))
            if r.get(qn("w:type"), "default") == "default"
        ),
        None,
    )
    if ref is None:
        return False
    rid = ref.get(f"{{{r_ns}}}id")
    rels = pkg.root("word/_rels/document.xml.rels")
    target = next((r.get("Target") for r in rels if r.get("Id") == rid), None)
    if not target:
        return False
    part = "word/" + target.lstrip("/")
    root = pkg.root(part)
    p = root.find(qn("w:p"))
    if p is None:
        return False
    # content width from this section
    width_twips = 9360
    info = list_sections(pkg)
    if section < len(info):
        s = info[section]
        if "page_width_pt" in s and "margins_pt" in s:
            width_twips = int(
                (s["page_width_pt"] - s["margins_pt"]["left"] - s["margins_pt"]["right"]) * 20
            )
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = etree.Element(qn("w:pPr"))
        p.insert(0, ppr)
    jc = ppr.find(qn("w:jc"))
    if jc is not None:
        jc.set(qn("w:val"), "left")
    tabs = ppr.find(qn("w:tabs"))
    if tabs is None:
        tabs = etree.SubElement(ppr, qn("w:tabs"))
    tab = etree.SubElement(tabs, qn("w:tab"))
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(width_twips))
    # insert a tab run before the PAGE field
    fld = p.find(qn("w:fldSimple"))
    if fld is not None:
        tr = etree.Element(qn("w:r"))
        etree.SubElement(tr, qn("w:tab"))
        fld.addprevious(tr)
    pkg.mark_dirty(part)
    return True


def apply_manuscript_format(
    pkg: DocxPackage,
    style: str,
    *,
    running_head: str | None = None,
    author_last_name: str | None = None,
) -> dict:
    """Stage 3: page-level manuscript conventions for the styles that define
    them: APA 7 (student/professional), MLA 9, Chicago 17 / Turabian.

    Only conventions with a clear public definition are encoded; everything
    contested or context-dependent lands in not_applied. IEEE/Vancouver/ASA/
    Harvard page formats are journal-template-specific and are refused."""
    from .furniture import add_page_numbers, list_sections, set_header_footer, set_section_properties
    from .text import set_paragraph_format

    key = _MANUSCRIPT_STYLES.get((style or "").strip().lower())
    if key is None:
        raise WordMcpError(
            f"no defined manuscript format for {style!r}. Supported: apa7 "
            "(=apa7-student), apa7-professional, mla9, chicago17 (=turabian). "
            "IEEE/Vancouver/ASA/Harvard page formats are journal-specific — "
            "use the journal's template."
        )

    applied: list[str] = []
    not_applied: list[str] = []

    n_sections = max(len(list_sections(pkg)), 1)
    for s in range(n_sections):
        set_section_properties(
            pkg, section=s,
            margins_pt={"top": 72, "bottom": 72, "left": 72, "right": 72},
        )
    applied.append(f"1-inch margins on {n_sections} section(s)")

    # double spacing
    paras = _body_paragraph_elements(pkg)
    outline_idx = {h["paragraph_index"] for h in get_outline(pkg)}
    if key == "chicago":
        body_idx = [
            i for i, el in paras
            if i not in outline_idx and paragraph_text(el).strip()
        ]
    else:
        body_idx = [i for i, el in paras if paragraph_text(el).strip()]
    if body_idx:
        set_paragraph_format(
            pkg, body_idx,
            {"line_spacing": 2.0, "space_before_pt": 0, "space_after_pt": 0},
        )
    applied.append(f"double spacing on {len(body_idx)} paragraph(s)")

    if key in ("apa-student", "apa-professional"):
        _set_heading_style_fmt(pkg, 1, align="center", bold=True, italic=False, indent_first_pt=None)
        _set_heading_style_fmt(pkg, 2, align="left", bold=True, italic=False, indent_first_pt=None)
        _set_heading_style_fmt(pkg, 3, align="left", bold=True, italic=True, indent_first_pt=None)
        _set_heading_style_fmt(pkg, 4, align="left", bold=True, italic=False, indent_first_pt=36)
        _set_heading_style_fmt(pkg, 5, align="left", bold=True, italic=True, indent_first_pt=36)
        applied.append("APA heading styles L1-L5 (formatting)")
        not_applied.append(
            "APA L4/L5 run-in headings: the heading text cannot be merged "
            "into the following paragraph automatically"
        )
        not_applied.append(
            "heading TEXT casing left unchanged (title-casing proper nouns "
            "needs human judgment)"
        )
        not_applied.append(
            "0.5in paragraph indent not applied: title page, abstract, and "
            "block quotes cannot be told apart from body text reliably"
        )
        if key == "apa-professional":
            head = running_head
            if not head:
                if pkg.has_part("docProps/core.xml"):
                    core = pkg.root("docProps/core.xml")
                    el = core.find("{http://purl.org/dc/elements/1.1/}title")
                    if el is not None and el.text:
                        head = el.text
                if head:
                    not_applied.append(
                        "running head derived from the document title; "
                        "confirm it is the intended shortened title"
                    )
                else:
                    not_applied.append(
                        "running head skipped: no running_head given and the "
                        "document has no title property"
                    )
            if head:
                head = head.upper()[:50]
                set_header_footer(
                    pkg, "header", head, section=0, alignment="left",
                    include_page_number=True,
                )
                if _patch_running_head(pkg, 0):
                    applied.append(f"running head '{head}' + right page number")
                else:
                    applied.append(f"running head '{head}' (page number inline)")
        else:
            set_header_footer(
                pkg, "header", "", section=0, alignment="right",
                include_page_number=True,
            )
            applied.append("page number top right (student paper: no running head)")

    elif key == "mla":
        name = author_last_name
        if not name and pkg.has_part("docProps/core.xml"):
            core = pkg.root("docProps/core.xml")
            el = core.find("{http://purl.org/dc/elements/1.1/}creator")
            if el is not None and el.text:
                name = el.text.split()[-1]
                not_applied.append(
                    "header surname derived from the document author "
                    "property; confirm"
                )
        if name:
            set_header_footer(
                pkg, "header", f"{name} ", section=0, alignment="right",
                include_page_number=True,
            )
            applied.append(f"MLA header '{name} <page>' top right")
        else:
            not_applied.append(
                "MLA name/page header skipped: no author_last_name given and "
                "no author metadata"
            )
        not_applied.append(
            "0.5in first-line indent not applied: the MLA heading block and "
            "title cannot be told apart from body text reliably"
        )

    elif key == "chicago":
        _set_normal_size(pkg, 12)
        applied.append("12pt base (Normal style)")
        _notes._ensure_styles(pkg, "footnote")
        applied.append(
            "footnote text single-spaced 10pt (FootnoteText style)"
        )
        add_page_numbers(pkg, section=0, position="footer", alignment="center")
        applied.append("page number bottom center")
        not_applied.append(
            "Turabian front-matter conventions (title page, roman-numeral "
            "front matter) are document-specific"
        )

    # hanging indent for the reference list (all three formats use one)
    try:
        ref_start, ref_end = _locate_reference_list(pkg)
        entry_idx = [
            i for i, el in _body_paragraph_elements(pkg)
            if i > ref_start and (ref_end is None or i < ref_end)
            and paragraph_text(el).strip()
        ]
        if entry_idx:
            set_paragraph_format(
                pkg, entry_idx,
                {"indent_left_pt": 36, "first_line_indent_pt": -36},
            )
            applied.append(
                f"hanging indent on {len(entry_idx)} reference entries"
            )
    except TargetNotFound:
        not_applied.append("no reference list found; hanging indent skipped")

    return {
        "style": style,
        "resolved_as": key,
        "applied": applied,
        "not_applied": not_applied,
        "note": (
            "Only publicly well-defined conventions are encoded; every "
            "skipped or approximated item is listed in not_applied."
        ),
    }
