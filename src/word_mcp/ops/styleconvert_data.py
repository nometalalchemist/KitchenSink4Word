"""Style tables and emitters for publication-style conversion.

Everything here is PURE: functions take a parsed entry model (built by
ops/styleconvert.py) and return text segments, never touching a document.
A segment is a (text, italic) tuple; the caller renders segments as runs.

Honesty notes baked into the emitters:
- Emission is only as good as the parse. Emitters never invent data: a
  missing field is omitted (with the style's punctuation adjusted), never
  guessed.
- Title-case <-> sentence-case transforms are heuristic. Sentence-casing a
  title-case title cannot know proper nouns; tokens with internal or full
  capitals are preserved and the entry is flagged for review.
- Styles that prefer full given names (Chicago, MLA, ASA) get initials when
  the source style (e.g. APA) only recorded initials — flagged, not invented.
"""

from __future__ import annotations

import re

# --------------------------------------------------------------- style table

# Canonical style keys -> metadata.
#   system: how in-text citations work in that style.
#   ordering: how the reference list is ordered.
STYLE_INFO = {
    "apa7": {
        "label": "APA 7",
        "system": "author-date",
        "ordering": "alphabetical",
        "ref_heading": "References",
        "title_case": "sentence",
    },
    "chicago17-author-date": {
        "label": "Chicago 17 author-date",
        "system": "author-date",
        "ordering": "alphabetical",
        "ref_heading": "References",
        "title_case": "title",
    },
    "chicago17-notes": {
        # Turabian is Chicago notes-bibliography for practical purposes;
        # the alias table below maps "turabian" here.
        "label": "Chicago 17 notes-bibliography (Turabian)",
        "system": "notes",
        "ordering": "alphabetical",
        "ref_heading": "Bibliography",
        "title_case": "title",
    },
    "mla9": {
        "label": "MLA 9",
        "system": "author-page",
        "ordering": "alphabetical",
        "ref_heading": "Works Cited",
        "title_case": "title",
    },
    "harvard": {
        "label": "Harvard (generic)",
        "system": "author-date",
        "ordering": "alphabetical",
        "ref_heading": "References",
        "title_case": "sentence",
    },
    "ieee": {
        "label": "IEEE",
        "system": "numbered-bracket",
        "ordering": "citation-order",
        "ref_heading": "References",
        "title_case": "sentence",
    },
    "vancouver": {
        "label": "Vancouver",
        "system": "numbered-superscript",
        "ordering": "citation-order",
        "ref_heading": "References",
        "title_case": "sentence",
    },
    "asa": {
        "label": "ASA (Sage)",
        "system": "author-date",
        "ordering": "alphabetical",
        "ref_heading": "References",
        "title_case": "title",
    },
}

STYLE_ALIASES = {
    "apa": "apa7",
    "apa7": "apa7",
    "apa 7": "apa7",
    "chicago-author-date": "chicago17-author-date",
    "chicago17-author-date": "chicago17-author-date",
    "chicago author-date": "chicago17-author-date",
    "chicago-ad": "chicago17-author-date",
    "chicago-notes": "chicago17-notes",
    "chicago17-notes": "chicago17-notes",
    "chicago notes-bibliography": "chicago17-notes",
    "chicago-nb": "chicago17-notes",
    "turabian": "chicago17-notes",
    "mla": "mla9",
    "mla9": "mla9",
    "mla 9": "mla9",
    "harvard": "harvard",
    "ieee": "ieee",
    "vancouver": "vancouver",
    "asa": "asa",
    "asa-sage": "asa",
    "sage": "asa",
}


def canonical_style(name: str) -> str | None:
    return STYLE_ALIASES.get((name or "").strip().lower())


# ------------------------------------------------------------------- casing

# Small words kept lowercase in title case unless first/last/after-colon.
_SMALL_WORDS = {
    "a", "an", "the", "and", "but", "or", "nor", "for", "so", "yet",
    "as", "at", "by", "in", "of", "off", "on", "onto", "out", "per",
    "to", "up", "via", "with", "from", "into", "over", "upon",
}

_WORD = re.compile(r"[^\s]+")


def _cap_token(tok: str) -> str:
    for i, ch in enumerate(tok):
        if ch.isalpha():
            return tok[:i] + ch.upper() + tok[i + 1:]
    return tok


def title_case(s: str) -> str:
    """Chicago-ish headline case. Tokens with internal capitals (DPRK,
    McNamara) are preserved as-is."""
    words = s.split(" ")
    out = []
    force_next = True  # first word, and word after a colon/period/dash
    for i, w in enumerate(words):
        if not w:
            out.append(w)
            continue
        bare = re.sub(r"[^\w'\-]", "", w)
        has_inner_cap = any(c.isupper() for c in bare[1:])
        last = i == len(words) - 1
        if has_inner_cap:
            out.append(w)
        elif force_next or last or bare.lower() not in _SMALL_WORDS:
            out.append(_cap_token(w))
        else:
            out.append(w.lower())
        force_next = bool(re.search(r"[:.?!—-]$", w))
    return " ".join(out)


def sentence_case(s: str) -> str:
    """Heuristic sentence case: lowercase every plainly-capitalized word
    except the first and the first after a colon/period/question mark.
    Tokens that are ALL-CAPS or mixed-case (acronyms, McX names) are
    preserved — proper nouns written as plain Xxxx words CANNOT be detected
    and will be lowercased; callers must flag the entry for review."""
    words = s.split(" ")
    out = []
    force_cap = True
    for w in words:
        if not w:
            out.append(w)
            continue
        bare = re.sub(r"[^\w'\-]", "", w)
        has_inner_cap = any(c.isupper() for c in bare[1:])
        if force_cap:
            out.append(_cap_token(w))
        elif has_inner_cap:
            out.append(w)  # acronym / mixed caps: keep
        else:
            out.append(w.lower())
        force_cap = bool(re.search(r"[:.?!]$", w))
    return " ".join(out)


def looks_title_case(s: str) -> bool:
    """True when most non-small words beyond the first are capitalized."""
    words = [re.sub(r"[^\w'\-]", "", w) for w in s.split(" ")[1:] if w]
    words = [w for w in words if w and w.lower() not in _SMALL_WORDS]
    if not words:
        return False
    caps = sum(1 for w in words if w[0].isupper())
    return caps / len(words) > 0.5


def cased_title(title: str, target_case: str, flags: list[str]) -> str:
    """Convert a title toward the target case convention, flagging heuristic
    downcasing (proper nouns cannot be detected)."""
    if target_case == "title":
        if not looks_title_case(title):
            return title_case(title)
        return title
    # sentence case wanted
    if looks_title_case(title):
        flags.append(
            "title sentence-cased heuristically; check proper nouns: "
            + title[:60]
        )
        return sentence_case(title)
    return title


# -------------------------------------------------------------------- names


def _to_initials(given: str) -> str:
    """'Ian' -> 'I.'; 'John A.' -> 'J. A.'; 'J. A.' unchanged; 'Jean-Paul'
    -> 'J.-P.'"""
    if not given:
        return ""
    parts = given.replace(".", ". ").split()
    out = []
    for p in parts:
        if not p:
            continue
        if "-" in p:
            out.append(
                "-".join(f"{seg[0]}." for seg in p.split("-") if seg)
            )
        else:
            out.append(p[0] + ".")
    return " ".join(out)


def _given_is_initials(given: str) -> bool:
    return bool(re.fullmatch(r"(?:[A-Z]\.?[\s\-]*)+", given or ""))


def fmt_author(a: dict, mode: str) -> str:
    """mode: inverted-initials | inverted-full | natural-full |
    natural-initials | vancouver."""
    if "literal" in a:
        return a["literal"]
    fam, giv = a.get("family", ""), a.get("given", "")
    if mode == "inverted-initials":
        ini = _to_initials(giv)
        return f"{fam}, {ini}" if ini else fam
    if mode == "inverted-full":
        return f"{fam}, {giv}" if giv else fam
    if mode == "natural-full":
        return f"{giv} {fam}".strip()
    if mode == "natural-initials":
        ini = _to_initials(giv)
        return f"{ini} {fam}".strip()
    if mode == "vancouver":
        ini = _to_initials(giv).replace(".", "").replace(" ", "")
        return f"{fam} {ini}".strip()
    raise ValueError(mode)


def authors_string(style: str, e: dict, flags: list[str]) -> str:
    """Reference-list author string per style. Never invents given names:
    styles preferring full names fall back to whatever the source had."""
    auths = e.get("authors") or []
    if not auths:
        return ""
    if style in ("chicago17-author-date", "chicago17-notes", "mla9", "asa"):
        if any(
            _given_is_initials(a.get("given", "")) for a in auths
            if "literal" not in a
        ):
            flags.append(
                "source gave initials only; this style prefers full given "
                "names (cannot be reconstructed)"
            )

    tail = " et al." if e.get("et_al") else ""

    if style in ("apa7",):
        parts = [fmt_author(a, "inverted-initials") for a in auths]
        if len(parts) == 1:
            s = parts[0]
        else:
            s = ", ".join(parts[:-1]) + ", & " + parts[-1]
        return s + tail
    if style == "harvard":
        parts = [fmt_author(a, "inverted-initials") for a in auths]
        s = parts[0] if len(parts) == 1 else (
            ", ".join(parts[:-1]) + " and " + parts[-1]
        )
        return s + tail
    if style in ("chicago17-author-date", "chicago17-notes", "asa"):
        parts = [fmt_author(auths[0], "inverted-full")] + [
            fmt_author(a, "natural-full") for a in auths[1:]
        ]
        s = parts[0] if len(parts) == 1 else (
            ", ".join(parts[:-1]) + ", and " + parts[-1]
        )
        return s + tail
    if style == "mla9":
        if len(auths) >= 3:
            return fmt_author(auths[0], "inverted-full") + ", et al."
        parts = [fmt_author(auths[0], "inverted-full")] + [
            fmt_author(a, "natural-full") for a in auths[1:]
        ]
        s = parts[0] if len(parts) == 1 else parts[0] + ", and " + parts[1]
        return s + tail
    if style == "ieee":
        parts = [fmt_author(a, "natural-initials") for a in auths]
        s = parts[0] if len(parts) == 1 else (
            ", ".join(parts[:-1]) + " and " + parts[-1]
        )
        return s + tail
    if style == "vancouver":
        shown = auths[:6]
        parts = [fmt_author(a, "vancouver") for a in shown]
        s = ", ".join(parts)
        if len(auths) > 6 or e.get("et_al"):
            s += ", et al"
        return s
    raise ValueError(style)


def intext_label(e: dict, *, amp: bool = False, max_names: int = 2) -> str:
    """In-text author label: 'Hurd', 'Hurd & Lake' / 'Hurd and Lake',
    'Hurd et al.' (3+ authors)."""
    auths = e.get("authors") or []
    if not auths:
        return "Anonymous"
    fams = [
        a.get("family") or a.get("literal", "") for a in auths
    ]
    if e.get("et_al") or len(fams) > max_names:
        return f"{fams[0]} et al."
    if len(fams) == 1:
        return fams[0]
    joiner = " & " if amp else " and "
    return joiner.join(fams)


# ----------------------------------------------------------- entry emission

_ENDASH = "–"


def norm_pages(pages: str | None) -> str | None:
    """Tidy spacing around a page-range separator but PRESERVE the source's
    own dash character — silently swapping "50-52" to "50–52" made
    round-trips text-unstable (v1.5 adversarial finding F5). Style guides do
    prefer the en dash; that is the author's call, not a silent rewrite."""
    if not pages:
        return None
    return re.sub(r"\s*([-–—])\s*", r"\1", pages.strip())


class SegBuilder:
    def __init__(self):
        self.parts: list[tuple[str, bool]] = []

    def add(self, text: str | None, italic: bool = False) -> None:
        if text:
            self.parts.append((text, italic))

    def segments(self) -> list[tuple[str, bool]]:
        # merge adjacent same-format parts
        merged: list[tuple[str, bool]] = []
        for text, it in self.parts:
            if merged and merged[-1][1] == it:
                merged[-1] = (merged[-1][0] + text, it)
            else:
                merged.append((text, it))
        return merged

    def text(self) -> str:
        return "".join(t for t, _ in self.parts)


def _editors_string(style: str, e: dict) -> str | None:
    eds = e.get("editors") or []
    if not eds:
        return None
    if style == "apa7":
        parts = [
            f"{_to_initials(a.get('given', ''))} {a.get('family', '')}".strip()
            if "literal" not in a else a["literal"]
            for a in eds
        ]
        joined = parts[0] if len(parts) == 1 else (
            ", ".join(parts[:-1]) + ", & " + parts[-1]
        )
        suffix = " (Ed.)," if len(eds) == 1 else " (Eds.),"
        return joined + suffix
    parts = [fmt_author(a, "natural-full") for a in eds]
    return (
        parts[0] if len(parts) == 1 else ", ".join(parts[:-1]) + " and " + parts[-1]
    )


def _doi_url_tail(e: dict) -> str | None:
    return e.get("doi") or e.get("url")


def emit_entry(
    style: str, e: dict, number: int | None = None
) -> tuple[list[tuple[str, bool]], list[str]]:
    """Render a parsed entry in the target style.

    Returns (segments, flags). segments = [(text, italic)]. Only entries
    with parse_confidence == 'full' should reach this function; the caller
    enforces that."""
    flags: list[str] = []
    b = SegBuilder()
    info = STYLE_INFO[style]
    case = info["title_case"]
    etype = e.get("type", "unknown")
    title = cased_title(e.get("title") or "", case, flags)
    pages = norm_pages(e.get("pages"))
    year = e.get("year") or "n.d."
    vol, iss = e.get("volume"), e.get("issue")
    container = e.get("container")
    pub, place = e.get("publisher"), e.get("place")
    tailref = _doi_url_tail(e)
    A = authors_string(style, e, flags)

    def end_period(builder: SegBuilder) -> None:
        txt = builder.text()
        if txt and not txt.rstrip().endswith((".", "?", "!")):
            builder.add(".")

    if style == "apa7":
        yr = year + (f", {e['year_extra']}" if e.get("year_extra") else "")
        b.add(f"{A} " if A else "")
        b.add(f"({yr}). ")
        if etype == "article":
            b.add(f"{title}. ")
            b.add(container, italic=True)
            if vol:
                b.add(", ")
                b.add(vol, italic=True)
                if iss:
                    b.add(f"({iss})")
                if pages:
                    b.add(f", {pages}")
                b.add(".")
            else:
                b.add(".")
        elif etype == "chapter":
            b.add(f"{title}. In ")
            eds = _editors_string("apa7", e)
            if eds:
                b.add(f"{eds} ")
            b.add(container, italic=True)
            if pages:
                b.add(f" (pp. {pages})")
            b.add(". ")
            b.add(pub)
            end_period(b)
        elif etype == "book":
            b.add(title, italic=True)
            b.add(". ")
            b.add(pub)
            end_period(b)
        else:  # report / web / unknown-but-full
            b.add(f"{title}. ")
            b.add(container or pub)
            end_period(b)
        if tailref:
            b.add(f" {tailref}")

    elif style in ("chicago17-author-date", "asa"):
        if A:
            b.add(A)
            b.add(f" {year}. " if A.endswith(".") else f". {year}. ")
        else:
            b.add(f"{year}. ")
        if etype == "article":
            b.add(f"“{title}.” ")
            b.add(container, italic=True)
            if vol:
                if style == "asa":
                    b.add(f" {vol}")
                    if iss:
                        b.add(f"({iss})")
                    if pages:
                        b.add(f":{pages}")
                else:
                    b.add(f" {vol}")
                    if iss:
                        b.add(f" ({iss})")
                    if pages:
                        b.add(f": {pages}")
            b.add(".")
        elif etype == "chapter":
            b.add(f"“{title}.” In ")
            b.add(container, italic=True)
            eds = _editors_string(style, e)
            if eds:
                b.add(f", edited by {eds}")
            if pages:
                b.add(f", {pages}")
            b.add(". ")
            if place:
                b.add(f"{place}: ")
            b.add(pub)
            end_period(b)
        else:
            b.add(title, italic=True)
            b.add(". ")
            if place:
                b.add(f"{place}: ")
            b.add(pub)
            end_period(b)
        if tailref:
            b.add(f" {tailref}")

    elif style == "chicago17-notes":
        # Bibliography form (notes themselves are emitted by emit_note).
        if A:
            b.add(A)
            b.add(" " if A.endswith(".") else ". ")
        if etype == "article":
            b.add(f"“{title}.” ")
            b.add(container, italic=True)
            if vol:
                b.add(f" {vol}")
                if iss:
                    b.add(f", no. {iss}")
            b.add(f" ({year})")
            if pages:
                b.add(f": {pages}")
            b.add(".")
        elif etype == "chapter":
            b.add(f"“{title}.” In ")
            b.add(container, italic=True)
            eds = _editors_string(style, e)
            if eds:
                b.add(f", edited by {eds}")
            if pages:
                b.add(f", {pages}")
            b.add(". ")
            if place:
                b.add(f"{place}: ")
            b.add(pub)
            b.add(f", {year}.")
        else:
            b.add(title, italic=True)
            b.add(". ")
            if place:
                b.add(f"{place}: ")
            b.add(pub)
            b.add(f", {year}.")
        if tailref:
            b.add(f" {tailref}")

    elif style == "mla9":
        if A:
            b.add(A)
            b.add(" " if A.endswith(".") else ". ")
        if etype == "article":
            b.add(f"“{title}.” ")
            b.add(container, italic=True)
            if vol:
                b.add(f", vol. {vol}")
            if iss:
                b.add(f", no. {iss}")
            b.add(f", {year}")
            if pages:
                b.add(f", pp. {pages}")
            b.add(".")
        elif etype == "chapter":
            b.add(f"“{title}.” ")
            b.add(container, italic=True)
            eds = _editors_string(style, e)
            if eds:
                b.add(f", edited by {eds}")
            b.add(f", {pub}, {year}" if pub else f", {year}")
            if pages:
                b.add(f", pp. {pages}")
            b.add(".")
        else:
            b.add(title, italic=True)
            b.add(". ")
            b.add(pub)
            b.add(f", {year}.")
        if tailref:
            b.add(f" {tailref}")

    elif style == "harvard":
        b.add(f"{A} " if A else "")
        b.add(f"({year}) ")
        if etype == "article":
            b.add(f"‘{title}’, ")
            b.add(container, italic=True)
            if vol:
                b.add(f", {vol}")
                if iss:
                    b.add(f"({iss})")
            if pages:
                b.add(f", pp. {pages}")
            b.add(".")
        elif etype == "chapter":
            b.add(f"‘{title}’, in ")
            eds = _editors_string(style, e)
            if eds:
                b.add(f"{eds} (ed.) ")
            b.add(container, italic=True)
            b.add(". ")
            if place:
                b.add(f"{place}: ")
            b.add(pub)
            if pages:
                b.add(f", pp. {pages}")
            end_period(b)
        else:
            b.add(title, italic=True)
            b.add(". ")
            if place:
                b.add(f"{place}: ")
            b.add(pub)
            end_period(b)
        if tailref:
            b.add(f" {tailref}")

    elif style == "ieee":
        if number is not None:
            b.add(f"[{number}] ")
        b.add(f"{A}, " if A else "")
        if etype == "article":
            b.add(f"“{title},” ")
            b.add(container, italic=True)
            if vol:
                b.add(f", vol. {vol}")
            if iss:
                b.add(f", no. {iss}")
            if pages:
                b.add(f", pp. {pages}")
            b.add(f", {year}.")
        elif etype == "chapter":
            b.add(f"“{title},” in ")
            b.add(container, italic=True)
            if place:
                b.add(f". {place}: {pub}" if pub else f". {place}")
            elif pub:
                b.add(f". {pub}")
            b.add(f", {year}")
            if pages:
                b.add(f", pp. {pages}")
            b.add(".")
        else:
            b.add(title, italic=True)
            b.add(". ")
            if place:
                b.add(f"{place}: ")
            b.add(pub)
            b.add(f", {year}.")
        if tailref:
            b.add(f" {tailref}")

    elif style == "vancouver":
        if number is not None:
            b.add(f"{number}. ")
        b.add(f"{A}. " if A else "")
        b.add(f"{title}. ")
        if etype == "article":
            b.add(f"{container}. " if container else "")
            b.add(f"{year}")
            if vol:
                b.add(f";{vol}")
                if iss:
                    b.add(f"({iss})")
                if pages:
                    b.add(f":{pages.replace(_ENDASH, '-')}")
            b.add(".")
        else:
            if place:
                b.add(f"{place}: ")
            if pub:
                b.add(f"{pub}; ")
            b.add(f"{year}.")
        if tailref:
            b.add(f" {tailref}")

    else:
        raise ValueError(f"no emitter for style {style!r}")

    return b.segments(), flags


# ------------------------------------------------------------ note emission


def short_title(title: str) -> str:
    """First few significant words of a title, for Chicago short notes."""
    words = title.split(":")[0].split()
    while words and words[0].lower() in ("the", "a", "an"):
        words = words[1:]
    words = words[:4]
    while words and words[-1].lower() in _SMALL_WORDS:
        words.pop()
    return " ".join(words).rstrip(",.;")


def emit_note(
    e: dict, locator: str | None, *, first_use: bool
) -> list[tuple[str, bool]]:
    """Chicago 17 footnote text for one cited work (long note on first use,
    short note after). Segments as (text, italic)."""
    b = SegBuilder()
    etype = e.get("type", "unknown")
    flags: list[str] = []
    title_tc = cased_title(e.get("title") or "", "title", flags)
    year = e.get("year") or "n.d."
    pages = norm_pages(e.get("pages"))
    loc = norm_pages(locator)
    auths = e.get("authors") or []

    if not first_use:
        fams = [a.get("family") or a.get("literal", "") for a in auths]
        if len(fams) > 3 or e.get("et_al"):
            name = f"{fams[0]} et al." if fams else "Anonymous"
        else:
            name = ", ".join(fams[:-1]) + (" and " + fams[-1] if len(fams) > 1 else fams[0] if fams else "Anonymous")
        b.add(f"{name}, ")
        st = short_title(title_tc)
        if etype in ("article", "chapter"):
            b.add(f"“{st},” " if loc else f"“{st}.”")
        else:
            b.add(st, italic=True)
            b.add(", " if loc else ".")
        if loc:
            b.add(f"{loc}.")
        return b.segments()

    # Long note: authors in natural order.
    parts = [fmt_author(a, "natural-full") for a in auths]
    if len(parts) > 3 or e.get("et_al"):
        name = parts[0] + " et al." if parts else "Anonymous"
    elif len(parts) >= 2:
        name = ", ".join(parts[:-1]) + " and " + parts[-1]
    else:
        name = parts[0] if parts else "Anonymous"
    b.add(f"{name}, ")
    if etype == "article":
        b.add(f"“{title_tc},” ")
        b.add(e.get("container"), italic=True)
        if e.get("volume"):
            b.add(f" {e['volume']}")
            if e.get("issue"):
                b.add(f", no. {e['issue']}")
        b.add(f" ({year})")
        cite_pages = loc or pages
        if cite_pages:
            b.add(f": {cite_pages}")
        b.add(".")
    elif etype == "chapter":
        b.add(f"“{title_tc},” in ")
        b.add(e.get("container"), italic=True)
        inner = []
        if e.get("place"):
            inner.append(f"{e['place']}: {e.get('publisher') or ''}".rstrip(": "))
        elif e.get("publisher"):
            inner.append(e["publisher"])
        inner.append(year)
        b.add(f" ({', '.join(inner)})")
        if loc or pages:
            b.add(f", {loc or pages}")
        b.add(".")
    else:
        b.add(title_tc, italic=True)
        inner = []
        if e.get("place"):
            inner.append(f"{e['place']}: {e.get('publisher') or ''}".rstrip(": "))
        elif e.get("publisher"):
            inner.append(e["publisher"])
        inner.append(year)
        b.add(f" ({', '.join(inner)})")
        if loc:
            b.add(f", {loc}")
        b.add(".")
    return b.segments()


# ---------------------------------------------------- in-text item emission


def emit_intext_item(
    style: str, e: dict, locator: str | None, number: int | None
) -> str:
    """One cited work as it appears INSIDE a parenthetical/bracket for the
    target style (without the surrounding parens/brackets)."""
    loc = norm_pages(locator)
    if style == "apa7":
        label = intext_label(e, amp=True)
        s = f"{label}, {e.get('year') or 'n.d.'}"
        if loc:
            s += f", pp. {loc}" if _ENDASH in loc or "," in loc else f", p. {loc}"
        return s
    if style == "chicago17-author-date":
        label = intext_label(e)
        s = f"{label} {e.get('year') or 'n.d.'}"
        if loc:
            s += f", {loc}"
        return s
    if style == "harvard":
        label = intext_label(e)
        s = f"{label}, {e.get('year') or 'n.d.'}"
        if loc:
            s += f", pp. {loc}" if _ENDASH in loc or "," in loc else f", p. {loc}"
        return s
    if style == "asa":
        label = intext_label(e)
        s = f"{label} {e.get('year') or 'n.d.'}"
        if loc:
            s += f":{loc}"
        return s
    if style == "mla9":
        label = intext_label(e)
        return f"{label} {loc}" if loc else label
    if style == "ieee":
        return f"[{number}, p. {loc}]" if loc else f"[{number}]"
    raise ValueError(style)


def emit_intext_year_only(style: str, e: dict, locator: str | None) -> str:
    """The parenthetical part of a NARRATIVE citation ('Hurd (1999)') for
    author-date targets."""
    loc = norm_pages(locator)
    year = e.get("year") or "n.d."
    if style == "apa7" or style == "harvard":
        s = year
        if loc:
            s += f", pp. {loc}" if _ENDASH in loc or "," in loc else f", p. {loc}"
        return f"({s})"
    if style == "chicago17-author-date":
        return f"({year}, {loc})" if loc else f"({year})"
    if style == "asa":
        return f"({year}:{loc})" if loc else f"({year})"
    raise ValueError(style)
