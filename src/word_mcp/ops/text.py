"""Text and formatting edit operations.

All mutating functions call pkg.mark_dirty() on the parts they touch and return
a summary dict. Nothing here writes to disk; the caller decides when to save.
"""

from __future__ import annotations

import copy
import re

from lxml import etree

from ..core.errors import AmbiguousTarget, TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from . import _regex, _runmap
from .read import body_items, paragraph_text

# ------------------------------------------------------------ search & replace

_TEXT_PARTS = {
    "body": ["word/document.xml"],
    "footnotes": ["word/footnotes.xml", "word/endnotes.xml"],
}


def _replace_parts(pkg: DocxPackage, scope: str) -> list[str]:
    parts: list[str] = []
    if scope in ("body", "all"):
        parts.append("word/document.xml")
    if scope in ("footnotes", "all"):
        parts += [
            p
            for p in ("word/footnotes.xml", "word/endnotes.xml")
            if pkg.has_part(p)
        ]
    if scope in ("headers", "all"):
        parts += [
            p
            for p in pkg.part_names()
            if re.fullmatch(r"word/(header|footer)\d+\.xml", p)
        ]
    if not parts:
        raise WordMcpError(f"unknown scope: {scope!r}")
    return parts


def search_and_replace(
    pkg: DocxPackage,
    replacements: list[dict],
    *,
    scope: str = "body",
    max_replacements: int | None = None,
    track: bool = False,
    author: str = "Claude",
) -> dict:
    """Batch replace. Each item: {find, replace, regex?: bool}.

    Matching is per paragraph (a match cannot span paragraphs). Replacement
    inherits the formatting of the first character it replaces. Later pairs see
    the results of earlier pairs.

    max_replacements is a blast-radius guard: if the total match count would
    exceed it, NOTHING is replaced and the error reports the count. Use it
    with broad regexes (.*-class patterns can rewrite the whole document).
    """
    counts = {item["find"]: 0 for item in replacements}

    # Zero-width-only regex pre-scan: a pattern that MATCHES but only at
    # zero width (pure lookarounds like (?<=foo)) used to fall through to
    # "0 replaced" with no explanation, since zero-length matches are
    # always skipped. Refuse loudly before touching anything, consistent
    # with the empty-regex preview refusal (field test, 2026-09-03).
    if any(item.get("regex") for item in replacements):
        zw_raw: dict[int, int] = {}
        zw_eff: dict[int, int] = {}
        for part in _replace_parts(pkg, scope):
            for p in pkg.root(part).iter(qn("w:p")):
                if _runmap._in_textbox(p, pkg.root(part)):
                    continue
                text, _ = _runmap.build_map(p)
                for i, item in enumerate(replacements):
                    if not item.get("regex"):
                        continue
                    for m in _regex.finditer(item["find"], text):
                        zw_raw[i] = zw_raw.get(i, 0) + 1
                        if m.start() != m.end():
                            zw_eff[i] = zw_eff.get(i, 0) + 1
        offenders = [
            replacements[i]["find"]
            for i in zw_raw
            if not zw_eff.get(i)
        ]
        if offenders:
            raise WordMcpError(
                f"regex pattern(s) {offenders} matched only zero-width "
                "positions (lookarounds such as (?<=x) match a position, "
                "not text), and zero-length matches are always skipped. "
                "Nothing was changed; include the text to be replaced in "
                "the pattern (e.g. (?<=x)y instead of a bare lookbehind)."
            )

    if max_replacements is not None:
        projected = 0
        for part in _replace_parts(pkg, scope):
            for p in pkg.root(part).iter(qn("w:p")):
                if _runmap._in_textbox(p, pkg.root(part)):
                    continue
                text, _ = _runmap.build_map(p)
                for item in replacements:
                    if item.get("regex"):
                        projected += sum(
                            1
                            for m in _regex.finditer(item["find"], text)
                            if m.start() != m.end()
                        )
                    elif item["find"]:
                        projected += text.count(item["find"])
        if projected > max_replacements:
            raise WordMcpError(
                f"would make {projected} replacements, over the "
                f"max_replacements guard of {max_replacements}; nothing was "
                "changed — narrow the pattern or raise the limit"
            )
    for part in _replace_parts(pkg, scope):
        dirty = False
        for p in pkg.root(part).iter(qn("w:p")):
            # Paragraphs nested inside text boxes are a separate story that
            # every read tool (and preview_replace) excludes — editing them
            # here made preview and reality diverge. Box content is edited
            # only via set_textbox_text.
            if _runmap._in_textbox(p, pkg.root(part)):
                continue
            for item in replacements:
                find, repl = item["find"], item["replace"]
                use_regex = bool(item.get("regex"))
                # One snapshot per (paragraph, pair): find every non-overlapping
                # match, then apply right-to-left so earlier offsets stay valid
                # and replacement text is never re-matched (a replacement that
                # contains its own find string must not loop).
                text, segments = _runmap.build_map(p)
                if use_regex:
                    matches = [
                        (m.start(), m.end(), m.expand(repl))
                        for m in _regex.finditer(find, text)
                        if m.start() != m.end()
                    ]
                else:
                    if not find:
                        continue
                    matches = []
                    pos = 0
                    while True:
                        pos = text.find(find, pos)
                        if pos < 0:
                            break
                        matches.append((pos, pos + len(find), repl))
                        pos += len(find)
                for start, end, actual in reversed(matches):
                    if track:
                        from . import _tracked

                        _tracked.tracked_replace_range(
                            p,
                            start,
                            end,
                            actual,
                            author=author,
                            rev_id=_tracked.next_rev_id(pkg.root(part)),
                            date=_tracked.now_iso(),
                        )
                    else:
                        _runmap.replace_range(p, segments, start, end, actual)
                    counts[find] += 1
                    dirty = True
        if dirty:
            pkg.mark_dirty(part)
    result = {"replaced": counts, "total": sum(counts.values())}
    if track:
        result["tracked_as"] = author
    return result


# ------------------------------------------------------------ paragraph CRUD


def _materialize_implicit_paragraph(pkg: DocxPackage) -> etree._Element | None:
    """A body with no direct w:p children (python-docx fresh documents hold
    only the trailing sectPr) still DISPLAYS one empty paragraph in Word.
    Materialize that paragraph so index 0 addresses what the user sees.
    Returns the new element, or None when the body already has paragraphs."""
    body = pkg.body()
    if body.find(qn("w:p")) is not None:
        return None
    p = etree.Element(qn("w:p"))
    sectpr = body.find(qn("w:sectPr"))
    if sectpr is not None:
        sectpr.addprevious(p)
    else:
        body.append(p)
    return p


def _body_paragraph(pkg: DocxPackage, index: int) -> etree._Element:
    paras = [el for kind, idx, el in body_items(pkg) if kind == "paragraph"]
    if not paras and index == 0:
        # Empty-body document: index 0 is the implicit paragraph Word shows.
        made = _materialize_implicit_paragraph(pkg)
        if made is not None:
            return made
    if 0 <= index < len(paras):
        return paras[index]
    raise TargetNotFound(
        f"no body paragraph with index {index}; the document has "
        f"{len(paras)} body paragraph(s)"
        + (f" (valid indices 0-{len(paras) - 1})" if paras else "")
    )


def _resolve_anchor(pkg: DocxPackage, anchor_text: str) -> etree._Element:
    """Unique body paragraph whose PLAIN text contains anchor_text. Matching
    is against visible text, never the XML representation — write '&', '<',
    '>' as the literal characters, not as XML entities."""
    matches = [
        el
        for kind, _, el in body_items(pkg)
        if kind == "paragraph" and anchor_text in paragraph_text(el)
    ]
    if not matches:
        msg = f"anchor text not found: {anchor_text!r}"
        if re.search(r"&amp;|&lt;|&gt;", anchor_text):
            msg += (
                " — note: anchors match the paragraph's PLAIN text, not its "
                "XML; write XML entities as the literal character "
                "('&' not '&amp;', '<' not '&lt;', '>' not '&gt;')"
            )
        raise TargetNotFound(msg)
    if len(matches) > 1:
        raise AmbiguousTarget(
            f"anchor text appears in {len(matches)} paragraphs; "
            "use a longer, unique anchor or a paragraph index"
        )
    return matches[0]


# XML 1.0-forbidden control characters plus DEL (0x7F). C0 controls make
# lxml refuse anyway; DEL is XML-legal but Word strips it silently on open,
# so accepting it would be a silent data change (field test, 2026-09-03).
_CTRL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")


def _check_storable_text(text: str, what: str = "text") -> None:
    m = _CTRL_CHARS_RE.search(text)
    if m:
        raise WordMcpError(
            f"{what} contains control character 0x{ord(m.group(0)):02X}, "
            "which Word cannot store (C0 controls and DEL 0x7F are "
            "refused; use \\n for line breaks and \\t for tabs)"
        )


def _make_paragraph(
    text: str, *, style: str | None = None, formatting: dict | None = None
) -> etree._Element:
    _check_storable_text(text, "paragraph text")
    p = etree.Element(qn("w:p"))
    if style or formatting:
        ppr = etree.SubElement(p, qn("w:pPr"))
        if style:
            etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), style)
    run = etree.SubElement(p, qn("w:r"))
    if formatting:
        run.append(_make_rpr(formatting))
    for i, line in enumerate(text.split("\n")):
        if i:
            etree.SubElement(run, qn("w:br"))
        if line:
            t = etree.SubElement(run, qn("w:t"))
            t.text = line
            _runmap._preserve_space(t)
    return p


def _paragraph_format_clone(
    p: etree._Element,
) -> tuple[etree._Element | None, etree._Element | None]:
    """Clonable formatting of a body paragraph for inherit_format /
    copy_format_from: (deep copy of its direct pPr minus numPr/sectPr/
    outlineLvl, deep copy of its terminal run's rPr). w:outlineLvl is
    STRUCTURAL, not visual: cloning it from a heading anchor silently made
    new body paragraphs show up in the outline and TOC (field test,
    2026-09-03), so it is excluded; set it deliberately with
    set_paragraph_format's outline_level. Revision markers (w:ins/w:del/
    w:rPrChange/w:pPrChange) are never cloned — copying the anchor's tracked
    state onto brand-new paragraphs would mark them inserted/deleted by
    someone else. Either element may be None when the paragraph carries no
    direct formatting at that level."""
    ppr = p.find(qn("w:pPr"))
    ppr_clone = copy.deepcopy(ppr) if ppr is not None else None
    if ppr_clone is not None:
        for tag in ("w:numPr", "w:sectPr", "w:pPrChange", "w:outlineLvl"):
            for node in ppr_clone.findall(qn(tag)):
                ppr_clone.remove(node)
        mark_rpr = ppr_clone.find(qn("w:rPr"))
        if mark_rpr is not None:
            for tag in ("w:ins", "w:del", "w:moveFrom", "w:moveTo",
                        "w:rPrChange"):
                for node in mark_rpr.findall(qn(tag)):
                    mark_rpr.remove(node)
            if len(mark_rpr) == 0:
                ppr_clone.remove(mark_rpr)
        if len(ppr_clone) == 0:
            ppr_clone = None
    # Terminal run = the LAST visible run (tracked-deleted and text-box runs
    # excluded), so a reference entry inherits the font its neighbor ends in.
    last_run = None
    for r in p.iter(qn("w:r")):
        if _runmap._in_deleted(r) or _runmap._in_textbox(r, p):
            continue
        last_run = r
    rpr = last_run.find(qn("w:rPr")) if last_run is not None else None
    if rpr is None and ppr is not None:
        # Runless (empty) paragraph: the paragraph-mark rPr is the only
        # record of its character formatting.
        rpr = ppr.find(qn("w:rPr"))
    rpr_clone = copy.deepcopy(rpr) if rpr is not None else None
    if rpr_clone is not None:
        for tag in ("w:ins", "w:del", "w:rPrChange"):
            for node in rpr_clone.findall(qn(tag)):
                rpr_clone.remove(node)
        if len(rpr_clone) == 0:
            rpr_clone = None
    return ppr_clone, rpr_clone


def insert_paragraphs(
    pkg: DocxPackage,
    paragraphs: list[dict],
    *,
    after_index: int | None = None,
    before_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    inherit_format: bool = False,
    copy_format_from: int | None = None,
    track: bool = False,
    author: str = "Claude",
) -> dict:
    """Insert paragraphs. Each item: {text, style?, formatting?}. With track,
    content and paragraph marks are recorded as insertions by `author`.

    inherit_format=True clones the ANCHOR paragraph's direct formatting (its
    pPr minus numPr/sectPr, plus its terminal run's rPr) onto every inserted
    paragraph, so e.g. a new reference entry picks up its neighbors' hanging
    indent, spacing, and font in one call. The anchor is the paragraph named
    by after_index/after_anchor; with before_index or at_end use
    copy_format_from=<body paragraph index> instead (same clone, explicit
    source; mutually exclusive with inherit_format). Explicit per-item
    style/formatting values still win over the clone."""
    specified = sum(
        x is not None for x in (after_index, before_index, after_anchor)
    ) + bool(at_end)
    if specified != 1:
        raise WordMcpError(
            "specify exactly one of after_index, before_index, after_anchor, at_end"
        )
    if inherit_format and copy_format_from is not None:
        raise WordMcpError(
            "inherit_format and copy_format_from are mutually exclusive; "
            "pass one (both name a paragraph to clone formatting from)"
        )
    fmt_source = None
    fmt_source_index: int | None = None
    if copy_format_from is not None:
        fmt_source = _body_paragraph(pkg, copy_format_from)
        fmt_source_index = copy_format_from
    elif inherit_format:
        if after_index is not None:
            fmt_source = _body_paragraph(pkg, after_index)
            fmt_source_index = after_index
        elif after_anchor is not None:
            fmt_source = _resolve_anchor(pkg, after_anchor)
            for kind, idx, el in body_items(pkg):
                if kind == "paragraph" and el is fmt_source:
                    fmt_source_index = idx
                    break
        else:
            raise WordMcpError(
                "inherit_format clones the after_index/after_anchor anchor "
                "paragraph; with before_index or at_end use "
                "copy_format_from=<body paragraph index> instead"
            )
    new_els = [
        _make_paragraph(
            item["text"],
            style=item.get("style"),
            formatting=item.get("formatting"),
        )
        for item in paragraphs
    ]
    if fmt_source is not None:
        ppr_clone, rpr_clone = _paragraph_format_clone(fmt_source)
        for item, el in zip(paragraphs, new_els):
            if ppr_clone is not None:
                new_ppr = copy.deepcopy(ppr_clone)
                if item.get("style"):
                    # Explicit per-item style wins over the cloned pStyle.
                    ps = new_ppr.find(qn("w:pStyle"))
                    if ps is None:
                        ps = etree.Element(qn("w:pStyle"))
                        new_ppr.insert(0, ps)  # pStyle leads pPr (schema)
                    ps.set(qn("w:val"), item["style"])
                old_ppr = el.find(qn("w:pPr"))
                if old_ppr is not None:
                    el.remove(old_ppr)
                el.insert(0, new_ppr)
            if rpr_clone is not None:
                for run in el.findall(qn("w:r")):
                    new_rpr = copy.deepcopy(rpr_clone)
                    if item.get("formatting"):
                        # Explicit per-item keys overlay the cloned rPr.
                        _apply_fmt(new_rpr, item["formatting"])
                    old_rpr = run.find(qn("w:rPr"))
                    if old_rpr is not None:
                        run.remove(old_rpr)
                    run.insert(0, new_rpr)
    body = pkg.body()
    if at_end:
        # Before the trailing sectPr if present.
        sectpr = body.find(qn("w:sectPr"))
        for el in new_els:
            if sectpr is not None:
                sectpr.addprevious(el)
            else:
                body.append(el)
    elif after_anchor is not None:
        ref = _resolve_anchor(pkg, after_anchor)
        for el in reversed(new_els):
            ref.addnext(el)
    elif after_index is not None:
        ref = _body_paragraph(pkg, after_index)
        for el in reversed(new_els):
            ref.addnext(el)
    else:
        ref = _body_paragraph(pkg, before_index)
        for el in new_els:
            ref.addprevious(el)
    if track:
        from . import _tracked

        root = pkg.root()
        date = _tracked.now_iso()
        for el in new_els:
            rid = _tracked.next_rev_id(root)
            _tracked.wrap_paragraph_content_inserted(
                el, author=author, rev_id=rid, date=date
            )
            _tracked.mark_paragraph_mark(
                el, "ins", author=author, rev_id=rid + 1, date=date
            )
    pkg.mark_dirty()
    result = {"inserted": len(new_els)}
    if fmt_source is not None:
        result["format_cloned_from"] = fmt_source_index
    if track:
        result["tracked_as"] = author
    return result


def _expect_guard(el, index: int, expect: str | None) -> None:
    """Validate-at-execution for an index-addressed destructive call: the
    paragraph must still carry the text the caller expected there, or
    nothing happens. See com/live_ops._expect_guard for the race this
    closes (concurrency matrix H4, 2026-09-05)."""
    if expect is None:
        return
    from .read import paragraph_text

    current = paragraph_text(el)
    if expect not in current:
        raise TargetNotFound(
            f"delete_paragraphs: paragraph {index} does not contain the "
            f"expected text {expect[:80]!r}; its current text begins: "
            f"{current[:120]!r}. Paragraph indices shift after "
            "insert/delete operations, including ones made by another "
            "agent or by the user in Word — re-read with get_text and "
            "retry. Nothing was changed."
        )


def delete_paragraphs(
    pkg: DocxPackage,
    start: int,
    end: int | None = None,
    *,
    track: bool = False,
    author: str = "Claude",
    expect_start: str | None = None,
    expect_end: str | None = None,
) -> dict:
    """Delete body paragraphs [start, end] inclusive (end defaults to start).
    With track, nothing is removed — content and paragraph marks are recorded
    as deletions by `author`, pending accept/reject.

    expect_start / expect_end: the paragraphs actually at those indices must
    still contain the given text, or the deletion refuses with nothing
    removed. Indices go stale whenever anything inserts or deletes above
    them, which on a shared document means another agent's edit landing
    between the read that produced the index and this call (concurrency
    matrix H4). The guard is opt-in on both routes so the file route and the
    live route behave identically."""
    end = start if end is None else end
    if end < start:
        raise WordMcpError("end must be >= start")
    # Empty-body document (fresh create_document): index 0 addresses the
    # implicit paragraph Word displays, so materialize it before ranging.
    _materialize_implicit_paragraph(pkg)
    targets = [
        el
        for kind, idx, el in body_items(pkg)
        if kind == "paragraph" and start <= idx <= end
    ]
    if len(targets) != end - start + 1:
        raise TargetNotFound(
            f"paragraph range {start}-{end} exceeds document "
            f"({len(targets)} of {end - start + 1} found)"
        )
    _expect_guard(targets[0], start, expect_start)
    _expect_guard(targets[-1], end, expect_end)
    if track:
        from . import _tracked

        root = pkg.root()
        date = _tracked.now_iso()
        for el in targets:
            if el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
                raise WordMcpError(
                    "a paragraph in the range carries a section break; "
                    "tracked deletion of section paragraphs is not supported"
                )
            rid = _tracked.next_rev_id(root)
            n = _tracked.wrap_paragraph_content_deleted(
                el, author=author, rev_id=rid, date=date
            )
            _tracked.mark_paragraph_mark(
                el, "del", author=author, rev_id=rid + n, date=date
            )
        pkg.mark_dirty()
        return {"deleted_tracked": len(targets), "tracked_as": author}

    body = pkg.body()
    remaining = sum(1 for k, _, _ in body_items(pkg) if k == "paragraph")
    # Word semantics: a document always keeps at least one paragraph. When
    # the range covers every paragraph, delete them and leave one fresh empty
    # paragraph behind (what Ctrl+A + Delete does in Word).
    deletes_all = remaining - len(targets) < 1
    # Field-balance guard: deleting a range that cuts through a complex field
    # (unbalanced w:fldChar begin/end) corrupts the document.
    depth = 0
    for el in targets:
        for fc in el.iter(qn("w:fldChar")):
            t = fc.get(qn("w:fldCharType"))
            if t == "begin":
                depth += 1
            elif t == "end":
                depth -= 1
    if depth != 0:
        raise WordMcpError(
            "the paragraph range cuts through a field (TOC, PAGEREF, SEQ...); "
            "deleting it would corrupt the document — widen the range to cover "
            "the whole field or use delete_toc for TOCs"
        )
    for el in targets:
        # A paragraph carrying the section's sectPr must not vanish silently.
        if el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            raise WordMcpError(
                f"paragraph {start + targets.index(el)} carries a section break; "
                "delete_paragraphs refuses it — remove the section first"
            )
        body.remove(el)
    if deletes_all:
        _materialize_implicit_paragraph(pkg)
    pkg.mark_dirty()
    from .notes import purge_orphans

    purged = purge_orphans(pkg)["purged"]
    result = {"deleted": len(targets)}
    if deletes_all:
        result["note"] = (
            "the range covered every body paragraph; one empty paragraph "
            "was left behind (a document always keeps at least one)"
        )
    if purged:
        result["note_definitions_purged"] = purged
    return result


def replace_paragraph_text(
    pkg: DocxPackage, index: int, new_text: str, *, expect: str | None = None
) -> dict:
    """Swap a paragraph's entire text, keeping its style and first run's format.

    expect guards against stale indices (indices shift after insert/delete
    operations): when given, the paragraph's CURRENT text must contain it or
    the call refuses with nothing changed. The old text is always returned
    as replaced_text so callers can verify the right paragraph was hit."""
    p = _body_paragraph(pkg, index)
    from .read import paragraph_text

    current = paragraph_text(p)
    if expect is not None and expect not in current:
        raise WordMcpError(
            f"paragraph {index} does not contain the expected text "
            f"{expect[:80]!r}; its current text begins: {current[:120]!r}. "
            "Paragraph indices shift after insert/delete operations - "
            "re-read with get_text, or use search_and_replace for "
            "text-anchored replacement. Nothing was changed."
        )
    first_rpr = None
    first_run = p.find(qn("w:r"))
    if first_run is not None:
        rpr = first_run.find(qn("w:rPr"))
        if rpr is not None:
            first_rpr = copy.deepcopy(rpr)
    for r in p.findall(qn("w:r")):
        p.remove(r)
    for wrapper in p.findall(qn("w:hyperlink")) + p.findall(qn("w:ins")):
        p.remove(wrapper)
    run = etree.SubElement(p, qn("w:r"))
    if first_rpr is not None:
        run.append(first_rpr)
    t = etree.SubElement(run, qn("w:t"))
    t.text = new_text
    _runmap._preserve_space(t)
    pkg.mark_dirty()
    from .notes import purge_orphans

    purged = purge_orphans(pkg)["purged"]
    result = {"replaced_paragraph": index, "replaced_text": current}
    if purged:
        result["note_definitions_purged"] = purged
    return result


def add_heading(
    pkg: DocxPackage,
    text: str,
    level: int = 1,
    *,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
) -> dict:
    if not 1 <= level <= 9:
        raise WordMcpError("heading level must be 1-9")
    style_id = ensure_heading_style(pkg, level)
    return insert_paragraphs(
        pkg,
        [{"text": text, "style": style_id}],
        after_index=after_index,
        after_anchor=after_anchor,
        at_end=at_end,
    )


def add_page_break(pkg: DocxPackage, *, after_index: int) -> dict:
    p = _body_paragraph(pkg, after_index)
    new_p = etree.Element(qn("w:p"))
    run = etree.SubElement(new_p, qn("w:r"))
    etree.SubElement(run, qn("w:br")).set(qn("w:type"), "page")
    p.addnext(new_p)
    pkg.mark_dirty()
    return {"page_break_after": after_index}


# ---------------------------------------------------------------- formatting

_TOGGLES = {"bold": "w:b", "italic": "w:i", "underline": "w:u", "strike": "w:strike"}

_CHAR_FMT_KEYS = set(_TOGGLES) | {
    "font", "size_pt", "color", "highlight", "superscript", "subscript",
    "small_caps", "all_caps", "hidden", "double_strike", "char_spacing_pt",
    "kerning_pt", "position_pt", "language", "east_asian_language",
}

_PARA_FMT_KEYS = {
    "alignment", "space_before_pt", "space_after_pt", "line_spacing",
    "indent_left_pt", "indent_right_pt", "first_line_indent_pt",
    "keep_with_next", "keep_lines_together", "page_break_before",
    "widow_control", "shading", "borders", "tab_stops", "outline_level",
}

# CT_PPr child sequence (python-docx _tag_seq, abridged to what we write).
_PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange",
]


def _ppr_get_or_add(ppr, local: str):
    existing = ppr.find(qn(f"w:{local}"))
    if existing is not None:
        return existing
    el = etree.Element(qn(f"w:{local}"))
    my_rank = _PPR_ORDER.index(local)
    for child in ppr:
        name = etree.QName(child).localname
        if name in _PPR_ORDER and _PPR_ORDER.index(name) > my_rank:
            child.addprevious(el)
            return el
    ppr.append(el)
    return el


def _check_keys(fmt: dict, allowed: set, what: str) -> None:
    unknown = set(fmt) - allowed
    if unknown:
        raise WordMcpError(
            f"unknown {what} key(s) {sorted(unknown)}; "
            f"allowed: {sorted(allowed)}"
        )


def _make_rpr(fmt: dict) -> etree._Element:
    rpr = etree.Element(qn("w:rPr"))
    _apply_fmt(rpr, fmt)
    return rpr


def _apply_fmt(rpr: etree._Element, fmt: dict) -> None:
    """Apply formatting keys to an rPr, respecting the schema's element order
    loosely (Word tolerates rPr child order in practice, but we keep toggles
    first, then fonts/size/color, matching common output)."""
    _check_keys(fmt, _CHAR_FMT_KEYS, "character-formatting")
    for key, tag in _TOGGLES.items():
        if key in fmt:
            el = rpr.find(qn(tag))
            if fmt[key]:
                if el is None:
                    el = etree.SubElement(rpr, qn(tag))
                if key == "underline":
                    el.set(qn("w:val"), "single")
                else:
                    el.attrib.pop(qn("w:val"), None)
            elif el is not None:
                rpr.remove(el)
    if "font" in fmt:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = etree.SubElement(rpr, qn("w:rFonts"))
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            rfonts.set(qn(attr), fmt["font"])
    if "size_pt" in fmt:
        for tag in ("w:sz", "w:szCs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = etree.SubElement(rpr, qn(tag))
            el.set(qn("w:val"), str(int(fmt["size_pt"] * 2)))
    if "color" in fmt:
        el = rpr.find(qn("w:color"))
        if el is None:
            el = etree.SubElement(rpr, qn("w:color"))
        el.set(qn("w:val"), fmt["color"].lstrip("#"))
    if "highlight" in fmt:
        el = rpr.find(qn("w:highlight"))
        if el is None:
            el = etree.SubElement(rpr, qn("w:highlight"))
        el.set(qn("w:val"), fmt["highlight"])
    for key, tag in (
        ("small_caps", "w:smallCaps"),
        ("all_caps", "w:caps"),
        ("hidden", "w:vanish"),
        ("double_strike", "w:dstrike"),
    ):
        if key in fmt:
            el = rpr.find(qn(tag))
            if fmt[key] and el is None:
                etree.SubElement(rpr, qn(tag))
            elif not fmt[key] and el is not None:
                rpr.remove(el)
    if "char_spacing_pt" in fmt:
        el = rpr.find(qn("w:spacing"))
        if el is None:
            el = etree.SubElement(rpr, qn("w:spacing"))
        el.set(qn("w:val"), str(int(fmt["char_spacing_pt"] * 20)))
    if "kerning_pt" in fmt:
        el = rpr.find(qn("w:kern"))
        if el is None:
            el = etree.SubElement(rpr, qn("w:kern"))
        el.set(qn("w:val"), str(int(fmt["kerning_pt"] * 2)))
    if "position_pt" in fmt:
        el = rpr.find(qn("w:position"))
        if el is None:
            el = etree.SubElement(rpr, qn("w:position"))
        el.set(qn("w:val"), str(int(fmt["position_pt"] * 2)))
    if "language" in fmt or "east_asian_language" in fmt:
        el = rpr.find(qn("w:lang"))
        if el is None:
            el = etree.SubElement(rpr, qn("w:lang"))
        if fmt.get("language"):
            el.set(qn("w:val"), fmt["language"])
        if fmt.get("east_asian_language"):
            el.set(qn("w:eastAsia"), fmt["east_asian_language"])
    if "superscript" in fmt or "subscript" in fmt:
        if fmt.get("superscript") and fmt.get("subscript"):
            raise WordMcpError(
                "superscript and subscript are mutually exclusive; pass one"
            )
        el = rpr.find(qn("w:vertAlign"))
        if el is None:
            el = etree.SubElement(rpr, qn("w:vertAlign"))
        el.set(
            qn("w:val"),
            "superscript" if fmt.get("superscript") else "subscript",
        )


def format_text(
    pkg: DocxPackage,
    *,
    paragraph_index: int | None = None,
    find: str | None = None,
    occurrence: int = 1,
    formatting: dict,
) -> dict:
    """Character-range formatting. Target either (paragraph_index + find) to
    format a substring, paragraph_index alone to format the whole paragraph,
    or find alone (searches the whole body; `occurrence` picks which match)."""
    if find is None and paragraph_index is None:
        raise WordMcpError("need paragraph_index, find, or both")

    if paragraph_index is not None:
        candidates: list[tuple[etree._Element, int | None]] = [
            (_body_paragraph(pkg, paragraph_index), paragraph_index)
        ]
    else:
        # Document order, INCLUDING paragraphs inside table cells (and any
        # other nested containers in document.xml). The body elements MUST
        # stay referenced while their id()s are looked up: lxml proxies are
        # ephemeral, and without the keepalive a recycled proxy address maps
        # a body paragraph to the wrong index (or to "table cell").
        _keepalive = [
            (el, idx) for kind, idx, el in body_items(pkg) if kind == "paragraph"
        ]
        body_idx = {id(el): idx for el, idx in _keepalive}
        candidates = [
            (p, body_idx.get(id(p))) for p in pkg.root().iter(qn("w:p"))
        ]
        del _keepalive

    seen = 0
    for p, idx in candidates:
        text, _ = _runmap.build_map(p)
        if find is None:
            spans = [(0, len(text))] if text else []
        else:
            spans = [
                (m.start(), m.end())
                for m in re.finditer(re.escape(find), text)
            ]
        for span in spans:
            seen += 1
            if seen == occurrence:
                runs = _runmap.split_for_range(p, span[0], span[1])
                for run in runs:
                    rpr = run.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = etree.Element(qn("w:rPr"))
                        run.insert(0, rpr)
                    _apply_fmt(rpr, formatting)
                pkg.mark_dirty()
                loc: dict = {"start": span[0], "end": span[1]}
                if idx is not None:
                    loc["paragraph"] = idx
                else:
                    loc["location"] = "table cell"
                return {"formatted": loc}
    raise TargetNotFound(
        f"occurrence {occurrence} of {find!r} not found"
        + (
            f" in paragraph {paragraph_index}"
            if paragraph_index is not None
            else " anywhere in the document (body and table cells searched)"
        )
    )


_ALIGN = {"left": "left", "center": "center", "right": "right", "justify": "both"}

# Numeric bounds mirroring Word's own UI limits (Word caps spacing and
# indents at 1584pt / 22in and line-spacing multiples at 132). Out-of-range
# values used to ride through into undefined rendering (field test,
# 2026-09-03): negative indents, 99999pt spacing.
_PT_MAX = 1584.0
_LINE_SPACING_MIN, _LINE_SPACING_MAX = 0.06, 132.0


def _num_ok(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def validate_paragraph_numeric(formatting: dict) -> None:
    """Range-check the numeric paragraph-formatting keys; raises with the
    valid bounds named. Shared by the file and live paths."""
    for key in ("space_before_pt", "space_after_pt"):
        if key in formatting:
            v = formatting[key]
            if not _num_ok(v) or not 0 <= v <= _PT_MAX:
                raise WordMcpError(
                    f"{key} must be a number from 0 to {_PT_MAX:.0f}pt "
                    f"(Word's maximum), got {v!r}"
                )
    if "line_spacing" in formatting:
        v = formatting["line_spacing"]
        if not _num_ok(v) or not (
            _LINE_SPACING_MIN <= v <= _LINE_SPACING_MAX
        ):
            raise WordMcpError(
                "line_spacing is a multiple of single spacing (1 = single, "
                f"2 = double) from {_LINE_SPACING_MIN} to "
                f"{_LINE_SPACING_MAX:.0f} (Word's maximum), got {v!r}"
            )
    for key in ("indent_left_pt", "indent_right_pt"):
        if key in formatting:
            v = formatting[key]
            if not _num_ok(v) or not 0 <= v <= _PT_MAX:
                raise WordMcpError(
                    f"{key} must be a number from 0 to {_PT_MAX:.0f}pt; "
                    f"got {v!r} (negative indents are refused; for a "
                    "hanging indent use a negative first_line_indent_pt)"
                )
    if "first_line_indent_pt" in formatting:
        v = formatting["first_line_indent_pt"]
        if not _num_ok(v) or not -_PT_MAX <= v <= _PT_MAX:
            raise WordMcpError(
                f"first_line_indent_pt must be a number from "
                f"-{_PT_MAX:.0f} to {_PT_MAX:.0f}pt (negative = hanging "
                f"indent), got {v!r}"
            )


def set_paragraph_format(
    pkg: DocxPackage, indices: list[int], formatting: dict
) -> dict:
    """Paragraph-level formatting for a batch of paragraphs. Keys: alignment,
    space_before_pt, space_after_pt, line_spacing, indent_left_pt,
    indent_right_pt, first_line_indent_pt, keep_with_next, outline_level.

    outline_level (0-8, or None to REMOVE the direct override) writes
    w:outlineLvl WITHOUT touching the paragraph's style or visual
    formatting. This is how academic templates express heading hierarchy on
    Normal-styled paragraphs with direct formatting — Word's navigation pane
    and TOC field harvesting (the \\o switch) honor outlineLvl, so setting it
    here is what puts such headings into the outline and the TOC, where
    apply_style('Heading N') would wreck the template's look. 0 is the top
    level (Heading 1 equivalent); body text simply has no outlineLvl."""
    _check_keys(formatting, _PARA_FMT_KEYS, "paragraph-formatting")
    validate_paragraph_numeric(formatting)
    if "outline_level" in formatting:
        lvl = formatting["outline_level"]
        if lvl is not None and (
            isinstance(lvl, bool) or not isinstance(lvl, int)
            or not 0 <= lvl <= 8
        ):
            raise WordMcpError(
                "outline_level must be an integer 0-8 (0 = top level, as in "
                "Heading 1) or null to remove the direct outlineLvl override"
            )
    for index in indices:
        p = _body_paragraph(pkg, index)
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("w:pPr"))
            p.insert(0, ppr)
        if "alignment" in formatting:
            val = _ALIGN.get(formatting["alignment"])
            if val is None:
                raise WordMcpError(f"alignment must be one of {list(_ALIGN)}")
            _ppr_get_or_add(ppr, "jc").set(qn("w:val"), val)
        for key, tag in (
            ("keep_lines_together", "keepLines"),
            ("page_break_before", "pageBreakBefore"),
        ):
            if key in formatting:
                el = ppr.find(qn(f"w:{tag}"))
                if formatting[key] and el is None:
                    _ppr_get_or_add(ppr, tag)
                elif not formatting[key] and el is not None:
                    ppr.remove(el)
        if "widow_control" in formatting:
            el = _ppr_get_or_add(ppr, "widowControl")
            el.set(qn("w:val"), "1" if formatting["widow_control"] else "0")
        if "outline_level" in formatting:
            lvl = formatting["outline_level"]
            if lvl is None:
                existing = ppr.find(qn("w:outlineLvl"))
                if existing is not None:
                    ppr.remove(existing)
            else:
                _ppr_get_or_add(ppr, "outlineLvl").set(qn("w:val"), str(lvl))
        if "shading" in formatting:
            shd = _ppr_get_or_add(ppr, "shd")
            shd.set(qn("w:val"), "clear")
            shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), formatting["shading"].lstrip("#"))
        if "borders" in formatting:
            spec = formatting["borders"]
            sides = (
                ["top", "bottom", "left", "right"]
                if spec in (True, "all")
                else list(spec)
            )
            bad = set(sides) - {"top", "bottom", "left", "right", "between", "bar"}
            if bad:
                raise WordMcpError(f"unknown border side(s): {sorted(bad)}")
            pbdr = _ppr_get_or_add(ppr, "pBdr")
            for child in list(pbdr):
                pbdr.remove(child)
            for side in ("top", "left", "bottom", "right", "between", "bar"):
                if side in sides:
                    b = etree.SubElement(pbdr, qn(f"w:{side}"))
                    b.set(qn("w:val"), "single")
                    b.set(qn("w:sz"), "4")
                    b.set(qn("w:space"), "4")
                    b.set(qn("w:color"), "auto")
        if "tab_stops" in formatting:
            tabs = _ppr_get_or_add(ppr, "tabs")
            for child in list(tabs):
                tabs.remove(child)
            leaders = {"none", "dot", "hyphen", "underscore", "middleDot"}
            aligns = {"left", "center", "right", "decimal", "bar"}
            for stop in formatting["tab_stops"]:
                align = stop.get("alignment", "left")
                leader = stop.get("leader", "none")
                if align not in aligns:
                    raise WordMcpError(f"tab alignment must be one of {sorted(aligns)}")
                if leader not in leaders:
                    raise WordMcpError(f"tab leader must be one of {sorted(leaders)}")
                t = etree.SubElement(tabs, qn("w:tab"))
                t.set(qn("w:val"), align)
                if leader != "none":
                    t.set(qn("w:leader"), leader)
                t.set(qn("w:pos"), str(int(stop["position_pt"] * 20)))
        if any(
            k in formatting
            for k in ("space_before_pt", "space_after_pt", "line_spacing")
        ):
            spacing = _ppr_get_or_add(ppr, "spacing")
            if "space_before_pt" in formatting:
                spacing.set(
                    qn("w:before"), str(int(formatting["space_before_pt"] * 20))
                )
            if "space_after_pt" in formatting:
                spacing.set(
                    qn("w:after"), str(int(formatting["space_after_pt"] * 20))
                )
            if "line_spacing" in formatting:
                spacing.set(
                    qn("w:line"), str(int(formatting["line_spacing"] * 240))
                )
                spacing.set(qn("w:lineRule"), "auto")
        if any(
            k in formatting
            for k in ("indent_left_pt", "indent_right_pt", "first_line_indent_pt")
        ):
            ind = _ppr_get_or_add(ppr, "ind")
            if "indent_left_pt" in formatting:
                ind.set(qn("w:left"), str(int(formatting["indent_left_pt"] * 20)))
            if "indent_right_pt" in formatting:
                ind.set(qn("w:right"), str(int(formatting["indent_right_pt"] * 20)))
            if "first_line_indent_pt" in formatting:
                val = formatting["first_line_indent_pt"]
                if val >= 0:
                    ind.set(qn("w:firstLine"), str(int(val * 20)))
                else:
                    ind.set(qn("w:hanging"), str(int(-val * 20)))
        if "keep_with_next" in formatting:
            kn = ppr.find(qn("w:keepNext"))
            if formatting["keep_with_next"] and kn is None:
                etree.SubElement(ppr, qn("w:keepNext"))
            elif not formatting["keep_with_next"] and kn is not None:
                ppr.remove(kn)
    pkg.mark_dirty()
    return {"formatted_paragraphs": indices}


_HEADING_SIZES_PT = {1: 16, 2: 14, 3: 13, 4: 12, 5: 12, 6: 12, 7: 12, 8: 12, 9: 12}


def ensure_heading_style(pkg: DocxPackage, level: int) -> str:
    """Return the styleId for heading `level`, creating the definition if the
    document lacks it. Word renders a pStyle reference to an undefined style as
    Normal text, so inserting headings without this is a silent failure."""
    from .read import list_styles

    styles = list_styles(pkg)
    for s in styles:
        if s["id"] == f"Heading{level}" or s["name"] == f"heading {level}":
            return s["id"]
    root = pkg.root("word/styles.xml")
    style = etree.SubElement(root, qn("w:style"))
    style.set(qn("w:type"), "paragraph")
    style.set(qn("w:styleId"), f"Heading{level}")
    etree.SubElement(style, qn("w:name")).set(qn("w:val"), f"heading {level}")
    etree.SubElement(style, qn("w:basedOn")).set(qn("w:val"), "Normal")
    etree.SubElement(style, qn("w:next")).set(qn("w:val"), "Normal")
    ppr = etree.SubElement(style, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:keepNext"))
    spacing = etree.SubElement(ppr, qn("w:spacing"))
    spacing.set(qn("w:before"), "240")
    spacing.set(qn("w:after"), "60")
    etree.SubElement(ppr, qn("w:outlineLvl")).set(qn("w:val"), str(level - 1))
    rpr = etree.SubElement(style, qn("w:rPr"))
    etree.SubElement(rpr, qn("w:b"))
    sz = str(_HEADING_SIZES_PT[level] * 2)
    etree.SubElement(rpr, qn("w:sz")).set(qn("w:val"), sz)
    etree.SubElement(rpr, qn("w:szCs")).set(qn("w:val"), sz)
    pkg.mark_dirty("word/styles.xml")
    return f"Heading{level}"


_CASE_TRANSFORMS = ("upper", "lower", "title", "sentence")


def change_case(
    pkg: DocxPackage,
    transform: str,
    *,
    indices: list[int] | None = None,
    find: str | None = None,
) -> dict:
    """Change text case: upper | lower | title | sentence. Target: paragraph
    indices, or every occurrence of `find` (case-insensitive), or both."""
    if transform not in _CASE_TRANSFORMS:
        raise WordMcpError(f"transform must be one of {_CASE_TRANSFORMS}")
    if indices is None and find is None:
        raise WordMcpError("give indices and/or find")

    def apply_str(s: str, state: dict) -> str:
        if transform == "upper":
            return s.upper()
        if transform == "lower":
            return s.lower()
        if transform == "title":
            # \w is unicode-aware: handles straße, émigré, Greek; scripts
            # without case (Korean, CJK) pass through unchanged.
            return re.sub(
                r"\w+(?:'\w+)?",
                lambda m: m.group(0)[0].upper() + m.group(0)[1:].lower(),
                s,
            )
        out = []
        for ch in s:
            if state["cap_next"] and ch.isalpha():
                out.append(ch.upper())
                state["cap_next"] = False
            else:
                out.append(ch.lower() if ch.isalpha() else ch)
            if ch in ".!?":
                state["cap_next"] = True
        return "".join(out)

    changed = 0
    targets = []
    if indices is not None:
        for i in indices:
            targets.append(_body_paragraph(pkg, i))
    for p in targets or [
        el for kind, _, el in body_items(pkg) if kind == "paragraph"
    ]:
        if targets:
            in_scope = True
        else:
            text, _ = _runmap.build_map(p)
            in_scope = find is not None and find.lower() in text.lower()
        if not in_scope:
            continue
        state = {"cap_next": True}
        for r in p.iter(qn("w:r")):
            if _runmap._in_deleted(r):
                continue
            for t in r.findall(qn("w:t")):
                if find is not None and indices is None:
                    t.text = re.sub(
                        re.escape(find),
                        lambda m: apply_str(m.group(0), {"cap_next": True}),
                        t.text or "",
                        flags=re.I,
                    )
                else:
                    t.text = apply_str(t.text or "", state)
        changed += 1
    if changed:
        pkg.mark_dirty()
    return {"transform": transform, "paragraphs_changed": changed}


def apply_style(pkg: DocxPackage, indices: list[int], style: str) -> dict:
    """Apply a paragraph style by id or name to a batch of paragraphs.
    Heading1-9 are auto-created if the document does not define them."""
    from .read import list_styles

    styles = list_styles(pkg)
    by_id = {s["id"] for s in styles}
    by_name = {s["name"]: s["id"] for s in styles if s["name"]}
    style_id = style if style in by_id else by_name.get(style)
    if style_id is None:
        m = re.fullmatch(r"[Hh]eading\s?([1-9])", style)
        if m:
            style_id = ensure_heading_style(pkg, int(m.group(1)))
    if style_id is None:
        raise TargetNotFound(
            f"style {style!r} not defined in this document; "
            f"available paragraph styles: "
            f"{sorted(s['id'] for s in styles if s['type'] == 'paragraph')[:30]}"
        )
    for index in indices:
        p = _body_paragraph(pkg, index)
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("w:pPr"))
            p.insert(0, ppr)
        pstyle = ppr.find(qn("w:pStyle"))
        if pstyle is None:
            pstyle = etree.Element(qn("w:pStyle"))
            ppr.insert(0, pstyle)
        pstyle.set(qn("w:val"), style_id)
    pkg.mark_dirty()
    return {"styled_paragraphs": indices, "style_id": style_id}
