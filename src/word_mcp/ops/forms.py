"""Form-field tooling: legacy form fields (w:ffData) and modern content
controls (w:sdt), listed, filled, and validated across the body and tables.

Fill semantics are strict: a dropdown value must be one of its options, a
checkbox takes a real boolean, and an ambiguous (duplicate) field name is
refused with locations rather than guessed at. Validation happens in full
BEFORE any mutation, so a refusal changes nothing.
"""

from __future__ import annotations

import copy
import re

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from ..core.package import DocxPackage, qn
from . import _runmap
from .read import paragraph_text, run_text

_BAD_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# sdtPr children that mark control types we must not touch as form fields
# (galleries, TOC wrappers, citations, pictures, group/equation controls).
_SDT_SKIP_TYPES = (
    "w:docPartObj", "w:docPartList", "w:citation", "w:bibliography",
    "w:picture", "w:group", "w:equation",
)

_CHECKED_GLYPH = "☒"
_UNCHECKED_GLYPH = "☐"


def _sval(parent: etree._Element, tag: str) -> str | None:
    el = parent.find(qn(tag))
    return el.get(qn("w:val")) if el is not None else None


def _context(p: etree._Element | None, limit: int = 80) -> str:
    if p is None:
        return ""
    return paragraph_text(p).strip()[:limit]


def _containing_paragraph(el: etree._Element) -> etree._Element | None:
    parent = el.getparent()
    while parent is not None:
        if etree.QName(parent).localname == "p":
            return parent
        parent = parent.getparent()
    return None


# ------------------------------------------------------------- legacy fields


def _scan_legacy_fields(p: etree._Element) -> list[dict]:
    """Legacy form fields (FORMTEXT/FORMCHECKBOX/FORMDROPDOWN) contained in
    this paragraph. Checkbox fields legitimately have no separate/result."""
    fields: list[dict] = []
    cur: dict | None = None
    for r in p.iter(qn("w:r")):
        fc = r.find(qn("w:fldChar"))
        if fc is not None:
            ftype = fc.get(qn("w:fldCharType"))
            if ftype == "begin":
                ffdata = fc.find(qn("w:ffData"))
                cur = (
                    {
                        "ffdata": ffdata,
                        "begin_run": r,
                        "separate_run": None,
                        "result_runs": [],
                        "stage": "instr",
                    }
                    if ffdata is not None
                    else None
                )
                continue
            if cur is None:
                continue
            if ftype == "separate":
                cur["separate_run"] = r
                cur["stage"] = "result"
            elif ftype == "end":
                fields.append(cur)
                cur = None
            continue
        if cur is not None and cur["stage"] == "result":
            cur["result_runs"].append(r)
    return fields


def _legacy_entry(field: dict, part: str, p: etree._Element) -> dict | None:
    ff = field["ffdata"]
    name = _sval(ff, "w:name")
    display = "".join(run_text(r) for r in field["result_runs"])
    base = {
        "name": name,
        "part": part,
        "context": _context(p),
        "_legacy": field,
    }
    text_input = ff.find(qn("w:textInput"))
    checkbox = ff.find(qn("w:checkBox"))
    ddlist = ff.find(qn("w:ddList"))
    if text_input is not None:
        base.update(kind="legacy_text", value=display)
    elif checkbox is not None:
        checked = _sval(checkbox, "w:checked")
        if checked is None:
            checked = _sval(checkbox, "w:default")
        base.update(kind="legacy_checkbox", value=checked in ("1", "true"))
    elif ddlist is not None:
        entries = [
            e.get(qn("w:val"), "") for e in ddlist.findall(qn("w:listEntry"))
        ]
        idx = int(_sval(ddlist, "w:result") or 0)
        base.update(
            kind="legacy_dropdown",
            value=entries[idx] if 0 <= idx < len(entries) else None,
            options=entries,
        )
    else:
        return None
    return base


# ------------------------------------------------------ content controls (sdt)


def _sdt_entry(sdt: etree._Element, part: str) -> dict | None:
    pr = sdt.find(qn("w:sdtPr"))
    content = sdt.find(qn("w:sdtContent"))
    if pr is None or content is None:
        return None
    for skip in _SDT_SKIP_TYPES:
        if pr.find(qn(skip)) is not None:
            return None
    tag = _sval(pr, "w:tag")
    alias = _sval(pr, "w:alias")
    placeholder = pr.find(qn("w:showingPlcHdr")) is not None
    text_value = "".join(run_text(r) for r in content.iter(qn("w:r")))
    base = {
        "name": tag or alias,
        "tag": tag,
        "alias": alias,
        "part": part,
        "context": _context(_containing_paragraph(sdt)) or text_value[:80],
        "placeholder_showing": placeholder,
        "_sdt": sdt,
    }
    checkbox = pr.find(qn("w14:checkbox"))
    dropdown = pr.find(qn("w:dropDownList"))
    combo = pr.find(qn("w:comboBox"))
    date = pr.find(qn("w:date"))
    if checkbox is not None:
        checked_el = checkbox.find(qn("w14:checked"))
        checked = (
            checked_el.get(qn("w14:val")) if checked_el is not None else None
        )
        base.update(kind="sdt_checkbox", value=checked in ("1", "true"))
    elif dropdown is not None or combo is not None:
        items = [
            {
                "display": li.get(qn("w:displayText"), ""),
                "value": li.get(qn("w:value"), ""),
            }
            for li in (dropdown if dropdown is not None else combo).findall(
                qn("w:listItem")
            )
        ]
        base.update(
            kind="sdt_dropdown" if dropdown is not None else "sdt_combo",
            value=text_value if not placeholder else "",
            options=[i["display"] for i in items],
            _items=items,
        )
    elif date is not None:
        base.update(kind="sdt_date", value=text_value if not placeholder else "")
    elif pr.find(qn("w:text")) is not None:
        base.update(kind="sdt_text", value=text_value if not placeholder else "")
    else:
        # No explicit type element = a rich-text control; fillable as text.
        base.update(kind="sdt_richtext", value=text_value if not placeholder else "")
    return base


# ------------------------------------------------------------------ inventory


def _inventory(pkg: DocxPackage) -> list[dict]:
    part = "word/document.xml"
    entries: list[dict] = []
    root = pkg.root(part)
    for p in root.iter(qn("w:p")):
        for field in _scan_legacy_fields(p):
            entry = _legacy_entry(field, part, p)
            if entry is not None:
                entries.append(entry)
    for sdt in root.iter(qn("w:sdt")):
        entry = _sdt_entry(sdt, part)
        if entry is not None:
            entries.append(entry)
    return entries


def _public(entry: dict) -> dict:
    return {k: v for k, v in entry.items() if not k.startswith("_")}


def list_form_fields(pkg: DocxPackage) -> dict:
    """Every fillable form field: legacy fields (FORMTEXT / FORMCHECKBOX /
    FORMDROPDOWN with names, values, options) and modern content controls
    (text, rich text, checkbox, dropdown/combo, date, with tag/alias),
    across the body and tables."""
    entries = [_public(e) for e in _inventory(pkg)]
    return {"count": len(entries), "fields": entries}


def _match_keys(entry: dict) -> set[str]:
    keys = {entry.get("name"), entry.get("tag"), entry.get("alias")}
    keys.discard(None)
    keys.discard("")
    return keys


def _check_text(name: str, value) -> str:
    text = str(value)
    if _BAD_CHARS_RE.search(text):
        raise WordMcpError(
            f"value for {name!r} contains control characters that cannot be "
            "stored in document text"
        )
    return text


def _as_bool(name: str, value) -> bool:
    if isinstance(value, bool):
        return value
    if value in (0, 1):
        return bool(value)
    raise WordMcpError(
        f"{name!r} is a checkbox; pass true/false, got {value!r}"
    )


# ---------------------------------------------------------------- fill: apply


def _new_text_run(ref_run: etree._Element | None, text: str) -> etree._Element:
    run = etree.Element(qn("w:r"))
    if ref_run is not None:
        rpr = ref_run.find(qn("w:rPr"))
        if rpr is not None:
            run.append(copy.deepcopy(rpr))
    t = etree.SubElement(run, qn("w:t"))
    t.text = text
    _runmap._preserve_space(t)
    return run


def _set_legacy_display(field: dict, text: str) -> None:
    ref = field["result_runs"][0] if field["result_runs"] else field["begin_run"]
    run = _new_text_run(ref, text)
    field["separate_run"].addnext(run)
    for r in field["result_runs"]:
        parent = r.getparent()
        if parent is not None:
            parent.remove(r)


_CONTENT_KINDS = ("r", "hyperlink", "ins", "sdt", "fldSimple", "bookmarkStart", "bookmarkEnd")


def _set_sdt_content(sdt: etree._Element, text: str) -> None:
    pr = sdt.find(qn("w:sdtPr"))
    if pr is not None:
        ph = pr.find(qn("w:showingPlcHdr"))
        if ph is not None:
            pr.remove(ph)
    content = sdt.find(qn("w:sdtContent"))
    ref_run = next(content.iter(qn("w:r")), None)
    run = _new_text_run(ref_run, text)
    paragraphs = content.findall(qn("w:p"))
    if paragraphs:
        target = paragraphs[0]
        for child in list(target):
            if etree.QName(child).localname in _CONTENT_KINDS:
                target.remove(child)
        target.append(run)
        for extra in paragraphs[1:]:
            content.remove(extra)
    else:
        if content.find(qn("w:tbl")) is not None:
            raise UnsupportedStructure(
                "content control holds a table; fill_form_fields cannot "
                "replace it with text"
            )
        for child in list(content):
            if etree.QName(child).localname in _CONTENT_KINDS:
                content.remove(child)
        content.append(run)


def _set_sdt_checkbox(sdt: etree._Element, value: bool) -> None:
    pr = sdt.find(qn("w:sdtPr"))
    checkbox = pr.find(qn("w14:checkbox"))
    checked = checkbox.find(qn("w14:checked"))
    if checked is None:
        checked = etree.SubElement(checkbox, qn("w14:checked"))
    checked.set(qn("w14:val"), "1" if value else "0")
    state = checkbox.find(
        qn("w14:checkedState") if value else qn("w14:uncheckedState")
    )
    glyph = _CHECKED_GLYPH if value else _UNCHECKED_GLYPH
    if state is not None and state.get(qn("w14:val")):
        try:
            glyph = chr(int(state.get(qn("w14:val")), 16))
        except ValueError:
            pass
    _set_sdt_content(sdt, glyph)


def _iso_full_date(text: str) -> str | None:
    import datetime

    try:
        d = datetime.date.fromisoformat(text[:10])
    except ValueError:
        return None
    return d.strftime("%Y-%m-%dT00:00:00Z")


def fill_form_fields(pkg: DocxPackage, values: dict, *, missing: str = "error") -> dict:
    """Set form-field values by name (legacy) or tag/alias (content control).
    Text fields get the value as text (legacy fields also update their
    default), checkboxes take booleans, dropdowns refuse values not in their
    options. missing: 'error' refuses (changing nothing) if a key matches no
    field; 'skip' ignores such keys and reports them. Duplicate field names
    are refused with their locations."""
    if missing not in ("error", "skip"):
        raise WordMcpError(
            "missing must be 'error' or 'skip' (a form key either matches a "
            "field or it does not)"
        )
    entries = _inventory(pkg)

    # -------- validation pass: everything checked before anything changes.
    plan: list[tuple[dict, object]] = []
    unknown: list[str] = []
    for key, value in values.items():
        matches = [e for e in entries if key in _match_keys(e)]
        if not matches:
            unknown.append(key)
            continue
        if len(matches) > 1:
            locs = [
                {"kind": e["kind"], "context": e["context"]} for e in matches
            ]
            raise AmbiguousTarget(
                f"{len(matches)} form fields match {key!r}: {locs}; give the "
                "fields unique names/tags, then retry — nothing was changed"
            )
        entry = matches[0]
        kind = entry["kind"]
        if kind in ("legacy_checkbox", "sdt_checkbox"):
            plan.append((entry, _as_bool(key, value)))
        elif kind == "legacy_dropdown":
            text = _check_text(key, value)
            if text not in entry["options"]:
                raise WordMcpError(
                    f"{text!r} is not an option of dropdown {key!r} "
                    f"(options: {entry['options']}); nothing was changed"
                )
            plan.append((entry, text))
        elif kind == "sdt_dropdown":
            text = _check_text(key, value)
            item = next(
                (
                    i
                    for i in entry["_items"]
                    if text in (i["display"], i["value"])
                ),
                None,
            )
            if item is None:
                raise WordMcpError(
                    f"{text!r} is not an option of dropdown {key!r} "
                    f"(options: {entry['options']}); nothing was changed"
                )
            plan.append((entry, item["display"] or item["value"]))
        elif kind in ("legacy_text", "sdt_text", "sdt_richtext", "sdt_combo", "sdt_date"):
            text = _check_text(key, value)
            if kind == "legacy_text" and entry["_legacy"]["separate_run"] is None:
                raise UnsupportedStructure(
                    f"legacy text field {key!r} has no result section "
                    "(malformed field); nothing was changed"
                )
            plan.append((entry, text))
        else:
            raise WordMcpError(
                f"field {key!r} has kind {kind!r} which fill_form_fields "
                "does not support; nothing was changed"
            )
    if unknown and missing == "error":
        raise WordMcpError(
            f"no form field matches {sorted(unknown)}; nothing was changed — "
            "check list_form_fields for names/tags or call with missing='skip'"
        )

    # -------- apply pass.
    filled: dict[str, str] = {}
    for entry, value in plan:
        kind = entry["kind"]
        if kind == "legacy_text":
            field = entry["_legacy"]
            text_input = field["ffdata"].find(qn("w:textInput"))
            default = text_input.find(qn("w:default"))
            if default is None:
                default = etree.SubElement(text_input, qn("w:default"))
                # Schema order in CT_FFTextInput: type, default, maxLength, format.
                type_el = text_input.find(qn("w:type"))
                if type_el is not None:
                    type_el.addnext(default)
                else:
                    text_input.insert(0, default)
            default.set(qn("w:val"), value)
            _set_legacy_display(field, value)
        elif kind == "legacy_checkbox":
            checkbox = entry["_legacy"]["ffdata"].find(qn("w:checkBox"))
            checked = checkbox.find(qn("w:checked"))
            if checked is None:
                checked = etree.SubElement(checkbox, qn("w:checked"))
            checked.set(qn("w:val"), "1" if value else "0")
        elif kind == "legacy_dropdown":
            field = entry["_legacy"]
            ddlist = field["ffdata"].find(qn("w:ddList"))
            result = ddlist.find(qn("w:result"))
            if result is None:
                result = etree.Element(qn("w:result"))
                ddlist.insert(0, result)
            result.set(qn("w:val"), str(entry["options"].index(value)))
            if field["separate_run"] is not None:
                _set_legacy_display(field, value)
        elif kind == "sdt_checkbox":
            _set_sdt_checkbox(entry["_sdt"], value)
        else:  # sdt_text / sdt_richtext / sdt_dropdown / sdt_combo / sdt_date
            _set_sdt_content(entry["_sdt"], value)
            if kind == "sdt_date":
                full = _iso_full_date(value)
                if full is not None:
                    date_el = entry["_sdt"].find(
                        f"{qn('w:sdtPr')}/{qn('w:date')}"
                    )
                    date_el.set(qn("w:fullDate"), full)
        name = entry.get("name") or entry.get("tag") or entry.get("alias")
        filled[name] = entry["kind"]
    if plan:
        pkg.mark_dirty()
    result = {"filled": filled, "count": len(plan)}
    if unknown:
        result["skipped_unknown"] = sorted(unknown)
    return result


# ------------------------------------------- content controls: full inventory

# sdtPr child -> reported type, in detection priority order. Anything with no
# marker element is a rich-text control.
_SDT_TYPE_MAP = (
    ("w14:checkbox", "checkbox"),
    ("w:dropDownList", "dropdown"),
    ("w:comboBox", "combo"),
    ("w:date", "date"),
    ("w:text", "text"),
    ("w:picture", "picture"),
    ("w:group", "group"),
    ("w:citation", "citation"),
    ("w:bibliography", "bibliography"),
    ("w:equation", "equation"),
    ("w:docPartObj", "gallery"),
    ("w:docPartList", "gallery_list"),
    ("w15:repeatingSection", "repeating_section"),
    ("w15:repeatingSectionItem", "repeating_section_item"),
)

# Types set_content_control_value can write without risking the control's
# bound machinery. Everything else is refused.
_SDT_WRITABLE = ("text", "richtext", "combo", "date", "dropdown", "checkbox")


def _sdt_type(pr: etree._Element) -> str:
    for tag, name in _SDT_TYPE_MAP:
        if pr.find(qn(tag)) is not None:
            return name
    return "richtext"


def _sdt_lock(pr: etree._Element) -> dict:
    val = _sval(pr, "w:lock")
    return {
        "lock": val,
        "content_locked": val in ("contentLocked", "sdtContentLocked"),
        "control_locked": val in ("sdtLocked", "sdtContentLocked"),
    }


def _all_sdts(pkg: DocxPackage) -> list[dict]:
    """EVERY well-formed content control in document order (including the
    gallery/citation/picture types the form inventory skips), with a stable
    per-scan index."""
    out = []
    for i, sdt in enumerate(pkg.root().iter(qn("w:sdt"))):
        pr = sdt.find(qn("w:sdtPr"))
        content = sdt.find(qn("w:sdtContent"))
        if pr is None or content is None:
            continue
        kind = _sdt_type(pr)
        text_value = "".join(run_text(r) for r in content.iter(qn("w:r")))
        placeholder = pr.find(qn("w:showingPlcHdr")) is not None
        entry = {
            "index": i,
            "tag": _sval(pr, "w:tag"),
            "alias": _sval(pr, "w:alias"),
            "type": kind,
            "value": text_value if not placeholder else "",
            "placeholder_showing": placeholder,
            "block": content.find(qn("w:p")) is not None
            or content.find(qn("w:tbl")) is not None,
            "context": _context(_containing_paragraph(sdt)) or text_value[:80],
            **_sdt_lock(pr),
            "_sdt": sdt,
            "_pr": pr,
        }
        if kind == "checkbox":
            checkbox = pr.find(qn("w14:checkbox"))
            checked_el = checkbox.find(qn("w14:checked"))
            checked = (
                checked_el.get(qn("w14:val")) if checked_el is not None else None
            )
            entry["value"] = checked in ("1", "true")
        elif kind in ("dropdown", "combo"):
            src = pr.find(qn("w:dropDownList"))
            if src is None:
                src = pr.find(qn("w:comboBox"))
            entry["_items"] = [
                {
                    "display": li.get(qn("w:displayText"), ""),
                    "value": li.get(qn("w:value"), ""),
                }
                for li in src.findall(qn("w:listItem"))
            ]
            entry["options"] = [it["display"] for it in entry["_items"]]
        out.append(entry)
    return out


def list_content_controls(pkg: DocxPackage) -> dict:
    """Every content control (SDT) in the document body, including the types
    fill_form_fields cannot fill: tag, alias, type (text / richtext /
    checkbox / dropdown / combo / date / picture / group / citation /
    bibliography / equation / gallery / repeating_section), current value,
    lock state, placeholder flag, and whether it is block-level. The index is
    the addressing handle for set_content_control_value."""
    entries = _all_sdts(pkg)
    return {"count": len(entries), "controls": [_public(e) for e in entries]}


def set_content_control_value(
    pkg: DocxPackage,
    value,
    *,
    tag: str | None = None,
    index: int | None = None,
) -> dict:
    """Set one content control's value, addressed by tag or by
    list_content_controls index. Text/rich-text/combo/date controls take a
    string, checkboxes a boolean, dropdowns one of their options. Refuses
    locked controls (contentLocked / sdtContentLocked) and types that cannot
    be safely written (gallery, repeating section, citation, bibliography,
    picture, group, equation); nothing is changed on refusal."""
    if (tag is None) == (index is None):
        raise WordMcpError("give exactly one of tag or index")
    entries = _all_sdts(pkg)
    if tag is not None:
        matches = [e for e in entries if e["tag"] == tag]
        if not matches:
            raise WordMcpError(
                f"no content control with tag {tag!r}; "
                "see list_content_controls"
            )
        if len(matches) > 1:
            raise AmbiguousTarget(
                f"{len(matches)} content controls share tag {tag!r} (indices "
                f"{[e['index'] for e in matches]}); address one by index"
            )
        entry = matches[0]
    else:
        entry = next((e for e in entries if e["index"] == index), None)
        if entry is None:
            raise WordMcpError(
                f"no content control with index {index} "
                f"({len(entries)} controls); see list_content_controls"
            )

    label = entry["tag"] or entry["alias"] or f"index {entry['index']}"
    if entry["content_locked"]:
        raise WordMcpError(
            f"content control {label!r} is locked "
            f"(w:lock={entry['lock']}); Word forbids editing its contents. "
            "Remove the lock in Word (Developer > Properties) first; "
            "nothing was changed"
        )
    kind = entry["type"]
    if kind not in _SDT_WRITABLE:
        raise UnsupportedStructure(
            f"content control {label!r} has type {kind!r}, which cannot be "
            "written safely (its content is bound to gallery/citation/"
            "picture machinery that a text write would disconnect); "
            "nothing was changed"
        )

    if kind == "checkbox":
        _set_sdt_checkbox(entry["_sdt"], _as_bool(label, value))
        stored = bool(value)
    elif kind == "dropdown":
        text = _check_text(label, value)
        item = next(
            (i for i in entry["_items"] if text in (i["display"], i["value"])),
            None,
        )
        if item is None:
            raise WordMcpError(
                f"{text!r} is not an option of dropdown {label!r} "
                f"(options: {entry['options']}); nothing was changed"
            )
        stored = item["display"] or item["value"]
        _set_sdt_content(entry["_sdt"], stored)
    else:
        stored = _check_text(label, value)
        _set_sdt_content(entry["_sdt"], stored)
        if kind == "date":
            full = _iso_full_date(stored)
            if full is not None:
                entry["_pr"].find(qn("w:date")).set(qn("w:fullDate"), full)
    pkg.mark_dirty()
    return {"set": label, "type": kind, "value": stored}


def insert_content_control(
    pkg: DocxPackage,
    *,
    tag: str,
    after_anchor: str,
    alias: str | None = None,
    text: str = "",
    occurrence: int = 1,
) -> dict:
    """Insert a new PLAIN-TEXT content control (inline SDT) immediately after
    `after_anchor` text. This is the one control type whose XML can be built
    verifiably safely; creating checkbox, dropdown, date, picture, gallery,
    or repeating controls is refused (fill and list still cover them when a
    template provides them). The tag must be unique in the document so the
    control stays addressable; `text` is the initial content."""
    import random

    from . import _runmap as _rm
    from .fields import _find_anchor_span

    if not tag or not tag.strip():
        raise WordMcpError("tag must be non-empty")
    tag = tag.strip()
    _check_text("tag", tag)
    if any(e["tag"] == tag for e in _all_sdts(pkg)):
        raise WordMcpError(
            f"a content control with tag {tag!r} already exists; tags must "
            "stay unique so controls remain addressable by tag"
        )
    if text:
        _check_text(tag, text)

    p, _, end = _find_anchor_span(pkg, after_anchor, occurrence)
    covered = _rm.split_for_range(p, end - 1, end)
    ref = covered[-1]

    sdt = etree.Element(qn("w:sdt"))
    pr = etree.SubElement(sdt, qn("w:sdtPr"))
    # CT_SdtPr schema order: alias, tag, id, ... , then the type element.
    if alias:
        etree.SubElement(pr, qn("w:alias")).set(qn("w:val"), alias)
    etree.SubElement(pr, qn("w:tag")).set(qn("w:val"), tag)
    etree.SubElement(pr, qn("w:id")).set(
        qn("w:val"), str(random.randint(1, 2**31 - 1))
    )
    etree.SubElement(pr, qn("w:text"))
    content = etree.SubElement(sdt, qn("w:sdtContent"))
    if text:
        content.append(_new_text_run(ref, text))
    ref.addnext(sdt)
    pkg.mark_dirty()
    return {
        "content_control_inserted": tag,
        "type": "text",
        "alias": alias,
        "initial_text": text,
    }


# ---------------------------------------------------------------- validation


_TEXT_KINDS = ("legacy_text", "sdt_text", "sdt_richtext", "sdt_combo", "sdt_date")


def _is_filled(entry: dict, *, require_checked: bool = False) -> bool:
    kind = entry["kind"]
    if entry.get("placeholder_showing"):
        return False
    if kind in _TEXT_KINDS:
        return bool(str(entry.get("value") or "").strip())
    if kind in ("legacy_checkbox", "sdt_checkbox"):
        return entry["value"] is True if require_checked else True
    if kind in ("legacy_dropdown", "sdt_dropdown"):
        return bool(str(entry.get("value") or "").strip())
    return True


def validate_form_completeness(
    pkg: DocxPackage, required: list[str] | None = None
) -> dict:
    """Report unfilled form fields: empty text, placeholder text still
    showing, and (for names in `required`) unchecked checkboxes. With
    `required`, exactly those fields are checked and missing ones are
    reported; without it, every field is checked."""
    entries = _inventory(pkg)
    unfilled: list[dict] = []
    missing_fields: list[str] = []
    if required is not None:
        for name in required:
            matches = [e for e in entries if name in _match_keys(e)]
            if not matches:
                missing_fields.append(name)
                continue
            for e in matches:
                if not _is_filled(e, require_checked=True):
                    unfilled.append(_public(e))
    else:
        for e in entries:
            if not _is_filled(e):
                unfilled.append(_public(e))
    return {
        "complete": not unfilled and not missing_fields,
        "checked": "required" if required is not None else "all",
        "fields_total": len(entries),
        "unfilled": unfilled,
        "missing_fields": missing_fields,
    }


def delete_content_control(
    pkg: DocxPackage,
    *,
    tag: str | None = None,
    index: int | None = None,
) -> dict:
    """Delete one content control (the WHOLE SDT including its content),
    addressed by tag or by list_content_controls index (exactly one),
    matching every other delete_element branch's remove-the-object
    semantics. Locked controls (either lock flavor) are refused; nothing is
    changed on refusal. v2 addition (V2_DESIGN Section 3.2); v1 had no
    content-control deletion path."""
    if (tag is None) == (index is None):
        raise WordMcpError("give exactly one of tag or index")
    entries = _all_sdts(pkg)
    if tag is not None:
        matches = [e for e in entries if e["tag"] == tag]
        if not matches:
            raise TargetNotFound(
                f"no content control with tag {tag!r}; see "
                "list_elements(type='content_controls')"
            )
        if len(matches) > 1:
            raise AmbiguousTarget(
                f"{len(matches)} content controls share tag {tag!r} (indices "
                f"{[e['index'] for e in matches]}); address one by index"
            )
        entry = matches[0]
    else:
        entry = next((e for e in entries if e["index"] == index), None)
        if entry is None:
            raise TargetNotFound(
                f"no content control with index {index} "
                f"({len(entries)} controls); see "
                "list_elements(type='content_controls')"
            )
    label = entry["tag"] or entry["alias"] or f"index {entry['index']}"
    if entry["content_locked"] or entry["control_locked"]:
        raise WordMcpError(
            f"content control {label!r} is locked (w:lock={entry['lock']}); "
            "Word forbids deleting it. Remove the lock in Word (Developer > "
            "Properties) first; nothing was changed"
        )
    sdt = entry["_sdt"]
    sdt.getparent().remove(sdt)
    pkg.mark_dirty()
    return {
        "deleted_content_control": label,
        "type": entry["type"],
        "was_block": entry["block"],
    }
