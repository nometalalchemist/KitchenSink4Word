"""Document protection (editing restrictions) with Word-compatible password
hashing.

The hash algorithm reproduces Word 365's rsaAES / SHA-512 / 100000-spin
documentProtection byte-for-byte (verified against Word-generated files, see
research/20260827_v12 topic 4): the password runs through the legacy Office
15-char verifier-key derivation, is re-encoded as zero-padded uppercase hex in
UTF-16LE, then salted-SHA-512 iterated with an appended little-endian counter.

Protection is an intent/integrity control, NOT encryption: the settings XML
can be stripped by anyone with a zip tool. For the committee workflow the
valuable mode is edit="trackedChanges" — every edit the recipient makes is
forced to be a tracked change.
"""

from __future__ import annotations

import base64
import hashlib
import os

from lxml import etree

from ..core.errors import WordMcpError
from ..core.package import NSMAP, DocxPackage, qn

_EDIT_MODES = ("readOnly", "comments", "trackedChanges", "forms")

_INIT = [
    0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E,
    0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3,
]
_MATRIX = [
    [0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09],
    [0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF],
    [0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0],
    [0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40],
    [0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5],
    [0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A],
    [0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9],
    [0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0],
    [0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC],
    [0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10],
    [0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168],
    [0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C],
    [0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD],
    [0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC],
    [0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4],
]


def _legacy_key_utf16le(password: str) -> bytes:
    pw = password[:15]
    b = []
    for ch in pw:
        t = ord(ch)
        v = t & 0xFF
        if v == 0:
            v = (t & 0xFF00) >> 8
        b.append(v)
    n = len(b)
    hi = _INIT[n - 1]
    for i in range(n):
        row = 15 - n + i
        for bit in range(7):
            if b[i] & (1 << bit):
                hi ^= _MATRIX[row][bit]
    lo = 0
    for i in range(n - 1, -1, -1):
        lo = (((lo >> 14) & 1) | ((lo << 1) & 0x7FFF)) ^ b[i]
    lo = (((lo >> 14) & 1) | ((lo << 1) & 0x7FFF)) ^ n ^ 0xCE4B
    combined = ((hi << 16) + lo) & 0xFFFFFFFF
    key4 = bytes((combined >> (i * 8)) & 0xFF for i in range(4))
    hexstr = "".join(f"{x:02X}" for x in key4)  # zero-padded, uppercase
    return hexstr.encode("utf-16-le")


def word_protection_hash(
    password: str, salt: bytes | None = None, spin_count: int = 100000
) -> tuple[str, str]:
    """(hash_b64, salt_b64) matching Word's rsaAES/SHA-512 documentProtection."""
    if salt is None:
        salt = os.urandom(16)
    key = _legacy_key_utf16le(password)
    h = hashlib.sha512(salt + key).digest()  # H0, not counted in spin
    for i in range(spin_count):
        h = hashlib.sha512(h + i.to_bytes(4, "little")).digest()
    return base64.b64encode(h).decode(), base64.b64encode(salt).decode()


def _settings_insert(root: etree._Element, el: etree._Element) -> None:
    """Place documentProtection after zoom/trackChanges, before
    defaultTabStop/compat (the verified practical anchors)."""
    for tag in ("w:defaultTabStop", "w:autoHyphenation", "w:compat", "w:rsids"):
        anchor = root.find(qn(tag))
        if anchor is not None:
            anchor.addprevious(el)
            return
    root.append(el)


def set_document_protection(
    pkg: DocxPackage,
    *,
    edit: str = "trackedChanges",
    password: str | None = None,
    restrict_formatting: bool = False,
) -> dict:
    """Restrict editing. edit: readOnly | comments | trackedChanges | forms.
    trackedChanges = every edit the recipient makes is forced to be a tracked
    change (the committee-review mode). With a password, turning protection
    off requires it; without, anyone can lift it in one click (still useful as
    a strong default). This is NOT encryption."""
    if edit not in _EDIT_MODES:
        raise WordMcpError(f"edit must be one of {_EDIT_MODES}")
    if not pkg.has_part("word/settings.xml"):
        _create_settings_part(pkg)
    root = pkg.root("word/settings.xml")
    existing = root.find(qn("w:documentProtection"))
    if existing is not None:
        root.remove(existing)

    dp = etree.Element(qn("w:documentProtection"))
    dp.set(qn("w:edit"), edit)
    if restrict_formatting:
        dp.set(qn("w:formatting"), "1")
    dp.set(qn("w:enforcement"), "1")
    if password:
        hash_b64, salt_b64 = word_protection_hash(password)
        dp.set(qn("w:cryptProviderType"), "rsaAES")
        dp.set(qn("w:cryptAlgorithmClass"), "hash")
        dp.set(qn("w:cryptAlgorithmType"), "typeAny")
        dp.set(qn("w:cryptAlgorithmSid"), "14")  # SHA-512
        dp.set(qn("w:cryptSpinCount"), "100000")
        dp.set(qn("w:hash"), hash_b64)
        dp.set(qn("w:salt"), salt_b64)
    _settings_insert(root, dp)
    pkg.mark_dirty("word/settings.xml")
    return {
        "protection": edit,
        "password_protected": bool(password),
        "restrict_formatting": restrict_formatting,
    }


def _create_settings_part(pkg: DocxPackage) -> None:
    """Minimal word/settings.xml + content-type override + rel — bare OOXML
    from other producers can lack it entirely."""
    root = etree.Element(qn("w:settings"), nsmap={"w": NSMAP["w"]})
    pkg.set_raw_part(
        "word/settings.xml",
        etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        ),
    )
    ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    ct_root = pkg.root("[Content_Types].xml")
    if not any(
        o.get("PartName") == "/word/settings.xml"
        for o in ct_root.findall(f"{{{ct_ns}}}Override")
    ):
        override = etree.SubElement(ct_root, f"{{{ct_ns}}}Override")
        override.set("PartName", "/word/settings.xml")
        override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.settings+xml",
        )
        pkg.mark_dirty("[Content_Types].xml")
    rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rel_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/"
        "relationships/settings"
    )
    rels_root = pkg.root("word/_rels/document.xml.rels")
    if not any(r.get("Type") == rel_type for r in rels_root):
        existing = {r.get("Id") for r in rels_root}
        n = 1
        while f"rId{n}" in existing:
            n += 1
        rel = etree.SubElement(rels_root, f"{{{rel_ns}}}Relationship")
        rel.set("Id", f"rId{n}")
        rel.set("Type", rel_type)
        rel.set("Target", "settings.xml")
        pkg.mark_dirty("word/_rels/document.xml.rels")


def remove_document_protection(pkg: DocxPackage) -> dict:
    """Lift the editing restriction (works regardless of password — the hash
    only gates Word's UI, not the XML)."""
    if not pkg.has_part("word/settings.xml"):
        # nothing to remove — explicit no-op, consistent with the
        # already-unprotected case
        return {"removed": False, "note": "document has no settings.xml"}
    root = pkg.root("word/settings.xml")
    dp = root.find(qn("w:documentProtection"))
    if dp is None:
        return {
            "protection_removed": False,
            "note": "document was not protected; nothing changed",
        }
    root.remove(dp)
    pkg.mark_dirty("word/settings.xml")
    return {"protection_removed": True}


def get_protection(pkg: DocxPackage) -> dict:
    if not pkg.has_part("word/settings.xml"):
        return {"protected": False}
    dp = pkg.root("word/settings.xml").find(qn("w:documentProtection"))
    if dp is None:
        return {"protected": False}
    return {
        "protected": dp.get(qn("w:enforcement")) in ("1", "true", None),
        "edit": dp.get(qn("w:edit")),
        "password_protected": dp.get(qn("w:hash")) is not None,
        "restrict_formatting": dp.get(qn("w:formatting")) in ("1", "true"),
    }
