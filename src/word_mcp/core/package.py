"""DocxPackage: safe load/save layer for .docx files.

Design guarantees:
- Parts that were never touched are written back byte-for-byte identical.
- Saves are atomic: temp file -> validation -> os.replace. The original is never
  left half-written, and validation failure leaves it untouched.
- Auto-backup on by default: before each in-place save the current content is
  rotated into stable slots (prev.docx / anchor.docx) inside a hidden
  .ks4w-backups/ folder next to the document (see core.safesave).
- Files locked by Word are detected before any work happens.
"""

from __future__ import annotations

import io
import os
import zipfile
from pathlib import Path

from lxml import etree

from .errors import (
    DocumentCorrupt,
    DocumentLocked,
    DocumentNotFound,
    DocumentProtected,
)
from .sandbox import check_path

NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
}


def qn(tag: str) -> str:
    """'w:tbl' -> '{http://...}tbl' (Clark notation)."""
    prefix, local = tag.split(":")
    return f"{{{NSMAP[prefix]}}}{local}"


def _is_ole_encrypted(head: bytes) -> bool:
    # OLE compound file magic = password-protected / legacy binary, not a ZIP.
    return head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")


class DocxPackage:
    """One .docx opened for inspection or editing."""

    def __init__(self, path: str | os.PathLike):
        check_path(path, "open document")  # no-op unless KS4W_ALLOWED_ROOTS is set
        self.path = Path(path)
        self._raw: dict[str, bytes] = {}  # part name -> original bytes
        self._order: list[str] = []  # original entry order, preserved on save
        self._trees: dict[str, etree._ElementTree] = {}  # parsed parts
        self._dirty: set[str] = set()  # parts whose tree must be re-serialized
        self._load()

    # ---------- loading ----------

    def _load(self) -> None:
        if not self.path.exists():
            raise DocumentNotFound(f"No file at {self.path}")
        self._check_lock()
        head = self.path.read_bytes()[:8] if self.path.stat().st_size >= 8 else b""
        if _is_ole_encrypted(head):
            raise DocumentProtected(
                f"{self.path.name} is password-protected or a legacy .doc; "
                "remove the password in Word first."
            )
        try:
            with zipfile.ZipFile(self.path) as zf:
                bad = zf.testzip()
                if bad is not None:
                    raise DocumentCorrupt(
                        f"{self.path.name}: corrupt ZIP entry '{bad}'."
                    )
                for info in zf.infolist():
                    self._raw[info.filename] = zf.read(info.filename)
                    self._order.append(info.filename)
        except zipfile.BadZipFile as exc:
            raise DocumentCorrupt(
                f"{self.path.name} is not a valid .docx (bad ZIP): {exc}"
            ) from exc
        if "word/document.xml" not in self._raw:
            raise DocumentCorrupt(
                f"{self.path.name} has no word/document.xml; not a Word document."
            )

    def _check_lock(self) -> None:
        """Word holds an exclusive lock on open docs; detect it up front."""
        owner_file = self.path.with_name("~$" + self.path.name[-153:])
        try:
            with open(self.path, "r+b"):
                pass
        except PermissionError:
            hint = " (Word owner file present)" if owner_file.exists() else ""
            raise DocumentLocked(
                f"{self.path.name} is open in Word or locked by another process{hint}. "
                "Dual-mode tools route to the open document automatically "
                "(live='auto'); com_save_open_document / "
                "com_close_open_document save or release it. This call has "
                "no live route (or live routing was turned off) and needs "
                "the file closed."
            ) from None

    # ---------- part access ----------

    def has_part(self, name: str) -> bool:
        return name in self._raw

    def part_names(self) -> list[str]:
        return list(self._order)

    def tree(self, name: str = "word/document.xml") -> etree._ElementTree:
        """Parsed XML for a part. Parsing is cached; call mark_dirty() after edits."""
        if name not in self._trees:
            if name not in self._raw:
                raise KeyError(f"part not in package: {name}")
            parser = etree.XMLParser(remove_blank_text=False, strip_cdata=False)
            self._trees[name] = etree.ElementTree(
                etree.fromstring(self._raw[name], parser=parser)
            )
        return self._trees[name]

    def root(self, name: str = "word/document.xml") -> etree._Element:
        return self.tree(name).getroot()

    def body(self) -> etree._Element:
        body = self.root().find(qn("w:body"))
        if body is None:
            raise DocumentCorrupt("document.xml has no w:body")
        return body

    def raw_part(self, name: str) -> bytes:
        return self._raw[name]

    def set_raw_part(self, name: str, data: bytes) -> None:
        """Add or replace a part with raw bytes (images, new XML parts)."""
        if name not in self._raw:
            self._order.append(name)
        self._raw[name] = data
        self._trees.pop(name, None)
        self._dirty.discard(name)  # raw bytes are authoritative now

    def mark_dirty(self, name: str = "word/document.xml") -> None:
        if name not in self._trees:
            raise RuntimeError(f"mark_dirty before tree() for {name}")
        self._dirty.add(name)

    # ---------- saving ----------

    def _serialize(self, name: str) -> bytes:
        tree = self._trees[name]
        return etree.tostring(
            tree, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    def save(self, dest: str | os.PathLike | None = None, *, do_backup: bool = True) -> Path:
        """Atomic save. dest=None means save in place (with slot backup rotation
        by default: prev.docx/anchor.docx under .ks4w-backups/, see core.safesave).
        do_backup=False skips the rotation only; the atomic validated save always
        applies."""
        from . import safesave

        dest_path = Path(dest) if dest else self.path
        check_path(dest_path, "save document")
        in_place = dest_path.resolve() == self.path.resolve()

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in self._order:
                data = self._serialize(name) if name in self._dirty else self._raw[name]
                zf.writestr(name, data)
        payload = buf.getvalue()

        # Validate the payload before touching the destination.
        self._validate_payload(payload)

        tmp = dest_path.with_name(dest_path.name + ".word-mcp-tmp")
        tmp.write_bytes(payload)
        try:
            # Capture the current (pre-mutation) content into the backup slots
            # via hardlink-then-replace: the document never leaves its own
            # path, and the old bytes survive the final replace below under
            # the prev slot's directory entry.
            if in_place and do_backup and dest_path.exists():
                safesave.rotate_slots(dest_path)
            safesave.replace_with_retry(tmp, dest_path)
        except Exception:
            tmp.unlink(missing_ok=True)
            raise
        # After a successful save the written bytes are the new baseline.
        if in_place:
            for name in self._dirty:
                self._raw[name] = self._serialize(name)
            self._dirty.clear()
        return dest_path

    @staticmethod
    def _validate_payload(payload: bytes) -> None:
        """Structural sanity: valid ZIP, well-formed XML in every dirty-able part."""
        from .errors import ValidationFailed

        try:
            with zipfile.ZipFile(io.BytesIO(payload)) as zf:
                names = set(zf.namelist())
                if "word/document.xml" not in names:
                    raise ValidationFailed("output lost word/document.xml")
                if "[Content_Types].xml" not in names:
                    raise ValidationFailed("output lost [Content_Types].xml")
                for name in names:
                    if name.endswith((".xml", ".rels")):
                        try:
                            etree.fromstring(zf.read(name))
                        except etree.XMLSyntaxError as exc:
                            raise ValidationFailed(
                                f"output part {name} is not well-formed XML: {exc}"
                            ) from exc
        except zipfile.BadZipFile as exc:  # pragma: no cover
            raise ValidationFailed(f"output is not a valid ZIP: {exc}") from exc
