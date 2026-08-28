"""Images: insert, list, replace, resize. Inline drawings via DrawingML."""

from __future__ import annotations

import struct
from pathlib import Path

from lxml import etree

from ..core.errors import TargetNotFound, WordMcpError
from ..core.package import DocxPackage, qn
from ..core.sandbox import check_path

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

_EXT_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".tiff": "image/tiff",
    ".emf": "image/x-emf",
    ".wmf": "image/x-wmf",
}

EMU_PER_INCH = 914400
EMU_PER_PT = 12700


def _image_size_px(data: bytes, ext: str) -> tuple[int, int] | None:
    """Native pixel size for PNG/JPEG/GIF without external deps."""
    try:
        if ext == ".png" and data[:8] == b"\x89PNG\r\n\x1a\n":
            w, h = struct.unpack(">II", data[16:24])
            return w, h
        if ext == ".gif":
            w, h = struct.unpack("<HH", data[6:10])
            return w, h
        if ext in (".jpg", ".jpeg"):
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                    h, w = struct.unpack(">HH", data[i + 5 : i + 9])
                    return w, h
                seg_len = struct.unpack(">H", data[i + 2 : i + 4])[0]
                i += 2 + seg_len
    except Exception:
        pass
    return None


def _next_docpr_id(pkg: DocxPackage) -> int:
    ids = [
        int(el.get("id", "0"))
        for el in pkg.root().iter(f"{{{_WP}}}docPr")
    ]
    return max(ids, default=0) + 1


def add_image(
    pkg: DocxPackage,
    image_path: str,
    *,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    width_pt: float | None = None,
    alignment: str = "center",
) -> dict:
    """Insert an inline image in its own centered paragraph. Width defaults to
    the native size capped at 6.5in; height keeps aspect ratio."""
    check_path(image_path, "read image file")
    src = Path(image_path)
    if not src.exists():
        raise TargetNotFound(f"image file not found: {image_path}")
    ext = src.suffix.lower()
    if ext not in _EXT_TYPES:
        raise WordMcpError(
            f"unsupported image type {ext}; use {sorted(_EXT_TYPES)}"
        )
    if width_pt is not None and width_pt <= 0:
        raise WordMcpError("width_pt must be positive")
    data = src.read_bytes()

    # Media part.
    n = 1
    while pkg.has_part(f"word/media/image{n}{ext}") or any(
        name.startswith(f"word/media/image{n}.") for name in pkg.part_names()
    ):
        n += 1
    media_part = f"word/media/image{n}{ext}"
    pkg.set_raw_part(media_part, data)

    # Content type default for the extension.
    ct_root = pkg.root("[Content_Types].xml")
    ext_name = ext.lstrip(".")
    if not any(
        d.get("Extension") == ext_name
        for d in ct_root.findall(f"{{{_CT_NS}}}Default")
    ):
        default = etree.SubElement(ct_root, f"{{{_CT_NS}}}Default")
        default.set("Extension", ext_name)
        default.set("ContentType", _EXT_TYPES[ext])
        pkg.mark_dirty("[Content_Types].xml")

    # Relationship.
    rels_part = "word/_rels/document.xml.rels"
    rels_root = pkg.root(rels_part)
    existing = {r.get("Id") for r in rels_root}
    i = 1
    while f"rId{i}" in existing:
        i += 1
    rid = f"rId{i}"
    rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
    rel.set("Id", rid)
    rel.set(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    )
    rel.set("Target", "media/" + media_part.rsplit("/", 1)[1])
    pkg.mark_dirty(rels_part)

    # Dimensions.
    px = _image_size_px(data, ext)
    if px:
        native_w_emu = int(px[0] / 96 * EMU_PER_INCH)
        native_h_emu = int(px[1] / 96 * EMU_PER_INCH)
    else:
        native_w_emu = native_h_emu = int(3 * EMU_PER_INCH)
    if width_pt:
        cx = int(width_pt * EMU_PER_PT)
    else:
        cx = min(native_w_emu, int(6.5 * EMU_PER_INCH))
    cy = int(cx * native_h_emu / native_w_emu)

    # Inline drawing.
    docpr_id = _next_docpr_id(pkg)
    p = etree.Element(qn("w:p"))
    ppr = etree.SubElement(p, qn("w:pPr"))
    jc = etree.SubElement(ppr, qn("w:jc"))
    jc.set(
        qn("w:val"),
        {"left": "left", "center": "center", "right": "right"}[alignment],
    )
    r = etree.SubElement(p, qn("w:r"))
    drawing = etree.SubElement(r, qn("w:drawing"))
    inline = etree.SubElement(drawing, f"{{{_WP}}}inline")
    for attr in ("distT", "distB", "distL", "distR"):
        inline.set(attr, "0")
    extent = etree.SubElement(inline, f"{{{_WP}}}extent")
    extent.set("cx", str(cx))
    extent.set("cy", str(cy))
    docpr = etree.SubElement(inline, f"{{{_WP}}}docPr")
    docpr.set("id", str(docpr_id))
    docpr.set("name", f"Picture {docpr_id}")
    graphic = etree.SubElement(inline, f"{{{_A}}}graphic")
    gdata = etree.SubElement(graphic, f"{{{_A}}}graphicData")
    gdata.set("uri", _PIC)
    pic = etree.SubElement(gdata, f"{{{_PIC}}}pic")
    nvpr = etree.SubElement(pic, f"{{{_PIC}}}nvPicPr")
    cnv = etree.SubElement(nvpr, f"{{{_PIC}}}cNvPr")
    cnv.set("id", str(docpr_id))
    cnv.set("name", src.name)
    etree.SubElement(nvpr, f"{{{_PIC}}}cNvPicPr")
    blipfill = etree.SubElement(pic, f"{{{_PIC}}}blipFill")
    blip = etree.SubElement(blipfill, f"{{{_A}}}blip")
    blip.set(f"{{{_R_NS}}}embed", rid)
    stretch = etree.SubElement(blipfill, f"{{{_A}}}stretch")
    etree.SubElement(stretch, f"{{{_A}}}fillRect")
    sppr = etree.SubElement(pic, f"{{{_PIC}}}spPr")
    xfrm = etree.SubElement(sppr, f"{{{_A}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{_A}}}off")
    off.set("x", "0")
    off.set("y", "0")
    ext_el = etree.SubElement(xfrm, f"{{{_A}}}ext")
    ext_el.set("cx", str(cx))
    ext_el.set("cy", str(cy))
    geom = etree.SubElement(sppr, f"{{{_A}}}prstGeom")
    geom.set("prst", "rect")
    etree.SubElement(geom, f"{{{_A}}}avLst")

    from .text import _body_paragraph, _resolve_anchor

    body = pkg.body()
    if at_end or (after_index is None and after_anchor is None):
        sectpr = body.find(qn("w:sectPr"))
        if sectpr is not None:
            sectpr.addprevious(p)
        else:
            body.append(p)
    elif after_anchor is not None:
        _resolve_anchor(pkg, after_anchor).addnext(p)
    else:
        _body_paragraph(pkg, after_index).addnext(p)
    pkg.mark_dirty()
    return {
        "image_added": media_part,
        "width_pt": round(cx / EMU_PER_PT, 1),
        "height_pt": round(cy / EMU_PER_PT, 1),
    }


def list_images(pkg: DocxPackage) -> list[dict]:
    rels_root = (
        pkg.root("word/_rels/document.xml.rels")
        if pkg.has_part("word/_rels/document.xml.rels")
        else None
    )
    rid_target = (
        {r.get("Id"): r.get("Target") for r in rels_root}
        if rels_root is not None
        else {}
    )
    out = []
    for i, blip in enumerate(pkg.root().iter(f"{{{_A}}}blip")):
        rid = blip.get(f"{{{_R_NS}}}embed")
        entry = {"index": i, "rel_id": rid, "target": rid_target.get(rid)}
        inline = blip.getparent()
        while inline is not None and not inline.tag.endswith("}inline"):
            inline = inline.getparent()
        if inline is not None:
            extent = inline.find(f"{{{_WP}}}extent")
            if extent is not None:
                entry["width_pt"] = round(int(extent.get("cx")) / EMU_PER_PT, 1)
                entry["height_pt"] = round(int(extent.get("cy")) / EMU_PER_PT, 1)
        out.append(entry)
    return out


def resize_image(pkg: DocxPackage, image_index: int, *, width_pt: float) -> dict:
    """Resize by index (as reported by list_images), keeping aspect ratio."""
    blips = list(pkg.root().iter(f"{{{_A}}}blip"))
    if not 0 <= image_index < len(blips):
        raise TargetNotFound(
            f"image index {image_index} out of range ({len(blips)} images)"
        )
    if width_pt <= 0:
        raise WordMcpError("width_pt must be positive")
    blip = blips[image_index]
    container = blip.getparent()
    while container is not None and not container.tag.endswith("}inline"):
        container = container.getparent()
    if container is None:
        raise WordMcpError(
            "image is floating/anchored, not inline; resize it in Word"
        )
    extent = container.find(f"{{{_WP}}}extent")
    old_cx = int(extent.get("cx"))
    old_cy = int(extent.get("cy"))
    new_cx = int(width_pt * EMU_PER_PT)
    new_cy = int(new_cx * old_cy / old_cx)
    extent.set("cx", str(new_cx))
    extent.set("cy", str(new_cy))
    for xfrm_ext in container.iter(f"{{{_A}}}ext"):
        xfrm_ext.set("cx", str(new_cx))
        xfrm_ext.set("cy", str(new_cy))
    pkg.mark_dirty()
    return {
        "resized": image_index,
        "width_pt": width_pt,
        "height_pt": round(new_cy / EMU_PER_PT, 1),
    }


def replace_image(pkg: DocxPackage, image_index: int, new_image_path: str) -> dict:
    """Swap an image's bytes, keeping placement and display size."""
    check_path(new_image_path, "read image file")
    src = Path(new_image_path)
    if not src.exists():
        raise TargetNotFound(f"image file not found: {new_image_path}")
    ext = src.suffix.lower()
    if ext not in _EXT_TYPES:
        raise WordMcpError(f"unsupported image type {ext}")
    blips = list(pkg.root().iter(f"{{{_A}}}blip"))
    if not 0 <= image_index < len(blips):
        raise TargetNotFound(f"image index {image_index} out of range")
    rid = blips[image_index].get(f"{{{_R_NS}}}embed")
    rels_root = pkg.root("word/_rels/document.xml.rels")
    target = next((r.get("Target") for r in rels_root if r.get("Id") == rid), None)
    if target is None:
        raise TargetNotFound(f"no relationship for image {image_index}")
    part = "word/" + target.lstrip("/")
    old_ext = "." + part.rsplit(".", 1)[1].lower()
    if old_ext != ext:
        raise WordMcpError(
            f"replacement must be the same type as the original ({old_ext}); "
            "or add a new image and delete this one"
        )
    pkg.set_raw_part(part, src.read_bytes())
    return {"replaced": part}
