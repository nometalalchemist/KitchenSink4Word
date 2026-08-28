"""Generate bundle/icon.png for the KitchenSink4Word .mcpb bundle.

Pure stdlib (zlib + struct): 512x512 RGB PNG (Claude Desktop's recommended
icon size), faucet-blue square with a white water-drop mark. Rendered with
4x4 supersampling for smooth edges.

Usage: python -X utf8 make_icon.py
"""

import math
import struct
import zlib
from pathlib import Path

SIZE = 512          # output pixels
SS = 4              # supersample factor
BG = (27, 105, 166)     # faucet blue
FG = (247, 251, 253)    # near-white drop

# Drop geometry in output-pixel space: apex above a circle, joined by the
# exact tangent wedge so the silhouette is a clean teardrop.
APEX = (SIZE * 0.5, SIZE * 0.211)
CENTER = (SIZE * 0.5, SIZE * 0.617)
RADIUS = SIZE * 0.2227


def inside_drop(x: float, y: float) -> bool:
    ax, ay = APEX
    cx, cy = CENTER
    r = RADIUS
    # circle part
    dx, dy = x - cx, y - cy
    if dx * dx + dy * dy <= r * r:
        return True
    # tangent wedge from apex toward the circle
    px, py = x - ax, y - ay
    vx, vy = cx - ax, cy - ay
    d = math.hypot(vx, vy)
    if d <= r:
        return False
    t_len = math.sqrt(d * d - r * r)   # apex-to-tangent-point length
    p_len = math.hypot(px, py)
    if p_len == 0.0:
        return True
    if p_len > t_len:
        return False
    cos_a = t_len / d                  # cos of wedge half-angle
    cos_p = (px * vx + py * vy) / (p_len * d)
    return cos_p >= cos_a


def render() -> bytes:
    n = SIZE * SS
    scale = 1.0 / SS
    rows = []
    for oy in range(SIZE):
        row = bytearray()
        row.append(0)  # PNG filter type 0 (None)
        for ox in range(SIZE):
            hit = 0
            for sy in range(SS):
                for sx in range(SS):
                    x = (ox * SS + sx + 0.5) * scale
                    y = (oy * SS + sy + 0.5) * scale
                    if inside_drop(x, y):
                        hit += 1
            cov = hit / (SS * SS)
            row.extend(
                round(BG[i] + (FG[i] - BG[i]) * cov) for i in range(3)
            )
        rows.append(bytes(row))
    return b"".join(rows)


def write_png(path: Path, raw: bytes) -> None:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", SIZE, SIZE, 8, 2, 0, 0, 0)  # 8-bit RGB
    png = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw, 9))
        + chunk(b"IEND", b"")
    )
    path.write_bytes(png)


if __name__ == "__main__":
    out = Path(__file__).parent / "icon.png"
    write_png(out, render())
    print(f"wrote {out} ({out.stat().st_size} bytes)")
