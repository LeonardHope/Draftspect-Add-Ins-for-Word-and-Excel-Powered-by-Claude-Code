#!/usr/bin/env python3
"""Generate the Office add-in icons (manifest IconUrl / HighResolutionIconUrl).

Same two-point "sparkle" mark as the menu-bar/tray icon
(scripts/build-tray-icon.py), but as a full-color app tile: white sparkle
on the coral brand background with rounded corners — what Office shows in
the Add-ins gallery and the task-pane header.

Outputs taskpane/icon-32.png and taskpane/icon-80.png (the sizes the
manifests reference).

Regenerate:  python3 scripts/build-icons.py
"""
import math
from pathlib import Path

from PIL import Image, ImageDraw

OUT_DIR = Path(__file__).resolve().parent.parent / "taskpane"
BG = (217, 119, 87, 255)  # coral brand background
FG = (255, 255, 255, 255)  # white sparkle


def _star_points(cx, cy, r_out, r_in):
    """8-vertex 4-point star: sharp cardinal points, pinched diagonal waist."""
    pts = []
    for i in range(8):
        ang = -math.pi / 2 + i * (math.pi / 4)  # start at top, clockwise
        r = r_out if i % 2 == 0 else r_in
        pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    return pts


def make(size: int, out: Path) -> None:
    # Supersample 4x then downscale for clean antialiased edges at small sizes.
    ss = 4
    S = size * ss
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # Rounded-rectangle coral tile.
    radius = max(2, S // 6)
    d.rounded_rectangle([(0, 0), (S - 1, S - 1)], radius, fill=BG)

    # Primary sparkle lower-left of center; companion sparkle upper-right —
    # identical composition to the tray mark so the brand reads consistently.
    pcx, pcy = S * 0.42, S * 0.56
    p_out = S * 0.42
    d.polygon(_star_points(pcx, pcy, p_out, p_out * 0.16), fill=FG)

    scx, scy = S * 0.76, S * 0.24
    s_out = S * 0.18
    d.polygon(_star_points(scx, scy, s_out, s_out * 0.18), fill=FG)

    img = img.resize((size, size), Image.LANCZOS)
    img.save(out, "PNG")
    print(f"wrote {out} ({size}x{size})")


if __name__ == "__main__":
    # 32 = IconUrl, 64 = HighResolutionIconUrl. These are the sizes the
    # Office manifest spec documents; an off-spec size (we used to ship 80)
    # can make Office reject the high-res icon and fall back to a blank
    # tile.
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    make(32, OUT_DIR / "icon-32.png")
    make(64, OUT_DIR / "icon-64.png")
