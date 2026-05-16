#!/usr/bin/env python3
"""Generate the macOS menu-bar / Windows tray icon.

A two-point "sparkle" mark — a large 4-point star with a small companion
star upper-right. Reads instantly as "AI assistant", crisp at 22 px, and
is not a letter or any prior project's glyph.

Output (black shape on transparent — a macOS *template* image; the OS
recolors it for light/dark menu bars, and app/main.mjs calls
setTemplateImage(true) on darwin):

  app/tray-icon.png       22x22  (1x)
  app/tray-icon@2x.png    44x44  (Retina; Electron auto-selects it)

Regenerate:  python3 scripts/build-tray-icon.py
"""
import math
from pathlib import Path
from PIL import Image, ImageDraw

OUT_DIR = Path(__file__).resolve().parent.parent / "app"
FG = (0, 0, 0, 255)  # template: pure black + alpha; OS handles the color


def _star_points(cx, cy, r_out, r_in):
    """8-vertex 4-point star: sharp cardinal points, pinched diagonal waist."""
    pts = []
    for i in range(8):
        ang = -math.pi / 2 + i * (math.pi / 4)  # start at top, go clockwise
        r = r_out if i % 2 == 0 else r_in
        pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    return pts


def make(size: int, out: Path) -> None:
    # Supersample 4x then downscale for clean antialiased edges at tiny sizes.
    ss = 4
    S = size * ss
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # Primary sparkle: lower-left of center, fills most of the canvas.
    pcx, pcy = S * 0.42, S * 0.56
    p_out = S * 0.46
    d.polygon(_star_points(pcx, pcy, p_out, p_out * 0.16), fill=FG)

    # Companion sparkle: upper-right, ~40% the size — the "shimmer".
    scx, scy = S * 0.76, S * 0.24
    s_out = S * 0.20
    d.polygon(_star_points(scx, scy, s_out, s_out * 0.18), fill=FG)

    img = img.resize((size, size), Image.LANCZOS)
    img.save(out, "PNG")
    print(f"wrote {out} ({size}x{size})")


if __name__ == "__main__":
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    make(22, OUT_DIR / "tray-icon.png")
    make(44, OUT_DIR / "tray-icon@2x.png")
