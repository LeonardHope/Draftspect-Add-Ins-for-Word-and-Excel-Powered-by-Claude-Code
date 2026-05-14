#!/usr/bin/env python3
"""Generate placeholder add-in icons.

Outputs taskpane/icon-32.png and taskpane/icon-80.png — referenced by the
Office manifests (IconUrl / HighResolutionIconUrl). Plain "OC" mark on a
soft-coral background so it's clearly a placeholder; swap in real artwork
before any non-internal release.
"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT_DIR = Path(__file__).resolve().parent.parent / "taskpane"
BG = (217, 119, 87)   # coral
FG = (255, 255, 255)


def make(size: int, out: Path) -> None:
    img = Image.new("RGBA", (size, size), BG)
    draw = ImageDraw.Draw(img)

    # Rounded corners (mask-and-paste so the alpha is correct).
    radius = max(2, size // 6)
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle([(0, 0), (size, size)], radius, fill=255)
    rounded = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    rounded.paste(img, (0, 0), mask)

    draw = ImageDraw.Draw(rounded)
    text = "OC"
    # Find a font size that fits; macOS ships SF Pro / Helvetica.
    for candidate in [
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/SFNS.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    ]:
        if Path(candidate).exists():
            font_path = candidate
            break
    else:
        font_path = None

    font_size = max(8, int(size * 0.5))
    if font_path:
        font = ImageFont.truetype(font_path, font_size)
    else:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pos = ((size - tw) / 2 - bbox[0], (size - th) / 2 - bbox[1])
    draw.text(pos, text, fill=FG, font=font)

    rounded.save(out, "PNG")
    print(f"wrote {out} ({size}×{size})")


if __name__ == "__main__":
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    make(32, OUT_DIR / "icon-32.png")
    make(80, OUT_DIR / "icon-80.png")
