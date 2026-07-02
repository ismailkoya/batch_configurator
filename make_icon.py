"""Generate batch_configurator.ico — a multi-resolution Windows icon.

Run once before PyInstaller. Output: batch_configurator.ico in this folder,
with 16/32/48/64/128/256 sizes embedded (Windows picks the best per context:
title bar, taskbar, Alt-Tab, file explorer thumbnail).

Style: solid amber disc (matches the configurator's accent colour) with white
'BC' wordmark in bold. Cheap, clean, recognisable at 16x16.

Requires Pillow:   pip install pillow
"""
from PIL import Image, ImageDraw, ImageFont

AMBER = (245, 158, 11, 255)   # #f59e0b — same accent as the web configurator
WHITE = (255, 255, 255, 255)
TEXT  = "BC"

def render(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d   = ImageDraw.Draw(img)
    # Background disc with a slim inset so the edge doesn't hit the icon bounds.
    inset = max(1, size // 32)
    d.ellipse((inset, inset, size - inset, size - inset), fill=AMBER)
    # Pick the boldest sans-serif font Windows ships.
    pt = int(size * 0.50)
    font = None
    for face in ("seguibd.ttf", "arialbd.ttf", "tahomabd.ttf", "arial.ttf"):
        try: font = ImageFont.truetype(face, pt); break
        except IOError: continue
    if font is None: font = ImageFont.load_default()
    # Centre the text via the actual bounding box.
    bbox = d.textbbox((0, 0), TEXT, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    tx = (size - tw) // 2 - bbox[0]
    ty = (size - th) // 2 - bbox[1]
    d.text((tx, ty), TEXT, fill=WHITE, font=font)
    return img


def main():
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    base = render(256)   # save() uses the largest, downsamples for the others
    base.save("batch_configurator.ico", sizes=sizes)
    print(f"Wrote batch_configurator.ico  ({len(sizes)} resolutions)")


if __name__ == "__main__":
    main()
