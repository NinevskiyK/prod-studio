"""Turn the raw app screenshots (phone-on-lavender JPEGs) into clean floating
phone PNGs: transparent margin + soft baked drop-shadow. Output -> assets/."""
import os
from PIL import Image, ImageDraw, ImageFilter

SRC = os.path.join(os.path.dirname(__file__), "..", "spenday")
OUT = os.path.join(os.path.dirname(__file__), "assets")
os.makedirs(OUT, exist_ok=True)

# raw filename -> semantic name
MAP = {
    "2.jpeg": "feed",
    "3.jpeg": "feed2",
    "4.jpeg": "scenario",
    "5.jpeg": "scenario2",
    "6.jpeg": "vote",
    "7.jpeg": "vote2",
    "8.jpeg": "build",
    "9.jpeg": "build2",
    "10.jpeg": "calc",
    "11.jpeg": "calc2",
    "Unknown.jpeg": "profile",
}

PAD = 70           # transparent padding around the phone for the shadow
SH_BLUR = 24       # shadow softness
SH_OFFY = 16       # shadow vertical offset
SH_ALPHA = 0.40    # shadow opacity


def cutout(path):
    """Flood-fill the lavender margin from the 4 corners -> transparent."""
    im = Image.open(path).convert("RGB")
    w, h = im.size
    key = (255, 0, 255)
    flood = im.copy()
    for seed in [(1, 1), (w - 2, 1), (1, h - 2), (w - 2, h - 2)]:
        ImageDraw.floodfill(flood, seed, key, thresh=70)
    rgba = im.convert("RGBA")
    px_f = flood.load()
    px_a = rgba.load()
    for y in range(h):
        for x in range(w):
            if px_f[x, y] == key:
                px_a[x, y] = (0, 0, 0, 0)
    # crop to the opaque bbox
    bbox = rgba.getbbox()
    return rgba.crop(bbox)


def with_shadow(phone):
    w, h = phone.size
    canvas = Image.new("RGBA", (w + 2 * PAD, h + 2 * PAD), (0, 0, 0, 0))
    # shadow from alpha silhouette
    alpha = phone.split()[3]
    shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    sh_alpha = alpha.point(lambda a: int(a * SH_ALPHA))
    black = Image.new("RGBA", phone.size, (20, 18, 50, 255))
    black.putalpha(sh_alpha)
    shadow.paste(black, (PAD, PAD + SH_OFFY), black)
    shadow = shadow.filter(ImageFilter.GaussianBlur(SH_BLUR))
    canvas = Image.alpha_composite(canvas, shadow)
    canvas.paste(phone, (PAD, PAD), phone)
    return canvas


for raw, name in MAP.items():
    p = os.path.join(SRC, raw)
    if not os.path.exists(p):
        print("MISSING", raw)
        continue
    phone = cutout(p)
    out = with_shadow(phone)
    out.save(os.path.join(OUT, f"phone_{name}.png"))
    print(f"{raw:14s} -> phone_{name}.png  {out.size[0]}x{out.size[1]}")

# QA composite: phones on navy + white to check for halos
qa = Image.new("RGBA", (1400, 700), (30, 27, 74, 255))
ImageDraw.Draw(qa).rectangle([700, 0, 1400, 700], fill=(255, 255, 255, 255))
for i, name in enumerate(["feed", "scenario", "calc"]):
    ph = Image.open(os.path.join(OUT, f"phone_{name}.png"))
    ph = ph.resize((int(ph.size[0] * 620 / ph.size[1]), 620))
    qa.alpha_composite(ph, (60 + i * 220, 40))
    qa.alpha_composite(ph, (760 + i * 220, 40))
qa.convert("RGB").save(os.path.join(OUT, "_qa_cutout.png"))
print("QA -> assets/_qa_cutout.png")
