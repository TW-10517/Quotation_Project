"""Render the QuoteGrid logo into ../assets: logo.png (1024px), logo_128.png, logo.ico."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter

OUT = Path(__file__).resolve().parent.parent / "assets"
S = 1024 * 4  # supersample, then downscale for clean edges

INDIGO = (124, 92, 255)
TEAL = (63, 208, 201)
DEEP = (58, 34, 170)


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def rounded_mask(size, box, r):
    m = Image.new("L", size, 0)
    ImageDraw.Draw(m).rounded_rectangle(box, radius=r, fill=255)
    return m


img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

# Background tile: diagonal indigo -> teal gradient
grad = Image.new("RGB", (S, S))
gp = grad.load()
for y in range(S):
    for x in range(0, S, 4):
        c = lerp(INDIGO, TEAL, (x + y) / (2 * S))
        for k in range(4):
            gp[x + k, y] = c
pad = int(S * 0.04)
img.paste(grad, (0, 0), rounded_mask((S, S), (pad, pad, S - pad, S - pad), int(S * 0.22)))

# Soft top highlight
hl = Image.new("RGBA", (S, S), (255, 255, 255, 0))
ImageDraw.Draw(hl).ellipse((-S * 0.3, -S * 0.75, S * 1.3, S * 0.45), fill=(255, 255, 255, 34))
hl.putalpha(Image.composite(hl.getchannel("A"), Image.new("L", (S, S), 0),
                            rounded_mask((S, S), (pad, pad, S - pad, S - pad), int(S * 0.22))))
img = Image.alpha_composite(img, hl)

# Document sheet with shadow
dx0, dy0, dx1, dy1 = int(S * 0.24), int(S * 0.17), int(S * 0.70), int(S * 0.80)
fold = int(S * 0.13)
sheet = [(dx0, dy0), (dx1 - fold, dy0), (dx1, dy0 + fold), (dx1, dy1), (dx0, dy1)]

shadow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
ImageDraw.Draw(shadow).polygon([(x + S * 0.02, y + S * 0.03) for x, y in sheet], fill=(20, 10, 60, 110))
img = Image.alpha_composite(img, shadow.filter(ImageFilter.GaussianBlur(S * 0.025)))

d = ImageDraw.Draw(img)
d.polygon(sheet, fill=(255, 255, 255, 255))
d.polygon([(dx1 - fold, dy0), (dx1 - fold, dy0 + fold), (dx1, dy0 + fold)], fill=(214, 206, 255, 255))

# Table grid on the sheet: header bar + rows + one column divider
gx0, gx1 = dx0 + int(S * 0.06), dx1 - int(S * 0.06)
gy = dy0 + int(S * 0.20)
row_h = int(S * 0.085)
lw = int(S * 0.016)
d.rounded_rectangle((gx0, gy, gx1, gy + row_h), radius=int(S * 0.02), fill=INDIGO)
for i in range(1, 4):
    y = gy + row_h * (i + 1)
    d.rounded_rectangle((gx0, y - lw // 2, gx1, y + lw // 2), radius=lw // 2, fill=(200, 192, 245))
colx = gx0 + int((gx1 - gx0) * 0.42)
d.rounded_rectangle((colx - lw // 2, gy + row_h, colx + lw // 2, gy + row_h * 4), radius=lw // 2, fill=(200, 192, 245))
# Highlighted "extracted" cell
d.rounded_rectangle((colx + lw * 2, gy + row_h * 2 + lw * 2, gx1 - lw, gy + row_h * 3 - lw * 2),
                    radius=int(S * 0.012), fill=TEAL)

# Extract badge: circle with downward arrow, bottom-right
cx, cy, r = int(S * 0.73), int(S * 0.75), int(S * 0.125)
badge_sh = Image.new("RGBA", (S, S), (0, 0, 0, 0))
ImageDraw.Draw(badge_sh).ellipse((cx - r + S * 0.01, cy - r + S * 0.02, cx + r + S * 0.01, cy + r + S * 0.02),
                                 fill=(20, 10, 60, 120))
img = Image.alpha_composite(img, badge_sh.filter(ImageFilter.GaussianBlur(S * 0.02)))
d = ImageDraw.Draw(img)
d.ellipse((cx - r - lw * 2, cy - r - lw * 2, cx + r + lw * 2, cy + r + lw * 2), fill=(255, 255, 255, 255))
d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=DEEP)
aw = int(r * 0.20)
d.rounded_rectangle((cx - aw // 2, cy - int(r * 0.55), cx + aw // 2, cy + int(r * 0.15)), radius=aw // 2, fill="white")
ah = int(r * 0.48)
d.polygon([(cx - ah, cy - int(r * 0.02)), (cx + ah, cy - int(r * 0.02)), (cx, cy + int(r * 0.52))], fill="white")

final = img.resize((1024, 1024), Image.LANCZOS)
final.save(OUT / "logo.png")
final.resize((128, 128), Image.LANCZOS).save(OUT / "logo_128.png")
final.save(OUT / "logo.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("wrote", OUT / "logo.png", OUT / "logo.ico")
