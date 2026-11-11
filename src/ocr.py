"""
Offline OCR using the text recognition built into Windows 10/11
(Windows.Media.Ocr). Nothing is uploaded; recognition runs on this PC with
the OCR languages Windows already has installed (e.g. English, Japanese).

    available() -> bool
    ocr_image(pil_image) -> (text, tables)

Pipeline
  1. clean up: grayscale, dark backgrounds inverted, contrast stretched
  2. pass 1 measures the text height; the image is rescaled so lines are
     ~36 px tall (the size Windows OCR reads best) and read again
  3. tilt is corrected with the angle Windows reports; upside-down or
     sideways pages are detected and turned
  4. ruled tables are rebuilt from the lines drawn in the image, merged
     cells included; text without a table is split into cells by gaps
  5. typical OCR slips inside numbers are corrected ($l25 -> $125, NIOO -> N100)
"""
from __future__ import annotations

import asyncio
import io
import re

try:
    from winrt.windows.globalization import Language
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream
    _OK = True
except Exception:          # not Windows, or winrt packages missing
    _OK = False

try:
    import numpy as np
except Exception:
    np = None

from PIL import Image, ImageOps

_CJK = re.compile(r"[぀-ヿ㐀-鿿]")
TARGET_TEXT_HEIGHT = 36      # px; Windows OCR is most accurate around 30-40 px
WORK_LIMIT = 6000            # px, longest side we hand to the engine


def available() -> bool:
    return _OK and bool(OcrEngine.available_recognizer_languages)


def _languages() -> list[str]:
    tags = [l.language_tag for l in OcrEngine.available_recognizer_languages]
    en = [t for t in tags if t.lower().startswith("en")]
    ja = [t for t in tags if t.lower().startswith("ja")]
    return en[:1] + ja[:1] or tags[:1]


# ---------------------------------------------------------------- engine call
async def _recognize_async(png: bytes, tag: str):
    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(png)
    await writer.store_async()
    writer.detach_stream()
    stream.seek(0)
    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()
    engine = OcrEngine.try_create_from_language(Language(tag))
    result = await engine.recognize_async(bitmap)
    words = []
    for line in result.lines:
        for w in line.words:
            r = w.bounding_rect
            words.append((w.text, r.x, r.y, r.width, r.height))
    angle = result.text_angle
    return words, (angle.value if angle is not None and hasattr(angle, "value") else angle)


def _recognize(img: Image.Image, tag: str):
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="PNG", compress_level=0)   # in-memory only: speed over size
    return asyncio.run(_recognize_async(buf.getvalue(), tag))


# ---------------------------------------------------------------- helpers
def _clean(img: Image.Image, ink: bool = True) -> Image.Image:
    img = ImageOps.exif_transpose(img)
    if np is not None and img.mode in ("RGB", "RGBA", "P", "CMYK"):
        lum = img.convert("L")
    else:
        lum = img.convert("L")
    small = lum.resize((max(1, lum.width // 8), max(1, lum.height // 8)))
    hist = small.histogram()
    mean = sum(i * n for i, n in enumerate(hist)) / max(1, sum(hist))
    g = lum
    if ink and np is not None and img.mode in ("RGB", "RGBA", "P", "CMYK"):
        rgb = np.asarray(img.convert("RGB"))
        # coloured ink on light pages: darkest channel keeps it dark
        if mean >= 110:
            g = Image.fromarray(rgb.min(axis=2))
    if mean < 110:                       # light text on a dark background
        g = ImageOps.invert(g)
    return ImageOps.autocontrast(g, cutoff=1)


def _fit(img: Image.Image, scale: float) -> Image.Image:
    limit = min(WORK_LIMIT, OcrEngine.max_image_dimension)
    scale = min(scale, limit / max(img.size))
    if abs(scale - 1) < 0.05:
        return img
    return img.resize((max(1, int(img.width * scale)), max(1, int(img.height * scale))), Image.LANCZOS)


def _median_height(words) -> float:
    hs = sorted(w[4] for w in words if len(w[0]) > 1)
    return hs[len(hs) // 2] if hs else 0


_VOWEL = re.compile(r"[aeiouyAEIOUY]")


def _quality(words) -> float:
    """How much of the result looks like real text (for picking rotations)."""
    good = 0
    for w in words:
        t = w[0]
        t = t.strip(".,:;()[]\"'")
        if (re.fullmatch(r"[A-Za-z]{2,}", t) and _VOWEL.search(t)) or \
           re.fullmatch(r"(US)?[$¥]?[0-9][0-9,./:%-]*[A-Za-z]{0,3}", t) or \
           re.fullmatch(r"(?=.*\d)(?=.*[A-Z])[A-Z0-9][A-Za-z0-9./-]{2,}", t) or _CJK.search(t):
            good += len(t)
    return good


_KNOWN = set("""price usd us cpu lcd model total the and of for sku intel amd core ultra memory battery
keyboard weight color colour design board date quotation quote unit qty item remark remarks spec ssd ram
dram ddr4 ddr5 lpddr5 webcam camera wlan adapter adaptor power cord panel display inch cover aluminum
plastic metal moq barebone gpu vram rtx black silver grey gray hz storage type option options with
company ltd inc co technology limited payment term terms fob shipment warranty description amount
""".split())


def _known(words) -> int:
    """Real dictionary words read: upside-down text gives look-alike junk."""
    return sum(1 for w in words if w[0].strip(".,:;()[]'\"").lower() in _KNOWN)


def _join(parts: list[str]) -> str:
    """Windows OCR returns each Japanese character as a word: no space
    between two CJK characters, a space everywhere else."""
    out = ""
    for p in parts:
        if out and not (_CJK.search(out[-1]) and _CJK.search(p[:1])):
            out += " "
        out += p
    return out


_DIGITISH = str.maketrans({"O": "0", "o": "0", "D": "0", "I": "1", "l": "1", "|": "1", "S": "5", "B": "8", "Z": "2"})


def _fix_text(s: str) -> str:
    """Correct OCR slips where the context says 'this is a number'."""
    s = re.sub(r"(\d)\s+([,.])\s*(\d)", r"\1\2\3", s)                       # 1 ,000 -> 1,000
    s = re.sub(r"\bUS\s*[S5]\s*(?=[0-9OIl])", "US$", s)                     # USS 63 -> US$63
    # money: "$l25.OO" -> "$125.00"
    s = re.sub(r"([$¥]\s*)([0-9OoIlSB][0-9OoIlSB,.]*)",
               lambda m: m.group(1) + m.group(2).translate(_DIGITISH)
               if re.search(r"[0-9]", m.group(2)) or len(m.group(2)) >= 2 else m.group(0), s)
    # Intel N-series / Core numbers: "NIOO" -> "N100", "N2OO" -> "N200"
    s = re.sub(r"\bN([0-9OIlC]{3})\b", lambda m: "N" + m.group(1).translate(_DIGITISH).replace("C", "0")
               if sum(c.isdigit() for c in m.group(1)) >= 1 or re.search(r"lake|intel|cpu|celeron", s, re.I)
               else m.group(0), s)
    # resolutions: "192OX1O8O" -> "1920X1080"
    s = re.sub(r"\b([0-9OIl]{3,4})\s*([xX×*])\s*([0-9OIl]{3,4})\b",
               lambda m: m.group(1).translate(_DIGITISH) + m.group(2) + m.group(3).translate(_DIGITISH)
               if sum(c.isdigit() for c in m.group(1) + m.group(3)) >= 4 else m.group(0), s)
    # thousands group read as letters: "1 ,OOO" -> "1,000"
    s = re.sub(r"(\d)\s*,\s*([0-9OoIl]{3})\b(?!\w)",
               lambda m: m.group(1) + "," + m.group(2).translate(_DIGITISH)
               if any(c.isdigit() or c in "Oo" for c in m.group(2)) else m.group(0), s)
    s = re.sub(r"(?<![A-Za-z])[I1l]\.?[JU]S\s*\$", "US$", s)               # I.JS$ -> US$
    # CPU model misreads: "RS 7640HS" -> "R5 7640HS", "2051-1" -> "205H"
    s = re.sub(r"\bR[S5]\s+([0-9]{4}[A-Z]{0,3})\b", r"R5 \1", s)
    s = re.sub(r"\b([0-9]{3,4})[1Il]-[1Il](?=[\s,)]|$)", r"\1H", s)
    s = re.sub(r"\b([0-9]{3,4})[1Il]-[1Il]([SX])\b", r"\1H\2", s)
    # Core™ read as "Coret", "CoretM", "Coret\"", "Coretta" before Ultra
    s = re.sub(r"\bCore\s*t\S{0,3}\s*(?=Ultra)", "Core ", s)
    s = re.sub(r"\bUltra\s*([3579])(?=\s)", r"Ultra \1", s)
    # misread separator between known panel widths/heights: "1920<1080" -> "1920x1080"
    s = re.sub(r"\b(1280|1366|1600|1920|2240|2560|2880|3200|3840)\s*[<>(){}:;~′'\"]\s*"
               r"(720|768|900|1080|1200|1400|1440|1600|1800|2000|2160|2400)\b", r"\1x\2", s)
    return s


# ---------------------------------------------------------------- clean-up
BLOCK = 24


def _mask_photos(img: Image.Image) -> Image.Image:
    """White out photos/pictures: they produce garbage words and their edges
    look like table lines. Text is black/white after clean-up, photos are
    full of light-grey shading."""
    if np is None:
        return img
    a = np.asarray(img).copy()
    H, W = a.shape
    by, bx = H // BLOCK, W // BLOCK
    if not by or not bx:
        return img
    blocks = a[:by * BLOCK, :bx * BLOCK].reshape(by, BLOCK, bx, BLOCK).swapaxes(1, 2).astype(np.float32)
    in_grey = (blocks > 150) & (blocks < 235)
    grey = in_grey.mean(axis=(2, 3))
    dark = (blocks < 140).mean(axis=(2, 3))
    # shading spread among the grey pixels: photos vary, flat cell fills don't
    gsum = np.where(in_grey, blocks, 0).sum(axis=(2, 3))
    gcnt = np.maximum(in_grey.sum(axis=(2, 3)), 1)
    gmean = gsum / gcnt
    gvar = np.where(in_grey, (blocks - gmean[..., None, None]) ** 2, 0).sum(axis=(2, 3)) / gcnt
    photo = (grey > 0.30) & (np.sqrt(gvar) > 9)
    # grow into neighbouring solid areas (dark screens, logos inside photos)
    for _ in range(4):
        grown = photo.copy()
        grown[1:] |= photo[:-1]; grown[:-1] |= photo[1:]
        grown[:, 1:] |= photo[:, :-1]; grown[:, :-1] |= photo[:, 1:]
        photo = photo | (grown & ((dark > 0.6) | ((grey > 0.25) & (np.sqrt(gvar) > 9))))
    # keep only large connected regions: a photo, not a few shaded cells
    photo = _large_regions(photo, min_blocks=12, keep=lambda comp: _is_photo(a, comp))
    if not photo.any():
        return img
    for y, x in zip(*np.nonzero(photo)):
        a[y * BLOCK:(y + 1) * BLOCK, x * BLOCK:(x + 1) * BLOCK] = 255
    return Image.fromarray(a)


def _is_photo(a, comp) -> bool:
    """A photo's mid-tones spread over many grey levels; a coloured cell fill
    (even with JPEG noise) sits on one or two."""
    ys = [y for y, _ in comp]; xs = [x for _, x in comp]
    box = a[min(ys) * BLOCK:(max(ys) + 1) * BLOCK, min(xs) * BLOCK:(max(xs) + 1) * BLOCK]
    if (box >= 235).mean() > 0.5:          # mostly paper: shaded cells / coloured text, not a picture
        return False
    vals = box[(box > 150) & (box < 235)]
    if vals.size < 500:
        return False
    h = np.histogram(vals, bins=17, range=(150, 235))[0]
    return h.max() / h.sum() < 0.22


def _large_regions(mask, min_blocks: int, keep=None):
    out = np.zeros_like(mask)
    seen = np.zeros_like(mask)
    H, W = mask.shape
    for y0, x0 in zip(*np.nonzero(mask)):
        if seen[y0, x0]:
            continue
        stack, comp = [(y0, x0)], []
        seen[y0, x0] = True
        while stack:
            y, x = stack.pop()
            comp.append((y, x))
            for ny, nx in ((y + 1, x), (y - 1, x), (y, x + 1), (y, x - 1)):
                if 0 <= ny < H and 0 <= nx < W and mask[ny, nx] and not seen[ny, nx]:
                    seen[ny, nx] = True
                    stack.append((ny, nx))
        if len(comp) >= min_blocks and (keep is None or keep(comp)):
            for y, x in comp:
                out[y, x] = True
    return out


def _erase_lines(img: Image.Image, hs, vs) -> Image.Image:
    """Remove table ruling lines so letters touching them are not clipped."""
    if not hs and not vs:
        return img
    from PIL import ImageDraw
    out = img.copy()
    d = ImageDraw.Draw(out)
    for y, s, e in hs:
        d.rectangle([s, y - 2, e, y + 2], fill=255)
    for x, s, e in vs:
        d.rectangle([x - 2, s, x + 2, e], fill=255)
    return out


# ---------------------------------------------------------------- recognition
def _read_best(ocr_img: Image.Image, layout_img: Image.Image, tag: str):
    """Pass 1 at a moderate size -> measure text -> rescale -> deskew -> rotate.
    Every resize/turn is applied to both images so word positions match the
    layout image (which still has its table lines)."""
    def scaled(scale):
        a = _fit(ocr_img, scale)
        f = a.width / ocr_img.width
        b = layout_img if a is ocr_img else layout_img.resize(a.size, Image.LANCZOS)
        return a, b, f
    work, layout, f = scaled(1.0)
    words, angle = _recognize(work, tag)
    h = _median_height(words)
    if h:
        w2, l2, f2 = scaled(f * TARGET_TEXT_HEIGHT / h)
        if abs(w2.width - work.width) / work.width > 0.2:
            work, layout = w2, l2
            words, angle = _recognize(work, tag)
    if angle and 0.4 < abs(angle) < 20:
        cand = work.rotate(angle, expand=True, fillcolor=255, resample=Image.BICUBIC)
        w2, _ = _recognize(cand, tag)
        if _quality(w2) >= _quality(words):
            work, words = cand, w2
            layout = layout.rotate(angle, expand=True, fillcolor=255, resample=Image.BICUBIC)
    q0 = _quality(words)
    letters = sum(len(w[0]) for w in words) or 1
    k0 = _known(words)
    cjk = sum(1 for w in words if _CJK.search(w[0]))
    if q0 / letters < 0.5 or len(words) < 8 or (k0 < 3 and cjk < 10):   # looks wrong: try turning
        for rot in (180, 90, 270):
            cand = work.rotate(rot, expand=True, fillcolor=255)
            w2, _ = _recognize(cand, tag)
            k2, q2 = _known(w2), _quality(w2)
            if (k2 >= k0 + 3 and q2 > q0 * 0.8) or (k2 >= k0 and q2 > max(q0 * 1.3, 20)):
                work, words, q0, k0 = cand, w2, q2, k2
                layout = layout.rotate(rot, expand=True, fillcolor=255)
                if k0 >= 5:
                    break                                # clearly upright now
    return layout, words, work


def _pick(results: dict[str, tuple]):
    """English unless the page is mostly Japanese. Words one engine skipped
    are taken from the other where the main result has nothing."""
    en = next((r for t, r in results.items() if not t.lower().startswith("ja")), None)
    ja = next((r for t, r in results.items() if t.lower().startswith("ja")), None)
    if ja is None:
        return en
    if en is None:
        return ja
    ja_text = "".join(w[0] for w in ja[1])
    cjk = len(_CJK.findall(ja_text))
    en_words = sum(1 for w in en[1] if re.fullmatch(r"[A-Za-z]{3,}", w[0]))
    main, other = (ja, en) if (cjk >= 10 and cjk / max(1, len(ja_text)) > 0.3 and cjk > en_words * 2) else (en, ja)
    def overlaps(a, b) -> bool:
        ix = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
        iy = min(a[2] + a[4], b[2] + b[4]) - max(a[2], b[2])
        return ix > 0 and iy > 0 and ix * iy > 0.2 * min(a[3] * a[4], b[3] * b[4])
    extra = [w for w in other[1]
             if (main is ja or not _CJK.search(w[0])) and not any(overlaps(w, m) for m in main[1])]
    return main[0], main[1] + extra


# ---------------------------------------------------------------- layout
def _ruling_lines(img: Image.Image):
    """Horizontal and vertical table lines as (pos, start, end) segments."""
    if np is None:
        return [], []
    g = np.asarray(img.convert("L"))
    a = g < 140
    H, W = a.shape

    def segments(mask2d, min_len):
        out = []
        for i, row in enumerate(mask2d):
            if not row.any():
                continue
            padded = np.concatenate(([0], row.astype(np.int8), [0]))
            d = np.diff(padded)
            starts, ends = np.where(d == 1)[0], np.where(d == -1)[0]
            for s, e in zip(starts, ends):
                if e - s >= min_len:
                    out.append((i, s, e))
        return out

    def merge(segs, gap):
        segs.sort()
        merged: list[list[int]] = []
        for p, s, e in segs:
            for m in merged:
                if abs(m[0] - p) <= gap and s <= m[2] + gap and e >= m[1] - gap:
                    m[1], m[2] = min(m[1], s), max(m[2], e)
                    break
            else:
                merged.append([p, s, e])
        return [tuple(m) for m in merged]

    hs = merge(segments(a, max(60, W // 12)), 3)
    vs = merge(segments(a.T, max(40, H // 40)), 3)
    return hs, vs


def _table_from_lines(words, hs, vs):
    """Assign words to the cells formed by ruling lines (merged cells kept)."""
    if len(hs) < 2 or len(vs) < 2:
        return None, words
    ys = sorted({h[0] for h in hs})
    xs = sorted({v[0] for v in vs})

    def collapse(vals, tol=6):
        out = []
        for v in vals:
            if not out or v - out[-1] > tol:
                out.append(v)
        return out
    ys, xs = collapse(ys), collapse(xs)
    if len(ys) < 2 or len(xs) < 2:
        return None, words
    top, bottom, left, right = ys[0], ys[-1], xs[0], xs[-1]

    def h_covers(y, x):      # is there a horizontal line at y passing x?
        return any(abs(p - y) <= 6 and s - 6 <= x <= e + 6 for p, s, e in hs)

    def v_covers(x, y):
        return any(abs(p - x) <= 6 and s - 6 <= y <= e + 6 for p, s, e in vs)

    nr, nc = len(ys) - 1, len(xs) - 1
    parent = list(range(nr * nc))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i
    for r in range(nr):
        ym = (ys[r] + ys[r + 1]) / 2
        for c in range(nc):
            xm = (xs[c] + xs[c + 1]) / 2
            if c + 1 < nc and not v_covers(xs[c + 1], ym):
                parent[find(r * nc + c + 1)] = find(r * nc + c)
            if r + 1 < nr and not h_covers(ys[r + 1], xm):
                parent[find((r + 1) * nc + c)] = find(r * nc + c)

    cell_words: dict[int, list] = {}
    outside = []
    for w in words:
        cx, cy = w[1] + w[3] / 2, w[2] + w[4] / 2
        if not (left <= cx <= right and top <= cy <= bottom):
            outside.append(w)
            continue
        c = max(i for i in range(nc) if xs[i] <= cx)
        r = max(i for i in range(nr) if ys[i] <= cy)
        cell_words.setdefault(find(min(r, nr - 1) * nc + min(c, nc - 1)), []).append(w)

    grid = [[""] * nc for _ in range(nr)]
    for root, ws in cell_words.items():
        r, c = divmod(root, nc)            # text goes to the merged cell's top-left
        # wrapped lines in one cell join with a space, like text-PDF cells
        grid[r][c] = _join([_join([w[0] for w in line]) for line in _lines(ws)])
    grid = [row for row in grid if any(row)]
    # drop columns that are empty everywhere (thin gaps between double lines)
    keep = [j for j in range(nc) if any(row[j] for row in grid)]
    grid = [[row[j] for j in keep] for row in grid]
    if len(grid) < 2 or len(keep) < 2:
        return None, words
    return grid, outside


def _lines(words):
    """Group words into text lines (top to bottom, left to right)."""
    if not words:
        return []
    h = _median_height(words) or 10
    rows: list[list] = []
    for w in sorted(words, key=lambda w: (w[2] + w[4] / 2, w[1])):
        cy = w[2] + w[4] / 2
        if rows and abs(rows[-1][-1][2] + rows[-1][-1][4] / 2 - cy) < h * 0.55:
            rows[-1].append(w)
        else:
            rows.append([w])
    return [sorted(r, key=lambda w: w[1]) for r in rows]


def _gap_grid(words):
    """Text without ruling lines: a wide gap starts a new cell; cells that
    share a left edge share a column."""
    rows = _lines(words)
    if not rows:
        return []
    h = _median_height(words) or 10
    cell_rows = []
    for row in rows:
        cells, cur, right = [], [], None
        for w in row:
            if cur and w[1] - right > h * 1.6:
                cells.append((cur[0][1], _join([x[0] for x in cur])))
                cur = []
            cur.append(w)
            right = w[1] + w[3]
        cells.append((cur[0][1], _join([x[0] for x in cur])))
        cell_rows.append(cells)
    anchors: list[float] = []
    for x in sorted(x for cells in cell_rows for x, _ in cells):
        if not anchors or x - anchors[-1] > h * 2:
            anchors.append(x)
    grid = []
    for cells in cell_rows:
        line = [""] * len(anchors)
        for x, t in cells:
            col = max(i for i, a in enumerate(anchors) if a <= x + h)
            line[col] = (line[col] + " " + t).strip()
        grid.append(line)
    return grid


# ---------------------------------------------------------------- public API
_GARBAGE = re.compile(r"^[Ⅰ-ⅿ|!¦︱丨]+$")      # roman-numeral glyphs etc. from hatching


def _despeckle(img: Image.Image) -> Image.Image:
    """Remove salt-and-pepper specks (fax/old scanner noise) only when the page
    has them: a median filter would thin fine strokes on a clean page."""
    if np is None:
        return img
    from PIL import ImageFilter
    med = img.filter(ImageFilter.MedianFilter(3))
    specks = (np.abs(np.asarray(img, dtype=np.int16) - np.asarray(med, dtype=np.int16)) > 150).mean()
    return med if specks > 0.004 else img


def _skew_angle(img: Image.Image) -> float:
    """Page tilt in degrees from horizontal projection sharpness (text rows
    and table lines give the sharpest profile when level)."""
    if np is None:
        return 0.0
    small = img.copy()
    small.thumbnail((1000, 1000))
    inv = ImageOps.invert(small)
    def score(a):
        prof = np.asarray(inv.rotate(a, resample=Image.BILINEAR, fillcolor=0), dtype=np.float32).sum(axis=1)
        return float(np.square(np.diff(prof)).sum())
    best, best_s = 0.0, score(0.0)
    for step, span in ((0.5, 6.0), (0.1, 0.5)):
        centre = best
        for a in np.arange(centre - span, centre + span + 1e-6, step):
            s = score(float(a))
            if s > best_s * 1.02 or (s > best_s and abs(a) > 1e-6 and best != 0.0):
                best, best_s = float(a), s
    return best if abs(best) >= 0.3 else 0.0


def _all_words(img: Image.Image, ink: bool = True) -> dict[str, tuple]:
    clean = _despeckle(_clean(img, ink))
    angle = _skew_angle(clean)
    if angle:
        clean = clean.rotate(angle, expand=True, fillcolor=255, resample=Image.BICUBIC)
    clean = _mask_photos(clean)
    hs, vs = _ruling_lines(clean)
    ocr_img = _erase_lines(clean, hs, vs)
    tags = _languages()
    # the first language finds the best size/turn; the others read that same image once
    layout, words, work = _read_best(ocr_img, clean, tags[0])
    out = {tags[0]: (layout, words)}
    for tag in tags[1:]:
        out[tag] = (layout, _recognize(work, tag)[0])
    return {t: (l, [w for w in ws if not _GARBAGE.match(w[0])]) for t, (l, ws) in out.items()}


def all_text(img) -> str:
    """Text from every installed OCR language (used to spot logos)."""
    if not available():
        return ""
    texts = [_join([w[0] for w in words]) for _, words in _all_words(img, ink=False).values()]
    # logos are big text: also read once at the original size, without rescaling
    plain = _fit(_clean(img, ink=False), 1.0)
    texts.append(_join([w[0] for w in _recognize(plain, _languages()[0])[0]]))
    return "\n".join(texts)


def ocr_image(img) -> tuple[str, list[list[list[str]]]]:
    if not available():
        raise RuntimeError("Windows OCR is not available on this PC")
    work, words = _pick(_all_words(img))
    words = [(_fix_text(w[0]),) + tuple(w[1:]) for w in words]
    hs, vs = _ruling_lines(work)
    table, rest = _table_from_lines(words, hs, vs)
    tables = []
    if table:
        tables.append([[_fix_text(c) for c in row] for row in table])
    loose = _gap_grid(rest if table else words)
    if loose and not table:
        tables.append([[_fix_text(c) for c in row] for row in loose])
    # plain text in reading order (whole page) for prose-style parsing
    text = "\n".join(_fix_text(" | ".join(c for c in row if c)) for row in _gap_grid(words))
    return text, tables
