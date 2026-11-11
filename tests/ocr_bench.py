"""
OCR accuracy benchmark.

Each real quotation PDF (from QUOTEGRID_SAMPLES, default ~/Downloads) is
turned into "scanned" images under harder conditions, read back through
OCR, and the extracted fields are compared with the text-based result.
Nothing is written outside a temp folder.

    python tests/ocr_bench.py
"""
from __future__ import annotations

import io
import os
import random
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import quote_extract as q  # noqa: E402
import pypdfium2 as pdfium  # noqa: E402
from PIL import Image, ImageFilter, ImageOps  # noqa: E402

SAMPLES = Path(os.environ.get("QUOTEGRID_SAMPLES", Path.home() / "Downloads"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
try:                     # local, git-ignored list of real quotation files
    from real_samples import OCR_BENCH_FILES as FILES  # noqa: E402
except ImportError:
    FILES = []
SKIP_COLS = {"その他"}       # free text: not comparable cell for cell

random.seed(7)


def _scan(img: Image.Image, cond: str) -> Image.Image:
    img = img.convert("RGB")
    if cond == "150dpi":
        return img
    if cond == "100dpi":
        return img.resize((int(img.width * 2 / 3), int(img.height * 2 / 3)))
    if cond == "tilt2":
        return img.rotate(2.0, expand=True, fillcolor="white", resample=Image.BICUBIC)
    if cond == "jpeg":
        buf = io.BytesIO(); img.save(buf, "JPEG", quality=25); return Image.open(buf).convert("RGB")
    if cond == "dark":
        return ImageOps.invert(img)
    if cond == "noisy":
        g = img.convert("L").filter(ImageFilter.GaussianBlur(0.8))
        px = g.load()
        for _ in range(g.width * g.height // 40):
            x, y = random.randrange(g.width), random.randrange(g.height)
            px[x, y] = random.choice((0, 255))
        return g.convert("RGB")
    if cond == "upside":
        return img.rotate(180)
    raise ValueError(cond)


CONDS = ["150dpi", "100dpi", "tilt2", "jpeg", "dark", "noisy", "upside"]


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="quotegrid_bench_"))
    totals: dict[str, list[int]] = {c: [0, 0] for c in CONDS}
    misses: dict[str, list[str]] = {c: [] for c in CONDS}
    t0 = time.time()
    for name in FILES:
        pdf_path = SAMPLES / name
        if not pdf_path.is_file():
            continue
        truth = q.extract(pdf_path)
        doc = pdfium.PdfDocument(str(pdf_path))
        page = doc[0].render(scale=150 / 72).to_pil()      # quotations here are 1 page
        doc.close()
        for cond in CONDS:
            p = tmp / f"{abs(hash(name)) % 10**6}_{cond}.png"
            _scan(page, cond).save(p)
            try:
                got = q.extract(p)
            except Exception as e:
                got = []
                misses[cond].append(f"{name}: {e}")
            for i, t in enumerate(truth):
                g = got[i] if i < len(got) else {}
                for col, v in t.items():
                    if col in SKIP_COLS or not v:
                        continue
                    totals[cond][1] += 1
                    if g.get(col, "") == v:
                        totals[cond][0] += 1
                    elif len(misses[cond]) < 6:
                        misses[cond].append(f"{name[:18]} r{i} {col}: want {v!r} got {g.get(col, '')!r}")
    all_ok = sum(v[0] for v in totals.values()); all_n = sum(v[1] for v in totals.values())
    for c in CONDS:
        ok, n = totals[c]
        print(f"{c:8s} {ok:3d}/{n:<3d} {100 * ok / max(1, n):5.1f}%")
        for m in misses[c]:
            print("           -", m)
    print(f"TOTAL    {all_ok}/{all_n} {100 * all_ok / max(1, all_n):.1f}%   ({time.time() - t0:.0f}s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
