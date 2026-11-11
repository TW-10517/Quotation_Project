"""
quote_extract.py  -  QuoteGrid's ODM-quotation field extractor.

Returns list[dict[str, str]] — one dict per product variant. Most quotations
produce a single row; a quote that prices multiple CPU options side-by-side
(not as "opt X" in a remark) produces one row per option.

Design
------
Most ODM quotations we see are line-item style: a table with an Item column
(Barebone / CPU / DDR / Memory / SSD / Keyboard / RJ45 / LCD), a free-prose
Spec/Description column, a Price column, and often a Remark column that
mentions an alternate option. The extractor:

  1. Walks every table, picks out "line-item rows" by matching the Item cell
     against ITEM_ALIASES.
  2. For each line item, maps its Price to the right output column (Barebone
     price, CPU price, LPDDR5 8/12/16GB price by GB in the spec, ...).
  3. Parses prose out of the Barebone spec cell: material, weight, LCD
     resolution, RJ45 presence, battery/camera/WiFi go to その他.
  4. Parses the CPU spec+remark to produce "CPU Platform", "CPU Name" and,
     if the remark has an alt CPU with a price, a "(opt NAME)" suffix.
  5. Computes B/B+CPU from Barebone + CPU if both are numeric.
  6. Finds Barebone Name (model code) and ODM supplier from the text.
"""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path

import pdfplumber
try: import docx
except ImportError: docx = None
try: import openpyxl
except ImportError: openpyxl = None
try: import ocr            # offline Windows OCR, for images and scanned pages
except Exception: ocr = None


# ---------------------------------------------------------------------------
# Column order
# ---------------------------------------------------------------------------
COLUMNS = [
    "ODM supplier", "Barebone Name", "CPU Platform", "CPU Name", "MOQ",
    "Barebone", "CPU", "B/B+CPU",
    "LPDDR5 8GB", "LPDDR5 12GB", "LPDDR5 16GB",
    "JP Keyboad", "RJ45", "A", "C", "D",
    "weight", "LCD", "その他",
]

# Known ODM suppliers (extend freely)
KNOWN_SUPPLIERS = [
    "SIXUNITED", "CLEVO", "Tongfang", "Compal", "Quanta", "Wistron",
    "Pegatron", "Inventec", "Mitac", "ECS", "ASUS", "MSI", "Gigabyte",
    "Mechrevo", "Hasee", "Uniwill", "Haier", "TongFang",
    "Emdoor", "Huaqin", "Longcheer", "Wingtech", "LCFC", "Bitland", "Chuwi",
    "Foxconn", "Topstar", "Thundersoft", "Getac", "Avita", "Positivo", "Arima",
]

# Panel names -> pixel resolution, used when a quotation gives only the name
# ("14\" FHD+" -> 1920*1200)
PANEL_RESOLUTIONS = [
    (r"UHD\+|WQUXGA",              "3840*2400"),
    (r"UHD|\b4K\b",                "3840*2160"),
    (r"\b3K\b",                    "2880*1800"),
    (r"\b2\.8K\b",                 "2880*1800"),
    (r"QHD\+|WQXGA",               "2560*1600"),
    (r"\b2\.5K\b",                 None),          # depends on 16:10 vs 16:9
    (r"\bW?QHD\b",                 "2560*1440"),
    (r"FHD\+|WUXGA",               "1920*1200"),
    (r"\bFHD\b|\bFull\s*HD\b|1080p", "1920*1080"),
    (r"\bHD\+",                    "1600*900"),
    (r"\bHD\b",                    "1366*768"),
]

# Item-cell keywords -> item kind. Matched with re.search so "Name Barebone"
# (from merged cells) still classifies as barebone.
ITEM_ALIASES: dict[str, list[str]] = {
    # checked first: "B/B+CPU" must not be read as just "CPU" or "Barebone"
    "bbcpu":     [r"b\s*/\s*b\s*\+\s*cpu", r"bare\s*bone\s*\+\s*cpu", r"bb\s*\+\s*cpu"],
    "barebone":  [r"\bbarebone\b", r"\bbare\s*bone\b", r"\bchassis\b", r"\bbase\s*model\b", r"ベアボーン"],
    "cpu":       [r"(?<!\+)(?<!\+ )\bcpu\b", r"\bprocessor\b", r"プロセッサ"],
    "memory":    [r"\b(?:lp)?ddr[0-9x]*\b", r"\bmemory\b", r"\bram\b", r"メモリ"],
    "ssd":       [r"\bssd\b", r"\bnvme\b", r"\bstorage\b", r"ストレージ"],
    "jp_kb":     [r"jp\s*keyboard", r"japan(ese)?\s*keyboard", r"jis\s*keyboard", r"日本語"],
    "rj45":      [r"\brj[\s-]?45\b", r"\blan\s*port\b"],
    "lcd":       [r"\blcd\b", r"\bpanel\b", r"\bdisplay\b", r"液晶"],
    "a_cover":   [r"\ba[\s-]*cover\b", r"Aカバー"],
    "c_cover":   [r"\bc[\s-]*cover\b", r"Cカバー"],
    "d_cover":   [r"\bd[\s-]*cover\b", r"Dカバー"],
    "weight":    [r"\bweight\b", r"重量"],
    "moq":       [r"\bmoq\b"],
}

# CPU platform names (ordered — specific first)
PLATFORMS = [
    # Short codes are kept as the quotation writes them ("WCL")
    (r"\bWCL\b",              "WCL"),
    (r"\bPTL\b",              "PTL"),
    (r"\bARL\b",              "ARL"),
    (r"\bLNL\b",              "LNL"),
    (r"\bMTL\b",              "MTL"),
    (r"\bRPL\b",              "RPL"),
    (r"\bADL[\s-]*N\b",       "ADL-N"),
    (r"\bTWL\b",              "TWL"),
    (r"Wildcat\s*Lake",       "Wildcat Lake"),
    (r"Alder\s*Lake[\s-]*N",  "Alder Lake-N"),
    (r"Twin\s*Lake",          "Twin Lake"),
    (r"Arrow\s*Lake[\s-]*H?", "Arrow Lake"),
    (r"Lunar\s*Lake",         "Lunar Lake"),
    (r"Meteor\s*Lake[\s-]*H?", "Meteor Lake"),
    (r"Raptor\s*Lake[\s-]*H?", "Raptor Lake"),
    (r"Alder\s*Lake[\s-]*H?", "Alder Lake"),
    (r"Panther\s*Lake",       "Panther Lake"),
    (r"Tiger\s*Lake",          "Tiger Lake"),
    (r"Strix\s*Point",         "Strix Point"),
    (r"Hawk\s*Point",          "Hawk Point"),
    (r"Phoenix",               "Phoenix"),
    (r"Rembrandt",             "Rembrandt"),
]

# Specific CPU names
CPU_RE = re.compile(
    r"\b(?:"
    r"N[0-9]{3,4}"                                  # N100, N150, N200
    r"|Core\s*(?:Ultra\s*)?[3579]\s*[0-9]{3,4}[A-Z]{0,3}"
    r"|i[3579]-[0-9]{4,5}[A-Z]{0,3}"
    r"|Ryzen\s*(?:AI\s*)?[3579]\s*[0-9]{3,4}[A-Z]{0,3}"
    r"|R[3579]\s*[0-9]{4}[A-Z]{0,3}"                 # R5 7640HS
    r")\b",
    re.IGNORECASE,
)

MONEY_NUM = r"[0-9]{1,3}(?:,[0-9]{3})+(?:\.[0-9]+)?|[0-9]+(?:\.[0-9]+)?"
MONEY_RE = re.compile(r"(?:US\$|USD|\$|JPY|JP¥|¥)\s*(" + MONEY_NUM + r")", re.IGNORECASE)

_WS = re.compile(r"\s+")


def _norm(s) -> str:
    if s is None: return ""
    return _WS.sub(" ", str(s)).strip()


# ---------------------------------------------------------------------------
# Readers
# ---------------------------------------------------------------------------
def _unspace(line: str) -> str:
    """Designer PDFs space out headings: 'N X 1 7 A B 1 -2' -> 'NX17AB1-2'."""
    toks = line.split()
    if len(toks) >= 4 and sum(len(t) <= 2 for t in toks) / len(toks) >= 0.75:
        return "".join(toks)
    return line


OCR_MIN_PAGE_CHARS = 30     # a page with less real text than this is treated as an image


def _render_page(path: Path, index: int, max_side: int = 4000):
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(str(path))
    try:
        page = doc[index]
        w, h = page.get_size()
        scale = min(3.0, max_side / max(w, h))
        return page.render(scale=scale).to_pil()
    finally:
        doc.close()


def _read_pdf(path: Path) -> tuple[str, list[list[list[str]]]]:
    parts, tables = [], []
    with pdfplumber.open(str(path)) as pdf:
        if len(pdf.pages) > MAX_PDF_PAGES:
            raise ValueError(f"PDF has {len(pdf.pages)} pages; quotations over {MAX_PDF_PAGES} pages are not read")
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if len(text.strip()) < OCR_MIN_PAGE_CHARS and page.images and ocr and ocr.available():
                # scanned / image-only page: read it with OCR
                o_text, o_tables = ocr.ocr_image(_render_page(path, i))
                if len(o_text.strip()) > len(text.strip()):
                    parts.append(o_text)
                    tables.extend(o_tables)
                    continue
            parts.append("\n".join(_unspace(l) for l in text.splitlines()))
            try:
                for t in page.extract_tables() or []:
                    tables.append([[(_norm(c) if c else "") for c in r] for r in t])
            except Exception:
                pass
    return "\n".join(parts), tables


IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")
MAX_IMAGE_PIXELS = 60_000_000


def _read_image(path: Path) -> tuple[str, list[list[list[str]]]]:
    if not (ocr and ocr.available()):
        raise RuntimeError("reading images needs Windows OCR, which is not available on this PC")
    from PIL import Image
    Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS     # refuse decompression bombs
    with Image.open(path) as img:
        img.load()
        return ocr.ocr_image(img)


def _logo_supplier(path: Path) -> str:
    """Supplier printed only as a logo: OCR the first and last page images."""
    if not (ocr and ocr.available()):
        return ""
    ext = path.suffix.lower()
    try:
        if ext == ".pdf":
            with pdfplumber.open(str(path)) as pdf:
                n = len(pdf.pages)
                if not any(p.images for p in pdf.pages):
                    return ""
            imgs = [_render_page(path, i, 2500) for i in sorted({0, n - 1})]
        elif ext in IMAGE_EXTS:
            from PIL import Image
            with Image.open(path) as im:
                imgs = [im.copy()]
        else:
            return ""
        for img in imgs:
            t = ocr.all_text(img)
            for s in KNOWN_SUPPLIERS:
                # OCR may double a letter at the edge of a logo ("CLEVOO")
                if re.search(rf"\b{re.escape(s)}", t, re.IGNORECASE):
                    return s
    except Exception:
        return ""
    return ""


def _read_docx(path: Path) -> tuple[str, list[list[list[str]]]]:
    if docx is None: raise RuntimeError("python-docx not installed")
    d = docx.Document(str(path))
    parts = [p.text for p in d.paragraphs if p.text]
    tables = [[[_norm(c.text) for c in r.cells] for r in t.rows] for t in d.tables]
    return "\n".join(parts), tables


def _read_txt(path: Path) -> tuple[str, list[list[list[str]]]]:
    return path.read_text(encoding="utf-8", errors="replace"), []


def _read_csv(path: Path) -> tuple[str, list[list[list[str]]]]:
    rows = []
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as f:
        for row in csv.reader(f):
            rows.append([_norm(c) for c in row])
    return "\n".join(" | ".join(r) for r in rows), ([rows] if rows else [])


def _read_xlsx(path: Path) -> tuple[str, list[list[list[str]]]]:
    if openpyxl is None: raise RuntimeError("openpyxl not installed")
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    tables, text_parts = [], []
    try:
        for ws in wb.worksheets:
            rows = []
            for row in ws.iter_rows(values_only=True, max_row=MAX_SHEET_ROWS, max_col=MAX_SHEET_COLS):
                cells = [_norm("" if v is None else str(v)) for v in row]
                if any(cells): rows.append(cells)
            if rows:
                # drop columns that are empty in every row (sheets often report 100+)
                width = max(max((i + 1 for i, c in enumerate(r) if c), default=0) for r in rows)
                rows = [r[:width] for r in rows]
                tables.append(rows)
                text_parts.append("\n".join(" | ".join(r) for r in rows))
    finally:
        wb.close()
    return "\n".join(text_parts), tables


MAX_SHEET_ROWS = 5000
MAX_SHEET_COLS = 200
MAX_PDF_PAGES = 60


READERS = {
    ".pdf": _read_pdf, ".docx": _read_docx, ".txt": _read_txt, ".md": _read_txt,
    ".csv": _read_csv, ".xlsx": _read_xlsx, ".xlsm": _read_xlsx,
    **{e: _read_image for e in IMAGE_EXTS},
}


# ---------------------------------------------------------------------------
# Line-item detection
# ---------------------------------------------------------------------------
# Fuzzy label matching: a label that is >= 80% similar to one of these terms
# counts as that field ("Keybaord", "Proccessor", "Weigth", "LCD Pannel").
# Used only after the exact patterns fail.
FUZZY_THRESHOLD = 0.80
FUZZY_TERMS: dict[str, list[str]] = {
    "bbcpu":    ["b/b+cpu", "barebone+cpu", "bare bone+cpu", "bb+cpu"],
    "barebone": ["barebone", "bare bone", "chassis", "base model"],
    "cpu":      ["processor"],
    "memory":   ["memory", "lpddr5", "ddr5", "dram"],
    "keyboard": ["keyboard"],
    "jp_kb":    ["jp keyboard", "japanese keyboard"],
    "rj45":     ["lan port", "ethernet"],
    "lcd":      ["display", "panel", "screen"],
    "weight":   ["weight"],
    "moq":      ["minimum order quantity", "min order qty"],
    "material": ["material"],
    "price":    ["price", "unit price"],
}
# Real words that look like field names but are something else
_FUZZY_BLOCK = {"height", "width", "depth", "length", "eight", "weights", "panels",
                "displayport", "prices", "presser"}


def _fuzzy_kind(label: str, allowed) -> str | None:
    from difflib import SequenceMatcher
    s = re.sub(r"\s+", " ", re.sub(r"[^a-z0-9+/ ]", " ", _norm(label).lower())).strip()
    if not s or len(s) > 40:
        return None
    words = s.split()
    best, best_kind = 0.0, None
    for kind, terms in FUZZY_TERMS.items():
        if kind not in allowed:
            continue
        for term in terms:
            n = len(term.split())
            # compare the whole label and every same-length word window
            cands = [s] + [" ".join(words[i:i + n]) for i in range(len(words) - n + 1)]
            for c in cands:
                if c in _FUZZY_BLOCK or len(c) < 5:
                    continue
                r = SequenceMatcher(None, c, term).ratio()
                if r >= FUZZY_THRESHOLD and r > best:
                    best, best_kind = r, kind
    return best_kind


def _match_item(cell: str) -> str | None:
    cell = _norm(cell).lower()
    if not cell: return None
    for kind, pats in ITEM_ALIASES.items():
        for p in pats:
            if re.search(p, cell, re.IGNORECASE): return kind
    kind = _fuzzy_kind(cell, ITEM_ALIASES.keys() | {"keyboard"})
    return "jp_kb" if kind == "keyboard" and re.search(r"\bjp\b|japan|jis", cell) else \
           (None if kind == "keyboard" else kind)


def _find_header(grid: list[list[str]]) -> tuple[int, dict[str, int]] | None:
    """Find a header row with columns like Item / Spec / Price / Remark."""
    targets = {
        "item":   [r"^item(\s*name)?$", r"^(part|component|description\s*item)$", r"^部品$", r"^項目$", r"^品名$"],
        "spec":   [r"spec", r"description", r"仕様"],
        "price":  [r"^pric", r"cost", r"amount", r"usd", r"単価", r"価格", r"金額"],
        "remark": [r"remark", r"^\W*\w?\s*mark$", r"note", r"備考"],
        "qty":    [r"^(qty|quantity)$", r"数量", r"m[oq]?q"],
        "model":  [r"^model(\s*name)?$", r"^model\s*no\.?$", r"型番"],
    }
    for r_i, row in enumerate(grid):
        found: dict[str, int] = {}
        for c_i, cell in enumerate(row):
            c_norm = _norm(cell).lower()
            for key, pats in targets.items():
                if any(re.search(p, c_norm, re.IGNORECASE) for p in pats):
                    found.setdefault(key, c_i)
        # Must have at least Item + Price (or Item + Spec) to be useful
        if "price" in found and ("item" in found or "spec" in found):
            return r_i, found
    return None


def _collect_line_items(grids: list[list[list[str]]]) -> list[dict]:
    """Each item: {kind, item_cell, spec, price, remark, model, qty}."""
    items: list[dict] = []
    for grid in grids:
        hdr = _find_header(grid)
        if not hdr: continue
        hdr_row, cols = hdr
        # Header text lost (blank in a scan): an unlabelled column right of
        # Price that holds text is the remarks column
        if "remark" not in cols and "price" in cols:
            used = set(cols.values())
            for c in range(cols["price"] + 1, len(grid[hdr_row])):
                if c not in used and not _norm(grid[hdr_row][c]) and \
                   any(_norm(r[c]) for r in grid[hdr_row + 1:] if c < len(r)):
                    cols["remark"] = c
                    break
        # Carry-forward last non-empty value for sparse columns (common in merged tables)
        last: dict[str, str] = {}
        for r in grid[hdr_row + 1:]:
            def cell(key):
                idx = cols.get(key)
                if idx is None or idx >= len(r): return ""
                v = _norm(r[idx])
                if v:
                    last[key] = v
                    return v
                # Only the (merged) Model cell legitimately spans rows
                return last.get(key, "") if key == "model" else ""

            item_cell = cell("item")
            spec = cell("spec")
            price = cell("price")
            remark = cell("remark")
            qty = cell("qty")
            model = cell("model")
            # Classify by item cell (or spec if item empty/same as spec)
            kind = _match_item(item_cell) or _match_item(spec[:40])
            if not kind and not (item_cell or spec or price or remark):
                continue
            items.append({
                "kind": kind,
                "item": item_cell,
                "spec": spec,
                "price": price,
                "remark": remark,
                "qty": qty,
                "model": model,
            })
    return items


# ---------------------------------------------------------------------------
# Formatters
# ---------------------------------------------------------------------------
def _parse_money(s: str) -> float | None:
    if not s: return None
    m = MONEY_RE.search(s) or re.search(r"\b(" + MONEY_NUM + r")\s*(?:USD|US\$)?\b", s)
    if not m: return None
    try: return float(m.group(1).replace(",", ""))
    except ValueError: return None


def _fmt_money(v: float) -> str:
    if v == int(v): return f"${int(v)}"
    return f"${v:.2f}"


def _find_price(s: str) -> str:
    v = _parse_money(s)
    return _fmt_money(v) if v is not None else _norm(s)


# ---------------------------------------------------------------------------
# Prose parsers
# ---------------------------------------------------------------------------
_MATERIAL_TO_COVERS = [
    (r"\b(full|all)\s*plastic\b|プラスチック", "Plastic"),
    (r"\b(full|all)?\s*aluminum\b|\balu(minium)?\b|アルミ", "Aluminum"),
    (r"\b(full|all)?\s*magnesium\b|マグネシウム", "Magnesium"),
    (r"\bcarbon\s*(fiber|fibre)\b|カーボン", "Carbon"),
    (r"\bmetal\b", "Metal"),
]


def _parse_barebone_spec(spec: str, result: dict[str, str], other: list[str]) -> None:
    s = spec
    # Material -> A/C/D (uniform unless specific A/C/D given below)
    for pat, name in _MATERIAL_TO_COVERS:
        if re.search(pat, s, re.IGNORECASE):
            for col in ("A", "C", "D"):
                if not result.get(col): result[col] = name
            break
    # Explicit per-cover material: "A cover: Metal", "C-cover Plastic"
    for cover in ("A", "C", "D"):
        m = re.search(rf"\b{cover}[\s-]*cover\s*[:：-]?\s*([A-Za-z ]{{3,25}})", s, re.IGNORECASE)
        if m: result[cover] = _norm(m.group(1)).title()

    # Weight
    m = re.search(r"\b(?:weight|重量)\s*[:：]?\s*([<≤≈~]?)\s*([0-9]+(?:\.[0-9]+)?)\s*(kgs?|g)\b", s, re.IGNORECASE)
    if m:
        unit_out = "Kg" if m.group(3).lower().startswith("kg") else "g"
        result["weight"] = f"{m.group(1)}{m.group(2)}{unit_out}"

    # LCD resolution — strip "14 inch" prefix, keep "1920×1080"
    m = re.search(r"\b([0-9]{3,4})\s*[xX×]\s*([0-9]{3,4})\b", s)
    if m:
        result["LCD"] = f"{m.group(1)}*{m.group(2)}"

    # RJ45 presence
    if re.search(r"\bwithout\s*rj[\s-]?45\b|\bno\s*rj[\s-]?45\b", s, re.IGNORECASE):
        if not result.get("RJ45"): result["RJ45"] = "×"
    elif re.search(r"\bwith\s*rj[\s-]?45\b|\brj[\s-]?45\b", s, re.IGNORECASE):
        if not result.get("RJ45"): result["RJ45"] = "〇"

    # NOTE: battery/camera/wifi/charger are routine spec-sheet info and
    # belong in "Basic specs" rather than その他. Skipped intentionally.
    _ = other  # unused but kept in signature for future notes


def _parse_cpu_spec(spec: str, remark: str, price: str, result: dict[str, str]) -> None:
    text = f"{spec} {remark}".strip()
    # Platform — pick primary from spec, keep opt if remark differs
    for pat, name in PLATFORMS:
        if re.search(pat, spec, re.IGNORECASE):
            result["CPU Platform"] = name
            break
    # Primary CPU name from spec
    m = CPU_RE.search(spec)
    primary = _norm(m.group(0)) if m else ""
    # Optional CPU from remark + its price
    opt_m = CPU_RE.search(remark)
    opt_price = _parse_money(remark)
    if primary and opt_m and _norm(opt_m.group(0)).lower() != primary.lower():
        result["CPU Name"] = f"{primary} (opt {_norm(opt_m.group(0))})"
    else:
        result["CPU Name"] = primary or _norm(opt_m.group(0)) if opt_m else primary

    # Price
    p = _parse_money(price)
    if p is not None:
        if opt_m and opt_price is not None and _norm(opt_m.group(0)).lower() != primary.lower():
            result["CPU"] = f"{_fmt_money(p)} ({_norm(opt_m.group(0))} {_fmt_money(opt_price)})"
        else:
            result["CPU"] = _fmt_money(p)


def _fmt_moq(raw: str) -> str:
    raw = _norm(raw)
    m = re.search(r"([0-9]+)\s*[Kk]\b", raw)
    if m: return f"{m.group(1)}K"
    m = re.search(r"([0-9,]+)\s*(pcs|units?|台|個)?", raw)
    if m:
        try:
            n = int(m.group(1).replace(",", ""))
            if n >= 1000 and n % 1000 == 0: return f"{n // 1000}K"
            return str(n)
        except ValueError: pass
    return raw


# ---------------------------------------------------------------------------
# Main extraction
# ---------------------------------------------------------------------------
def _rectangular(table: list[list]) -> list[list[str]]:
    if not table: return []
    ncols = max(len(r) for r in table)
    return [[_norm(r[c]) if c < len(r) else "" for c in range(ncols)] for r in table]


def _find_supplier(text: str) -> str:
    for s in KNOWN_SUPPLIERS:
        if re.search(rf"\b{re.escape(s)}\b", text, re.IGNORECASE):
            return s
    return ""


def _find_moq(text: str) -> str:
    m = re.search(r"\bMOQ\s*(?:is|:|：|\|)?\s*([0-9]+\s*[Kk]\b|[0-9,]+(?:\s*pcs)?)", text, re.IGNORECASE) \
        or re.search(r"\bM[OQ]Q\s*([0-9,]+)\s*(?:pcs|台)?", text, re.IGNORECASE)
    return _fmt_moq(m.group(1)) if m else ""


# ---------------------------------------------------------------------------
# Multi-product tables (one row per product)
#
#   A) Repeated label blocks - products side by side, each block repeats the
#      same label column:   NX14 | LCD | 14"...   NX15 | LCD | 15.6"...
#   B) One label column, one column per model:
#                 | NS50MU | NS70MU
#      LCD        | 15.6"  | 17.3"
# Both checks are strict so ordinary single-product quotations are unaffected.
# ---------------------------------------------------------------------------
_HEADER_WORDS = re.compile(
    r"^(items?|specs?|specification|description|price|unit\s*price|qty|quantity|"
    r"remarks?|notes?|options?|model|no\.?|項目|仕様|価格|備考|数量)$", re.IGNORECASE)

# label -> field kind, for key/value product blocks (first match wins)
_KV_LABELS: list[tuple[str, list[str]]] = [
    ("ignore",   [r"touch", r"タッチ", r"\bv\s*ram\b", r"\bgpu\b"]),  # not LCD / system memory
    ("bbcpu",    ITEM_ALIASES["bbcpu"]),
    ("barebone", ITEM_ALIASES["barebone"]),
    ("cpu",      ITEM_ALIASES["cpu"]),
    ("memory",   ITEM_ALIASES["memory"] + [r"\bdram\b"]),
    ("keyboard", [r"key\s*board", r"\bkb\b", r"キーボード"]),
    ("rj45",     ITEM_ALIASES["rj45"]),
    ("lcd",      ITEM_ALIASES["lcd"]),
    ("weight",   ITEM_ALIASES["weight"]),
    ("moq",      ITEM_ALIASES["moq"]),
    ("material", [r"material", r"^design$", r"材質", r"素材"]),
    ("a_cover",  ITEM_ALIASES["a_cover"]),
    ("c_cover",  ITEM_ALIASES["c_cover"]),
    ("d_cover",  ITEM_ALIASES["d_cover"]),
    ("price",    [r"^(unit\s*|total\s*|fob\s*)?pric", r"^cost$", r"^(価格|単価|金額)$"]),
]
# Row label that names the products in a model-per-column table
# also common OCR misreads of "SKU" (SW, 5KU, SKIJ)
_MODEL_ROW = re.compile(r"^(s\s?k\s?(u|ij|v)|5ku|sw|model(\s*name)?|product(\s*name)?|機種|型番|モデル)$", re.IGNORECASE)


def _kv_kind(label: str) -> str | None:
    label = _norm(label)
    for kind, pats in _KV_LABELS:
        if any(re.search(p, label, re.IGNORECASE) for p in pats):
            return None if kind == "ignore" else kind
    return _fuzzy_kind(label, {k for k, _ in _KV_LABELS} - {"jp_kb"})


def _short_label(s: str) -> str:
    s = _norm(s)
    if not s or len(s) > 30 or MONEY_RE.search(s) or re.fullmatch(r"[0-9.,\s%]+", s):
        return ""
    return s


def _is_model_like(s: str) -> bool:
    s = _norm(s)
    return (bool(s) and len(s) <= 40 and not _HEADER_WORDS.match(s) and not MONEY_RE.search(s)
            and bool(re.search(r"[A-Za-z]", s)) and bool(re.search(r"[0-9]", s)))


def _block_pairs(grid, r0, label_col, value_cols) -> list[tuple[str, str, bool]]:
    """(label, value, is_default) triples from row r0 down. A blank label
    continues the one above it: those are option lines (is_default False)."""
    pairs, label = [], ""
    for r in range(r0, len(grid)):
        first = bool(grid[r][label_col])
        if first:
            label = grid[r][label_col]
        if not label:
            continue
        for c in value_cols:
            if grid[r][c]:
                pairs.append((label, grid[r][c], first))
    return pairs


def _products_repeated_labels(grid) -> list[tuple[str, list]]:
    ncols = len(grid[0]) if grid else 0
    hits: dict[tuple, int] = {}
    first_row: dict[tuple, int] = {}
    for r, row in enumerate(grid):
        cols_by_text: dict[str, list[int]] = {}
        for c, cell in enumerate(row):
            t = _short_label(cell).lower()
            if t:
                cols_by_text.setdefault(t, []).append(c)
        for cols in cols_by_text.values():
            if len(cols) >= 2:
                key = tuple(cols)
                hits[key] = hits.get(key, 0) + 1
                first_row.setdefault(key, r)
    # Identical values repeat across blocks too; labels sit LEFT of values,
    # so take the left-most column set that repeats on 3+ rows.
    candidates = [k for k, n in hits.items() if n >= 3]
    if not candidates:
        return []
    label_cols = min(candidates, key=lambda k: k[0])
    r0 = first_row[label_cols]
    # The block must actually be labelled with quotation fields
    if sum(1 for r in range(r0, len(grid)) if _kv_kind(grid[r][label_cols[0]])) < 2:
        return []
    products = []
    for i, lc in enumerate(label_cols):
        end = label_cols[i + 1] if i + 1 < len(label_cols) else ncols
        name = ""
        for r in range(r0 - 1, -1, -1):
            cells = [grid[r][c] for c in range(lc, end) if grid[r][c]]
            if cells:
                name = cells[0]
                break
        products.append((name, _block_pairs(grid, r0, lc, range(lc + 1, end))))
    return products


def _products_model_columns(grid, named_only: bool = False) -> list[tuple[str, list]]:
    ncols = len(grid[0]) if grid else 0
    for lc in range(min(2, ncols)):
        label_rows = [r for r in range(len(grid)) if _kv_kind(grid[r][lc])]
        if len(label_rows) < 3:
            continue
        named = [r for r in range(len(grid)) if _MODEL_ROW.match(_norm(grid[r][lc]))]
        if named:                      # "SKU | AB12CD3E | XY5ZZQ"
            hr = named[0]
            r0 = hr + 1
            heads = [(c, grid[hr][c]) for c in range(lc + 1, ncols) if grid[hr][c]]
        elif named_only:
            continue
        else:                          # model names in the row above the labels
            r0 = label_rows[0]
            heads = []
            for hr in range(r0 - 1, -1, -1):
                heads = [(c, grid[hr][c]) for c in range(lc + 1, ncols) if grid[hr][c]]
                if heads:
                    break
        # under an explicit SKU/Model label, any code-like token is a name
        # (OCR may have turned its only digit into a letter: XY5ZZQ -> XYSZZQ)
        ok = (lambda n: _is_model_like(n) or bool(re.fullmatch(r"[A-Z0-9][A-Z0-9_./()-]{3,24}", _norm(n))))\
            if named else _is_model_like
        if len(heads) < (1 if named_only else 2) or not all(ok(n) for _, n in heads):
            continue
        return [(name, _block_pairs(grid, r0, lc, [c])) for c, name in heads]
    return []


def _detect_products(grids) -> list[tuple[str, list]]:
    for grid in grids:
        # An explicit "SKU"/"Model" row is the strongest signal, so try it first:
        # identical spec values side by side can look like repeated labels.
        for finder in (lambda g: _products_model_columns(g, named_only=True),
                       _products_repeated_labels, _products_model_columns):
            prods = finder(grid)
            names = [_norm(n) for n, _ in prods]
            if len(prods) >= 2 and all(names) and len(set(names)) == len(names):
                return prods
    # A single product laid out as "SKU | X" + "label | value" rows
    for grid in grids:
        prods = _products_model_columns(grid, named_only=True)
        if len(prods) == 1 and _norm(prods[0][0]) and \
           sum(1 for l, _, _ in prods[0][1] if _kv_kind(l)) >= 3:
            return prods
    return []


def _panel_resolution(s: str) -> str:
    """'1920x1200' as written, else the standard resolution of a panel name."""
    res = re.search(r"\b([0-9]{3,4})\s*[xX×*]\s*([0-9]{3,4})\b", s)
    if res:
        return f"{res.group(1)}*{res.group(2)}"
    for pat, px in PANEL_RESOLUTIONS:
        if re.search(pat, s, re.IGNORECASE):
            if px is None:   # 2.5K
                return "2560*1600" if re.search(r"16\s*:\s*10", s) else "2560*1440"
            return px
    return ""


def _val_lcd_spec(s: str) -> str:
    px = _panel_resolution(s)
    if px:
        return px
    size = re.search(r"\b([0-9]{2}(?:\.[0-9])?)\s*(?:\"|”|″|inch|インチ)", s, re.IGNORECASE)
    return f'{size.group(1)}"' if size else ""


def _cpu_options(values: list[str]) -> list[tuple[str, str]]:
    """[(cpu_name, price), ...] from lines like 'Intel WCL Core 7 350U $520'."""
    out = []
    for v in values:
        v = re.sub(r"[™®]|\bprocessors?\b", " ", v, flags=re.IGNORECASE)
        v = re.sub(r"\bUltra\s*([3579])\b", r"Ultra \1", v)
        m = CPU_RE.search(v)
        p = _val_money(v)
        if m or p:
            out.append((_norm(m.group(0)) if m else "", p))
    return out


_DELTA_RE = re.compile(r"([+\-－])\s*(?:US\$|USD|\$)\s*([0-9][0-9,]*(?:\.[0-9]+)?)", re.IGNORECASE)
_REMARK_LABEL = re.compile(r"^(remarks?|notes?|備考|注記)$", re.IGNORECASE)
# Kinds whose option lines are already folded into their own column
_OPTION_HANDLED = {"bbcpu", "cpu", "memory"}


def _extras(pairs) -> list[str]:
    """Option lines with a price change, and remarks, for the その他 column."""
    out: list[str] = []
    for label, value, is_default in pairs:
        label, value = _norm(label), _norm(value)
        if _REMARK_LABEL.match(label):
            out.append(value)
            continue
        if _kv_kind(label) in _OPTION_HANDLED:
            continue
        d = _DELTA_RE.search(value)
        if not d:
            continue
        amount = f"{'-' if d.group(1) in '-－' else '+'}{_fmt_money(float(d.group(2).replace(',', '')))}"
        desc = _norm(re.sub(r"\(?\s*" + re.escape(d.group(0)) + r"\s*\)?", " ", value))
        desc = re.sub(r"\s*\*.*$", "", desc)          # drop "*For Intel"-style footnotes
        if not desc or desc.lower() == label.lower():
            desc = label
        out.append(f"{desc} {amount}")
    seen: set[str] = set()
    return [x for x in out if not (x in seen or seen.add(x))]


def _rows_from_pairs(name: str, pairs, doc_text: str) -> list[dict[str, str]]:
    """One row per product; several complete B/B+CPU prices (one per CPU)
    become one row each so every row carries a single price."""
    base = _row_from_pairs(name, pairs, doc_text)
    bb_vals = [_norm(v) for l, v, _ in pairs if _kv_kind(l) == "bbcpu"]
    opts = _cpu_options(bb_vals)
    if len(opts) < 2 or not all(n and p for n, p in opts):
        return [base]
    rows = []
    for n, p in opts:
        r = dict(base)
        r["CPU Name"], r["B/B+CPU"] = n, p
        rows.append(r)
    return rows


def _row_from_pairs(name: str, pairs, doc_text: str) -> dict[str, str]:
    res = {c: "" for c in COLUMNS}
    res["Barebone Name"] = _norm(name)
    res["ODM supplier"] = _find_supplier(doc_text)

    groups: dict[str, list[str]] = {}
    for label, value, _ in pairs:
        kind = _kv_kind(label)
        if kind:
            groups.setdefault(kind, []).append(_norm(value))
    res["その他"] = "; ".join(_extras(pairs))

    # CPU options: first line is the default, the rest become "(opt ...)"
    for kind, price_col in (("bbcpu", "B/B+CPU"), ("cpu", "CPU")):
        opts = _cpu_options(groups.get(kind, []))
        if not opts:
            continue
        (n0, p0), rest = opts[0], opts[1:]
        if p0:
            res[price_col] = p0 + (f" ({' / '.join(f'{n} {p}'.strip() for n, p in rest)})" if rest else "")
        if n0 and not res["CPU Name"]:
            others = [n for n, _ in rest if n]
            res["CPU Name"] = n0 + (f" (opt {' / '.join(others)})" if others else "")

    cpu_text = " ".join(groups.get("bbcpu", []) + groups.get("cpu", []))
    for pat, plat in PLATFORMS:
        if re.search(pat, cpu_text or doc_text, re.IGNORECASE):
            res["CPU Platform"] = plat
            break

    if groups.get("barebone"):
        res["Barebone"] = next((p for p in map(_val_money, groups["barebone"]) if p), "")

    for v in groups.get("memory", []):
        if not re.search(r"lpddr", v, re.IGNORECASE):
            continue
        gb = re.search(r"\b(8|12|16)\s*GB\b", v, re.IGNORECASE)
        col = f"LPDDR5 {gb.group(1)}GB" if gb else ""
        if col and not res[col]:
            res[col] = _val_money(v) or "〇"

    for v in groups.get("keyboard", []):
        if re.search(r"\b(jp|jis|japan(ese)?)\b|日本語", v, re.IGNORECASE):
            # JP layout as the default costs nothing extra
            res["JP Keyboad"] = _val_money(v) or ("$0.0" if re.search(r"default|標準", v, re.IGNORECASE) else "〇")
            break

    if groups.get("rj45"):
        res["RJ45"] = _val_rj45(groups["rj45"][0])
    if groups.get("lcd"):
        res["LCD"] = _val_lcd_spec(groups["lcd"][0])
    if groups.get("weight"):
        res["weight"] = _val_weight(groups["weight"][0])
    if groups.get("material"):
        v = groups["material"][0]
        # "ACD: Carbon Fiber/ B: Mylar" names the covers explicitly
        split = re.findall(r"\b([ABCD]{1,4})\s*[:：]\s*([^/;,]+)", v)
        if split:
            for letters, mat_txt in split:
                for ch in letters:
                    if ch in "ACD":
                        res[ch] = _val_material(mat_txt) or _norm(mat_txt)
        else:
            mat = _val_material(v)
            if mat and not mat.startswith("$"):
                res["A"] = res["C"] = res["D"] = mat
    for kind, col in (("a_cover", "A"), ("c_cover", "C"), ("d_cover", "D")):
        if groups.get(kind):
            res[col] = _val_material(groups[kind][0]) or groups[kind][0]

    res["MOQ"] = _fmt_moq(groups["moq"][0]) if groups.get("moq") else _find_moq(doc_text)

    # A single "Price" row for a spec that already includes the CPU is the B/B+CPU price
    if groups.get("price") and not res["B/B+CPU"]:
        p = next((x for x in map(_val_money, groups["price"]) if x), "")
        if res["CPU Name"] and not res["CPU"]:
            res["B/B+CPU"] = p
        elif not res["Barebone"]:
            res["Barebone"] = p

    if not res["B/B+CPU"]:
        bb = _parse_money(res["Barebone"])
        cp = _parse_money(re.sub(r"\s*\(.*?\)\s*$", "", res["CPU"]))
        if bb is not None and cp is not None:
            res["B/B+CPU"] = _fmt_money(bb + cp)
    return res


def _split_cpu_rows(row: dict[str, str]) -> list[dict[str, str]]:
    """'N100 (opt N150)' with CPU '$58 (N150 $63)' -> one row per CPU, each with
    its own CPU price and B/B+CPU = Barebone + CPU."""
    m = re.fullmatch(r"(.+?)\s*\(opt\s+(.+)\)", row["CPU Name"])
    if not m:
        return [row]
    names = [m.group(1)] + [n.strip() for n in m.group(2).split("/") if n.strip()]

    def option_prices(cell: str) -> dict[str, str]:
        pm = re.fullmatch(r"(\$[0-9.,]+)\s*\((.+)\)", cell)
        if not pm:
            return {}
        out = {names[0]: pm.group(1)}
        for part in pm.group(2).split("/"):
            om = re.fullmatch(r"\s*(.+?)\s+(\$[0-9.,]+)\s*", part)
            if om:
                out[om.group(1)] = om.group(2)
        return out

    cpu_p, bb_p = option_prices(row["CPU"]), option_prices(row["B/B+CPU"])
    if not cpu_p and not bb_p:
        # option named without its own price: keep one row
        return [row]
    base = _parse_money(row["Barebone"])
    rows = []
    for n in names:
        r = dict(row)
        r["CPU Name"] = n
        r["CPU"] = cpu_p.get(n, "") if cpu_p else row["CPU"]
        if bb_p:
            r["B/B+CPU"] = bb_p.get(n, "")
        else:
            c = _parse_money(r["CPU"])
            r["B/B+CPU"] = _fmt_money(base + c) if base is not None and c is not None else ""
        rows.append(r)
    return rows


# ---------------------------------------------------------------------------
# Safe reading: refuse files that are too big or would expand into gigabytes
# ---------------------------------------------------------------------------
MAX_FILE_BYTES = 50 * 1024 * 1024        # 50 MB on disk
MAX_UNZIPPED_BYTES = 300 * 1024 * 1024   # .docx/.xlsx are zip files
MAX_ZIP_RATIO = 200                      # typical office files compress < 20x


def read_document(path: Path) -> tuple[str, list[list[list[str]]]]:
    ext = path.suffix.lower()
    reader = READERS.get(ext)
    if reader is None:
        raise ValueError(f"unsupported file type: {ext}")
    size = path.stat().st_size
    if size > MAX_FILE_BYTES:
        raise ValueError(f"file is larger than {MAX_FILE_BYTES // (1024 * 1024)} MB")
    if ext in (".docx", ".xlsx", ".xlsm"):
        import zipfile
        try:
            with zipfile.ZipFile(path) as z:
                total = sum(i.file_size for i in z.infolist())
        except zipfile.BadZipFile:
            raise ValueError("file is damaged or not a real Office document")
        if total > MAX_UNZIPPED_BYTES or (size and total / size > MAX_ZIP_RATIO):
            raise ValueError("file expands to an unsafe size (possible zip bomb)")
    text, tables = reader(path)
    return _normalize(text), [[[_normalize(c) for c in row] for row in t] for t in tables]


def _normalize(s: str) -> str:
    """Write CPU names one way: 'Intel® Core™ 7 processor 350' -> 'Intel Core 7 350'."""
    if not s:
        return s
    s = re.sub(r"[™®]", " ", s)
    # OCR spells the symbols out: "CoreTMUltra", "Intel(R)"
    s = re.sub(r"\b(Core|Intel|Pentium|Celeron|Ryzen)\s*(?:TM|\(TM\)|\(R\)|®)(?=\s|[A-Z0-9])", r"\1 ", s)
    s = re.sub(r"\b(Core\s*(?:Ultra\s*)?[3579])\s*processors?\s+(?=[0-9])", r"\1 ", s, flags=re.IGNORECASE)
    s = re.sub(r"\b(Core)\s*Ultra\s*([3579])(?=\s|\b)", r"\1 Ultra \2", s, flags=re.IGNORECASE)
    return re.sub(r"[ \t]{2,}", " ", s)


# ---------------------------------------------------------------------------
# Output safety and self-checks
# ---------------------------------------------------------------------------
def _safe_cell(v: str) -> str:
    """Values pasted into Excel that start with = + - @ become formulas.
    Text from a quotation must never run as one, so prefix those with '."""
    v = (v or "").replace("\t", " ").replace("\r", " ").replace("\n", " ").strip()
    if v[:1] in ("=", "@") or (v[:1] in ("+", "-") and not re.match(r"[+\-]\s*[$¥]?\s*[0-9]", v)):
        return "'" + v
    return v


def _squash(s: str) -> str:
    return re.sub(r"[^0-9a-z]", "", s.lower())


def check_row(row: dict[str, str], text: str) -> dict[str, str]:
    """Columns whose value could not be confirmed against the document,
    with the reason. Nothing is changed - the UI highlights these cells."""
    warn: dict[str, str] = {}
    doc = _squash(text)
    doc_numbers = {float(n.replace(",", "")) for n in re.findall(r"[0-9][0-9,]*(?:\.[0-9]+)?", text)
                   if n.replace(",", "").replace(".", "", 1).isdigit()}

    name = row.get("Barebone Name", "")
    if name and _squash(name) and _squash(name) not in doc:
        warn["Barebone Name"] = "not found word-for-word in the document"
    # CPUs are written many ways ("Core™Ultra 7 Processors 356H"): check the model number
    for cpu_name in re.split(r"\(opt |/", row.get("CPU Name", "")):
        tok = re.findall(r"[0-9]{3,5}[A-Z]{0,3}", cpu_name)
        if tok and _squash(tok[-1]) not in doc:
            warn["CPU Name"] = f"{cpu_name.strip(' )')} not found in the document"
            break

    bb, cpu, total = (_parse_money(re.sub(r"\s*\(.*\)$", "", row.get(c, "")))
                      for c in ("Barebone", "CPU", "B/B+CPU"))
    if bb is not None and cpu is not None and total is not None and abs(bb + cpu - total) > 0.01:
        warn["B/B+CPU"] = f"does not equal Barebone + CPU ({_fmt_money(bb + cpu)})"

    for col in ("Barebone", "CPU", "B/B+CPU", "LPDDR5 8GB", "LPDDR5 12GB", "LPDDR5 16GB", "JP Keyboad"):
        v = row.get(col, "")
        if col in warn or not v:
            continue
        if col == "B/B+CPU" and bb is not None and cpu is not None:
            continue                    # calculated, already checked above
        for amt in re.findall(r"\$\s*([0-9][0-9,]*(?:\.[0-9]+)?)", v):
            n = float(amt.replace(",", ""))
            if n and n not in doc_numbers:
                warn[col] = f"${amt} does not appear in the document"
                break
    return warn


_COMPANY_RE = re.compile(
    r"^\s*([A-Z][A-Za-z0-9&\-]{2,30})(?:\s+[A-Z][A-Za-z]+){0,3}\s+"
    r"(?:Technology|Technologies|Electronics|Computer|Computers|Industrial|Intelligent)\b"
    r"(?:\s+(?:Co\.?|Corp\.?|Corporation|Inc\.?|Ltd\.?|Limited|,))*", re.MULTILINE)
_NOT_SUPPLIER_LINE = re.compile(r"^\s*(to|attn|attention|customer|buyer|bill\s*to|ship\s*to|dear|宛先|御中)\b",
                                re.IGNORECASE)


def _supplier_from_letterhead(text: str) -> str:
    """'Uniwill Technology Inc.' at the top of a quotation -> 'Uniwill'.
    Only the first few lines (the letterhead), never a To:/Attn: line."""
    for line in text.splitlines()[:8]:
        if _NOT_SUPPLIER_LINE.search(line):
            continue
        m = _COMPANY_RE.match(line)
        if m:
            return m.group(1)
    return ""


def _unique(values) -> str:
    vals = {v for v in values if v}
    return vals.pop() if len(vals) == 1 else ""


def _finalize(rows: list[dict[str, str]], text: str) -> None:
    """Fill blanks from the whole document - but only when the document holds
    exactly one candidate, so values from different products never mix."""
    lines = [l for l in text.splitlines() if l.strip()]
    supplier = _find_supplier(text) or _supplier_from_letterhead(text)

    lcd_lines = [l for l in lines if re.search(r"\b(lcd|panel|display|screen)\b|液晶|画面", l, re.IGNORECASE)
                 and not re.search(r"touch|タッチ", l, re.IGNORECASE)]
    lcd = _unique(_panel_resolution(l) for l in lcd_lines)
    weight = _unique(_val_weight(l) for l in lines if re.search(r"\bweight\b|重量|重さ", l, re.IGNORECASE))
    rj = [l for l in lines if re.search(r"rj[\s-]?45", l, re.IGNORECASE)]
    rj45 = _unique(_val_rj45(l) for l in rj)
    mat_lines = [l for l in lines if re.search(r"\b(material|housing|chassis|casing|design)\b|材質|素材|筐体",
                                               l, re.IGNORECASE)]
    material = _unique(m for m in (_val_material(l) for l in mat_lines) if m and not m.startswith("$"))

    for r in rows:
        if not r["ODM supplier"]:
            r["ODM supplier"] = supplier
        lcd_now = r["LCD"]
        if not re.search(r"[0-9]\*[0-9]", lcd_now):
            r["LCD"] = _panel_resolution(lcd_now) or lcd or lcd_now
        if not r["weight"]:
            r["weight"] = weight
        if not r["RJ45"]:
            r["RJ45"] = rj45
        if material and not (r["A"] or r["C"] or r["D"]):
            r["A"] = r["C"] = r["D"] = material


def extract_checked(path: Path) -> list[dict]:
    """[{'row': {...}, 'warn': {column: reason}}] - used by the app."""
    text, tables = read_document(path)
    rows = [r for row in _extract_from(text, tables) for r in _split_cpu_rows(row)]
    _finalize(rows, text)
    if any(not r["ODM supplier"] for r in rows):
        logo = _logo_supplier(path)
        if logo:
            text += f"\n{logo}"          # so the self-check can confirm it
            for r in rows:
                r["ODM supplier"] = r["ODM supplier"] or logo
    out = []
    for r in rows:
        warn = check_row(r, text)
        out.append({"row": {c: _safe_cell(r.get(c, "")) for c in COLUMNS}, "warn": warn})
    return out


def extract(path: Path) -> list[dict[str, str]]:
    return [x["row"] for x in extract_checked(path)]


def _extract_from(text: str, tables) -> list[dict[str, str]]:
    grids = [_rectangular(t) for t in tables if t]

    # Several products in one file -> one row each
    products = _detect_products(grids)
    if products:
        return [row for name, pairs in products for row in _rows_from_pairs(name, pairs, text)]

    result: dict[str, str] = {c: "" for c in COLUMNS}
    other_notes: list[str] = []

    # 1) Supplier
    result["ODM supplier"] = _find_supplier(text)

    # 2) Model code — try a row in table flagged as "Model", then regex scan
    for grid in grids:
        hdr = _find_header(grid)
        if not hdr: continue
        _, cols = hdr
        if "model" in cols:
            for row in grid:
                v = row[cols["model"]] if cols["model"] < len(row) else ""
                if v and re.search(r"[A-Z]{2,4}[0-9]{2,4}", v):
                    result["Barebone Name"] = v
                    break
            if result.get("Barebone Name"): break
    if not result["Barebone Name"]:
        m = re.search(r"\b([A-Za-z]{1,4}[0-9][A-Za-z0-9]*[-–][A-Za-z0-9][A-Za-z0-9\-]{2,25})\b", text)
        if m: result["Barebone Name"] = m.group(1)
    if not result["Barebone Name"]:
        # a datasheet title: the first line is just the model code ("NX17AB1-2")
        for line in [l.strip() for l in text.splitlines() if l.strip()][:5]:
            if re.fullmatch(r"[A-Za-z]{1,5}[0-9][A-Za-z0-9]*(?:[-–/][A-Za-z0-9]+)*", line) and len(line) >= 5 \
               and not CPU_RE.fullmatch(line):
                result["Barebone Name"] = line
                break

    # 3) Line items
    items = _collect_line_items(grids)

    # 4) Barebone line
    for it in items:
        if it["kind"] != "barebone": continue
        if it["price"] and not result["Barebone"]:
            result["Barebone"] = _find_price(it["price"])
        _parse_barebone_spec(it["spec"], result, other_notes)
        # MOQ may live in the barebone's qty cell
        if it["qty"] and not result["MOQ"]:
            result["MOQ"] = _fmt_moq(it["qty"])
        break

    # 5) CPU
    cpu_item = next((it for it in items if it["kind"] == "cpu"), None)
    if cpu_item:
        _parse_cpu_spec(cpu_item["spec"], cpu_item["remark"], cpu_item["price"], result)

    # 6) Memory — map spec GB to the right LPDDR5 column ONLY if the spec
    # actually says LPDDR (regular SO-DIMM DDR4/5 does not belong here).
    for it in items:
        if it["kind"] != "memory": continue
        spec = it["spec"]
        if not re.search(r"\blpddr\b", spec, re.IGNORECASE) \
           and not re.search(r"\blpddr\b", it["item"], re.IGNORECASE):
            continue
        gb_m = re.search(r"\b(8|12|16)\s*GB\b", spec, re.IGNORECASE)
        if not gb_m: continue
        col = f"LPDDR5 {gb_m.group(1)}GB"
        if col in result and not result[col]:
            result[col] = _find_price(it["price"])

    # 7) JP Keyboard / RJ45 / LCD / A/C/D as explicit line items
    for it in items:
        if it["kind"] == "bbcpu" and not result["B/B+CPU"] and it["price"]:
            result["B/B+CPU"] = _find_price(it["price"])
        if it["kind"] == "jp_kb" and not result["JP Keyboad"]:
            result["JP Keyboad"] = _find_price(it["price"]) if it["price"] else _norm(it["spec"])
        if it["kind"] == "rj45" and not result["RJ45"]:
            result["RJ45"] = _find_price(it["price"]) if it["price"] else "〇"
        if it["kind"] == "lcd" and not result["LCD"]:
            m = re.search(r"\b([0-9]{3,4})\s*[xX×]\s*([0-9]{3,4})\b", it["spec"])
            result["LCD"] = f"{m.group(1)}*{m.group(2)}" if m else _norm(it["spec"])
        if it["kind"] in ("a_cover", "c_cover", "d_cover"):
            tgt = {"a_cover": "A", "c_cover": "C", "d_cover": "D"}[it["kind"]]
            if not result[tgt]: result[tgt] = _norm(it["spec"]) or "〇"

    # 8) MOQ from terms — prefer this over qty-cell parse if clear
    m = re.search(r"\bMOQ\s*(?:is|:|：|\|)?\s*([0-9]+\s*[Kk]\b|[0-9,]+(?:\s*pcs)?)", text, re.IGNORECASE)
    if m: result["MOQ"] = _fmt_moq(m.group(1))
    else:
        m = re.search(r"\bM[OQ]Q\s*([0-9,]+)\s*(?:pcs|台)?", text, re.IGNORECASE)
        if m: result["MOQ"] = _fmt_moq(m.group(1))

    # 9) B/B+CPU sum
    bb = _parse_money(result.get("Barebone", ""))
    # for CPU, strip the "(opt ...)" part
    cpu_main = re.sub(r"\s*\(.*?\)\s*$", "", result.get("CPU", "") or "")
    cpu_v = _parse_money(cpu_main)
    if bb is not None and cpu_v is not None and not result["B/B+CPU"]:
        result["B/B+CPU"] = _fmt_money(bb + cpu_v)

    # 10) その他 — join accumulated notes (barebone extras etc.)
    if other_notes and not result["その他"]:
        result["その他"] = " / ".join(other_notes)

    # 11) Universal fallback: for every still-empty field, scan the full
    # corpus (text + flattened table cells) for an anchor keyword and
    # grab the nearest matching value. No placeholders — just what's
    # actually in the document.
    corpus_lines = _build_corpus_lines(text, grids)
    _universal_fallback(corpus_lines, result)

    # 12) Always return ONE row per file. No variant splitting — if the
    # document has multiple products the user can duplicate the row in
    # the UI and edit, which is far safer than guessing from layout.
    return [result]


def _build_corpus_lines(text: str, grids: list[list[list[str]]]) -> list[str]:
    """Every text line plus every table row (cells joined by ' | ')."""
    lines = [_norm(l) for l in text.splitlines()]
    for g in grids:
        for row in g:
            cells = [c for c in row if c]
            if cells:
                lines.append(" | ".join(cells))
    return [l for l in lines if l]


def _val_money(s: str) -> str:
    # Prose lines: only trust amounts with a currency marker ("6. WIFI" is not $6)
    m = MONEY_RE.search(s)
    if not m: return ""
    try: return _fmt_money(float(m.group(1).replace(",", "")))
    except ValueError: return ""


def _val_weight(s: str) -> str:
    m = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*(kgs?|g)\b", s, re.IGNORECASE)
    if not m: return ""
    return f"{m.group(1)}{'Kg' if m.group(2).lower().startswith('kg') else 'g'}"


def _val_lcd(s: str) -> str:
    px = _panel_resolution(s)
    if px: return px
    m = re.search(r"\b([0-9]{2}(?:\.[0-9])?)\s*(?:\"|inch|インチ|”)", s, re.IGNORECASE)
    return f'{m.group(1)}"' if m else ""


def _val_material(s: str) -> str:
    for pat, name in _MATERIAL_TO_COVERS:
        if re.search(pat, s, re.IGNORECASE): return name
    return _val_money(s)


def _val_jp(s: str) -> str:
    """JP keyboard adder: its price, or $0.0 when it costs nothing extra."""
    if not re.search(r"\bjp\b|\bjis\b|japan|日本語", s, re.IGNORECASE):
        return ""
    if re.search(r"no\s*(extra|additional)\s*(cost|charge)|free|included|standard|default|標準|無償|込み", s, re.IGNORECASE):
        return "$0.0"
    return _val_money(s)


def _val_rj45(s: str) -> str:
    if re.search(r"\b(without|no|w/o)\s*rj[\s-]?45\b|rj[\s-]?45\s*(なし|無)", s, re.IGNORECASE):
        return "×"
    return _val_money(s) or "〇"


def _val_after_anchor(s: str, anchor: str) -> str:
    """Text that follows the anchor keyword on the same line (for その他 etc.)."""
    m = re.search(anchor + r"\s*[:：]?\s*(.+)", s, re.IGNORECASE)
    return _norm(m.group(1))[:150] if m else ""


# column -> (anchor regexes, value function). Only used for still-empty fields.
UNIVERSAL: dict[str, tuple[list[str], callable]] = {
    "Barebone":    ([r"\bbare\s*bone\b", r"\bbarebone\b", r"ベアボーン"], _val_money),
    "CPU":         ([r"(?<!\+)(?<!\+ )\bcpu\b", r"\bprocessor\b"], _val_money),
    "B/B+CPU":     ([r"b/?b\s*\+\s*cpu", r"barebone\s*\+\s*cpu"], _val_money),
    "LPDDR5 8GB":  ([r"lpddr5x?\s*8\s*gb", r"8\s*gb\s*lpddr5"], _val_money),
    "LPDDR5 12GB": ([r"lpddr5x?\s*12\s*gb", r"12\s*gb\s*lpddr5"], _val_money),
    "LPDDR5 16GB": ([r"lpddr5x?\s*16\s*gb", r"16\s*gb\s*lpddr5"], _val_money),
    "JP Keyboad":  ([r"jp\s*k(ey)?b(oard)?", r"japan(ese)?\s*keyboard", r"\bjis\b", r"日本語",
                     r"keyboard\s*[:：]\s*(jp|japan)"], lambda s: _val_jp(s)),
    "RJ45":        ([r"rj[\s-]?45"], _val_rj45),
    "A":           ([r"\ba[\s-]*(cover|case|shell)\b", r"Aカバー", r"A面"], _val_material),
    "C":           ([r"\bc[\s-]*(cover|case|shell)\b", r"Cカバー", r"C面"], _val_material),
    "D":           ([r"\bd[\s-]*(cover|case|shell)\b", r"Dカバー", r"D面"], _val_material),
    "weight":      ([r"\bweight\b", r"重量", r"重さ"], _val_weight),
    "LCD":         ([r"\blcd\b", r"\bpanel\b", r"\bdisplay\b", r"\bscreen\b", r"液晶", r"画面"], _val_lcd),
}


_MATERIAL_WORDS = r"(alumin(?:i)?um|alu|plastic|metal|magnesium|carbon(?:\s*fib(?:er|re))?|アルミ|プラスチック)"


def _covers_from_line(line: str) -> dict[str, str]:
    """'A cover aluminum, C/D plastic' -> {'A': 'Aluminum', 'C': 'Plastic', 'D': 'Plastic'}"""
    out: dict[str, str] = {}
    for m in re.finditer(r"(?<![A-Za-z])([ACD](?:\s*[/&,+]\s*[ACD])*)(?:[\s-]*(?:cover|case|shell|面|カバー))?"
                         r"\s*[:：=-]?\s*" + _MATERIAL_WORDS, line, re.IGNORECASE):
        word = m.group(2)
        mat = _val_material(word)
        if not mat or mat.startswith("$"):
            mat = "Plastic" if re.search(r"plastic|プラスチック", word, re.IGNORECASE) else _norm(word).title()
        for ch in re.findall(r"[ACD]", m.group(1).upper()):
            out.setdefault(ch, mat)
    return out


def _segment_after(line: str, anchor_match: re.Match) -> str:
    """Text from the anchor to the next ' / ', ';' or '|' - so on
    'LPDDR5 8GB +$18 / LPDDR5 16GB +$34' each size gets its own price."""
    rest = line[anchor_match.start():]
    # "Label | value | value": the cells after the label belong to it
    if re.match(r"[^|]{0,40}\|", rest):
        return rest
    return re.split(r"\s/\s|;|\|", rest, maxsplit=1)[0]


def _universal_fallback(lines: list[str], result: dict[str, str]) -> None:
    for col, (anchors, fn) in UNIVERSAL.items():
        if result.get(col): continue
        for i, line in enumerate(lines):
            m = next((m for a in anchors if (m := re.search(a, line, re.IGNORECASE))), None)
            if not m:
                continue
            # value next to the label first; then the next line, but only when
            # the label ends its line (value printed below it)
            seg = _segment_after(line, m)
            cands = [seg]
            if not re.sub(r"[\s:：]", "", line[m.end():]):
                # label alone on its line. In datasheets the label sits in the
                # middle of its bullet list, so the first value is ABOVE it.
                if i > 0 and lines[i - 1].lstrip().startswith(("•", "・", "-")):
                    cands.append(lines[i - 1])
                if i + 1 < len(lines):
                    cands.append(lines[i + 1])
            for cand in cands:
                v = fn(cand)
                if v:
                    result[col] = v
                    break
            if result.get(col): break

    # "A cover aluminum, C/D plastic"
    for line in lines:
        if re.search(r"cover|housing|material|chassis|casing|材質|筐体|カバー|面", line, re.IGNORECASE):
            for ch, mat in _covers_from_line(line).items():
                if not result[ch]:
                    result[ch] = mat
    # "Metal (A/D)" - material first, covers in brackets
    for line in lines:
        m = re.search(_MATERIAL_WORDS + r"\s*\(\s*([ACD](?:\s*[/&,]\s*[ACD])*)\s*\)?", line, re.IGNORECASE)
        if m and re.fullmatch(r"[ACD](?:\s*[/&,]\s*[ACD])*", m.group(2).strip(), re.IGNORECASE):
            mat = _val_material(m.group(1))
            if not mat or mat.startswith("$"):
                mat = _norm(m.group(1)).title()
            for ch in re.findall(r"[ACD]", m.group(2).upper()):
                if not result[ch]:
                    result[ch] = mat

    # Material stated once for the whole chassis -> A/C/D
    if not (result["A"] or result["C"] or result["D"]):
        for line in lines:
            if re.search(r"\bmaterial\b|材質|素材", line, re.IGNORECASE):
                v = _val_material(line)
                if v and not v.startswith("$"):
                    result["A"] = result["C"] = result["D"] = v
                    break

    # Weight / LCD anywhere in the doc (no anchor needed for distinctive formats)
    blob = "\n".join(lines)
    if not result["weight"]:
        m = re.search(r"\b([0-9]\.[0-9]{1,2})\s*kgs?\b", blob, re.IGNORECASE)
        if m: result["weight"] = f"{m.group(1)}Kg"
    if not result["LCD"]:
        result["LCD"] = _val_lcd(blob)

    # CPU name/platform anywhere
    if not result["CPU Name"]:
        # every CPU the document lists: the first is the default, the rest options
        names: list[str] = []
        for m in CPU_RE.finditer(blob):
            n = _norm(m.group(0))
            if n not in names:
                names.append(n)
        if names:
            result["CPU Name"] = names[0] + (f" (opt {' / '.join(names[1:])})" if len(names) > 1 else "")
    if not result["CPU Platform"]:
        for pat, name in PLATFORMS:
            if re.search(pat, blob, re.IGNORECASE):
                result["CPU Platform"] = name
                break
        if not result["CPU Platform"] and re.search(r"\bADL[\s-]*N\b", blob):
            result["CPU Platform"] = "Alder Lake-N"

    # Recompute B/B+CPU if the fallback just found the parts
    if not result["B/B+CPU"]:
        bb = _parse_money(result["Barebone"])
        cp = _parse_money(re.sub(r"\s*\(.*?\)\s*$", "", result["CPU"]))
        if bb is not None and cp is not None:
            result["B/B+CPU"] = _fmt_money(bb + cp)


# ---------------------------------------------------------------------------
# Inspect (for the app's debug modal)
# ---------------------------------------------------------------------------
def inspect(path: Path) -> dict:
    ext = path.suffix.lower()
    reader = READERS.get(ext)
    if reader is None:
        return {"ok": False, "error": f"unsupported type {ext}"}
    try:
        text, tables = read_document(path)
        checked = extract_checked(path)
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}
    grids = [_rectangular(t) for t in tables if t]
    items = _collect_line_items(grids)
    return {
        "ok": True, "file": path.name,
        "text": text, "tables": grids,
        "line_items": items, "columns": COLUMNS,
        "extracted": [x["row"] for x in checked],
        "warnings": [x["warn"] for x in checked],
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print("usage: python quote_extract.py <file>", file=sys.stderr); return 2
    path = Path(argv[1]).expanduser().resolve()
    if not path.is_file():
        print(f"not a file: {path}", file=sys.stderr); return 1
    rows = extract(path)
    header = "\t".join(COLUMNS)
    lines = ["\t".join(r.get(c, "") for c in COLUMNS) for r in rows]
    out = path.with_suffix(".tsv")
    out.write_text(header + "\n" + "\n".join(lines) + "\n", encoding="utf-8-sig")
    # Reconfigure stdout to utf-8 so em-dash / fullwidth chars print on JP consoles
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")   # type: ignore[attr-defined]
    except Exception: pass
    print(header)
    for ln in lines: print(ln)
    print(f"\n[wrote {len(rows)} row(s)] {out}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
