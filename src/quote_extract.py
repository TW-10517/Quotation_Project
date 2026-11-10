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
    "SIXUNITED", "Clevo", "Tongfang", "Compal", "Quanta", "Wistron",
    "Pegatron", "Inventec", "Mitac", "ECS", "ASUS", "MSI", "Gigabyte",
    "Mechrevo", "Hasee", "Uniwill", "Haier", "TongFang",
]

# Item-cell keywords -> item kind. Matched with re.search so "Name Barebone"
# (from merged cells) still classifies as barebone.
ITEM_ALIASES: dict[str, list[str]] = {
    "barebone":  [r"\bbarebone\b", r"\bbare\s*bone\b", r"\bchassis\b", r"\bbase\s*model\b", r"ベアボーン"],
    "cpu":       [r"\bcpu\b", r"\bprocessor\b", r"プロセッサ"],
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
    r"|Core\s*(?:Ultra\s*)?[579]\s*[0-9]{3,4}[A-Z]{0,3}"
    r"|i[3579]-[0-9]{4,5}[A-Z]{0,3}"
    r"|Ryzen\s*(?:AI\s*)?[3579]\s*[0-9]{3,4}[A-Z]{0,3}"
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
def _read_pdf(path: Path) -> tuple[str, list[list[list[str]]]]:
    parts, tables = [], []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            parts.append(page.extract_text() or "")
            try:
                for t in page.extract_tables() or []:
                    tables.append([[(_norm(c) if c else "") for c in r] for r in t])
            except Exception:
                pass
    return "\n".join(parts), tables


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
    for ws in wb.worksheets:
        rows = []
        for row in ws.iter_rows(values_only=True):
            cells = [_norm("" if v is None else str(v)) for v in row]
            if any(cells): rows.append(cells)
        if rows:
            tables.append(rows)
            text_parts.append("\n".join(" | ".join(r) for r in rows))
    return "\n".join(text_parts), tables


READERS = {
    ".pdf": _read_pdf, ".docx": _read_docx, ".txt": _read_txt, ".md": _read_txt,
    ".csv": _read_csv, ".xlsx": _read_xlsx, ".xlsm": _read_xlsx,
}


# ---------------------------------------------------------------------------
# Line-item detection
# ---------------------------------------------------------------------------
def _match_item(cell: str) -> str | None:
    cell = _norm(cell).lower()
    if not cell: return None
    for kind, pats in ITEM_ALIASES.items():
        for p in pats:
            if re.search(p, cell, re.IGNORECASE): return kind
    return None


def _find_header(grid: list[list[str]]) -> tuple[int, dict[str, int]] | None:
    """Find a header row with columns like Item / Spec / Price / Remark."""
    targets = {
        "item":   [r"^item(\s*name)?$", r"^(part|component|description\s*item)$", r"^部品$", r"^項目$", r"^品名$"],
        "spec":   [r"spec", r"description", r"仕様"],
        "price":  [r"^pric", r"cost", r"amount", r"usd", r"単価", r"価格", r"金額"],
        "remark": [r"remark", r"note", r"備考"],
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
        unit_out = "kg" if m.group(3).lower().startswith("kg") else "g"
        result["weight"] = f"{m.group(1)}{m.group(2)} {unit_out}"

    # LCD resolution — strip "14 inch" prefix, keep "1920×1080"
    m = re.search(r"\b([0-9]{3,4})\s*[x×]\s*([0-9]{3,4})\b", s)
    if m:
        result["LCD"] = f"{m.group(1)}×{m.group(2)}"

    # RJ45 presence
    if re.search(r"\bwithout\s*rj[\s-]?45\b|\bno\s*rj[\s-]?45\b", s, re.IGNORECASE):
        if not result.get("RJ45"): result["RJ45"] = "×"
    elif re.search(r"\bwith\s*rj[\s-]?45\b|\brj[\s-]?45\b", s, re.IGNORECASE):
        if not result.get("RJ45"): result["RJ45"] = "○"

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


def extract(path: Path) -> list[dict[str, str]]:
    ext = path.suffix.lower()
    reader = READERS.get(ext)
    if reader is None:
        raise ValueError(f"unsupported file type: {ext}")
    text, tables = reader(path)
    grids = [_rectangular(t) for t in tables if t]

    result: dict[str, str] = {c: "" for c in COLUMNS}
    other_notes: list[str] = []

    # 1) Supplier
    for s in KNOWN_SUPPLIERS:
        if re.search(rf"\b{re.escape(s)}\b", text, re.IGNORECASE):
            result["ODM supplier"] = s.upper() if s.isupper() or s.isalpha() and len(s) <= 10 else s
            if s.upper() == "SIXUNITED": result["ODM supplier"] = "SIXUNITED"
            break

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
        if it["kind"] == "jp_kb" and not result["JP Keyboad"]:
            result["JP Keyboad"] = _find_price(it["price"]) if it["price"] else _norm(it["spec"])
        if it["kind"] == "rj45" and not result["RJ45"]:
            result["RJ45"] = _find_price(it["price"]) if it["price"] else "○"
        if it["kind"] == "lcd" and not result["LCD"]:
            m = re.search(r"\b([0-9]{3,4})\s*[x×]\s*([0-9]{3,4})\b", it["spec"])
            result["LCD"] = f"{m.group(1)}×{m.group(2)}" if m else _norm(it["spec"])
        if it["kind"] in ("a_cover", "c_cover", "d_cover"):
            tgt = {"a_cover": "A", "c_cover": "C", "d_cover": "D"}[it["kind"]]
            if not result[tgt]: result[tgt] = _norm(it["spec"]) or "○"

    # 8) MOQ from terms — prefer this over qty-cell parse if clear
    m = re.search(r"\bMOQ\s*(?:is|:|：)?\s*([0-9]+\s*[Kk]\b|[0-9,]+(?:\s*pcs)?)", text, re.IGNORECASE)
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
    return f"{m.group(1)} {'kg' if m.group(2).lower().startswith('kg') else 'g'}"


def _val_lcd(s: str) -> str:
    m = re.search(r"\b([0-9]{3,4})\s*[x×*]\s*([0-9]{3,4})\b", s)
    if m: return f"{m.group(1)}×{m.group(2)}"
    m = re.search(r"\b([0-9]{2}(?:\.[0-9])?)\s*(?:\"|inch|インチ|”)", s, re.IGNORECASE)
    return f'{m.group(1)}"' if m else ""


def _val_material(s: str) -> str:
    for pat, name in _MATERIAL_TO_COVERS:
        if re.search(pat, s, re.IGNORECASE): return name
    return _val_money(s)


def _val_rj45(s: str) -> str:
    if re.search(r"\b(without|no|w/o)\s*rj[\s-]?45\b|rj[\s-]?45\s*(なし|無)", s, re.IGNORECASE):
        return "×"
    return _val_money(s) or "○"


def _val_after_anchor(s: str, anchor: str) -> str:
    """Text that follows the anchor keyword on the same line (for その他 etc.)."""
    m = re.search(anchor + r"\s*[:：]?\s*(.+)", s, re.IGNORECASE)
    return _norm(m.group(1))[:150] if m else ""


# column -> (anchor regexes, value function). Only used for still-empty fields.
UNIVERSAL: dict[str, tuple[list[str], callable]] = {
    "Barebone":    ([r"\bbare\s*bone\b", r"\bbarebone\b", r"ベアボーン"], _val_money),
    "CPU":         ([r"\bcpu\b", r"\bprocessor\b"], _val_money),
    "B/B+CPU":     ([r"b/?b\s*\+\s*cpu", r"barebone\s*\+\s*cpu"], _val_money),
    "LPDDR5 8GB":  ([r"lpddr5x?\s*8\s*gb", r"8\s*gb\s*lpddr5"], _val_money),
    "LPDDR5 12GB": ([r"lpddr5x?\s*12\s*gb", r"12\s*gb\s*lpddr5"], _val_money),
    "LPDDR5 16GB": ([r"lpddr5x?\s*16\s*gb", r"16\s*gb\s*lpddr5"], _val_money),
    "JP Keyboad":  ([r"jp\s*k(ey)?b(oard)?", r"japan(ese)?\s*keyboard", r"jis", r"日本語"], _val_money),
    "RJ45":        ([r"rj[\s-]?45"], _val_rj45),
    "A":           ([r"\ba[\s-]*(cover|case|shell)\b", r"Aカバー", r"A面"], _val_material),
    "C":           ([r"\bc[\s-]*(cover|case|shell)\b", r"Cカバー", r"C面"], _val_material),
    "D":           ([r"\bd[\s-]*(cover|case|shell)\b", r"Dカバー", r"D面"], _val_material),
    "weight":      ([r"\bweight\b", r"重量", r"重さ"], _val_weight),
    "LCD":         ([r"\blcd\b", r"\bpanel\b", r"\bdisplay\b", r"\bscreen\b", r"液晶", r"画面"], _val_lcd),
}


def _universal_fallback(lines: list[str], result: dict[str, str]) -> None:
    for col, (anchors, fn) in UNIVERSAL.items():
        if result.get(col): continue
        for i, line in enumerate(lines):
            if not any(re.search(a, line, re.IGNORECASE) for a in anchors):
                continue
            # same line first, then the next line (value often sits below its label)
            for cand in (line, lines[i + 1] if i + 1 < len(lines) else ""):
                v = fn(cand)
                if v:
                    result[col] = v
                    break
            if result.get(col): break

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
        if m: result["weight"] = f"{m.group(1)} kg"
    if not result["LCD"]:
        result["LCD"] = _val_lcd(blob)

    # CPU name/platform anywhere
    if not result["CPU Name"]:
        m = CPU_RE.search(blob)
        if m: result["CPU Name"] = _norm(m.group(0))
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
        text, tables = reader(path)
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}
    grids = [_rectangular(t) for t in tables if t]
    items = _collect_line_items(grids)
    return {
        "ok": True, "file": path.name,
        "text": text, "tables": grids,
        "line_items": items, "columns": COLUMNS,
        "extracted": extract(path),
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
