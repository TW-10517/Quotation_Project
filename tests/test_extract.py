"""
Regression tests for the extractor.

    python tests/test_extract.py          (or: python -m pytest tests)

Only synthetic, made-up quotations are used. They are generated in a temp
folder and deleted afterwards; no real supplier data is stored anywhere.
"""
from __future__ import annotations

import atexit
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import quote_extract as q  # noqa: E402

TMP = Path(tempfile.mkdtemp(prefix="quotegrid_test_"))
atexit.register(shutil.rmtree, TMP, ignore_errors=True)


def _write(name: str, content: str) -> Path:
    p = TMP / name
    p.write_text(content, encoding="utf-8")
    return p


def _rows(path: Path) -> list[dict]:
    return q.extract(path)


def _expect(row: dict, **want) -> None:
    for k, v in want.items():
        col = {"BB_CPU": "B/B+CPU", "Name": "Barebone Name", "Supplier": "ODM supplier",
               "Platform": "CPU Platform", "CPU_Name": "CPU Name", "JP": "JP Keyboad"}.get(k, k)
        assert row[col] == v, f"{col}: expected {v!r}, got {row[col]!r}"


# ---------------------------------------------------------------- synthetic
def test_single_product_spec_sheet():
    p = _write("single.csv",
               "Clevo Co. Ltd quotation\n"
               "SKU,NB14X2\n"
               "CPU,Intel MTL Core Ultra 5 125U\n"
               'LCD,"14"" FHD+ IPS"\n'
               "Weight,1.35kg\n"
               "Material,Aluminum\n"
               "Price,USD 410\n")
    rows = _rows(p)
    assert len(rows) == 1, rows
    _expect(rows[0], Supplier="CLEVO", Name="NB14X2", Platform="MTL", CPU_Name="Core Ultra 5 125U",
            BB_CPU="$410", LCD="1920*1200", weight="1.35Kg", A="Aluminum", C="Aluminum", D="Aluminum")


def test_two_models_side_by_side_never_mix():
    p = _write("two.csv",
               "Model,AB14,AB16\n"
               "CPU,Core 5 220H,Core 7 250H\n"
               'Display,"14"" 1920x1200","16"" 2560x1600"\n'
               "Weight,1.4 kg,1.8 kg\n"
               "Price,$500,$640\n")
    rows = _rows(p)
    assert [r["Barebone Name"] for r in rows] == ["AB14", "AB16"], rows
    _expect(rows[0], CPU_Name="Core 5 220H", LCD="1920*1200", weight="1.4Kg", BB_CPU="$500")
    _expect(rows[1], CPU_Name="Core 7 250H", LCD="2560*1600", weight="1.8Kg", BB_CPU="$640")


def test_misspelled_labels():
    p = _write("typos.csv",
               "SKU,ZX15\n"
               "Proccessor,Intel N150\n"
               "Weigth,1.6 kg\n"
               "Displya,15.6 FHD\n"
               "Keybaord,JP keyboard (default)\n")
    r = _rows(p)[0]
    _expect(r, Name="ZX15", CPU_Name="N150", weight="1.6Kg", LCD="1920*1080", JP="$0.0")


def test_cpu_options_become_rows():
    p = _write("opts.csv",
               "Model,KX14\n"
               "B/B+CPU,Core 7 350U $520\n"
               ",Core 5 320U $430\n"
               "LCD,14 FHD\n"
               "Keyboard,JP\n")
    rows = _rows(p)
    assert [(r["CPU Name"], r["B/B+CPU"]) for r in rows] == [("Core 7 350U", "$520"), ("Core 5 320U", "$430")], rows


def test_formula_text_is_neutralised():
    p = _write("evil.csv", 'SKU,"=HYPERLINK(""http://x"")"\nCPU,N100\nLCD,14 FHD\nWeight,1.3kg\n')
    for r in _rows(p):
        for v in r.values():
            assert not v.startswith(("=", "@")), v


def test_unknown_fields_stay_blank():
    p = _write("sparse.csv", "SKU,QQ13\nCPU,N100\nLCD,13.3 FHD\nWeight,1.2kg\n")
    r = _rows(p)[0]
    for col in ("ODM supplier", "MOQ", "Barebone", "CPU", "B/B+CPU", "RJ45", "A", "C", "D"):
        assert r[col] == "", f"{col} should be blank, got {r[col]!r}"


def test_recipient_company_is_not_the_supplier():
    p = _write("letter.txt", "To: Example Trading Co., Ltd\nModel: PQ14-AB1\nBarebone: $150\nCPU: N150 $60\n")
    assert _rows(p)[0]["ODM supplier"] == ""


def test_prose_quotation():
    p = _write("prose.txt",
               "Hasee Computer Co., Ltd\n"
               "Quotation for model HS14-N150\n"
               "Barebone price: US$ 152\n"
               "CPU: Intel Twin Lake N150, US$ 61\n"
               "Memory: LPDDR5 8GB onboard +$18 / LPDDR5 16GB +$34\n"
               "Keyboard: JP layout, no extra cost\n"
               "I/O: USB-C x2, HDMI, RJ45 x1\n"
               "Housing: A cover aluminum, C/D plastic\n"
               "Weight: approx. 1.38 kg\n"
               "Panel: 14 inch FHD+ IPS\n"
               "MOQ: 2,000 pcs\n")
    _expect(_rows(p)[0], Supplier="Hasee", Name="HS14-N150", Platform="Twin Lake", CPU_Name="N150",
            MOQ="2K", Barebone="$152", CPU="$61", BB_CPU="$213", **{"LPDDR5 8GB": "$18", "LPDDR5 16GB": "$34"},
            JP="$0.0", RJ45="〇", A="Aluminum", C="Plastic", D="Plastic", weight="1.38Kg", LCD="1920*1200")


def test_image_quotation_via_ocr():
    if not (q.ocr and q.ocr.available()):
        print("  (skipped: Windows OCR not available)"); return
    from PIL import Image, ImageDraw, ImageFont
    font = ImageFont.truetype("arial.ttf", 30)
    img = Image.new("RGB", (1400, 520), "white")
    d = ImageDraw.Draw(img)
    rows = [("Model", "KQ14-N150"), ("Barebone", "US$ 155"), ("CPU", "Intel N150 US$ 62"),
            ("Weight", "1.42 kg"), ("LCD", "14 inch 1920x1200"), ("MOQ", "1,000 pcs")]
    d.text((40, 20), "Emdoor Technology Co., Ltd", fill="black", font=font)
    for i, (k, v) in enumerate(rows):
        d.text((40, 90 + i * 64), k, fill="black", font=font)
        d.text((420, 90 + i * 64), v, fill="black", font=font)
    p = TMP / "scan.png"
    img.save(p)
    _expect(_rows(p)[0], Supplier="Emdoor", Name="KQ14-N150", CPU_Name="N150", Barebone="$155", CPU="$62",
            BB_CPU="$217", weight="1.42Kg", LCD="1920*1200", MOQ="1K")


def test_oversized_file_refused():
    p = TMP / "huge.txt"
    with p.open("wb") as f:
        f.seek(q.MAX_FILE_BYTES + 1); f.write(b"\0")
    try:
        q.extract(p)
    except ValueError:
        return
    raise AssertionError("oversized file was read")


if __name__ == "__main__":
    failed = 0
    for name, fn in list(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn(); print("PASS", name)
            except Exception as e:
                failed += 1; print("FAIL", name, "-", e)
    sys.exit(1 if failed else 0)
