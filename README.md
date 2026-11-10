<p align="center"><img src="assets/logo.png" width="96" alt="QuoteGrid logo"></p>

# QuoteGrid

Desktop app for Windows that reads ODM laptop quotations and turns each one into a single spreadsheet row, ready to paste into Excel.

Drop one or many quotations (PDF, DOCX, TXT, CSV, XLSX) onto the window. Each file becomes one row under these columns:

`ODM supplier · Barebone Name · CPU Platform · CPU Name · MOQ · Barebone · CPU · B/B+CPU · LPDDR5 8GB · LPDDR5 12GB · LPDDR5 16GB · JP Keyboad · RJ45 · A · C · D · weight · LCD · その他`

Only values found in the document are filled in; anything not mentioned is left blank. `B/B+CPU` is calculated from the Barebone and CPU prices.

## Download

Get `QuoteGrid-vX.Y.Z-win64.zip` from [Releases](../../releases), unzip, and double-click `QuoteGrid.exe`. No Python or internet connection needed. It needs Microsoft Edge WebView2, which Windows 11 already includes.

## Using it

- **Hold** a row for half a second to copy it.
- **Tick** rows (Shift+click for a range, Ctrl+A for all), then press **Ctrl+C** or **Copy selected**.
- Row copies have no header, so they paste straight under existing Excel rows. **Copy all** includes the header.
- Every cell is editable before copying.
- **Inspect one…** shows exactly what the app read from a file, which helps when a field comes back blank.

Files are processed entirely on your PC; nothing is uploaded.

## Run from source

```powershell
pip install -r requirements.txt
python src\quotegrid_app.py
```

Command-line extraction of a single file:

```powershell
python src\quote_extract.py "Quotation XYZ.pdf"
```

## Build the .exe

```powershell
.\build.ps1 -Version 1.0.0
```

Output: `release\QuoteGrid\QuoteGrid.exe` and `release\QuoteGrid-v1.0.0-win64.zip`.

## Adding support for a new quotation layout

The patterns live at the top of `src/quote_extract.py`:

- `KNOWN_SUPPLIERS` — supplier names to detect
- `ITEM_ALIASES` — words used in a quotation's Item column (Barebone, CPU, DDR, …)
- `UNIVERSAL` — keyword → value rules used when no item table is found
- `PLATFORMS` / `CPU_RE` — CPU platform and model names
