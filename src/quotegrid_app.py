"""
QuoteGrid - desktop app.

Drop one or many quotations (PDF / DOCX / TXT / CSV / XLSX). Each file
becomes one row under the 19 fixed columns; every cell is editable; copy
rows straight into Excel.

Run from source:
    python src/quotegrid_app.py
"""

from __future__ import annotations

import base64
import json
import os
import sys
import tempfile
import uuid
from pathlib import Path

import webview

import quote_extract as _qe

APP_NAME = "QuoteGrid"
# In a PyInstaller bundle, data files are unpacked under sys._MEIPASS
_ASSETS = (Path(sys._MEIPASS) / "assets" if hasattr(sys, "_MEIPASS")
           else Path(__file__).resolve().parent.parent / "assets")
COLUMNS = _qe.COLUMNS
inspect_file = _qe.inspect
SUPPORTED_EXTS = tuple(_qe.READERS.keys())
MAX_FILES_PER_DROP = 100


# ---------------------------------------------------------------------------
# Privacy: QuoteGrid never talks to the network. Enforce that in-process so a
# future library change can't quietly send quotation data anywhere.
# ---------------------------------------------------------------------------
def _block_network() -> None:
    import socket
    loopback = {"127.0.0.1", "::1", "localhost"}
    orig_connect, orig_connect_ex = socket.socket.connect, socket.socket.connect_ex

    def _check(addr) -> None:
        host = addr[0] if isinstance(addr, tuple) else addr
        if host not in loopback:
            raise PermissionError("QuoteGrid works offline; network access is disabled")

    def connect(self, addr):
        _check(addr); return orig_connect(self, addr)

    def connect_ex(self, addr):
        _check(addr); return orig_connect_ex(self, addr)

    socket.socket.connect, socket.socket.connect_ex = connect, connect_ex


class _PrivateTemp:
    """Dropped files are written to a private temp folder under random names
    (never the original file name) and deleted as soon as they are read."""
    def __init__(self) -> None:
        self._dir = tempfile.TemporaryDirectory(prefix="quotegrid_")

    def write(self, name: str, data: bytes) -> Path:
        suffix = Path(name).suffix.lower()
        if suffix not in SUPPORTED_EXTS:
            raise ValueError(f"unsupported type ({suffix or 'none'})")
        p = Path(self._dir.name) / f"{uuid.uuid4().hex}{suffix}"
        p.write_bytes(data)
        return p

    def __enter__(self): return self
    def __exit__(self, *exc): self._dir.cleanup()


def _purge_old_temp() -> None:
    """Remove temp copies a crashed earlier session may have left behind."""
    import shutil
    tmp = Path(tempfile.gettempdir())
    for pattern in ("quotegrid_*", "qx_*"):
        for p in tmp.glob(pattern):
            try:
                shutil.rmtree(p) if p.is_dir() else p.unlink()
            except OSError:
                pass


def _win_set_clipboard(text: str) -> None:
    """Put unicode text on the Windows clipboard (works without a focused,
    user-activated page, unlike navigator.clipboard)."""
    import ctypes
    import time
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    u32 = ctypes.WinDLL("user32", use_last_error=True)
    k32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    k32.GlobalAlloc.restype = wintypes.HGLOBAL
    k32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    k32.GlobalLock.restype = wintypes.LPVOID
    k32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    k32.GlobalFree.argtypes = [wintypes.HGLOBAL]
    u32.OpenClipboard.argtypes = [wintypes.HWND]
    u32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    u32.SetClipboardData.restype = wintypes.HANDLE

    data = text.encode("utf-16-le") + b"\x00\x00"
    h = k32.GlobalAlloc(0x0002, len(data))  # GMEM_MOVEABLE
    if not h:
        raise ctypes.WinError(ctypes.get_last_error())
    p = k32.GlobalLock(h)
    ctypes.memmove(p, data, len(data))
    k32.GlobalUnlock(h)

    for _ in range(20):  # another app may briefly hold the clipboard
        if u32.OpenClipboard(None):
            break
        time.sleep(0.05)
    else:
        k32.GlobalFree(h)
        raise RuntimeError("clipboard is busy")
    try:
        u32.EmptyClipboard()
        if not u32.SetClipboardData(13, h):  # CF_UNICODETEXT; clipboard owns h on success
            k32.GlobalFree(h)
            raise ctypes.WinError(ctypes.get_last_error())
    finally:
        u32.CloseClipboard()


# ---------------------------------------------------------------------------
# Python API exposed to JS as window.pywebview.api.*
# ---------------------------------------------------------------------------
class Api:
    # Only methods without a leading underscore are callable from the page.
    # None of them accept a file path from the page: files arrive either
    # through the native Open dialog or as dropped bytes.
    def __init__(self) -> None:
        self._window: webview.Window | None = None

    def supported_exts(self) -> list[str]:
        return list(SUPPORTED_EXTS)

    # ---- browse via native dialog ----------------------------------------
    def pick_files(self) -> dict:
        if self._window is None:
            return {"ok": False, "error": "no window"}
        patterns = "*" + ";*".join(SUPPORTED_EXTS)
        result = self._window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=(f"Documents ({patterns})", "All files (*.*)"),
        )
        if not result:
            return {"ok": False, "cancelled": True}
        return self._extract_paths(list(result))

    # ---- extract from dropped bytes --------------------------------------
    def extract_dropped(self, files: list[dict]) -> dict:
        """files: list of {name, data_b64}."""
        if len(files) > MAX_FILES_PER_DROP:
            return {"ok": False, "error": f"drop at most {MAX_FILES_PER_DROP} files at a time"}
        out_rows, errors = [], []
        with _PrivateTemp() as tmp:
            for f in files:
                display = Path(str(f.get("name", "file"))).name   # name only, never a path
                try:
                    data = base64.b64decode(f["data_b64"], validate=True)
                    if len(data) > _qe.MAX_FILE_BYTES:
                        raise ValueError(f"file is larger than {_qe.MAX_FILE_BYTES // (1024 * 1024)} MB")
                    self._extract_one(tmp.write(display, data), display, out_rows, errors)
                except Exception as e:
                    errors.append({"file": display, "error": str(e)})
        return {"ok": True, "rows": out_rows, "errors": errors, "columns": COLUMNS}

    def _extract_paths(self, paths: list[str]) -> dict:
        out_rows, errors = [], []
        for p in paths:
            path = Path(p)
            if not path.is_file():
                errors.append({"file": path.name, "error": "not a file"})
            elif path.suffix.lower() not in SUPPORTED_EXTS:
                errors.append({"file": path.name, "error": f"unsupported type ({path.suffix})"})
            else:
                self._extract_one(path, path.name, out_rows, errors)
        return {"ok": True, "rows": out_rows, "errors": errors, "columns": COLUMNS}

    @staticmethod
    def _extract_one(path: Path, display: str, out_rows: list, errors: list) -> None:
        try:
            results = _qe.extract_checked(path)
        except Exception as e:
            errors.append({"file": display, "error": str(e) or type(e).__name__})
            return
        for idx, res in enumerate(results):
            out_rows.append({
                "id": uuid.uuid4().hex,
                "file": display,
                "variant": (idx + 1) if len(results) > 1 else None,
                "variant_count": len(results),
                "row": res["row"],
                "warn": res["warn"],
            })

    # ---- inspect: show the user exactly what the extractor saw ------------
    def inspect_dropped(self, name: str, data_b64: str) -> dict:
        display = Path(str(name)).name
        try:
            data = base64.b64decode(data_b64, validate=True)
        except Exception as e:
            return {"ok": False, "error": f"could not read the file: {e}"}
        if len(data) > _qe.MAX_FILE_BYTES:
            return {"ok": False, "error": "file is too large"}
        try:
            with _PrivateTemp() as tmp:
                r = inspect_file(tmp.write(display, data))
        except Exception as e:
            return {"ok": False, "error": str(e)}
        if r.get("ok"):
            r["file"] = display
        return r

    def save_inspect(self, data: dict) -> dict:
        """Write the inspect dump to a text file the user can share back."""
        if self._window is None:
            return {"ok": False, "error": "no window"}
        name = (data.get("file") or "inspect").rsplit(".", 1)[0] + ".inspect.txt"
        result = self._window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=name,
            file_types=("Text (*.txt)", "All files (*.*)"),
        )
        if not result:
            return {"ok": False, "cancelled": True}
        out = Path(result if isinstance(result, str) else result[0])
        lines: list[str] = []
        lines.append(f"FILE: {data.get('file','?')}")
        lines.append("=" * 72)
        lines.append("\n--- WHAT CAME OUT (extracted rows) ---")
        for i, r in enumerate(data.get("extracted") or []):
            lines.append(f"[row {i+1}]")
            for c in data.get("columns", []):
                lines.append(f"  {c}: {r.get(c,'')}")
        lines.append("\n--- LABEL HITS (cells the extractor recognised) ---")
        for h in data.get("label_hits", []):
            lines.append(f"  table {h['table']} r{h['row']} c{h['col']}: '{h['cell']}'  ->  {h['maps_to']}")
        lines.append("\n--- TABLES (every cell) ---")
        for ti, tbl in enumerate(data.get("tables", []) or []):
            lines.append(f"\n[table {ti}]")
            for ri, row in enumerate(tbl):
                for ci, cell in enumerate(row):
                    if cell:
                        lines.append(f"  r{ri} c{ci}: {cell}")
        lines.append("\n--- RAW TEXT ---")
        lines.append(data.get("text") or "")
        out.write_text("\n".join(lines), encoding="utf-8")
        return {"ok": True, "path": str(out)}

    # ---- save a TSV of the rows the UI is currently showing --------------
    def copy_text(self, text: str) -> dict:
        try:
            _win_set_clipboard(text.replace("\r\n", "\n").replace("\n", "\r\n"))
            return {"ok": True}
        except Exception as e:
            return {"ok": False, "error": str(e)}

    def save_tsv(self, rows: list[dict]) -> dict:
        if self._window is None:
            return {"ok": False, "error": "no window"}
        result = self._window.create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="quotations.tsv",
            file_types=("TSV files (*.tsv)", "All files (*.*)"),
        )
        if not result:
            return {"ok": False, "cancelled": True}
        out = Path(result if isinstance(result, str) else result[0])
        lines = ["\t".join(COLUMNS)]
        for r in rows:
            data = r.get("row", {})
            lines.append("\t".join(_qe._safe_cell(str(data.get(c) or "")) for c in COLUMNS))
        out.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
        return {"ok": True, "path": str(out), "count": len(rows)}


# ---------------------------------------------------------------------------
# Embedded UI
# ---------------------------------------------------------------------------
HTML = r"""
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<!-- The page may not load or send anything over the network -->
<meta http-equiv="Content-Security-Policy"
      content="default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data:; font-src data:; connect-src 'none'; form-action 'none'; frame-src 'none'; base-uri 'none'" />
<title>QuoteGrid</title>
<style>
  :root {
    --bg: #0b0d12;
    --bg-2: #11141b;
    --panel: #161a23;
    --panel-2: #1c2130;
    --panel-3: #222838;
    --border: #262c3a;
    --border-2: #333b4f;
    --text: #e6e9f2;
    --muted: #8a93a6;
    --muted-2: #6b7489;
    --accent: #7c5cff;
    --accent-2: #3fd0c9;
    --danger: #ff5a7a;
    --ok: #4ade80;
    --warn: #fbbf24;
    --shadow: 0 10px 40px rgba(0,0,0,0.45);
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; height: 100%; background: var(--bg);
    color: var(--text); font-family: -apple-system, "Segoe UI", Inter, Roboto, sans-serif;
    font-size: 14px; }
  body {
    background:
      radial-gradient(900px 500px at 10% -10%, rgba(124,92,255,0.18), transparent 60%),
      radial-gradient(800px 500px at 110% 10%, rgba(63,208,201,0.14), transparent 60%),
      var(--bg);
    min-height: 100vh;
  }
  .app { max-width: 1760px; margin: 0 auto; padding: 22px 26px 60px; }

  header.bar { display:flex; align-items:center; justify-content:space-between; margin-bottom:16px; }
  .brand { display:flex; align-items:center; gap:12px; }
  .logo { width:40px; height:40px; display:block;
    filter: drop-shadow(0 6px 14px rgba(124,92,255,0.45)); }
  .brand h1 { font-size:16px; margin:0; letter-spacing:.2px; }
  .brand .sub { color:var(--muted); font-size:12px; margin-top:2px; }
  .pill { padding:6px 10px; border-radius:999px; font-size:12px;
    background:var(--panel); color:var(--muted); border:1px solid var(--border); }

  .drop {
    position:relative; border:1.5px dashed rgba(124,92,255,.5);
    background: linear-gradient(180deg, rgba(124,92,255,.06), rgba(63,208,201,.03)), var(--panel);
    border-radius:16px; padding:22px; display:flex; align-items:center; gap:18px;
    cursor:pointer; transition: transform .15s ease, border-color .15s ease, box-shadow .15s ease, background .2s ease;
  }
  .drop:hover { border-color: var(--accent); box-shadow: var(--shadow); }
  .drop.drag { border-color: var(--accent-2);
    background: linear-gradient(180deg, rgba(63,208,201,.12), rgba(124,92,255,.06)), var(--panel);
    transform: scale(1.005); }
  .drop .icon { width:48px; height:48px; border-radius:14px; flex-shrink:0;
    background: linear-gradient(135deg, rgba(124,92,255,.25), rgba(63,208,201,.25));
    display:grid; place-items:center; border:1px solid var(--border); }
  .drop .icon svg { width:24px; height:24px; }
  .drop .copy { flex:1; }
  .drop h2 { margin:0 0 2px; font-size:15px; font-weight:600; }
  .drop p  { margin:0; color:var(--muted); font-size:12.5px; }
  .drop .actions { display:flex; gap:8px; }

  .btn { display:inline-flex; align-items:center; gap:8px;
    padding:9px 14px; border-radius:10px; cursor:pointer;
    background: linear-gradient(135deg, var(--accent), #5b3df0);
    color:white; border:none; font-weight:600; font-size:12.5px; letter-spacing:.2px;
    box-shadow: 0 6px 18px rgba(124,92,255,.35);
    transition: transform .08s ease, box-shadow .15s ease, filter .15s ease; }
  .btn:hover { filter: brightness(1.07); box-shadow: 0 8px 22px rgba(124,92,255,.45); }
  .btn:active { transform: translateY(1px); }
  .btn.ghost { background:var(--panel-2); color:var(--text); border:1px solid var(--border); box-shadow:none; }
  .btn.ghost:hover { border-color:var(--accent); }
  .btn.ok   { background: linear-gradient(135deg, #22c55e, #16a34a); box-shadow:0 6px 18px rgba(34,197,94,.35); }
  .btn.warn { background: linear-gradient(135deg, #f59e0b, #d97706); box-shadow:0 6px 18px rgba(245,158,11,.35); }
  .btn.danger { background: linear-gradient(135deg, #ef4444, #b91c1c); box-shadow:0 6px 18px rgba(239,68,68,.35); }
  .btn[disabled] { opacity:.5; cursor:not-allowed; filter:none; }

  .toolbar { display:flex; align-items:center; gap:10px; margin:16px 2px 10px; }
  .toolbar .counts { color:var(--muted); font-size:12.5px; margin-right:auto; }
  .toolbar .counts b { color:var(--text); }
  .toolbar input.search { background:var(--panel-2); color:var(--text);
    border:1px solid var(--border); border-radius:10px; padding:8px 12px;
    font-size:12.5px; outline:none; width:220px; font-family:inherit; }
  .toolbar input.search:focus { border-color:var(--accent); }

  /* table card */
  .tcard {
    background:var(--panel); border:1px solid var(--border); border-radius:14px;
    box-shadow: var(--shadow); overflow:hidden;
  }
  .scroll { overflow:auto; max-height: calc(100vh - 290px); }
  table.data {
    border-collapse: separate; border-spacing: 0; width: max-content; min-width:100%;
    font-size: 12.5px;
  }
  table.data th, table.data td {
    padding: 8px 10px; border-bottom:1px solid var(--border); vertical-align: middle;
    white-space: nowrap;
  }
  table.data thead th {
    position: sticky; top: 0; z-index: 2;
    background: #171c28; color: #cfd4e4;
    font-weight: 600; letter-spacing: .3px; text-align: left;
    border-bottom:1px solid var(--border-2);
  }
  table.data thead th.sel { position: sticky; left: 0; z-index: 3; width: 44px; min-width: 44px; }
  table.data tbody td.sel {
    position: sticky; left: 0; z-index: 1; width: 44px; min-width: 44px;
    background: var(--panel); text-align: center;
  }
  table.data thead th.src { position: sticky; left: 44px; z-index: 3; min-width: 220px; }
  table.data tbody td.src {
    position: sticky; left: 44px; z-index: 1;
    background: var(--panel); border-right:1px solid var(--border);
    min-width: 220px; max-width: 240px; overflow:hidden; text-overflow:ellipsis;
    cursor: pointer; user-select: none;
  }
  table.data tbody tr:hover td { background: #1b2130; }
  table.data tbody tr:hover td.src, table.data tbody tr:hover td.sel { background:#1e2435; }

  /* row selection */
  input.chk { appearance: none; width: 16px; height: 16px; margin: 0; cursor: pointer;
    border: 1.5px solid var(--border-2); border-radius: 5px; background: var(--panel-2);
    display: inline-grid; place-items: center; transition: all .12s ease; vertical-align: middle; }
  input.chk:hover { border-color: var(--accent); }
  input.chk:checked { background: var(--accent); border-color: var(--accent); }
  input.chk:checked::after { content: ""; width: 4px; height: 8px; border: solid #fff;
    border-width: 0 2px 2px 0; transform: translateY(-1px) rotate(45deg); }
  input.chk:indeterminate { background: var(--accent); border-color: var(--accent); }
  input.chk:indeterminate::after { content: ""; width: 8px; height: 2px; background: #fff; }
  table.data tbody tr.selected td { background: rgba(124,92,255,.13); }
  table.data tbody tr.selected td.src, table.data tbody tr.selected td.sel { background: #231f3d; }
  table.data tbody tr.selected td.sel { box-shadow: inset 3px 0 0 var(--accent); }

  /* long-press feedback: a bar fills under the row while held, then a flash */
  table.data tbody tr { position: relative; }
  table.data tbody tr td.src::after { content: ""; position: absolute; left: 0; bottom: 0;
    height: 2px; width: 0; background: linear-gradient(90deg, var(--accent), var(--accent-2)); }
  table.data tbody tr.pressing td.src::after { width: 100%; transition: width .55s linear; }
  @keyframes copiedFlash { 0% { background: rgba(63,208,201,.35); } 100% { background: transparent; } }
  table.data tbody tr.copied td { animation: copiedFlash .7s ease-out; }

  .hint { color: var(--muted-2); font-size: 11.5px; margin: 8px 4px 0; }
  .hint kbd { font-family: inherit; font-size: 11px; padding: 1px 6px; border-radius: 5px;
    border: 1px solid var(--border-2); background: var(--panel-2); color: var(--muted); }
  table.data td.actions-cell { text-align:right; }
  table.data td input[type=text] {
    width: 100%; min-width: 120px; background: transparent; color:var(--text);
    border: 1px solid transparent; border-radius: 6px; padding: 4px 6px;
    font-size: 12.5px; outline: none; font-family: inherit;
  }
  table.data td input[type=text]:hover { border-color: var(--border-2); background: var(--panel-2); }
  table.data td input[type=text]:focus { border-color: var(--accent); background: #1f2434; }
  table.data td.empty input[type=text] { color: var(--muted-2); }
  table.data td.warn input[type=text] { border-color: rgba(251,191,36,.55); background: rgba(251,191,36,.08); color: var(--warn); }

  .fileTag { display:inline-flex; align-items:center; gap:6px; font-size:12px; }
  .fileTag .dot { width:8px; height:8px; border-radius:50%; background:var(--accent-2); flex:none; }
  .fileTag .name { overflow:hidden; text-overflow:ellipsis; max-width:170px; }
  .varBadge { display:inline-block; padding:1px 6px; border-radius:6px;
    background: rgba(124,92,255,.18); color:#c6b9ff; font-size:11px; margin-left:6px; }

  .rowDel { background: transparent; border:none; color: var(--muted); cursor:pointer;
    padding: 4px 6px; border-radius: 6px; font-size: 14px; }
  .rowDel:hover { color: var(--danger); background: rgba(255,90,122,.08); }

  .empty-state {
    padding: 44px 20px; text-align:center; color: var(--muted);
  }

  /* errors */
  .errs { background: rgba(255,90,122,.07); border:1px solid rgba(255,90,122,.25);
    border-radius: 12px; padding: 10px 12px; margin-top: 10px; font-size: 12.5px; }
  .errs b { color: var(--danger); }
  .errs ul { margin: 6px 0 0 18px; padding:0; }

  /* toast */
  .toast { position:fixed; right:20px; bottom:20px; z-index:99;
    background:#1a1f2c; border:1px solid var(--border); color:var(--text);
    padding:12px 16px; border-radius:12px; box-shadow:var(--shadow);
    display:flex; align-items:center; gap:10px;
    transform:translateY(20px); opacity:0; pointer-events:none;
    transition: all .2s ease; }
  .toast.show { transform:none; opacity:1; }
  .toast .mark { width:10px; height:10px; border-radius:50%; background:var(--accent-2); }
  .toast.err .mark { background: var(--danger); }
  .toast.ok  .mark { background: var(--ok); }

  .hidden { display:none !important; }

  /* Inspect modal */
  .modal { position: fixed; inset: 0; background: rgba(0,0,0,.6);
    display: flex; align-items: center; justify-content: center; z-index: 100;
    padding: 30px; }
  .modal-card { background: var(--panel); border: 1px solid var(--border);
    border-radius: 14px; width: 100%; max-width: 1100px; max-height: 92vh;
    display: flex; flex-direction: column; box-shadow: var(--shadow); }
  .modal-card header { padding: 14px 18px; border-bottom: 1px solid var(--border);
    display: flex; justify-content: space-between; align-items: flex-start; gap: 16px; }
  .modal-body { padding: 10px 18px 18px; overflow: auto; }
  .modal-body section { margin-top: 14px; }
  .modal-body h4 { margin: 0 0 6px; font-size: 12.5px; color: var(--accent-2); letter-spacing: .4px; text-transform: uppercase; }
  .inspBlock { background: #0d1017; border: 1px solid var(--border); border-radius: 10px;
    padding: 10px 12px; font-family: "Cascadia Code", Consolas, ui-monospace, monospace;
    font-size: 12px; color: #c7cde0; max-height: 260px; overflow: auto; white-space: pre-wrap; }
  .pre { background: #0d1017; border: 1px solid var(--border); border-radius: 10px;
    padding: 10px 12px; font-family: "Cascadia Code", Consolas, ui-monospace, monospace;
    font-size: 11.5px; color: #c7cde0; max-height: 280px; overflow: auto; white-space: pre-wrap; margin: 0; }
  .inspTable { border-collapse: collapse; margin: 6px 0 14px; }
  .inspTable td { border: 1px solid var(--border); padding: 4px 8px; font-size: 11.5px; vertical-align: top; }
  .inspTable td.lbl { color: #c6b9ff; background: rgba(124,92,255,.08); }
  .emptyTag { display: inline-block; margin: 2px 4px 2px 0; padding: 2px 8px;
    border-radius: 999px; background: rgba(251,191,36,.12); color: #fcd34d;
    font-size: 11.5px; }
</style>
</head>
<body>
  <div class="app">
    <header class="bar">
      <div class="brand">
        <img class="logo" src="__LOGO__" alt="" />
        <div>
          <h1>QuoteGrid</h1>
          <div class="sub">Drop one or many quotations · PDF · DOCX · TXT · CSV · XLSX · images (PNG / JPG)</div>
        </div>
      </div>
      <div class="pill" id="statusPill">idle</div>
    </header>

    <div class="drop" id="drop" tabindex="0" role="button" aria-label="Drop files or click to browse">
      <div class="icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" color="#cfd4e4">
          <path d="M12 3v12"/><path d="m7 10 5 5 5-5"/><path d="M5 21h14"/>
        </svg>
      </div>
      <div class="copy">
        <h2>Drop files here — or click to browse</h2>
        <p>Supports <span id="extList">—</span>. Many files at once are fine; each file becomes one row.</p>
      </div>
      <div class="actions">
        <button class="btn ghost" id="inspectBtn" type="button" title="Debug one file: see every cell the extractor saw">Inspect one…</button>
        <button class="btn" id="pickBtn" type="button">Choose files</button>
      </div>
    </div>

    <!-- Inspect modal -->
    <div id="inspect" class="modal hidden">
      <div class="modal-card">
        <header>
          <div>
            <h3 style="margin:0;">Inspect <span id="inspFile" style="color:var(--muted);font-weight:500;"></span></h3>
            <div style="color:var(--muted); font-size:12px; margin-top:3px;">Every cell the extractor saw, plus which cells it recognized as labels. Share the dump to improve matching.</div>
          </div>
          <div style="display:flex; gap:8px;">
            <button class="btn ghost" id="inspSaveBtn" type="button">Save dump…</button>
            <button class="btn ghost" id="inspCloseBtn" type="button">Close</button>
          </div>
        </header>
        <div class="modal-body">
          <section>
            <h4>Extracted rows</h4>
            <div id="inspExtracted" class="inspBlock"></div>
          </section>
          <section>
            <h4>Columns that came back EMPTY</h4>
            <div id="inspEmpty" class="inspBlock"></div>
          </section>
          <section>
            <h4>Label cells recognized</h4>
            <div id="inspLabels" class="inspBlock"></div>
          </section>
          <section>
            <h4>Tables (every cell)</h4>
            <div id="inspTables" class="inspBlock"></div>
          </section>
          <section>
            <h4>Raw text</h4>
            <pre id="inspText" class="pre"></pre>
          </section>
        </div>
      </div>
    </div>

    <div class="errs hidden" id="errsBox"><b>Could not read:</b><ul id="errsList"></ul></div>

    <div class="toolbar">
      <div class="counts">
        <b id="rowCount">0</b> row(s) ·
        <b id="fileCount">0</b> file(s) ·
        <b id="filledCount">0</b>/<b id="totalCells">0</b> cells filled
      </div>
      <input class="search" id="search" type="search" placeholder="Filter rows…" />
      <button class="btn ghost" id="clearBtn" type="button">Clear all</button>
      <button class="btn ghost" id="saveBtn"  type="button">Save .tsv</button>
      <button class="btn ghost" id="copyAllBtn" type="button" title="All rows, with the column header">Copy all</button>
      <button class="btn ok"    id="copySelBtn" type="button" disabled title="Selected rows only, no header — paste under existing Excel rows">Copy selected</button>
    </div>

    <div class="tcard">
      <div class="scroll" id="scroll">
        <table class="data">
          <thead>
            <tr>
              <th class="sel"><input type="checkbox" class="chk" id="selAll" title="Select all visible rows" /></th>
              <th class="src">Source file</th>
              <!-- data columns injected here -->
              <th></th>
            </tr>
          </thead>
          <tbody id="tbody"></tbody>
        </table>
        <div class="empty-state" id="emptyState">
          Nothing loaded yet. Drop files above to populate the table.
        </div>
      </div>
    </div>
    <div class="hint">
      <kbd>Hold</kbd> a row to copy it ·
      tick rows (or <kbd>Shift</kbd>+click for a range) then <kbd>Ctrl</kbd>+<kbd>C</kbd> or <b>Copy selected</b> ·
      row copies have no header, so they paste straight under your Excel rows
    </div>
  </div>

  <div class="toast" id="toast"><span class="mark"></span><span id="toastMsg">…</span></div>

<script>
  const COLUMNS = __COLUMNS__;
  const SUPPORTED = __EXTS__;

  const $ = (id) => document.getElementById(id);
  const drop = $("drop"), pickBtn = $("pickBtn");
  const tbody = $("tbody"), scroll = $("scroll"), emptyState = $("emptyState");
  const copyAllBtn = $("copyAllBtn"), copySelBtn = $("copySelBtn");
  const saveBtn = $("saveBtn"), clearBtn = $("clearBtn"), selAll = $("selAll");
  const statusPill = $("statusPill"), toast = $("toast"), toastMsg = $("toastMsg");
  const errsBox = $("errsBox"), errsList = $("errsList");
  const search = $("search");

  // State: list of row objects { id, file, variant, variant_count, row: {col: val} }
  let ROWS = [];
  const SELECTED = new Set();   // row ids
  let anchorId = null;          // last row clicked, for Shift+click ranges
  const LONG_PRESS_MS = 550;

  function rowsToTsv(rows, withHeader) {
    const lines = rows.map(r => COLUMNS.map(c => (r.row[c] || "").replace(/[\t\r\n]+/g, " ")).join("\t"));
    if (withHeader) lines.unshift(COLUMNS.join("\t"));
    return lines.join("\n");
  }

  async function copyText(txt) {
    try {
      const res = await window.pywebview.api.copy_text(txt);
      if (res && res.ok) return true;
    } catch (_) {}
    try { await navigator.clipboard.writeText(txt); return true; } catch (_) {}
    return false;
  }

  async function copyRows(rows, withHeader, label) {
    if (!rows.length) return showToast("no rows to copy", "err");
    const ok = await copyText(rowsToTsv(rows, withHeader));
    if (!ok) return showToast("could not access the clipboard", "err");
    showToast(label || `copied ${rows.length} row(s) — paste into Excel`, "ok");
    rows.forEach(r => {
      const tr = tbody.querySelector(`tr[data-id="${r.id}"]`);
      if (!tr) return;
      tr.classList.remove("copied"); void tr.offsetWidth; tr.classList.add("copied");
    });
  }

  function selectedRows() { return ROWS.filter(r => SELECTED.has(r.id)); }
  function visibleIds() { return [...tbody.querySelectorAll("tr")].map(tr => tr.dataset.id); }

  function updateSelUI() {
    tbody.querySelectorAll("tr").forEach(tr => {
      const on = SELECTED.has(tr.dataset.id);
      tr.classList.toggle("selected", on);
      const cb = tr.querySelector("td.sel input");
      if (cb) cb.checked = on;
    });
    const vis = visibleIds();
    const nSel = vis.filter(id => SELECTED.has(id)).length;
    selAll.checked = vis.length > 0 && nSel === vis.length;
    selAll.indeterminate = nSel > 0 && nSel < vis.length;
    copySelBtn.disabled = SELECTED.size === 0;
    copySelBtn.textContent = SELECTED.size ? `Copy selected (${SELECTED.size})` : "Copy selected";
  }

  function toggleRow(id, shift) {
    if (shift && anchorId && anchorId !== id) {
      const vis = visibleIds();
      const a = vis.indexOf(anchorId), b = vis.indexOf(id);
      if (a >= 0 && b >= 0) {
        const turnOn = !SELECTED.has(id) || SELECTED.has(anchorId);
        vis.slice(Math.min(a, b), Math.max(a, b) + 1).forEach(x => turnOn ? SELECTED.add(x) : SELECTED.delete(x));
        anchorId = id; updateSelUI(); return;
      }
    }
    SELECTED.has(id) ? SELECTED.delete(id) : SELECTED.add(id);
    anchorId = id;
    updateSelUI();
  }

  function attachLongPress(tr, r) {
    let timer = null, sx = 0, sy = 0;
    const cancel = () => {
      if (timer) { clearTimeout(timer); timer = null; }
      tr.classList.remove("pressing");
    };
    tr.addEventListener("pointerdown", (e) => {
      if (e.button !== 0 || e.target.closest(".rowDel, td.sel")) return;
      sx = e.clientX; sy = e.clientY;
      tr.classList.add("pressing");
      timer = setTimeout(() => {
        timer = null;
        tr.classList.remove("pressing");
        tr._longPressed = true;            // swallow the click that follows release
        // Holding a row that's part of a multi-selection copies the whole selection
        const sel = selectedRows();
        if (SELECTED.has(r.id) && sel.length > 1) copyRows(sel, false);
        else copyRows([r], false, `copied row: ${r.row["Barebone Name"] || r.file}`);
      }, LONG_PRESS_MS);
    });
    tr.addEventListener("pointermove", (e) => {
      if (timer && Math.hypot(e.clientX - sx, e.clientY - sy) > 6) cancel();  // dragging / selecting text
    });
    ["pointerup", "pointerleave", "pointercancel"].forEach(ev => tr.addEventListener(ev, cancel));
    tr.addEventListener("click", (e) => {
      if (tr._longPressed) { tr._longPressed = false; e.preventDefault(); e.stopPropagation(); }
    }, true);
  }

  $("extList").textContent = SUPPORTED.join(" · ");

  // Build table header dynamically (columns from Python)
  (function buildHead() {
    const tr = document.querySelector("table.data thead tr");
    const srcTh = tr.querySelector("th.src");
    const endTh = tr.lastElementChild;
    COLUMNS.forEach((col) => {
      const th = document.createElement("th");
      th.textContent = col;
      tr.insertBefore(th, endTh);
    });
  })();

  function showToast(msg, kind) {
    toastMsg.textContent = msg;
    toast.className = "toast show " + (kind || "");
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => toast.className = "toast", 2400);
  }
  function setStatus(text, busy) {
    statusPill.textContent = text;
    statusPill.style.color = busy ? "#fff" : "";
    statusPill.style.background = busy ? "linear-gradient(135deg,#7c5cff,#3fd0c9)" : "";
    statusPill.style.borderColor = busy ? "transparent" : "";
  }

  function refreshStats() {
    const files = new Set(ROWS.map(r => r.file));
    $("rowCount").textContent = ROWS.length;
    $("fileCount").textContent = files.size;
    const total = ROWS.length * COLUMNS.length;
    const filled = ROWS.reduce((n, r) =>
      n + COLUMNS.filter(c => (r.row[c] || "").trim() !== "").length, 0);
    $("filledCount").textContent = filled;
    $("totalCells").textContent = total;
    emptyState.style.display = ROWS.length ? "none" : "block";
  }

  function renderRow(r) {
    const tr = document.createElement("tr");
    tr.dataset.id = r.id;
    if (SELECTED.has(r.id)) tr.classList.add("selected");

    const sel = document.createElement("td");
    sel.className = "sel";
    const cb = document.createElement("input");
    cb.type = "checkbox"; cb.className = "chk"; cb.checked = SELECTED.has(r.id);
    cb.title = "Select row (Shift+click for a range)";
    cb.addEventListener("click", (e) => { e.preventDefault(); toggleRow(r.id, e.shiftKey); });
    sel.appendChild(cb);
    sel.addEventListener("click", (e) => { if (e.target === sel) toggleRow(r.id, e.shiftKey); });
    tr.appendChild(sel);

    const src = document.createElement("td");
    src.className = "src";
    src.title = "Click to select · hold to copy this row";
    src.addEventListener("click", (e) => toggleRow(r.id, e.shiftKey));
    src.innerHTML = `<span class="fileTag"><span class="dot"></span><span class="name" title="${esc(r.file)}">${esc(r.file)}</span>${
      r.variant ? `<span class="varBadge">variant ${r.variant}/${r.variant_count}</span>` : ""
    }</span>`;
    tr.appendChild(src);

    COLUMNS.forEach((col) => {
      const td = document.createElement("td");
      const v = r.row[col] || "";
      if (!v) td.classList.add("empty");
      const inp = document.createElement("input");
      inp.type = "text";
      inp.value = v;
      inp.placeholder = "—";
      const why = r.warn && r.warn[col];
      if (why && v) { td.classList.add("warn"); inp.title = "Check this value: " + why; }
      inp.addEventListener("input", (e) => {
        r.row[col] = e.target.value;
        td.classList.toggle("empty", !e.target.value);
        td.classList.remove("warn"); inp.title = "";   // the user has reviewed it
        if (r.warn) delete r.warn[col];
        refreshStats();
      });
      td.appendChild(inp);
      tr.appendChild(td);
    });

    const actions = document.createElement("td");
    actions.className = "actions-cell";
    const del = document.createElement("button");
    del.className = "rowDel";
    del.title = "Delete row";
    del.textContent = "✕";
    del.addEventListener("click", () => {
      ROWS = ROWS.filter(x => x.id !== r.id);
      SELECTED.delete(r.id);
      tr.remove();
      refreshStats();
      updateSelUI();
    });
    actions.appendChild(del);
    tr.appendChild(actions);

    attachLongPress(tr, r);
    return tr;
  }

  function esc(s) { return String(s).replace(/[&<>"']/g, m => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[m])); }

  function renderAll() {
    tbody.innerHTML = "";
    const q = (search.value || "").toLowerCase().trim();
    ROWS.forEach(r => {
      if (q) {
        const hay = r.file.toLowerCase() + " " + COLUMNS.map(c => r.row[c] || "").join(" ").toLowerCase();
        if (!hay.includes(q)) return;
      }
      tbody.appendChild(renderRow(r));
    });
    refreshStats();
    updateSelUI();
  }

  function pushRows(newRows) {
    // Each incoming row already has a unique id from Python
    ROWS = ROWS.concat(newRows);
    renderAll();
    // Scroll to the end so the user sees the newest batch
    requestAnimationFrame(() => { scroll.scrollTop = scroll.scrollHeight; });
  }

  function showErrs(errs) {
    if (!errs || !errs.length) { errsBox.classList.add("hidden"); return; }
    errsList.innerHTML = errs.map(e => `<li><b>${esc(e.file)}</b>: ${esc(e.error)}</li>`).join("");
    errsBox.classList.remove("hidden");
  }

  async function callApi(fn, ...args) {
    setStatus("working…", true);
    try {
      const res = await window.pywebview.api[fn](...args);
      setStatus("idle", false);
      return res;
    } catch (e) {
      setStatus("error", false);
      showToast("API error: " + e, "err");
      throw e;
    }
  }

  async function onPick() {
    const res = await callApi("pick_files");
    if (!res) return;
    if (res.cancelled) return;
    if (!res.ok) return showToast(res.error || "failed", "err");
    pushRows(res.rows);
    showErrs(res.errors);
    const n = res.rows.length;
    showToast(n ? `added ${n} row(s)` : "nothing extracted", n ? "ok" : "err");
  }

  // Drag and drop (multi)
  ["dragenter","dragover"].forEach(ev => drop.addEventListener(ev, (e) => {
    e.preventDefault(); e.stopPropagation(); drop.classList.add("drag");
  }));
  ["dragleave","drop"].forEach(ev => drop.addEventListener(ev, (e) => {
    e.preventDefault(); e.stopPropagation(); drop.classList.remove("drag");
  }));
  drop.addEventListener("drop", async (e) => {
    const files = Array.from(e.dataTransfer.files || []);
    if (!files.length) return;
    const okFiles = files.filter(f => SUPPORTED.some(ext => f.name.toLowerCase().endsWith(ext)));
    const bad = files.filter(f => !okFiles.includes(f));
    if (bad.length) showToast(`skipped ${bad.length} unsupported file(s)`, "err");
    if (!okFiles.length) return;

    setStatus("reading files…", true);
    // Encode files as base64 and send in one batch
    const payload = [];
    for (const f of okFiles) {
      const buf = await f.arrayBuffer();
      const bytes = new Uint8Array(buf);
      let bin = "";
      const chunk = 0x8000;
      for (let i = 0; i < bytes.length; i += chunk) {
        bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
      }
      payload.push({ name: f.name, data_b64: btoa(bin) });
    }
    const res = await callApi("extract_dropped", payload);
    if (!res.ok) return showToast(res.error || "failed", "err");
    pushRows(res.rows);
    showErrs(res.errors);
    const n = res.rows.length;
    showToast(n ? `added ${n} row(s) from ${okFiles.length} file(s)` : "nothing extracted", n ? "ok" : "err");
  });

  drop.addEventListener("click", onPick);
  pickBtn.addEventListener("click", (e) => { e.stopPropagation(); onPick(); });
  drop.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); onPick(); }
  });

  // Bulk actions
  copyAllBtn.addEventListener("click", () =>
    copyRows(ROWS, true, `copied ${ROWS.length} row(s) with header — paste into Excel`));
  copySelBtn.addEventListener("click", () => copyRows(selectedRows(), false));

  selAll.addEventListener("click", () => {
    const vis = visibleIds();
    const allOn = vis.length && vis.every(id => SELECTED.has(id));
    vis.forEach(id => allOn ? SELECTED.delete(id) : SELECTED.add(id));
    updateSelUI();
  });

  document.addEventListener("keydown", (e) => {
    const inField = e.target.matches("input[type=text], input[type=search], textarea");
    // Ctrl+C copies selected rows — unless you're copying text inside a cell
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "c" && SELECTED.size) {
      const hasTextSel = inField && e.target.selectionStart !== e.target.selectionEnd;
      if (!hasTextSel) { e.preventDefault(); copyRows(selectedRows(), false); }
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "a" && !inField && ROWS.length) {
      e.preventDefault(); visibleIds().forEach(id => SELECTED.add(id)); updateSelUI();
    }
    if (e.key === "Escape" && SELECTED.size && !inField) { SELECTED.clear(); updateSelUI(); }
  });

  saveBtn.addEventListener("click", async () => {
    if (!ROWS.length) return showToast("no rows to save", "err");
    const res = await callApi("save_tsv", ROWS);
    if (res.cancelled) return;
    if (!res.ok) return showToast(res.error || "save failed", "err");
    showToast(`saved ${res.count} row(s) → ${res.path}`, "ok");
  });

  clearBtn.addEventListener("click", () => {
    if (!ROWS.length) return;
    ROWS = []; SELECTED.clear(); anchorId = null;
    renderAll(); errsBox.classList.add("hidden");
    showToast("cleared", "ok");
  });

  search.addEventListener("input", renderAll);

  // ---- Inspect ----------------------------------------------------------
  const inspBtn = $("inspectBtn");
  const inspModal = $("inspect");
  const inspCloseBtn = $("inspCloseBtn");
  const inspSaveBtn  = $("inspSaveBtn");
  const inspFileLbl  = $("inspFile");
  let LAST_INSPECT = null;

  function renderInspect(data) {
    LAST_INSPECT = data;
    inspFileLbl.textContent = data.file || "";
    // extracted rows
    const ex = $("inspExtracted");
    ex.innerHTML = "";
    (data.extracted || []).forEach((r, i) => {
      const t = document.createElement("div");
      t.style.marginBottom = "6px";
      t.innerHTML = `<b style="color:#c6b9ff">row ${i+1}</b><br>` +
        data.columns.map(c => `<span style="color:${r[c]?'#4ade80':'#fbbf24'}">${esc(c)}</span>: ${esc(r[c] || '—')}`).join("<br>");
      ex.appendChild(t);
    });

    // empty columns across all rows
    const emptySet = new Set();
    (data.extracted || []).forEach(r => data.columns.forEach(c => { if (!r[c]) emptySet.add(c); }));
    $("inspEmpty").innerHTML = emptySet.size
      ? [...emptySet].map(c => `<span class="emptyTag">${esc(c)}</span>`).join("")
      : "<i>nothing — all columns filled</i>";

    // label hits
    const lbls = $("inspLabels");
    lbls.innerHTML = (data.label_hits || []).length
      ? (data.label_hits.map(h =>
          `table ${h.table} r${h.row} c${h.col}: "${esc(h.cell)}"  →  <b style="color:#4ade80">${esc(h.maps_to)}</b>`
        ).join("\n"))
      : "<i>no labels recognised — none of the cells matched any known label pattern</i>";

    // tables
    const tc = $("inspTables");
    tc.innerHTML = "";
    (data.tables || []).forEach((tbl, ti) => {
      const h = document.createElement("div");
      h.innerHTML = `<b style="color:#c6b9ff">[table ${ti}]</b>`;
      tc.appendChild(h);
      const t = document.createElement("table");
      t.className = "inspTable";
      tbl.forEach(row => {
        const tr = document.createElement("tr");
        row.forEach(cell => {
          const td = document.createElement("td");
          const isLabel = LAST_INSPECT.label_hits.some(x => x.table === ti && tbl[x.row]?.[x.col] === cell && x.cell === cell);
          if (isLabel) td.className = "lbl";
          td.textContent = cell || "·";
          tr.appendChild(td);
        });
        t.appendChild(tr);
      });
      tc.appendChild(t);
    });

    $("inspText").textContent = data.text || "";
    inspModal.classList.remove("hidden");
  }

  async function pickOneForInspect() {
    // Reuse native dialog (single file) — easiest path
    const res = await callApi("pick_files");
    if (!res || res.cancelled) return;
    if (!res.ok) return showToast(res.error || "failed", "err");
    // pick_files returns rows, not inspect; need the raw file -> use drop handler path
    // Simpler: prompt the user to drag into the drop zone while holding Alt? No - just do a second dialog.
    showToast("Drop a single file onto the window to inspect it", "err");
  }

  // Click inspect -> trigger a hidden file input (single file)
  const inspPicker = document.createElement("input");
  inspPicker.type = "file";
  inspPicker.accept = SUPPORTED.join(",");
  inspPicker.style.display = "none";
  document.body.appendChild(inspPicker);
  inspBtn.addEventListener("click", (e) => { e.stopPropagation(); inspPicker.click(); });
  inspPicker.addEventListener("change", async () => {
    const f = inspPicker.files && inspPicker.files[0];
    if (!f) return;
    const buf = await f.arrayBuffer();
    const bytes = new Uint8Array(buf);
    let bin = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    const res = await callApi("inspect_dropped", f.name, btoa(bin));
    inspPicker.value = "";
    if (!res.ok) return showToast(res.error || "inspect failed", "err");
    renderInspect(res);
  });

  inspCloseBtn.addEventListener("click", () => inspModal.classList.add("hidden"));
  inspModal.addEventListener("click", (e) => { if (e.target === inspModal) inspModal.classList.add("hidden"); });
  inspSaveBtn.addEventListener("click", async () => {
    if (!LAST_INSPECT) return;
    const res = await callApi("save_inspect", LAST_INSPECT);
    if (res.cancelled) return;
    if (!res.ok) return showToast(res.error || "save failed", "err");
    showToast("saved → " + res.path, "ok");
  });

  refreshStats();
</script>
</body>
</html>
"""


def build_html() -> str:
    cols_js = json.dumps(COLUMNS, ensure_ascii=False)
    exts_js = json.dumps(list(SUPPORTED_EXTS))
    logo = _ASSETS / "logo_128.png"
    logo_uri = ("data:image/png;base64," + base64.b64encode(logo.read_bytes()).decode()) if logo.is_file() else ""
    return (HTML.replace("__COLUMNS__", cols_js)
                .replace("__EXTS__", exts_js)
                .replace("__LOGO__", logo_uri))


def main() -> int:
    _block_network()
    _purge_old_temp()
    api = Api()
    window = webview.create_window(
        title=APP_NAME,
        html=build_html(),
        js_api=api,
        width=1500,
        height=900,
        min_size=(1100, 680),
        background_color="#0b0d12",
    )
    api._window = window
    # private_mode: no cookies/cache/local storage kept on disk; no dev tools
    webview.start(gui="edgechromium" if sys.platform == "win32" else None,
                  private_mode=True, debug=False)
    _purge_old_temp()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
