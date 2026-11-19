"""
Cricket Tournament Payment Tracker
----------------------------------
A tiny web app that uses an Excel workbook as its ONLY data store.

* Every edit made in the web UI is written straight into the matching cell of
  the .xlsx file (formulas are never overwritten).
* Formulas in the workbook are evaluated on the server, so balances / totals in
  the UI always match what Excel will show.
* If someone edits the .xlsx in Excel and saves it, the web UI picks the change
  up automatically.

Run:  python app.py            (then open http://localhost:8000)
Env:  TRACKER_XLSX   path to the workbook (default: data/2025_Cricket_Tournaments.xlsx)
      Accounts: see auth.py. Captains can read + write, viewers can only read.
      EDIT_PIN       only used once: protects creation of the first (captain) account
      PORT           default 8000
"""
import datetime as dt
from copy import copy
import hashlib
import hmac
import json
import secrets
import glob
import os
import re
import shutil
import tempfile
import threading
import time

import openpyxl
from flask import Flask, abort, jsonify, request, send_file, send_from_directory
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from pycel import ExcelCompiler

import auth

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.abspath(os.environ.get(
    "TRACKER_XLSX", os.path.join(BASE, "data", "2025_Cricket_Tournaments.xlsx")))
EDIT_PIN = os.environ.get("EDIT_PIN", "").strip()
SETTINGS_FILE = os.path.join(os.path.dirname(XLSX), "crictrack-settings.json")
_fails = {}   # ip -> [count, locked_until]


def _settings():
    try:
        with open(SETTINGS_FILE) as f:
            return json.load(f)
    except Exception:
        return {}


def _write_settings(d):
    tmp = SETTINGS_FILE + ".tmp"
    with open(tmp, "w") as f:
        json.dump(d, f, indent=2)
    os.replace(tmp, SETTINGS_FILE)


def _hash_pin(pin, salt):
    return hashlib.pbkdf2_hmac("sha256", pin.encode(), bytes.fromhex(salt), 200_000).hex()


def _pin_source():
    if EDIT_PIN:
        return "env"
    return "file" if _settings().get("pin_hash") else ""


def _pin_required():
    return bool(_pin_source())


def _pin_ok(given):
    given = str(given or "").strip()
    if EDIT_PIN:
        return hmac.compare_digest(given, EDIT_PIN)
    st = _settings()
    if not st.get("pin_hash"):
        return True
    return hmac.compare_digest(_hash_pin(given, st["salt"]), st["pin_hash"])
BACKUP_DIR = os.path.join(os.path.dirname(XLSX), "backups")
KEEP_BACKUPS = 40
BACKUP_EVERY_SECONDS = 10 * 60

STATIC_DIR = os.path.join(BASE, "static")
app = Flask(__name__, static_folder=None)     # pages are served by the routes below (login required)
auth.init(app, os.path.dirname(XLSX), STATIC_DIR, lambda: _tournament_sheets())
lock = threading.RLock()
state = {"mtime": None, "wb": None, "values": {}, "layout": {}, "last_backup": 0.0}


# --------------------------------------------------------------------------- #
# Loading & formula evaluation
# --------------------------------------------------------------------------- #
def _is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def _norm(v):
    return v.strip().lower() if isinstance(v, str) else None


def _jsonable(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return {"t": "d", "v": v.strftime("%Y-%m-%d")}
    if isinstance(v, bool):
        return {"t": "b", "v": v}
    if isinstance(v, (int, float)):
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        elif isinstance(v, float):
            v = round(v, 2)
        return {"t": "n", "v": v}
    if v is None:
        return {"t": "e", "v": None}
    return {"t": "s", "v": str(v)}


def _evaluate_all(wb, xc):
    """Evaluate every formula cell with pycel. Returns {sheet: {ref: value}}."""
    out = {}
    for ws in wb.worksheets:
        vals = {}
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    try:
                        v = xc.evaluate(f"'{ws.title}'!{c.coordinate}")
                        if hasattr(v, "item"):          # numpy scalar
                            v = v.item()
                        if isinstance(v, str) and v.startswith("#"):
                            v = None
                    except Exception:
                        v = None
                    vals[c.coordinate] = v
        out[ws.title] = vals
    return out


def _load(force=False):
    """(Re)load the workbook if the file changed on disk."""
    with lock:
        mtime = os.path.getmtime(XLSX)
        if not force and state["mtime"] == mtime and state["wb"] is not None:
            return
        wb = openpyxl.load_workbook(XLSX)
        upgraded = _upgrade_sheets(wb)
        if upgraded:
            _save(wb)
            mtime = os.path.getmtime(XLSX)
            print(f"  Upgraded to fee-based match costs: {', '.join(upgraded)}")
        state["wb"] = wb
        state["xc"] = ExcelCompiler(filename=XLSX)
        state["values"] = _evaluate_all(wb, state["xc"])
        state["layout"] = {ws.title: _detect_layout(ws) for ws in wb.worksheets}
        state["mtime"] = mtime


# --------------------------------------------------------------------------- #
# Layout detection: figure out where the tables live on each sheet
# --------------------------------------------------------------------------- #
def _detect_layout(ws):
    grid = {}
    for row in ws.iter_rows():
        for c in row:
            if not isinstance(c, MergedCell) and not _is_blank(c.value):
                grid[(c.row, c.column)] = c.value
    g = lambda r, c: grid.get((r, c))
    used = set()

    def claim(r1, r2, c1, c2):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                used.add((r, c))

    def row_blank(r, cols):
        return all(_is_blank(g(r, c)) for c in cols)

    # merged ranges: info box + titles
    info, merged_titles = None, {}
    for mr in ws.merged_cells.ranges:
        tl = ws.cell(mr.min_row, mr.min_col).value
        if isinstance(tl, str) and tl.strip().lower().startswith("tournament"):
            lines = [l.strip() for l in tl.strip().splitlines() if l.strip()]
            facts = []
            for l in lines:
                k, _, v = l.partition(":")
                facts.append({"k": k.strip(), "v": v.strip().rstrip("-").rstrip("/").strip()})
            info = {"ref": ws.cell(mr.min_row, mr.min_col).coordinate, "facts": facts,
                    "box": [mr.min_row, mr.max_row, mr.min_col, mr.max_col]}
            claim(mr.min_row, mr.max_row, mr.min_col, mr.max_col)
        else:
            merged_titles[(mr.min_row, mr.min_col)] = (mr.min_col, mr.max_col)

    info_cols = set()
    if info:
        info_cols = set(range(info["box"][2], info["box"][3] + 1))

    players, lineups, installments = None, [], []

    for (r, c), v in sorted(grid.items()):
        n = _norm(v)
        # ---- per-match line-up blocks: SN | Player | Cost | Status
        if n == "sn" and _norm(g(r, c + 1)) == "player" and _norm(g(r, c + 2)) == "cost":
            cols = list(range(c, c + 4))
            title_ref = ws.cell(r - 1, c).coordinate if r > 1 else None
            title = g(r - 1, c) or ""
            rows, rr = [], r + 1
            while rr <= ws.max_row and not row_blank(rr, cols):
                rows.append(rr)
                rr += 1
            lineups.append({
                "title": str(title).strip(), "title_ref": title_ref, "header_row": r,
                "cols": [get_column_letter(x) for x in cols], "rows": rows,
                "player_rows": [x for x in rows if isinstance(g(x, c), (int, float))],
            })
            claim(r - 1, rr - 1, c, c + 3)
        # ---- main player table: SN | Player | Match 1 ...
        elif n == "sn" and _norm(g(r, c + 1)) == "player" and players is None:
            cols, cc = [], c
            while not _is_blank(g(r, cc)):
                cols.append(cc)
                cc += 1
            rows, rr = [], r + 1
            while rr <= ws.max_row and not row_blank(rr, cols):
                rows.append(rr)
                rr += 1
            headers = [str(g(r, x)).strip() for x in cols]
            players = {
                "header_row": r, "cols": [get_column_letter(x) for x in cols], "headers": headers,
                "rows": rows,
                "player_rows": [x for x in rows if isinstance(g(x, c), (int, float))],
                "name_col": get_column_letter(c + 1), "sn_col": get_column_letter(c),
                "match_cols": [get_column_letter(x) for x, h in zip(cols, headers)
                               if re.match(r"^match\s*\d", h, re.I)],
            }
            claim(r, rr - 1, cols[0], cols[-1])
        # ---- instalment table: Sno | Name | ...
        elif n == "sno" and _norm(g(r, c + 1)) == "name":
            cols, cc, gap = [], c, 0
            while cc not in info_cols and cc <= ws.max_column:
                has_data = any(not _is_blank(g(x, cc)) for x in range(r, r + 40))
                if _is_blank(g(r, cc)) and not has_data:
                    gap += 1
                    if gap >= 2:
                        break
                else:
                    gap = 0
                cols.append(cc)
                cc += 1
            while cols and _is_blank(g(r, cols[-1])) and \
                    not any(not _is_blank(g(x, cols[-1])) for x in range(r + 1, r + 40)):
                cols.pop()
            rows, rr = [], r + 1
            while rr <= ws.max_row and not row_blank(rr, cols):
                rows.append(rr)
                rr += 1
            installments.append({
                "header_row": r, "cols": [get_column_letter(x) for x in cols],
                "headers": [str(g(r, x)).strip() if not _is_blank(g(r, x)) else "" for x in cols],
                "rows": rows, "name_col": get_column_letter(c + 1),
            })
            claim(r, rr - 1, cols[0], cols[-1])

    # ---- everything else: group leftover cells into small "ledger" blocks
    left = [k for k in grid if k not in used]
    blocks, seen = [], set()
    for start in sorted(left):
        if start in seen:
            continue
        comp, stack = [], [start]
        seen.add(start)
        while stack:
            r, c = stack.pop()
            comp.append((r, c))
            for (r2, c2) in left:
                if (r2, c2) not in seen and abs(r2 - r) <= 2 and abs(c2 - c) <= 3:
                    seen.add((r2, c2))
                    stack.append((r2, c2))
        rs = [x[0] for x in comp]
        cs = [x[1] for x in comp]
        r1, r2, c1, c2 = min(rs), max(rs), min(cs), max(cs)
        # extend one column right so a label-only row can receive a value
        blocks.append({
            "rows": list(range(r1, r2 + 1)),
            "cols": [get_column_letter(x) for x in range(c1, c2 + 1)],
        })

    return {"info": info, "players": players, "lineups": lineups,
            "installments": installments, "ledger": blocks}


# --------------------------------------------------------------------------- #
# Serialisation
# --------------------------------------------------------------------------- #
def _sheet_payload(ws):
    vals = state["values"].get(ws.title, {})
    cells = {}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c, MergedCell):
                continue
            v = c.value
            if _is_blank(v):
                continue
            if isinstance(v, str) and v.startswith("="):
                d = _jsonable(vals.get(c.coordinate))
                d["f"] = v
            else:
                d = _jsonable(v)
            cells[c.coordinate] = d
    merged_tail = []
    for mr in ws.merged_cells.ranges:
        for row in ws.iter_rows(min_row=mr.min_row, max_row=mr.max_row,
                                min_col=mr.min_col, max_col=mr.max_col):
            for c in row:
                if isinstance(c, MergedCell):
                    merged_tail.append(c.coordinate)
    return {"name": ws.title, "cells": cells, "layout": state["layout"][ws.title],
            "locked_cells": merged_tail}


def _version():
    return str(state["mtime"])


# --------------------------------------------------------------------------- #
# Writing
# --------------------------------------------------------------------------- #
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
NUM_RE = re.compile(r"^-?\d+(\.\d+)?$")


def _coerce(raw, cell):
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return raw
    s = str(raw).strip()
    if s == "":
        return None
    if DATE_RE.match(s):
        return dt.datetime.strptime(s, "%Y-%m-%d")
    t = s.replace(",", "").replace("₹", "").strip()
    if NUM_RE.match(t):
        f = float(t)
        return int(f) if f.is_integer() else f
    return s


def _backup():
    now = time.time()
    if now - state["last_backup"] < BACKUP_EVERY_SECONDS:
        return
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    name = os.path.splitext(os.path.basename(XLSX))[0]
    shutil.copy2(XLSX, os.path.join(BACKUP_DIR, f"{name}-{stamp}.xlsx"))
    state["last_backup"] = now
    olds = sorted(glob.glob(os.path.join(BACKUP_DIR, f"{name}-*.xlsx")))
    for f in olds[:-KEEP_BACKUPS]:
        try:
            os.remove(f)
        except OSError:
            pass


def _save(wb):
    try:
        wb.calculation.fullCalcOnLoad = True   # Excel recalculates when opened
    except Exception:
        pass
    _backup()
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(XLSX))
    os.close(fd)
    try:
        wb.save(tmp)
        os.replace(tmp, XLSX)   # atomic: the Excel file is never half-written
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)


def _check_pin(given=None):
    """Every change to the workbook goes through here.
    Admins: anything. Captains: only their tournament (the 'sheet' in the request). Players: nothing."""
    if request.path == "/api/tournaments" and request.method == "POST":
        return auth.require_create()
    b = request.get_json(force=True, silent=True) or {}
    sheet = b.get("sheet")
    if not sheet or sheet in (PAYMENTS_SHEET, ORG_SHEET):
        return auth.require_admin()
    auth.require_write(sheet)


def _tournament_sheets():
    _load()
    return [t for t, lay in state["layout"].items()
            if lay.get("players") and t not in (PAYMENTS_SHEET, ORG_SHEET)]


def _who():
    u = auth.current_user()
    return u["name"] if u else ""


# --------------------------------------------------------------------------- #
# Routes
# --------------------------------------------------------------------------- #
@app.get("/")
def index():
    return send_from_directory(STATIC_DIR, "ledger.html")


@app.get("/favicon.ico")
@app.get("/icon.svg")
@app.get("/apple-touch-icon.png")
def icons():
    return send_from_directory(STATIC_DIR, request.path.lstrip("/"), max_age=7 * 24 * 3600)


@app.get("/sheets")
def sheets_view():
    return send_from_directory(STATIC_DIR, "index.html")


@app.get("/api/data")
def api_data():
    _load()
    with lock:
        wb = state["wb"]
        return jsonify({
            "file": os.path.basename(XLSX), "version": _version(),
            **auth.auth_fields(),
            "saved_at": dt.datetime.fromtimestamp(state["mtime"]).strftime("%d %b %Y, %I:%M %p"),
            "sheets": [dict(_sheet_payload(ws), can_edit=auth.can_write(ws.title) if ws.title not in (PAYMENTS_SHEET, ORG_SHEET) else auth.is_admin())
                       for ws in wb.worksheets
                       if auth.is_admin() or (ws.title in _tournament_sheets() and auth.can_read(ws.title))],
        })


@app.get("/api/version")
def api_version():
    _load()
    return jsonify({"version": _version()})


@app.post("/api/check-pin")
def api_check_pin():
    _check_pin()
    return jsonify({"ok": True})


@app.post("/api/cells")
def api_cells():
    """Body: {"sheet": "...", "changes": [{"ref": "D5", "value": "Yes"}, ...]}"""
    _check_pin()
    body = request.get_json(force=True, silent=True) or {}
    sheet, changes = body.get("sheet"), body.get("changes") or []
    if not changes:
        abort(400, "No changes")
    _load()
    with lock:
        wb = state["wb"]
        if sheet not in wb.sheetnames:
            abort(404, "Unknown sheet")
        ws = wb[sheet]
        for ch in changes:
            ref = str(ch.get("ref", "")).upper()
            try:
                col, row = coordinate_from_string(ref)
                column_index_from_string(col)
            except Exception:
                abort(400, f"Bad cell reference {ref}")
            cell = ws[ref]
            if isinstance(cell, MergedCell):
                abort(400, f"{ref} is part of a merged cell")
            if isinstance(cell.value, str) and cell.value.startswith("="):
                abort(400, f"{ref} holds a formula and is calculated automatically")
            val = _coerce(ch.get("value"), cell)
            cell.value = val
            try:
                addr = f"'{ws.title}'!{ref}"
                if addr.replace("'", "") not in state["xc"].cell_map:
                    state["xc"].evaluate(addr)      # register the cell with the formula engine
                state["xc"].set_value(addr, "" if val is None else val)
            except Exception:
                state["xc_dirty"] = True
            if isinstance(val, dt.datetime) and not cell.is_date:
                cell.number_format = "dd-mmm-yy"
        try:
            _save(wb)
        except PermissionError:
            _load(force=True)
            abort(423, "The Excel file is open/locked by another program (close it in Excel and retry).")
        state["mtime"] = os.path.getmtime(XLSX)
        if state.pop("xc_dirty", False):
            state["xc"] = ExcelCompiler(filename=XLSX)
        state["values"] = _evaluate_all(wb, state["xc"])
        state["layout"][ws.title] = _detect_layout(ws)
        return jsonify({"ok": True, "version": _version(),
                        "saved_at": dt.datetime.fromtimestamp(state["mtime"]).strftime("%d %b %Y, %I:%M %p"),
                        "sheet": dict(_sheet_payload(state["wb"][sheet]), can_edit=True)})


# --------------------------------------------------------------------------- #
# Ledger: one row per player per tournament, payments log, played toggles
# --------------------------------------------------------------------------- #
PAYMENTS_SHEET = "Payments"
PAY_TYPES = ("Advance", "Match fees", "Other")
PAY_HEADERS = ["ID", "Date", "Tournament", "Player", "Amount", "Mode", "Note",
               "Added to cell", "Recorded at", "Type", "Recorded by"]
COUNTIF_RE = re.compile(
    r"COUNTIF\(\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*,\s*\$?([A-Z]+)\$?(\d+)\s*\)", re.I)
SUM_RE = re.compile(r"^=\s*SUM\(\s*\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)\s*\)\s*$", re.I)


def _val(ws, ref):
    v = ws[ref].value
    if isinstance(v, str) and v.startswith("="):
        v = state["values"].get(ws.title, {}).get(ref)
    return v


def _n(v):
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return 0
    return round(float(v), 2)


def _ledger_cols(p):
    def find(*pats):
        for pat in pats:
            for h, c in zip(p["headers"], p["cols"]):
                if re.search(pat, h, re.I):
                    return c
        return None
    return (find(r"^to be paid$", r"^amount due$", r"^total$", r"^due"),
            find(r"paid so far|so far paid", r"^paid", r"install"))


def _xc_set(ws, ref, val):
    try:
        addr = f"'{ws.title}'!{ref}"
        if addr.replace("'", "") not in state["xc"].cell_map:
            state["xc"].evaluate(addr)
        state["xc"].set_value(addr, "" if val is None else val)   # pycel ignores None
    except Exception:
        state["xc_dirty"] = True


def _commit(touched_sheets):
    wb = state["wb"]
    try:
        _save(wb)
    except PermissionError:
        _load(force=True)
        abort(423, "The Excel file is open/locked by another program (close it in Excel and retry).")
    state["mtime"] = os.path.getmtime(XLSX)
    if state.pop("xc_dirty", False):
        state["xc"] = ExcelCompiler(filename=XLSX)
    state["values"] = _evaluate_all(wb, state["xc"])
    for s in touched_sheets:
        state["layout"][s] = _detect_layout(wb[s])


MATCH_HDR_RE = re.compile(r"^match\s*(\d+)\s*(?:[-–:]\s*(.*))?$", re.I)
COST_ROW_LABEL = "Cost per player"


def _cost_row(ws, p):
    """Row holding per-match cost (only on tournament sheets created by this app)."""
    for r in p["rows"]:
        if r not in p["player_rows"] and _norm(ws[f"{p['name_col']}{r}"].value) == COST_ROW_LABEL.lower():
            return r
    return None


def _match_titles(layout, ws=None):
    p, lus = layout["players"], layout["lineups"]
    crow = _cost_row(ws, p) if ws is not None else None
    fr = _footer_rows(ws, p) if ws is not None else {}
    titles = []
    for i, col in enumerate(p["match_cols"]):
        hdr = p["headers"][p["cols"].index(col)]
        m = MATCH_HDR_RE.match(hdr or "")
        label = (m.group(2) or "").strip() if m else ""
        if label:
            t = hdr
        elif i < len(lus) and lus[i]["title"]:
            t = lus[i]["title"]
        else:
            t = hdr or f"Match {i + 1}"
        item = {"col": col, "short": f"M{i + 1}", "title": t, "label": label}
        if crow:
            item["cost"] = _n(ws[f"{col}{crow}"].value)
        if ws is not None and fr.get("rate"):
            item.update(rate=_n(_val(ws, f"{col}{fr['rate']}")), team=int(_n(ws[f"{col}{fr['team']}"].value)) if fr.get("team") else DEFAULT_TEAM,
                        pool=_n(_val(ws, f"{col}{fr['pool']}")) if fr.get("pool") else 0,
                        played=int(_n(_val(ws, f"{col}{fr['count']}"))) if fr.get("count") else 0,
                        cost=_n(_val(ws, f"{col}{fr['cost']}")) if fr.get("cost") else 0)
        titles.append(item)
    return titles


def _amounts(text):
    return [int(x.replace(",", "")) for x in re.findall(r"\d[\d,]*", str(text or ""))]


def _tour_settings(info):
    adv = est_min = est_max = 0
    for f in (info or {}).get("facts", []):
        k = f["k"].lower()
        nums = _amounts(f["v"])
        if not nums:
            continue
        if k.startswith("advance"):
            adv = nums[0]
        elif k.startswith("expected"):
            est_min, est_max = nums[0], (nums[1] if len(nums) > 1 else nums[0])
    return adv, est_min, est_max


ORG_SHEET = "Organiser Payments"
ORG_HEADERS = ["ID", "Date", "Tournament", "Amount", "Mode", "Note", "Added to cell", "Date cell", "Recorded at",
               "Recorded by"]
ORG_BAL_RE = re.compile(
    r"^=\s*(?:SUM\(\s*)?\$?([A-Z]+)\$?(\d+)\s*\)?\s*-\s*SUM\(\s*\$?([A-Z]+)\$?(\d+)\s*"
    r"(?::\s*\$?([A-Z]+)\$?(\d+))?\s*\)\s*$", re.I)


def _is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def _org_block(ws):
    """Find the organiser block: a cell like  =Fee - SUM(paid1:paidN)  outside the player tables."""
    lay = state["layout"].get(ws.title) or {}
    for blk in lay.get("ledger", []):
        for r in blk["rows"]:
            for col in blk["cols"]:
                v = ws[f"{col}{r}"].value
                m = ORG_BAL_RE.match(v) if _is_formula(v) else None
                if not m:
                    continue
                c1, r1 = m.group(3).upper(), int(m.group(4))
                c2, r2 = (m.group(5) or c1).upper(), int(m.group(6) or r1)
                slots = [ws.cell(rr, cc).coordinate
                         for rr in range(r1, r2 + 1)
                         for cc in range(column_index_from_string(c1), column_index_from_string(c2) + 1)]
                return {"fee_ref": f"{m.group(1).upper()}{m.group(2)}", "slots": slots, "bal_ref": f"{col}{r}"}
    return None


def _org_summary(ws):
    org = _org_block(ws)
    if not org:
        return {"exists": False, "fee": 0, "paid": 0, "balance": 0, "fee_editable": True}
    fee = _n(_val(ws, org["fee_ref"]))
    paid = round(sum(_n(_val(ws, x)) for x in org["slots"]), 2)
    return {"exists": True, "fee": fee, "paid": paid, "balance": round(fee - paid, 2),
            "fee_editable": not _is_formula(ws[org["fee_ref"]].value),
            "fee_ref": org["fee_ref"], "bal_ref": org["bal_ref"]}


def _write_org_block(ws, col, row, fee, collected_range=None):
    """Standard organiser block: fee, 4 payment slots (amount + date), balance, collected."""
    from openpyxl.styles import Font, PatternFill
    L, V, Dc = col, col + 1, col + 2
    vl = get_column_letter(V)
    bold, head = Font(bold=True), PatternFill("solid", fgColor="F4EAD2")
    ws.cell(row, L, "Organiser fee").font = bold
    ws.cell(row, V, fee or None).number_format = "#,##0"
    ws.cell(row, Dc, "Date paid").font = Font(bold=True, color="888888")
    for i in range(4):
        ws.cell(row + 1 + i, L, f"Paid to organiser {i + 1}")
        ws.cell(row + 1 + i, V).number_format = "#,##0"
        ws.cell(row + 1 + i, Dc).number_format = "dd-mmm-yy"
    ws.cell(row + 5, L, "Balance to organiser").font = bold
    ws.cell(row + 5, V, f"={vl}{row}-SUM({vl}{row + 1}:{vl}{row + 4})").number_format = "#,##0"
    if collected_range:
        ws.cell(row + 6, L, "Collected from players").font = bold
        ws.cell(row + 6, V, f"=SUM({collected_range})").number_format = "#,##0"
        ws.cell(row + 7, L, "Kitty in hand").font = bold
        ws.cell(row + 7, V, f"={vl}{row + 6}-SUM({vl}{row + 1}:{vl}{row + 4})").number_format = "#,##0"
    for r in range(row, row + (8 if collected_range else 6)):
        for c in (L, V, Dc):
            ws.cell(r, c).fill = head
    for c, w in ((L, 22), (V, 12), (Dc, 12)):
        cur = ws.column_dimensions[get_column_letter(c)].width
        if not cur or cur < w:
            ws.column_dimensions[get_column_letter(c)].width = w


def _ensure_org_block(ws, fee):
    """Create an organiser block on a sheet that doesn't have one (below the info box, or to the right)."""
    lay = state["layout"].get(ws.title) or {}
    info, p = lay.get("info"), lay.get("players")
    col, row = None, None
    cand = []
    if info:
        cand.append((info["box"][2], info["box"][1] + 2))
    cand.append((ws.max_column + 2, 2))
    for c, r in cand:
        if all(_is_blank(ws.cell(rr, cc).value) for rr in range(r, r + 9) for cc in range(c, c + 3)):
            col, row = c, r
            break
    if col is None:
        col, row = ws.max_column + 2, 2
    rng = None
    if p:
        _, paid_c = _ledger_cols(p)
        if paid_c and p["player_rows"]:
            rng = f"{paid_c}{p['player_rows'][0]}:{paid_c}{p['player_rows'][-1]}"
    _write_org_block(ws, col, row, fee, rng)
    state["xc_dirty"] = True


def _org_ws(wb):
    from openpyxl.styles import Font, PatternFill
    if ORG_SHEET in wb.sheetnames:
        ows = wb[ORG_SHEET]
    else:
        ows = wb.create_sheet(ORG_SHEET)
        ows.append(ORG_HEADERS)
        ows.freeze_panes = "A2"
    for i, w in enumerate([9, 12, 28, 12, 12, 34, 14, 12, 20, 16], start=1):
        c = ows.cell(1, i)
        if _is_blank(c.value):            # older sheets get the new "Recorded by" column
            c.value = ORG_HEADERS[i - 1]
        ows.column_dimensions[get_column_letter(i)].width = w
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="A8771C")
    return ows


def _ledger_payload():
    wb = state["wb"]
    tournaments, entries = [], []
    for ws in wb.worksheets:
        lay = state["layout"].get(ws.title) or {}
        p = lay.get("players")
        if not p or ws.title in (PAYMENTS_SHEET, ORG_SHEET) or not auth.can_read(ws.title):
            continue
        due_c, paid_c = _ledger_cols(p)
        info = lay.get("info") or {}
        tname = next((f["v"] for f in info.get("facts", []) if f["k"].lower().startswith("tournament")), "")
        matches = _match_titles(lay, ws)
        free = sum(1 for r in p["player_rows"] if _is_blank(ws[f"{p['name_col']}{r}"].value))
        cust_c, extra_c = _hdr_col(p, "Custom amount"), _hdr_col(p, "Extra")
        adv, est_min, est_max = _tour_settings(info)
        tournaments.append({"sheet": ws.title, "title": tname, "matches": matches,
                            "facts": info.get("facts", []), "due_col": due_c, "paid_col": paid_c,
                            "editable_matches": _is_v2(ws, p), "shared_cost": _is_v2(ws, p),
                            "custom_ok": bool(_hdr_col(p, "Custom amount")), "free_slots": free,
                            "access": auth.access(ws.title), "can_edit": auth.can_write(ws.title),
                            "advance": adv, "est_min": est_min, "est_max": est_max,
                            "has_info": bool(info), "org": _org_summary(ws)})
        for r in p["player_rows"]:
            name = ws[f"{p['name_col']}{r}"].value
            if _is_blank(name):
                continue
            played = []
            for m in matches:
                c = ws[f"{m['col']}{r}"]
                raw = c.value
                v = str(_val(ws, c.coordinate) or "").strip().lower()
                if isinstance(raw, str) and raw.startswith("="):
                    mode = "lineup" if COUNTIF_RE.search(raw) else None
                else:
                    mode = "cell"
                played.append({"col": m["col"], "v": "yes" if v == "yes" else ("no" if v == "no" else ""),
                               "editable": bool(mode)})
            due = _n(_val(ws, f"{due_c}{r}")) if due_c else 0
            paid = _n(_val(ws, f"{paid_c}{r}")) if paid_c else 0
            ent = {"sheet": ws.title, "row": r, "name": str(name).strip(),
                   "played": played, "due": due, "paid": paid, "balance": round(paid - due, 2)}
            if cust_c:
                cv = ws[f"{cust_c}{r}"].value
                ent["custom"] = cv if isinstance(cv, (int, float)) and not isinstance(cv, bool) else None
                ent["base"] = round(sum(m.get("cost", 0) for m, pl in zip(matches, played) if pl["v"] == "yes")
                                    + _n(ws[f"{extra_c}{r}"].value), 2)
            entries.append(ent)
    payments = []
    if PAYMENTS_SHEET in wb.sheetnames:
        pws = wb[PAYMENTS_SHEET]
        for row in pws.iter_rows(min_row=2, values_only=True):
            if not row or _is_blank(row[0]):
                continue
            d = row[1]
            payments.append({
                "id": str(row[0]),
                "date": d.strftime("%Y-%m-%d") if isinstance(d, (dt.date, dt.datetime)) else (str(d or "")),
                "sheet": row[2] or "", "player": row[3] or "", "amount": _n(row[4]),
                "mode": row[5] or "", "note": row[6] or "", "cell": row[7] or "",
                "recorded": row[8].strftime("%d %b %Y, %I:%M %p") if isinstance(row[8], dt.datetime) else str(row[8] or ""),
                "type": (row[9] if len(row) > 9 else None) or "",
                "by": (row[10] if len(row) > 10 else None) or "",
            })
    payments = [x for x in payments if auth.can_read(x["sheet"])]
    payments.reverse()
    org_payments = []
    if ORG_SHEET in wb.sheetnames:
        for row in wb[ORG_SHEET].iter_rows(min_row=2, values_only=True):
            if not row or _is_blank(row[0]):
                continue
            d = row[1]
            org_payments.append({
                "id": str(row[0]), "date": d.strftime("%Y-%m-%d") if isinstance(d, (dt.date, dt.datetime)) else str(d or ""),
                "sheet": row[2] or "", "amount": _n(row[3]), "mode": row[4] or "", "note": row[5] or "",
                "cell": row[6] or "",
                "by": (row[9] if len(row) > 9 else None) or "",
                "recorded": row[8].strftime("%d %b %Y, %I:%M %p") if len(row) > 8 and isinstance(row[8], dt.datetime) else "",
            })
        org_payments = [x for x in org_payments if auth.can_read(x["sheet"])]
        org_payments.reverse()
    return {"ok": True, "file": os.path.basename(XLSX), "version": _version(),
            **auth.auth_fields(),
            "saved_at": dt.datetime.fromtimestamp(state["mtime"]).strftime("%d %b %Y, %I:%M %p"),
            "tournaments": tournaments, "entries": entries, "payments": payments,
            "org_payments": org_payments}


def _player_ctx(sheet, row):
    wb = state["wb"]
    if sheet not in wb.sheetnames:
        abort(404, "Unknown tournament")
    ws = wb[sheet]
    p = (state["layout"].get(sheet) or {}).get("players")
    try:
        row = int(row)
    except Exception:
        abort(400, "Bad row")
    if not p or row not in p["player_rows"]:
        abort(404, "Player not found in this tournament")
    name = ws[f"{p['name_col']}{row}"].value
    if _is_blank(name):
        abort(404, "Player not found")
    return ws, p, row, str(name).strip()


@app.get("/api/ledger")
def api_ledger():
    _load()
    with lock:
        return jsonify(_ledger_payload())


@app.post("/api/played")
def api_played():
    """Body: {sheet, row, col, value: "Yes"|"No"} — marks whether a player played a match."""
    _check_pin()
    b = request.get_json(force=True, silent=True) or {}
    want = "Yes" if str(b.get("value", "")).lower() == "yes" else "No"
    _load()
    with lock:
        ws, p, row, name = _player_ctx(b.get("sheet"), b.get("row"))
        col = str(b.get("col", "")).upper()
        if col not in p["match_cols"]:
            abort(400, "Not a match column")
        cell = ws[f"{col}{row}"]
        raw = cell.value
        if isinstance(raw, str) and raw.startswith("="):
            # This sheet derives Yes/No from the match line-up: edit the line-up instead.
            m = COUNTIF_RE.search(raw)
            if not m:
                abort(400, "This cell is calculated by a formula that can't be toggled from the web.")
            c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
            pcol = column_index_from_string(c1)
            refs = [ws.cell(rr, pcol) for rr in range(r1, r2 + 1)]
            same = [c for c in refs if str(c.value or "").strip().lower() == name.lower()]
            cost_col = pcol + 1
            hdr = _norm(ws.cell(r1 - 1, cost_col).value)
            if want == "Yes" and not same:
                free = next((c for c in refs if _is_blank(c.value)), None)
                if free is None:
                    abort(409, "This match line-up already has every slot filled on the sheet. Mark someone else as not played first.")
                free.value = name
                _xc_set(ws, free.coordinate, name)
                if hdr == "cost" and _is_blank(ws.cell(free.row, cost_col).value):
                    costs = [ws.cell(rr, cost_col).value for rr in range(r1, r2 + 1)]
                    costs = [c for c in costs if isinstance(c, (int, float))]
                    if costs:
                        common = max(set(costs), key=costs.count)
                        ws.cell(free.row, cost_col).value = common
                        _xc_set(ws, ws.cell(free.row, cost_col).coordinate, common)
            elif want == "No":
                for c in same:
                    c.value = None
                    _xc_set(ws, c.coordinate, None)
                    if hdr == "cost":
                        for off in (1, 2):
                            cc = ws.cell(c.row, pcol + off)
                            if not (isinstance(cc.value, str) and cc.value.startswith("=")):
                                cc.value = None
                                _xc_set(ws, cc.coordinate, None)
        else:
            cell.value = want
            _xc_set(ws, cell.coordinate, want)
        _commit([ws.title])
        return jsonify(_ledger_payload())


def _payments_ws(wb):
    from openpyxl.styles import Font, PatternFill
    if PAYMENTS_SHEET in wb.sheetnames:
        pws = wb[PAYMENTS_SHEET]
    else:
        pws = wb.create_sheet(PAYMENTS_SHEET)
        pws.append(PAY_HEADERS)
        pws.freeze_panes = "A2"
    for i, w in enumerate([9, 12, 28, 20, 11, 10, 34, 16, 20, 12, 16], start=1):
        c = pws.cell(1, i)
        if _is_blank(c.value):            # older Payments sheets get the new "Type" column
            c.value = PAY_HEADERS[i - 1]
        pws.column_dimensions[get_column_letter(i)].width = w
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1E6B47")
    return pws


def _paid_target(ws, paid_col, row):
    """The input cell a payment should be added to, so the sheet's own totals stay right."""
    c = ws[f"{paid_col}{row}"]
    if not (isinstance(c.value, str) and c.value.startswith("=")):
        return c
    m = SUM_RE.match(c.value)
    if m and int(m.group(2)) == row and int(m.group(4)) == row:
        a, z = column_index_from_string(m.group(1)), column_index_from_string(m.group(3))
        for ci in range(z, a - 1, -1):
            t = ws.cell(row, ci)
            if not (isinstance(t.value, str) and t.value.startswith("=")):
                return t
    abort(400, "Couldn't find where to add this payment on the sheet.")


def _parse_amount(raw, allow_zero=False):
    try:
        amount = float(str(raw if raw is not None else "").replace(",", "").replace("₹", "").strip() or "0")
    except ValueError:
        abort(400, "Enter a valid amount")
    if amount == 0 and not allow_zero:
        abort(400, "Amount can't be zero")
    return int(amount) if amount.is_integer() else round(amount, 2)


def _pay_type(t):
    t = str(t or "").strip().lower()
    return next((x for x in PAY_TYPES if x.lower() == t), "Other" if t else "")


def _apply_payment(ws, p, row, name, amount, pdate, mode, note, ptype=""):
    """Add `amount` to the player's paid cell and log it on the Payments sheet."""
    _, paid_c = _ledger_cols(p)
    if not paid_c:
        abort(400, "This sheet has no 'Paid' column")
    target = _paid_target(ws, paid_c, row)
    cur = target.value if isinstance(target.value, (int, float)) else 0
    newv = cur + amount
    newv = int(newv) if float(newv).is_integer() else round(newv, 2)
    target.value = newv
    if not target.number_format or target.number_format == "General":
        target.number_format = "#,##0"
    _xc_set(ws, target.coordinate, newv)
    pws = _payments_ws(state["wb"])
    ids = [int(str(v)[1:]) for (v,) in pws.iter_rows(min_row=2, max_col=1, values_only=True)
           if isinstance(v, str) and re.match(r"^P\d+$", v)]
    pid = f"P{(max(ids) + 1 if ids else 1):04d}"
    pws.append([pid, pdate, ws.title, name, amount, str(mode or "").strip()[:20],
                str(note or "").strip()[:200], target.coordinate, dt.datetime.now(), _pay_type(ptype) or None, _who() or None])
    pws.cell(pws.max_row, 2).number_format = "dd-mmm-yy"
    pws.cell(pws.max_row, 5).number_format = "#,##0"
    pws.cell(pws.max_row, 9).number_format = "dd-mmm-yy hh:mm"
    return pid


# ---------------------------- tournaments & players ------------------------ #
BAD_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")
PLAYER_CAPACITY = 30


def _style_header(c):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="1E6B47")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = Border(bottom=Side(style="thin", color="CCCCCC"))


# --------------------------------------------------------------------------- #
# Tournament sheet layout (v2): each match costs  rate x players-per-match  (e.g. 560 x 12),
# shared equally by the players who actually played it; optional custom amount per player.
# --------------------------------------------------------------------------- #
DEFAULT_RATE, DEFAULT_TEAM = 560, 12
V2_LABELS = (("count", "total players"), ("rate", "rate per player"), ("team", "players per match"),
             ("pool", "match cost"), ("cost", "cost per player"))
V2_TAIL = ["Matches Played", "Extra", "Custom amount", "Total", "Paid So far", "Balance"]
V1_TAIL = ["Matches Played", "Extra", "Total", "Paid So far", "Balance"]


def _footer_rows(ws, p):
    out = {}
    for r in p["rows"]:
        if r in p["player_rows"]:
            continue
        lab = _norm(ws[f"{p['name_col']}{r}"].value) or ""
        for k, pref in V2_LABELS:
            if k not in out and lab.startswith(pref):
                out[k] = r
    return out


def _is_v2(ws, p):
    return bool(p) and "rate" in _footer_rows(ws, p)


def _hdr_col(p, name):
    return next((c for c, h in zip(p["cols"], p["headers"]) if h.strip().lower() == name.lower()), None)


def _write_tournament_sheet(ws, spec):
    """Write a complete tournament sheet. spec = {lines, matches:[{label,rate,team}], players:[{row,name,
    yes:[..], extra, custom, paid}], fee, org_slots:[(amount, date)]}. Returns the cell map."""
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    n = len(spec["matches"])
    hr, first = 2, 3
    last = first + PLAYER_CAPACITY - 1
    foot = last + 1
    r_rate, r_team, r_pool, r_cost = foot + 1, foot + 2, foot + 3, foot + 4
    mcols = [get_column_letter(4 + i) for i in range(n)]
    ci = 4 + n
    played_c, extra_c, cust_c, total_c, paid_c, bal_c = (get_column_letter(ci + k) for k in range(6))
    headers = ["SN", "Player"] + [f"Match {i + 1}" + (f" - {m['label']}" if m.get("label") else "")
                                  for i, m in enumerate(spec["matches"])] + V2_TAIL
    for k, h in enumerate(headers):
        _style_header(ws.cell(hr, 2 + k, h))
    ws.row_dimensions[hr].height = 32
    thin = Side(style="thin", color="E3E3E3")
    money = "#,##0;[Red]-#,##0"
    for r in range(first, last + 1):
        ws[f"B{r}"] = r - first + 1
        ws[f"{played_c}{r}"] = f'=COUNTIF({mcols[0]}{r}:{mcols[-1]}{r},"Yes")'
        shares = "+".join(f'IF({m}{r}="Yes",{m}${r_cost},0)' for m in mcols)
        ws[f"{total_c}{r}"] = f"=IF(ISNUMBER({cust_c}{r}),{cust_c}{r},{shares})+SUM({extra_c}{r})"
        ws[f"{bal_c}{r}"] = f"={paid_c}{r}-{total_c}{r}"
        for col in (extra_c, cust_c, total_c, paid_c, bal_c):
            ws[f"{col}{r}"].number_format = money
        for k in range(len(headers)):
            ws.cell(r, 2 + k).border = Border(bottom=thin)
        for m in mcols:
            ws[f"{m}{r}"].alignment = Alignment(horizontal="center")
    for pl in spec.get("players", []):
        r = pl["row"]
        ws[f"C{r}"] = pl["name"]
        for m, v in zip(mcols, pl.get("yes", [])):
            ws[f"{m}{r}"] = v if v not in ("", None) else None
        for col, key in ((extra_c, "extra"), (cust_c, "custom"), (paid_c, "paid")):
            v = pl.get(key)
            ws[f"{col}{r}"] = v if isinstance(v, (int, float)) and not isinstance(v, bool) else None
    ws[f"C{foot}"] = "Total Players"
    for m in mcols:
        ws[f"{m}{foot}"] = f'=COUNTIF({m}{first}:{m}{last},"Yes")'
    for col in (extra_c, total_c, paid_c, bal_c):
        ws[f"{col}{foot}"] = f"=SUM({col}{first}:{col}{last})"
        ws[f"{col}{foot}"].number_format = money
    info_col, org_row = ci + 7, hr + 11
    fee_ref = f"{get_column_letter(info_col + 1)}{org_row}"
    for m, mm in zip(mcols, spec["matches"]):
        ws[f"{m}{r_team}"] = mm.get("team", DEFAULT_TEAM)
    _write_fee_formulas(ws, mcols, foot, r_rate, r_team, r_pool, r_cost, fee_ref)
    grey, yellow = PatternFill("solid", fgColor="F1F1EC"), PatternFill("solid", fgColor="FFF4CC")
    for r in (foot, r_rate, r_team, r_pool, r_cost):
        for k in range(len(headers)):
            ws.cell(r, 2 + k).fill = yellow if r in (r_rate, r_team) else grey
            ws.cell(r, 2 + k).font = Font(bold=True)
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{mcols[0]}{first}:{mcols[-1]}{last}")
    ic = ws.cell(hr, info_col, "\n".join(spec["lines"]))
    ic.alignment = Alignment(wrap_text=True, vertical="top")
    ic.fill = PatternFill("solid", fgColor="EAF4EE")
    ws.merge_cells(start_row=hr, start_column=info_col, end_row=hr + 9, end_column=info_col + 3)
    _write_org_block(ws, info_col, org_row, spec.get("fee") or 0, f"{paid_c}{first}:{paid_c}{last}")
    slots = [ws.cell(org_row + 1 + i, info_col + 1).coordinate for i in range(4)]
    dates = [ws.cell(org_row + 1 + i, info_col + 2).coordinate for i in range(4)]
    for (amt, d), sref, dref in zip(spec.get("org_slots", []), slots, dates):
        if amt:
            ws[sref] = amt
        if d:
            ws[dref] = d
            ws[dref].number_format = "dd-mmm-yy"
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 5
    ws.column_dimensions["C"].width = 26
    for m in mcols:
        ws.column_dimensions[m].width = 12
    for col in (played_c, extra_c, cust_c, total_c, paid_c, bal_c):
        ws.column_dimensions[col].width = 12
    for k in range(4):
        cur = ws.column_dimensions[get_column_letter(info_col + k)].width
        if not cur or cur < 12:
            ws.column_dimensions[get_column_letter(info_col + k)].width = 12
    ws.freeze_panes = "D3"
    return {"paid_c": paid_c, "slots": slots, "dates": dates}


def _write_fee_formulas(ws, mcols, foot, r_rate, r_team, r_pool, r_cost, fee_ref):
    """Match cost = tournament fee / number of matches; per player = match cost / players per match
    (e.g. 42000 / 6 = 7000; 7000 / 12 = 583). The match cost is shared by the players who actually played."""
    col, row = coordinate_from_string(fee_ref)
    fee = f"${col}${row}"
    n = len(mcols)
    ws[f"C{r_rate}"] = "Rate per player (fee ÷ matches ÷ players)"
    ws[f"C{r_team}"] = "Players per match"
    ws[f"C{r_pool}"] = "Match cost (fee ÷ matches)"
    ws[f"C{r_cost}"] = "Cost per player (shared)"
    for m in mcols:
        ws[f"{m}{r_pool}"] = f"=IF(ISNUMBER({fee}),{fee}/{n},0)"
        ws[f"{m}{r_rate}"] = f"=IF({m}{r_team}>0,ROUND({m}{r_pool}/{m}{r_team},0),0)"
        ws[f"{m}{r_cost}"] = f"=IF({m}{foot}>0,ROUND({m}{r_pool}/{m}{foot},0),0)"
        for r in (r_rate, r_pool, r_cost):
            ws[f"{m}{r}"].number_format = "#,##0"


def _fee_rule_line(fee, n, team):
    if fee and n:
        per = fee / n
        return (f"Per Match: Rs {fee:,.0f} / {n} = Rs {per:,.0f}; / {team} players = Rs {round(per / team):,} each, "
                f"shared by those who play")
    return "Per Match: tournament fee / matches, shared by those who play"


def _info_fee(ws, lay):
    for f in ((lay or {}).get("info") or {}).get("facts", []):
        if f["k"].lower().startswith("fee"):
            nums = _amounts(f["v"])
            if nums:
                return nums[0]
    return 0


def _refresh_rule_line(ws):
    """Keep the 'Per Match:' line in the tournament box in step with the fee and players per match."""
    lay = state["layout"].get(ws.title) or {}
    p, info = lay.get("players"), lay.get("info")
    if not p or not info:
        return
    fr, org = _footer_rows(ws, p), _org_block(ws)
    fee = _n(ws[org["fee_ref"]].value) if org else 0
    team = int(_n(ws[f"{p['match_cols'][0]}{fr['team']}"].value) or DEFAULT_TEAM) if fr.get("team") else DEFAULT_TEAM
    _set_info_line(ws, info, "Per Match", _fee_rule_line(fee, len(p["match_cols"]), team).split(": ", 1)[1])


def _needs_fee_upgrade(ws):
    """v2 sheets (match cost = rate x players) -> fee-based match cost."""
    lay = _detect_layout(ws)
    p = lay.get("players")
    if not p or not _is_v2(ws, p):
        return False
    fr = _footer_rows(ws, p)
    v = ws[f"{p['match_cols'][0]}{fr['pool']}"].value if fr.get("pool") else None
    return isinstance(v, str) and "*" in v


def _fee_upgrade(ws):
    lay = _detect_layout(ws)
    state["layout"][ws.title] = lay
    p = lay["players"]
    fr = _footer_rows(ws, p)
    org = _org_block(ws)
    if not org:
        _ensure_org_block(ws, _info_fee(ws, lay))
        lay = _detect_layout(ws)
        state["layout"][ws.title] = lay
        org = _org_block(ws)
    fv = ws[org["fee_ref"]].value
    if not _is_formula(fv) and not _n(fv):
        # no fee entered: keep today's total (rate x players for every match) as the fee, so nobody drops to Rs 0
        ws[org["fee_ref"]] = sum(_n(ws[f"{c}{fr['rate']}"].value) * (_n(ws[f"{c}{fr['team']}"].value) or DEFAULT_TEAM)
                                 for c in p["match_cols"]) or None
    _write_fee_formulas(ws, p["match_cols"], fr["count"], fr["rate"], fr["team"], fr["pool"], fr["cost"], org["fee_ref"])
    for r in (fr["rate"], fr["team"]):
        for c in p["cols"]:
            ws[f"{c}{r}"].font = copy(ws[f"{c}{fr['count']}"].font)
    _refresh_rule_line(ws)
    return ws.title


def _needs_upgrade(ws):
    lay = _detect_layout(ws)
    p = lay.get("players")
    if not p or [h.strip() for h in p["headers"][-5:]] != V1_TAIL or _is_v2(ws, p):
        return False
    return _cost_row(ws, p) is not None


def _upgrade_v1(wb, ws):
    """Rebuild an app-made v1 sheet (fixed cost per player) in the v2 shared-cost layout,
    keeping players in the same rows and re-pointing the payment logs."""
    title = ws.title
    lay = _detect_layout(ws)
    state["layout"][title] = lay
    p, info = lay["players"], lay.get("info") or {}
    hdr = dict(zip(p["cols"], p["headers"]))
    extra_c, paid_c = _hdr_col(p, "Extra"), _hdr_col(p, "Paid So far")
    matches = []
    for col in p["match_cols"]:
        m = MATCH_HDR_RE.match(hdr[col] or "")
        matches.append({"label": (m.group(2) or "").strip() if m else "", "rate": DEFAULT_RATE, "team": DEFAULT_TEAM})
    players = []
    for r in p["player_rows"]:
        name = ws[f"{p['name_col']}{r}"].value
        if _is_blank(name):
            continue
        lit = lambda ref: ws[ref].value if not _is_formula(ws[ref].value) else None
        players.append({"row": r, "name": str(name).strip(), "yes": [lit(f"{c}{r}") for c in p["match_cols"]],
                        "extra": lit(f"{extra_c}{r}"), "paid": lit(f"{paid_c}{r}")})
    lines = [l for l in str(ws[info["ref"]].value or "").splitlines() if l.strip()] if info else [f"Tournament: {title}"]
    lines = [l for l in lines if not l.lower().startswith("per match")]
    lines.insert(min(len(lines), 5), _fee_rule_line(_info_fee(ws, lay), len(matches), DEFAULT_TEAM))
    org = _org_block(ws)
    fee, org_slots, old_slots, old_dates = 0, [], [], []
    if org:
        fv = ws[org["fee_ref"]].value
        fee = fv if isinstance(fv, (int, float)) else 0
        for sref in org["slots"]:
            c = ws[sref]
            d = ws.cell(c.row, c.column + 1).value
            org_slots.append((c.value if isinstance(c.value, (int, float)) else None,
                              d if isinstance(d, (dt.date, dt.datetime)) else None))
            old_slots.append(sref)
            old_dates.append(ws.cell(c.row, c.column + 1).coordinate)
    if not fee:
        crow = _cost_row(ws, p)
        fee = _info_fee(ws, lay) or (sum(_n(ws[f"{c}{crow}"].value) for c in p["match_cols"]) * DEFAULT_TEAM if crow else 0)
    if len(org_slots) > 4:                       # fold any extra instalments into the 4th slot
        tot = sum(a or 0 for a, _ in org_slots[3:])
        org_slots = org_slots[:3] + [(tot, org_slots[3][1])]
    idx = wb.sheetnames.index(title)
    nws = wb.create_sheet(title[:25] + "~v2", idx)
    cm = _write_tournament_sheet(nws, {"lines": lines, "matches": matches, "players": players,
                                       "fee": fee, "org_slots": org_slots})
    wb.remove(ws)
    nws.title = title
    # re-point the logs so Undo keeps working
    if PAYMENTS_SHEET in wb.sheetnames:
        for row in wb[PAYMENTS_SHEET].iter_rows(min_row=2):
            if len(row) > 7 and row[2].value == title and isinstance(row[7].value, str):
                col, r = coordinate_from_string(row[7].value)
                if col == paid_c:
                    row[7].value = f"{cm['paid_c']}{r}"
    if ORG_SHEET in wb.sheetnames:
        for row in wb[ORG_SHEET].iter_rows(min_row=2):
            if len(row) > 7 and row[2].value == title:
                if row[6].value in old_slots:
                    i = min(old_slots.index(row[6].value), 3)
                    row[6].value = cm["slots"][i]
                    if row[7].value:
                        row[7].value = cm["dates"][i]
    return title


def _upgrade_sheets(wb):
    sheets = [ws for ws in wb.worksheets if ws.title not in (PAYMENTS_SHEET, ORG_SHEET)]
    v1 = [ws for ws in sheets if _needs_upgrade(ws)]
    v2 = [ws for ws in sheets if ws not in v1 and _needs_fee_upgrade(ws)]
    if not v1 and not v2:
        return []
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    shutil.copy2(XLSX, os.path.join(BACKUP_DIR, f"before-fee-based-cost-upgrade-{stamp}.xlsx"))
    done = [_upgrade_v1(wb, ws) for ws in v1]
    for ws in list(wb.worksheets):
        if ws.title in done:
            state["layout"][ws.title] = _detect_layout(ws)
            _refresh_rule_line(ws)
    return done + [_fee_upgrade(ws) for ws in v2]


@app.post("/api/tournaments")
def api_new_tournament():
    """Create a new tournament sheet (v2 layout).
    Body: {sheet, title, ground, fee, team_size, matches, rate, team, rates:[..], match_names:[..], advance, est_min, est_max}"""
    _check_pin()
    b = request.get_json(force=True, silent=True) or {}
    sheet = re.sub(r"\s+", " ", str(b.get("sheet") or "")).strip()
    title = str(b.get("title") or "").strip() or sheet
    ground = str(b.get("ground") or "").strip()
    if not sheet:
        abort(400, "Give the tournament a short name (used as the sheet name)")
    if len(sheet) > 31 or BAD_SHEET_CHARS.search(sheet) or sheet.endswith("~v2"):
        abort(400, "Short name must be 31 characters or less and can't contain [ ] : * ? / \\")
    try:
        n = int(b.get("matches") or 0)
    except ValueError:
        n = 0
    if not 1 <= n <= 15:
        abort(400, "Number of matches must be between 1 and 15")
    rate = _parse_amount(b.get("rate", b.get("cost", DEFAULT_RATE)), allow_zero=True) or DEFAULT_RATE
    try:
        team = int(b.get("team") or DEFAULT_TEAM)
    except ValueError:
        team = DEFAULT_TEAM
    if not 1 <= team <= 30:
        abort(400, "Players per match must be between 1 and 30")
    rates = [_parse_amount(c, allow_zero=True) for c in (b.get("rates") or b.get("costs") or [])]
    rates = [(rates[i] if i < len(rates) and rates[i] else rate) for i in range(n)]
    fee = _parse_amount(b.get("fee"), allow_zero=True)
    if not fee:
        abort(400, "Enter the tournament fee: match costs are worked out from it (fee ÷ matches ÷ players).")
    try:
        squad = int(b.get("team_size") or 0)
    except ValueError:
        squad = 0
    names = [str(x or "").strip()[:60] for x in (b.get("match_names") or [])]
    advance = _parse_amount(b.get("advance"), allow_zero=True)
    est_min = _parse_amount(b.get("est_min"), allow_zero=True)
    est_max = _parse_amount(b.get("est_max"), allow_zero=True) or est_min
    if est_min and est_max < est_min:
        est_min, est_max = est_max, est_min
    matches = [{"label": names[i] if i < len(names) else "", "rate": rates[i], "team": team} for i in range(n)]
    lines = [f"Tournament: {title}"]
    if ground:
        lines.append(f"Ground: {ground}")
    if fee:
        lines.append(f"Fee: Rs {fee:,}/-")
    if squad:
        lines.append(f"Team: {squad} Players")
    if fee and squad:
        lines.append(f"Per Head: Rs {round(fee / squad):,}/-")
    lines.append(_fee_rule_line(fee, n, team))
    lines.append(f"Matches: {n}")
    if advance:
        lines.append(f"Advance: Rs {advance:,}/- per player")
    if est_min:
        lines.append(f"Expected cost: Rs {est_min:,} - {est_max:,}/- per player" if est_max != est_min
                     else f"Expected cost: Rs {est_min:,}/- per player")
    _load()
    with lock:
        wb = state["wb"]
        if sheet.lower() in [x.lower() for x in wb.sheetnames] or sheet.lower() in (PAYMENTS_SHEET.lower(), ORG_SHEET.lower()):
            abort(409, f"A sheet called '{sheet}' already exists")
        logs = [wb.sheetnames.index(x) for x in (PAYMENTS_SHEET, ORG_SHEET) if x in wb.sheetnames]
        ws = wb.create_sheet(sheet, min(logs) if logs else len(wb.sheetnames))
        _write_tournament_sheet(ws, {"lines": lines, "matches": matches, "players": [], "fee": fee})
        state["xc_dirty"] = True
        _commit([sheet])
        if not auth.is_admin():
            auth.grant(auth.current_user()["username"], sheet, "captain")
        out = _ledger_payload()
        out["message"] = f"Tournament '{sheet}' created with {n} matches. Now add players."
        out["created"] = sheet
        return jsonify(out)


@app.post("/api/custom")
def api_custom_amount():
    """Set (or clear with amount "") a player's custom amount, which replaces their match-based amount.
    Body: {sheet, row, amount}"""
    _check_pin()
    b = request.get_json(force=True, silent=True) or {}
    raw = str(b.get("amount") if b.get("amount") is not None else "").strip()
    _load()
    with lock:
        ws, p, row, name = _player_ctx(b.get("sheet"), b.get("row"))
        col = _hdr_col(p, "Custom amount")
        if not col:
            abort(400, "This is an older tournament sheet without custom amounts.")
        val = None if raw == "" else _parse_amount(raw, allow_zero=True)
        ws[f"{col}{row}"] = val
        ws[f"{col}{row}"].number_format = "#,##0"
        _xc_set(ws, f"{col}{row}", val)
        _commit([ws.title])
        out = _ledger_payload()
        out["message"] = (f"{name}: custom amount ₹{val:,}" if val is not None
                          else f"{name}: back to the match-based amount")
        return jsonify(out)


@app.post("/api/players")
def api_add_players():
    """Add players to a tournament. Body: {sheet, players: [{name, paid, date, mode}]}"""
    from copy import copy
    from openpyxl.formula.translate import Translator
    _check_pin()
    b = request.get_json(force=True, silent=True) or {}
    sheet = b.get("sheet")
    plist = [x for x in (b.get("players") or []) if str((x or {}).get("name") or "").strip()]
    if not plist:
        abort(400, "Enter at least one player name")
    _load()
    with lock:
        wb = state["wb"]
        if sheet not in wb.sheetnames:
            abort(404, "Unknown tournament")
        ws = wb[sheet]
        p = (state["layout"].get(sheet) or {}).get("players")
        if not p:
            abort(400, "This sheet has no player table")
        ncol = p["name_col"]
        existing = {str(ws[f"{ncol}{r}"].value).strip().lower() for r in p["player_rows"]
                    if not _is_blank(ws[f"{ncol}{r}"].value)}
        free = [r for r in p["player_rows"] if _is_blank(ws[f"{ncol}{r}"].value)]
        named = [r for r in p["player_rows"] if not _is_blank(ws[f"{ncol}{r}"].value)]
        if len(plist) > len(free):
            abort(409, f"Only {len(free)} empty player row(s) left on this sheet. "
                       "Add more numbered rows in Excel (or create a new tournament) and try again.")
        seen = set()
        for x in plist:
            nm = re.sub(r"\s+", " ", str(x["name"])).strip()[:60]
            if nm.lower() in existing or nm.lower() in seen:
                abort(409, f"{nm} is already in this tournament")
            seen.add(nm.lower())
        template = named[-1] if named else (p["player_rows"][0] if p["player_rows"] else None)
        added, paid_total = [], 0
        for x, r in zip(plist, free):
            nm = re.sub(r"\s+", " ", str(x["name"])).strip()[:60]
            ws[f"{ncol}{r}"] = nm
            _xc_set(ws, f"{ncol}{r}", nm)
            # copy formulas/styles from an existing player row so totals work for the new row
            if template and template != r:
                for col in p["cols"]:
                    if col in (p["sn_col"],):
                        continue
                    t, d = ws[f"{col}{template}"], ws[f"{col}{r}"]
                    if t.has_style:
                        d._style = copy(t._style)
                    if col == ncol:
                        continue
                    d_is_formula = isinstance(d.value, str) and d.value.startswith("=")
                    if isinstance(t.value, str) and t.value.startswith("=") and (_is_blank(d.value) or d_is_formula):
                        newf = Translator(t.value, origin=t.coordinate).translate_formula(d.coordinate)
                        if d.value != newf:
                            d.value = newf
                            state["xc_dirty"] = True
            added.append((r, nm))
        for x, (r, nm) in zip(plist, added):
            amt = _parse_amount(x.get("paid"), allow_zero=True)
            if amt:
                try:
                    pdate = dt.datetime.strptime(str(x.get("date") or dt.date.today()), "%Y-%m-%d")
                except ValueError:
                    pdate = dt.datetime.now()
                ptype = _pay_type(x.get("type")) or "Advance"
                _apply_payment(ws, p, r, nm, amt, pdate, x.get("mode") or "",
                               "Advance" if ptype == "Advance" else "Opening payment", ptype)
                paid_total += amt
        _commit([sheet])
        out = _ledger_payload()
        out["message"] = (f"Added {len(added)} player{'s' if len(added) != 1 else ''} to {sheet}"
                          + (f" · ₹{paid_total:,} recorded" if paid_total else ""))
        return jsonify(out)


def _set_info_line(ws, info, key, text):
    """Replace (in place) or add a 'Key: value' line in the tournament info box (merged cell)."""
    cell = ws[info["ref"]]
    lines = [l for l in str(cell.value or "").splitlines()]
    idx = next((i for i, l in enumerate(lines) if l.strip().lower().startswith(key.lower() + ":")), None)
    lines = [l for i, l in enumerate(lines) if i == idx or not l.strip().lower().startswith(key.lower() + ":")]
    if idx is not None:
        if text:
            lines[idx] = f"{key}: {text}"
        else:
            lines.pop(idx)
    elif text:
        lines.append(f"{key}: {text}")
    cell.value = "\n".join(lines)


@app.post("/api/tournaments/settings")
def api_tournament_settings():
    """Body: {sheet, advance, est_min, est_max}. Stored as lines in the sheet's tournament info box."""
    _check_pin()
    b = request.get_json(force=True, silent=True) or {}
    advance = _parse_amount(b.get("advance"), allow_zero=True)
    est_min = _parse_amount(b.get("est_min"), allow_zero=True)
    est_max = _parse_amount(b.get("est_max"), allow_zero=True) or est_min
    if est_min and est_max < est_min:
        est_min, est_max = est_max, est_min
    _load()
    with lock:
        wb = state["wb"]
        sheet = b.get("sheet")
        if sheet not in wb.sheetnames:
            abort(404, "Unknown tournament")
        ws = wb[sheet]
        info = (state["layout"].get(sheet) or {}).get("info")
        if not info:
            abort(400, "This sheet has no 'Tournament:' info box to store the settings in.")
        _set_info_line(ws, info, "Advance", f"Rs {advance:,}/- per player" if advance else "")
        _set_info_line(ws, info, "Expected cost",
                       "" if not est_min else (f"Rs {est_min:,} - {est_max:,}/- per player" if est_max != est_min
                                               else f"Rs {est_min:,}/- per player"))
        if "org_fee" in b and str(b.get("org_fee") or "").strip() != "":
            fee = _parse_amount(b.get("org_fee"), allow_zero=True)
            org = _org_block(ws)
            if org:
                fc = ws[org["fee_ref"]]
                if _is_formula(fc.value):
                    if _n(_val(ws, org["fee_ref"])) != fee:
                        abort(400, f"The organiser fee on this sheet is a formula ({org['fee_ref']}); change it in Excel.")
                else:
                    fc.value = fee
                    _xc_set(ws, org["fee_ref"], fee)
            elif fee:
                _ensure_org_block(ws, fee)
                state["xc_dirty"] = True
                state["layout"][sheet] = _detect_layout(ws)
            _refresh_rule_line(ws)
        _commit([sheet])
        out = _ledger_payload()
        out["message"] = f"Settings saved for {sheet}"
        return jsonify(out)


@app.post("/api/tournaments/matches")
def api_edit_matches():
    """Rename matches and set the per-player rate and players per match.
    Body: {sheet, matches: [{col, label, rate, team}]}"""
    _check_pin()
    b = request.get_json(force=True, silent=True) or {}
    _load()
    with lock:
        wb = state["wb"]
        sheet = b.get("sheet")
        if sheet not in wb.sheetnames:
            abort(404, "Unknown tournament")
        ws = wb[sheet]
        p = (state["layout"].get(sheet) or {}).get("players")
        fr = _footer_rows(ws, p) if p else {}
        if not fr.get("rate"):
            abort(400, "Match costs on this older sheet are built into its formulas; change them in Excel.")
        for i, col in enumerate(p["match_cols"]):
            m = next((x for x in (b.get("matches") or []) if x.get("col") == col), None)
            if not m:
                continue
            label = str(m.get("label") or "").strip()[:60]
            h = f"Match {i + 1}" + (f" - {label}" if label else "")
            ws[f"{col}{p['header_row']}"] = h
            _xc_set(ws, f"{col}{p['header_row']}", h)
            try:
                team = int(m.get("team") or DEFAULT_TEAM)
            except ValueError:
                team = DEFAULT_TEAM
            if not 1 <= team <= 30:
                abort(400, "Players per match must be between 1 and 30")
            ws[f"{col}{fr['team']}"] = team
            _xc_set(ws, f"{col}{fr['team']}", team)
        _refresh_rule_line(ws)
        _commit([sheet])
        out = _ledger_payload()
        out["message"] = f"Matches updated for {sheet}"
        return jsonify(out)


@app.post("/api/payments")
def api_add_payment():
    """Body: {sheet, row, amount, date, mode, note}"""
    _check_pin()
    b = request.get_json(force=True, silent=True) or {}
    amount = _parse_amount(b.get("amount"))
    try:
        pdate = dt.datetime.strptime(str(b.get("date") or dt.date.today()), "%Y-%m-%d")
    except ValueError:
        abort(400, "Bad date")
    _load()
    with lock:
        ws, p, row, name = _player_ctx(b.get("sheet"), b.get("row"))
        _apply_payment(ws, p, row, name, amount, pdate, b.get("mode"), b.get("note"), b.get("type"))
        _commit([ws.title])
        out = _ledger_payload()
        out["message"] = f"₹{amount:,} recorded for {name}"
        return jsonify(out)


@app.post("/api/organiser")
def api_pay_organiser():
    """Record a payment made TO the organiser. Body: {sheet, amount, date, mode, note}"""
    _check_pin()
    b = request.get_json(force=True, silent=True) or {}
    amount = _parse_amount(b.get("amount"))
    try:
        pdate = dt.datetime.strptime(str(b.get("date") or dt.date.today()), "%Y-%m-%d")
    except ValueError:
        abort(400, "Bad date")
    _load()
    with lock:
        wb = state["wb"]
        sheet = b.get("sheet")
        if sheet not in wb.sheetnames:
            abort(404, "Unknown tournament")
        ws = wb[sheet]
        org = _org_block(ws)
        if not org:
            abort(400, "Set the organiser fee for this tournament first (Tournaments → Settings).")
        cells = [ws[x] for x in org["slots"] if not _is_formula(ws[x].value)]
        if not cells:
            abort(400, "The organiser payment cells on this sheet are formulas; record it in Excel.")
        slot = next((c for c in cells if _is_blank(c.value)), cells[-1])
        cur = slot.value if isinstance(slot.value, (int, float)) else 0
        nv = cur + amount
        nv = int(nv) if float(nv).is_integer() else round(nv, 2)
        slot.value = nv
        slot.number_format = "#,##0"
        _xc_set(ws, slot.coordinate, nv)
        dcell = ws.cell(slot.row, slot.column + 1)
        date_ref = ""
        if (_is_blank(dcell.value) or isinstance(dcell.value, (dt.date, dt.datetime))) and not isinstance(dcell, MergedCell):
            dcell.value = pdate
            dcell.number_format = "dd-mmm-yy"
            date_ref = dcell.coordinate
        ows = _org_ws(wb)
        ids = [int(str(v)[1:]) for (v,) in ows.iter_rows(min_row=2, max_col=1, values_only=True)
               if isinstance(v, str) and re.match(r"^O\d+$", v)]
        oid = f"O{(max(ids) + 1 if ids else 1):04d}"
        ows.append([oid, pdate, sheet, amount, str(b.get("mode") or "").strip()[:20],
                    str(b.get("note") or "").strip()[:200], slot.coordinate, date_ref, dt.datetime.now(), _who() or None])
        ows.cell(ows.max_row, 2).number_format = "dd-mmm-yy"
        ows.cell(ows.max_row, 4).number_format = "#,##0"
        ows.cell(ows.max_row, 9).number_format = "dd-mmm-yy hh:mm"
        _commit([sheet])
        out = _ledger_payload()
        bal = next((t["org"]["balance"] for t in out["tournaments"] if t["sheet"] == sheet), 0)
        out["message"] = f"₹{amount:,} paid to the organiser · balance ₹{bal:,.0f}"
        return jsonify(out)


@app.delete("/api/payments/<pid>")
def api_del_payment(pid):
    """Undo a payment recorded from the website: subtract it again and remove the log row.
    Allowed for admins and the captain of the payment's tournament."""
    _load()
    with lock:
        wb = state["wb"]
        if pid.startswith("O"):
            if ORG_SHEET not in wb.sheetnames:
                abort(404, "Payment not found")
            ows = wb[ORG_SHEET]
            hit = next((r for r in range(2, ows.max_row + 1) if str(ows.cell(r, 1).value) == pid), None)
            if not hit:
                abort(404, "Payment not found")
            sheet, amount = ows.cell(hit, 3).value, _n(ows.cell(hit, 4).value)
            auth.require_write(sheet)
            ref, dref = ows.cell(hit, 7).value, ows.cell(hit, 8).value
            touched = []
            if sheet in wb.sheetnames and ref:
                ws = wb[sheet]
                c = ws[ref]
                if isinstance(c.value, (int, float)):
                    nv = c.value - amount
                    nv = int(nv) if float(nv).is_integer() else round(nv, 2)
                    c.value = nv if nv != 0 else None
                    _xc_set(ws, ref, c.value)
                    if c.value is None and dref:
                        ws[dref].value = None
                    touched.append(sheet)
            ows.delete_rows(hit)
            _commit(touched)
            out = _ledger_payload()
            out["message"] = "Organiser payment removed"
            return jsonify(out)
        if PAYMENTS_SHEET not in wb.sheetnames:
            abort(404, "No payments yet")
        pws = wb[PAYMENTS_SHEET]
        hit = None
        for r in range(2, pws.max_row + 1):
            if str(pws.cell(r, 1).value) == pid:
                hit = r
                break
        if not hit:
            abort(404, "Payment not found")
        sheet, amount, ref = pws.cell(hit, 3).value, _n(pws.cell(hit, 5).value), pws.cell(hit, 8).value
        auth.require_write(sheet)
        touched = []
        if sheet in wb.sheetnames and ref:
            ws = wb[sheet]
            c = ws[ref]
            if isinstance(c.value, (int, float)):
                nv = c.value - amount
                nv = int(nv) if float(nv).is_integer() else round(nv, 2)
                c.value = nv if nv != 0 else None
                _xc_set(ws, ref, c.value)
                touched.append(sheet)
        pws.delete_rows(hit)
        _commit(touched)
        out = _ledger_payload()
        out["message"] = "Payment removed"
        return jsonify(out)


@app.get("/download")
def download():
    auth.require_admin()
    _load()
    return send_file(XLSX, as_attachment=True, download_name=os.path.basename(XLSX))


@app.errorhandler(403)
@app.errorhandler(400)
@app.errorhandler(401)
@app.errorhandler(404)
@app.errorhandler(409)
@app.errorhandler(429)
@app.errorhandler(423)
def _err(e):
    return jsonify({"ok": False, "error": getattr(e, "description", str(e))}), e.code


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    _load(force=True)
    print(f"\n  CricTrack running        ->  http://localhost:{port}")
    print(f"  Excel file               ->  {XLSX}")
    print("  Accounts                 ->  sign in at /login (first account = admin)\n")
    try:
        from waitress import serve
        serve(app, host="0.0.0.0", port=port, threads=4)
    except ImportError:
        app.run(host="0.0.0.0", port=port)
