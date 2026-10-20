"""
Inventory position layer for the Pricing App.

Loads the "Inventory Coil Projections" workbook into pricing.db and exposes each
product's inventory position so the quote screen can price more or less competitively.

Signal (business-confirmed):
  * Basis = PROJECTED months of supply = (on-hand lbs + coils on PO) / monthly
    consumption (6-month avg, falling back to 1-month). This reflects the forward
    position the projections workbook is built around, including coils already on order.
  * Per-SKU available lbs (on-hand minus reserved) from `Details` is shown alongside.
  * Flat $/cwt adjustment by position, tunable in Admin:
        long  (>= long_months)  -> adj_long_cwt   (typically negative = discount)
        short (<= short_months) -> adj_short_cwt  (typically positive = premium)
        balanced / unknown      -> 0
"""
import os
import sqlite3
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "pricing.db")
DEFAULT_WORKBOOK = os.path.join(HERE, "Inventory Coil Projections V1.xlsx")

# Direct SQL Server source (the same query that feeds the workbook). Windows auth =
# no stored password. Override the connection string via env var if needed.
SOURCE_DB_CONN = os.environ.get(
    "PRICING_INV_DB_CONN",
    "DRIVER={SQL Server};SERVER=10.0.1.50;DATABASE=Inventory_Planning;Trusted_Connection=yes")
SOURCE_QUERY_FILE = os.path.join(HERE, "inventory_query.sql")

# Inventory-position ranking, kept SEPARATELY for CTL (cut-to-length coil) and discrete
# PLATE so each can be tuned on its own. Rule per type: months-on-hand < short_months ->
# premium (+adj_short); > long_months -> discount (adj_long); in between -> baseline.
PRODUCT_TYPES = ("ctl", "plate")
RULE_KEYS = ("short_months", "long_months", "adj_short_cwt", "adj_long_cwt")
DEFAULT_RULE = {"short_months": 1.0, "long_months": 2.0, "adj_short_cwt": 2.0, "adj_long_cwt": -2.0}

# Inventory data older than this many days shows a "stale" warning to reps.
STALE_DAYS = 2

SCHEMA = """
CREATE TABLE IF NOT EXISTS inventory_sku (
    key TEXT PRIMARY KEY, grade TEXT, size TEXT, width TEXT,
    ohd_lbs REAL, reserved_lbs REAL, available_lbs REAL, incoming_lbs REAL, po_lbs REAL,
    avg_cost_cwt REAL, product_type TEXT
);
CREATE TABLE IF NOT EXISTS inventory_grade (
    grade TEXT PRIMARY KEY, ohd_lbs REAL, avail_lbs REAL, po_lbs REAL,
    cons_1m REAL, cons_6m REAL, forecast REAL, w1_moh REAL
);
CREATE TABLE IF NOT EXISTS inventory_sku_moh (
    key TEXT PRIMARY KEY, ohd_lbs REAL, avail_lbs REAL, po_lbs REAL,
    cons_1m REAL, cons_6m REAL, forecast REAL, w1_lbs REAL
);
CREATE TABLE IF NOT EXISTS inv_rule (k TEXT PRIMARY KEY, v REAL);
CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS inventory_pos (
    scope TEXT, key TEXT, grade TEXT, size TEXT, width TEXT, form TEXT,
    ohd_lbs REAL, available_lbs REAL, incoming_lbs REAL, po_lbs REAL,
    avg_cost_cwt REAL, months REAL, product_type TEXT,
    PRIMARY KEY (scope, key)
);
"""

# Form -> how its inventory position is matched + which rule it uses.
# Coil forms match on grade+width (fungible); discrete forms match an exact SKU.
COIL_FORMS = ("CC", "FPCO")        # CC = cut coil (CTL), FPCO = floor-plate coil
DISCRETE_FORMS = ("CP", "FP")      # CP = discrete plate, FP = floor-plate stock
FORM_CLASS = {"CC": "coil", "FPCO": "coil", "CP": "disc", "FP": "disc"}
FORM_PTYPE = {"CC": "ctl", "FPCO": "plate", "CP": "plate", "FP": "plate"}


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def _norm_grade(label):
    s = str(label).strip()
    if s == "101150":
        return "1011-50"
    if s.startswith("786"):
        return "786"
    return s


def _is_grade_label(label):
    if label in (None, "Grand Total"):
        return False
    s = str(label).strip()
    if s == "" or "|" in s or "MOH" in s or "month" in s.lower():   # skip leaf/header/junk rows
        return False
    return True


# Thickness description (as used in the MOH pivot Item key) -> price-list thickness code.
DESC_TO_CODE = {
    '3/16"': "316", '1/4"': "14", '5/16"': "516", '3/8"': "38", '7/16"': "716",
    '1/2"': "12", '9/16"': "916", '5/8"': "58", '11/16"': "1116", '3/4"': "34",
    '7/8"': "78", '1"': "1", '1-1/4"': '1-1/4"', '1-1/2"': '1-1/2"', '2"': '2"',
    "7 GA": "7G", "8 GA": "8G", "9 GA": "9G", "10 GA": "10G", "11 GA": "11G",
    "12 GA": "12G", "14 GA": "14G", "16 GA": "16G",
}


def parse_moh_item(item):
    """MOH pivot Item 'grade|thickness_desc|finish|width' -> price-list key, or None."""
    parts = str(item).split("|")
    if len(parts) != 4:
        return None
    grade, desc, finish, width = [p.strip() for p in parts]
    code = DESC_TO_CODE.get(desc)
    if not code or not width:
        return None
    grade = _norm_grade(grade)
    wpart = width + "-slit edge" if finish.upper() == "SLIT" else width
    return "-".join(p for p in (code, wpart, grade) if p)


def init_rule():
    conn = connect()
    conn.executescript(SCHEMA)
    have = {r["k"] for r in conn.execute("SELECT k FROM inv_rule")}
    for t in PRODUCT_TYPES:
        for k, v in DEFAULT_RULE.items():
            kk = t + "_" + k
            if kk not in have:
                conn.execute("INSERT INTO inv_rule VALUES (?,?)", (kk, v))
    conn.commit()
    conn.close()


def get_rules():
    """Both rule sets: {'ctl': {...}, 'plate': {...}}."""
    init_rule()
    conn = connect()
    stored = {r["k"]: r["v"] for r in conn.execute("SELECT k, v FROM inv_rule")}
    conn.close()
    return {t: {k: stored.get(t + "_" + k, DEFAULT_RULE[k]) for k in RULE_KEYS} for t in PRODUCT_TYPES}


def set_rule(product_type, updates):
    """Update one product type's rule (product_type in PRODUCT_TYPES)."""
    init_rule()
    t = product_type if product_type in PRODUCT_TYPES else "ctl"
    conn = connect()
    for k in RULE_KEYS:
        if k in updates and updates[k] is not None:
            conn.execute("INSERT OR REPLACE INTO inv_rule VALUES (?,?)", (t + "_" + k, float(updates[k])))
    conn.commit()
    conn.close()
    return get_rules()


def import_inventory(path=DEFAULT_WORKBOOK):
    import openpyxl
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

    det = {}
    ws = wb["Details"]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=15, values_only=True):
        grade, size, w = row[1], row[2], row[6]
        if grade is None or size is None or w is None:
            continue
        width = str(int(w)) if isinstance(w, (int, float)) else str(w).strip()
        g = _norm_grade(grade)   # 786-14G / 786-316 / 101150 -> 786 / 1011-50 (match price-list grade)
        key = "-".join(p for p in (str(size).strip(), width, g) if p)
        d = det.setdefault(key, {"grade": g, "size": str(size).strip(), "width": width,
                                 "ohd": 0.0, "res": 0.0, "inc": 0.0, "po": 0.0, "cst_lbs": 0.0,
                                 "cc": 0.0, "fpco": 0.0})
        d["ohd"] += _num(row[9]); d["res"] += _num(row[11])
        d["inc"] += _num(row[13]); d["po"] += _num(row[14])
        d["cst_lbs"] += _num(row[10]) * _num(row[9])   # ohd_cst($/cwt) x ohd_lbs -> lbs-weighted avg
        fw = _num(row[9]) + _num(row[14])              # ohd + PO lbs, for CTL/plate form dominance
        if str(row[0] or "").strip().upper() == "FPCO":
            d["fpco"] += fw                            # FPCO = discrete plate
        else:
            d["cc"] += fw                              # CC (or unknown) = CTL / cut coil
    sku_rows = [(k, d["grade"], d["size"], d["width"], d["ohd"], d["res"],
                 d["ohd"] - d["res"], d["inc"], d["po"],
                 round(d["cst_lbs"] / d["ohd"], 2) if d["ohd"] > 0 else None,
                 "plate" if d["fpco"] > d["cc"] else "ctl") for k, d in det.items()]

    grade_acc = {}        # grade subtotal rows (collapsed pivot)
    leaf_grade_acc = {}   # grades aggregated from per-SKU leaf rows (expanded pivot)
    sku_moh = {}          # per-SKU, keyed by price-list key: (ohd,avail,po,c1,c6,fc,w1_lbs)
    for sheet in ("Non65 MOH", "Gr65 MOH"):
        ws = wb[sheet]
        for r in range(4, ws.max_row + 1):
            label = ws.cell(r, 1).value
            if label in (None, "Grand Total") or str(label).strip() == "":
                continue
            s = str(label).strip()
            ohd, avail, po = _num(ws.cell(r, 2).value), _num(ws.cell(r, 3).value), _num(ws.cell(r, 5).value)
            c1, c6, fc = _num(ws.cell(r, 7).value), _num(ws.cell(r, 8).value), _num(ws.cell(r, 9).value)
            w1_moh, w1_lbs = _num(ws.cell(r, 10).value), _num(ws.cell(r, 23).value)  # col W = W1 Lbs
            if "|" in s:
                key = parse_moh_item(s)
                if not key:
                    continue
                row = (ohd, avail, po, c1, c6, fc, w1_lbs)
                # multiple finishes (plain + P&O) can map to one product key -> SUM, not overwrite
                sku_moh[key] = tuple(a + b for a, b in zip(sku_moh[key], row)) if key in sku_moh else row
                g = _norm_grade(s.split("|")[0]); acc = leaf_grade_acc
            elif _is_grade_label(s) and any((ohd, avail, po, c1, c6, fc, w1_moh)):
                g = _norm_grade(s); acc = grade_acc
            else:
                continue
            a = acc.setdefault(g, dict(ohd=0.0, avail=0.0, po=0.0, c1=0.0, c6=0.0, fc=0.0, moh=0.0, n=0))
            a["ohd"] += ohd; a["avail"] += avail; a["po"] += po; a["c1"] += c1
            a["c6"] += c6; a["fc"] += fc; a["moh"] += w1_moh; a["n"] += 1
    grade_acc.update(leaf_grade_acc)   # prefer leaf-aggregated grade rollups when present
    grade_rows = [(g, a["ohd"], a["avail"], a["po"], a["c1"], a["c6"], a["fc"],
                   (a["moh"] / a["n"] if a["n"] else 0.0)) for g, a in grade_acc.items()]
    sku_moh_rows = [(k,) + v for k, v in sku_moh.items()]

    conn = connect()
    conn.execute("DROP TABLE IF EXISTS inventory_sku")       # rebuilt each import (schema may change)
    conn.execute("DROP TABLE IF EXISTS inventory_sku_moh")
    conn.executescript(SCHEMA)
    conn.execute("DELETE FROM inventory_sku")
    conn.execute("DELETE FROM inventory_grade")
    conn.execute("DELETE FROM inventory_sku_moh")
    conn.executemany("INSERT OR REPLACE INTO inventory_sku VALUES (?,?,?,?,?,?,?,?,?,?,?)", sku_rows)
    conn.executemany("INSERT OR REPLACE INTO inventory_grade VALUES (?,?,?,?,?,?,?,?)", grade_rows)
    conn.executemany("INSERT OR REPLACE INTO inventory_sku_moh VALUES (?,?,?,?,?,?,?,?)", sku_moh_rows)
    conn.execute("INSERT OR REPLACE INTO meta VALUES ('last_inventory_import', ?)",
                 (datetime.now(timezone.utc).isoformat(timespec="seconds"),))
    conn.execute("INSERT OR REPLACE INTO meta VALUES ('inv_source_mtime', ?)",
                 (str(os.path.getmtime(path)),))
    conn.execute("INSERT OR REPLACE INTO meta VALUES ('inv_source_file', ?)",
                 (os.path.basename(path),))
    conn.commit()
    conn.close()
    init_rule()
    return {"skus": len(sku_rows), "grades": len(grade_rows), "sku_moh": len(sku_moh_rows)}


def db_available():
    """True if the SQL Server source is reachable/usable right now."""
    try:
        import pyodbc
    except Exception:
        return False
    if not os.path.exists(SOURCE_QUERY_FILE):
        return False
    try:
        cn = pyodbc.connect(SOURCE_DB_CONN, timeout=4)
        cn.close()
        return True
    except Exception:
        return False


def import_inventory_from_db():
    """Pull inventory LIVE from SQL Server (same query as the workbook) and populate
    inventory_sku + inventory_sku_moh directly - no spreadsheet export needed.
    Months-on-hand = (on-hand + PO) / 6-month-avg monthly consumption."""
    import pyodbc
    sql = open(SOURCE_QUERY_FILE, encoding="utf-8").read()
    cn = pyodbc.connect(SOURCE_DB_CONN, timeout=20)
    cur = cn.cursor(); cur.execute(sql)
    cols = [c[0] for c in cur.description]
    i = {c: n for n, c in enumerate(cols)}
    rows = cur.fetchall(); cn.close()
    monthly_cols = [c for c in ("Tot_C_1M", "c_lbs_2mth", "c_lbs_3mth", "c_lbs_4mth", "c_lbs_5mth", "c_lbs_6mth") if c in i]

    def n(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    def _blank():
        return {"ohd": 0.0, "res": 0.0, "inc": 0.0, "po": 0.0, "cst_lbs": 0.0,
                "cons6": 0.0, "cons1": 0.0, "grade": None, "size": None, "width": None, "form": None}

    det = {}
    coil_acc, disc_acc = {}, {}     # form-aware: coil keyed grade|width, discrete keyed thk-width-grade
    for r in rows:
        grade, size, w = r[i["grade"]], r[i["size"]], r[i["W"]]
        if grade is None or size is None or w is None:
            continue
        width = str(int(w)) if isinstance(w, (int, float)) else str(w).strip()
        try:
            width = str(int(float(width)))
        except ValueError:
            pass
        g = _norm_grade(grade)
        size_s = str(size).strip()
        key = "-".join(p for p in (size_s, width, g) if p)
        form = str(r[i["form"]] or "").strip().upper()
        ohd, po = n(r[i["ohd_lbs"]]), n(r[i["po_lbs_tot"]])
        res = n(r[i["tot_res_lbs"]]); inc = n(r[i["incoming_lbs"]])
        cst_lbs = n(r[i["ohd_cst"]]) * ohd
        cons6 = sum(n(r[i[c]]) for c in monthly_cols)
        cons1 = n(r[i["Tot_C_1M"]]) if "Tot_C_1M" in i else 0.0

        d = det.setdefault(key, {"grade": g, "size": size_s, "width": width,
                                 "ohd": 0.0, "res": 0.0, "inc": 0.0, "po": 0.0, "cst_lbs": 0.0,
                                 "cc": 0.0, "fpco": 0.0, "cons6": 0.0, "cons1": 0.0})
        d["ohd"] += ohd; d["res"] += res; d["inc"] += inc; d["po"] += po
        d["cst_lbs"] += cst_lbs; d["cons6"] += cons6; d["cons1"] += cons1
        if form == "FPCO":
            d["fpco"] += ohd + po
        else:
            d["cc"] += ohd + po

        # form-aware accumulation: coil forms -> grade|width, discrete forms -> exact SKU
        cls = FORM_CLASS.get(form)
        if cls == "coil":
            a = coil_acc.setdefault(g + "|" + width, _blank())
            a["grade"], a["width"] = g, width
        elif cls == "disc":
            a = disc_acc.setdefault(key, _blank())
            a["grade"], a["size"], a["width"] = g, size_s, width
        else:
            a = None
        if a is not None:
            a["ohd"] += ohd; a["res"] += res; a["inc"] += inc; a["po"] += po
            a["cst_lbs"] += cst_lbs; a["cons6"] += cons6; a["cons1"] += cons1
            a["form"] = form

    sku_rows, sku_moh_rows, grade_acc = [], [], {}
    for k, d in det.items():
        avail = d["ohd"] - d["res"]
        ptype = "plate" if d["fpco"] > d["cc"] else "ctl"
        cost = round(d["cst_lbs"] / d["ohd"], 2) if d["ohd"] > 0 else None
        sku_rows.append((k, d["grade"], d["size"], d["width"], d["ohd"], d["res"],
                         avail, d["inc"], d["po"], cost, ptype))
        monthly = d["cons6"] / (len(monthly_cols) or 1)          # 6-mo avg monthly consumption
        # reuse the sku_moh formula months = w1_lbs / forecast  ->  (ohd+po) / monthly
        sku_moh_rows.append((k, d["ohd"], avail, d["po"], d["cons1"], monthly, monthly, d["ohd"] + d["po"]))
        a = grade_acc.setdefault(d["grade"], dict(ohd=0.0, avail=0.0, po=0.0, c1=0.0, c6=0.0, fc=0.0, moh=0.0, n=0))
        a["ohd"] += d["ohd"]; a["avail"] += avail; a["po"] += d["po"]; a["c1"] += d["cons1"]; a["c6"] += monthly; a["fc"] += monthly; a["n"] += 1
    grade_rows = [(g, a["ohd"], a["avail"], a["po"], a["c1"], a["c6"], a["fc"], 0.0) for g, a in grade_acc.items()]

    def _months(ohd, po, monthly):
        tot = ohd + po
        if monthly and monthly > 0:
            return round(min(tot / monthly, 12.0), 1)
        return 12.0 if tot > 0 else 0.0

    def _pos_rows(acc, scope):
        out = []
        for k, a in acc.items():
            avail = a["ohd"] - a["res"]
            monthly = a["cons6"] / (len(monthly_cols) or 1)
            cost = round(a["cst_lbs"] / a["ohd"], 2) if a["ohd"] > 0 else None
            ptype = FORM_PTYPE.get((a.get("form") or "").upper(), "ctl")
            out.append((scope, k, a.get("grade"), a.get("size"), a.get("width"), a.get("form"),
                        a["ohd"], avail, a["inc"], a["po"], cost,
                        _months(a["ohd"], a["po"], monthly), ptype))
        return out
    pos_rows = _pos_rows(coil_acc, "coil") + _pos_rows(disc_acc, "sku")

    conn = connect()
    conn.execute("DROP TABLE IF EXISTS inventory_sku")
    conn.execute("DROP TABLE IF EXISTS inventory_sku_moh")
    conn.executescript(SCHEMA)
    for t in ("inventory_sku", "inventory_grade", "inventory_sku_moh", "inventory_pos"):
        conn.execute("DELETE FROM " + t)
    conn.executemany("INSERT OR REPLACE INTO inventory_sku VALUES (?,?,?,?,?,?,?,?,?,?,?)", sku_rows)
    conn.executemany("INSERT OR REPLACE INTO inventory_grade VALUES (?,?,?,?,?,?,?,?)", grade_rows)
    conn.executemany("INSERT OR REPLACE INTO inventory_sku_moh VALUES (?,?,?,?,?,?,?,?)", sku_moh_rows)
    conn.executemany("INSERT OR REPLACE INTO inventory_pos VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", pos_rows)
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    conn.execute("INSERT OR REPLACE INTO meta VALUES ('last_inventory_import', ?)", (now,))
    conn.execute("INSERT OR REPLACE INTO meta VALUES ('inv_source_mtime', ?)", (str(datetime.now().timestamp()),))
    conn.execute("INSERT OR REPLACE INTO meta VALUES ('inv_source_file', 'SQL Server: Inventory_Planning (live)')")
    conn.commit()
    conn.close()
    init_rule()
    return {"skus": len(sku_rows), "grades": len(grade_rows), "sku_moh": len(sku_moh_rows), "source": "database"}


def maybe_auto_import(path=DEFAULT_WORKBOOK):
    """Re-import automatically if the workbook file has been saved since the last import.

    Lets a person just overwrite the export file (Level 1 auto-detect); the app picks
    it up on the next page load / periodic check without anyone clicking Re-import.
    Returns True if it re-imported.
    """
    if not os.path.exists(path):
        return False
    cur = os.path.getmtime(path)
    conn = connect()
    try:
        row = conn.execute("SELECT v FROM meta WHERE k='inv_source_mtime'").fetchone()
    except sqlite3.OperationalError:
        row = None
    conn.close()
    last = float(row["v"]) if row and row["v"] else 0.0
    if (not has_data()) or cur > last + 1:   # 1s tolerance
        import_inventory(path)
        return True
    return False


def projected_months(gr):
    cons = gr["cons_6m"] if gr["cons_6m"] and gr["cons_6m"] > 0 else gr["cons_1m"]
    if not cons or cons <= 0:
        return None
    return round((gr["ohd_lbs"] + gr["po_lbs"]) / cons, 2)


def classify(mos, rule):
    """rule = one product type's dict. <short -> premium, >long -> discount, else baseline."""
    if mos is None:
        return "unknown", 0.0
    if mos > rule["long_months"]:
        return "long", rule["adj_long_cwt"]
    if mos < rule["short_months"]:
        return "short", rule["adj_short_cwt"]
    return "balanced", 0.0


def has_data():
    conn = connect()
    try:
        n = conn.execute("SELECT COUNT(*) FROM inventory_grade").fetchone()[0]
    except sqlite3.OperationalError:
        n = 0
    conn.close()
    return n > 0


def get_meta():
    conn = connect()
    try:
        rows = {r["k"]: r["v"] for r in conn.execute(
            "SELECT k, v FROM meta WHERE k IN ('last_inventory_import','inv_source_mtime','inv_source_file')")}
    except sqlite3.OperationalError:
        rows = {}
    conn.close()
    as_of, days_old, stale = None, None, False
    if rows.get("inv_source_mtime"):
        dt = datetime.fromtimestamp(float(rows["inv_source_mtime"]))
        as_of = dt.isoformat(timespec="seconds")
        days_old = round((datetime.now() - dt).total_seconds() / 86400.0, 1)
        stale = days_old > STALE_DAYS
    return {"last_inventory_import": rows.get("last_inventory_import"),
            "source_file": rows.get("inv_source_file"),
            "as_of": as_of, "days_old": days_old, "stale": stale, "stale_days": STALE_DAYS}


def _wkey(width):
    try:
        return str(int(float(width)))
    except (TypeError, ValueError):
        return str(width).strip()


def get_inventory_for(thickness, width, grade, form=None, rules=None):
    """Form-aware inventory position. Coil forms (CC/FPCO) match on grade+width;
    discrete forms (CP/FP) match an exact SKU (thickness+width+grade). No matching
    stock on hand -> the position defaults to SHORT (never a grade-level fallback)."""
    rules = rules or get_rules()
    form = (form or "").strip().upper()
    cls = FORM_CLASS.get(form)
    if cls is None:   # no explicit form: coil for gauges, discrete (plate) otherwise
        cls = "coil" if str(thickness).strip().upper().endswith("G") else "disc"
    scope = "coil" if cls == "coil" else "sku"
    wk = _wkey(width)
    poskey = (grade + "|" + wk) if scope == "coil" else \
        "-".join(p for p in (str(thickness).strip(), wk, grade) if p)
    conn = connect()
    try:
        rec = conn.execute("SELECT * FROM inventory_pos WHERE scope=? AND key=?",
                           (scope, poskey)).fetchone()
    except sqlite3.OperationalError:
        conn.close(); return None
    conn.close()
    ptype = FORM_PTYPE.get(form, "ctl" if cls == "coil" else "plate")
    rule = rules.get(ptype, rules["ctl"])
    if rec is None:
        # No stock on hand for this item -> SHORT.
        pos, adj = classify(0.0, rule)
        return {"grade": grade, "product_type": ptype, "form": form or None,
                "match_scope": scope, "matched": False,
                "sku_ohd_lbs": 0, "sku_available_lbs": 0, "avg_cost_cwt": None,
                "sku_incoming_lbs": None, "months_of_supply": 0.0,
                "position": pos, "adjustment_cwt": round(adj, 2), "basis": "no-stock"}
    mos = rec["months"] if rec["months"] is not None else 0.0
    pos, adj = classify(mos, rule)
    return {"grade": grade, "product_type": ptype, "form": rec["form"] or form or None,
            "match_scope": scope, "matched": True,
            "sku_ohd_lbs": round(rec["ohd_lbs"]) if rec["ohd_lbs"] is not None else None,
            "sku_available_lbs": round(rec["available_lbs"]) if rec["available_lbs"] is not None else None,
            "avg_cost_cwt": rec["avg_cost_cwt"],
            "sku_incoming_lbs": round(rec["incoming_lbs"]) if rec["incoming_lbs"] else None,
            "months_of_supply": mos, "position": pos, "adjustment_cwt": round(adj, 2),
            "basis": ("coil (grade+width)" if scope == "coil" else "sku")}


if __name__ == "__main__":
    print("Imported:", import_inventory())
    print("Rules:", get_rules())
