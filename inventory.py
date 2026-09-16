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

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pricing.db")
DEFAULT_WORKBOOK = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "Inventory Coil Projections V1.xlsx")

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
"""


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


def get_inventory_for(key, grade, rules=None):
    rules = rules or get_rules()
    conn = connect()
    try:
        sku = conn.execute("SELECT * FROM inventory_sku WHERE key=?", (key,)).fetchone()
        gr = conn.execute("SELECT * FROM inventory_grade WHERE grade=?", (grade,)).fetchone()
        smoh = conn.execute("SELECT * FROM inventory_sku_moh WHERE key=?", (key,)).fetchone()
    except sqlite3.OperationalError:
        conn.close(); return None
    conn.close()
    if not sku and not gr and not smoh:
        return None
    ptype = (sku["product_type"] if sku and sku["product_type"] else "ctl")   # CTL vs discrete plate
    rule = rules.get(ptype, rules["ctl"])
    # Prefer the MOH sheet's on-hand/available (what the CCO reads); fall back to Details.
    out = {"grade": grade, "product_type": ptype,
           "sku_ohd_lbs": round(smoh["ohd_lbs"]) if smoh else (round(sku["ohd_lbs"]) if sku else None),
           "sku_available_lbs": round(smoh["avail_lbs"]) if smoh else (round(sku["available_lbs"]) if sku else None),
           "avg_cost_cwt": (sku["avg_cost_cwt"] if sku and sku["avg_cost_cwt"] else None),
           "sku_incoming_lbs": (round(sku["incoming_lbs"]) if sku and sku["incoming_lbs"] else None),
           "grade_available_lbs": None, "on_order_lbs": None, "months_of_supply": None,
           "position": "unknown", "adjustment_cwt": 0.0, "basis": None}
    if smoh is not None:
        # Per-SKU months-on-hand = projected W1 lbs / monthly forecast (matches the sheet's W1 MOH);
        # dead stock (no forecast) caps at 12, mirroring the sheet.
        fc, w1 = smoh["forecast"], smoh["w1_lbs"]
        if fc and fc > 0:
            mos = round(min(w1 / fc, 12.0), 1)
        else:
            mos = 12.0 if (w1 > 0 or smoh["ohd_lbs"] > 0) else 0.0
        pos, adj = classify(mos, rule)
        out.update({"months_of_supply": mos, "position": pos, "adjustment_cwt": round(adj, 2),
                    "on_order_lbs": round(smoh["po_lbs"]), "basis": "sku"})
    elif gr:
        # Fall back to grade-level projected months when this SKU has no per-SKU MOH row
        mos = projected_months(gr)
        pos, adj = classify(mos, rule)
        out.update({"grade_available_lbs": round(gr["avail_lbs"]), "on_order_lbs": round(gr["po_lbs"]),
                    "months_of_supply": mos, "position": pos, "adjustment_cwt": round(adj, 2), "basis": "grade"})
    return out


if __name__ == "__main__":
    print("Imported:", import_inventory())
    print("Rules:", get_rules())
