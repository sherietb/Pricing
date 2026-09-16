"""
SQLite storage for the Pricing App.

Two ways to keep prices current (both supported):
  * Re-import: reload everything from the pricing workbook (openpyxl).
  * Admin edit: update an individual product price / adder in the DB.

Stdlib only (sqlite3). openpyxl is used solely for the workbook import.
"""
import glob
import json
import os
import sqlite3
from datetime import datetime, timezone

import pricing_engine as pe

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "pricing.db")


def latest_pricing_workbook():
    """Newest 'Pricing *.xlsx' in the app folder (ignores Excel ~$ lock files).

    Lets a user drop in a re-dated file (Pricing 7.14.26.xlsx -> Pricing 9.9.26.xlsx)
    without editing any code.
    """
    files = [p for p in glob.glob(os.path.join(HERE, "Pricing *.xlsx"))
             if not os.path.basename(p).startswith("~$")]
    if not files:
        return os.path.join(HERE, "Pricing.xlsx")  # nonexistent -> clear error on import
    return max(files, key=os.path.getmtime)


DEFAULT_WORKBOOK = latest_pricing_workbook()

# Outbound freight by destination. cwt assumes ~full truckload (all_in / 46,000 lb x 100).
# Regions each carry a FREIGHT $/cwt and a COMPETITIVE $/cwt. On a quote the rep applies
# freight only, competitive only, or both (they stack). freight_cwt = pass-through cost
# (excluded from CRU Spread); comp_cwt = a price move (counts toward margin).
# NOTE: DFW / East TX / Houston-South TX freight and ALL comp_cwt are PLACEHOLDERS - confirm.
REGIONS = [
    {"region": "DFW",              "freight_cwt": 0.75, "comp_cwt": 0.00},
    {"region": "East TX",          "freight_cwt": 1.25, "comp_cwt": 0.00},
    {"region": "OKC",              "freight_cwt": 1.70, "comp_cwt": 0.00},
    {"region": "Tulsa",            "freight_cwt": 2.05, "comp_cwt": 0.00},
    {"region": "Houston/South TX", "freight_cwt": 2.00, "comp_cwt": 0.00},
    {"region": "West TX",          "freight_cwt": 2.55, "comp_cwt": 0.00},
]

SCHEMA = """
CREATE TABLE IF NOT EXISTS products (
    key         TEXT PRIMARY KEY,
    thickness   TEXT NOT NULL,
    thickness_in REAL,
    grade       TEXT NOT NULL,
    width       TEXT,
    width_num   REAL,
    price_cwt   REAL,
    price_note  TEXT
);
CREATE TABLE IF NOT EXISTS stock_tiers  (tier_label TEXT, max_weight INTEGER, adder_cwt REAL);
CREATE TABLE IF NOT EXISTS length_tiers (tier_label TEXT, max_weight INTEGER, adder_cwt REAL);
CREATE TABLE IF NOT EXISTS extras (name TEXT PRIMARY KEY, adder_cwt REAL, selectable INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS freight (destination TEXT PRIMARY KEY, freight_cwt REAL, comp_cwt REAL);
CREATE TABLE IF NOT EXISTS quotes (
    id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT, customer TEXT, quote_no TEXT,
    destination TEXT, adder_basis TEXT, use_inventory INTEGER, grand_total REAL, payload TEXT
);
CREATE TABLE IF NOT EXISTS customers (
    name TEXT PRIMARY KEY, segment TEXT, playbook INTEGER, credit TEXT
);
CREATE TABLE IF NOT EXISTS cust_config (k TEXT PRIMARY KEY, v REAL);
"""

# Customer pricing layers ($/cwt), tunable in Admin. Playbook 1/2/3 and credit High/Low.
CUSTOMER_CONFIG = {"pb1_cwt": -2.0, "pb2_cwt": 0.0, "pb3_cwt": 2.0,
                   "credit_high_cwt": 2.0, "credit_low_cwt": 0.0}


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = connect()
    conn.executescript(SCHEMA)
    conn.commit()
    conn.close()


def _s(v):
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _tier_weight(label):
    digits = "".join(ch for ch in str(label) if ch.isdigit())
    return int(digits) if digits else None


def import_from_workbook(path=None):
    """Wipe and reload all pricing data from the Excel workbook.

    With no path, always picks the newest 'Pricing *.xlsx' currently in the folder.
    """
    import openpyxl

    if path is None:
        path = latest_pricing_workbook()
    if not os.path.exists(path):
        raise FileNotFoundError(path)

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["1011-36"]

    products, seen = [], set()
    for r in range(2, ws.max_row + 1):
        th, gr, wd, price = (ws.cell(r, 1).value, ws.cell(r, 2).value,
                             ws.cell(r, 3).value, ws.cell(r, 5).value)
        if th is None and gr is None:
            continue
        th_s, gr_s, wd_s = _s(th), _s(gr), _s(wd)
        key = "-".join(part for part in (th_s, wd_s, gr_s) if part)
        if key in seen:      # keep first occurrence
            continue
        seen.add(key)
        numeric = isinstance(price, (int, float))
        products.append((key, th_s, pe.thickness_to_inch(th_s), gr_s, wd_s,
                         pe.width_to_number(wd_s),
                         float(price) if numeric else None,
                         None if numeric else _s(price)))

    stock, length, target = [], [], None
    for r in range(3, 20):
        g, h = ws.cell(r, 7).value, ws.cell(r, 8).value
        if g is None:
            continue
        gl = str(g).lower()
        if "stock adders" in gl:
            target = stock; continue
        if "custom length" in gl:
            target = length; continue
        if target is not None and isinstance(h, (int, float)):
            target.append((_s(g), _tier_weight(g), float(h)))

    cost = wb["Cost"]
    extras = []
    for r in range(2, cost.max_row + 1):
        d, e = cost.cell(r, 4).value, cost.cell(r, 5).value
        if d is not None and isinstance(e, (int, float)):
            name = _s(d)
            selectable = 1 if name in pe.ALLOWED_EXTRAS else 0
            extras.append((name, float(e), selectable))

    conn = connect()
    conn.executescript(SCHEMA)
    for t in ("products", "stock_tiers", "length_tiers", "extras"):
        conn.execute(f"DELETE FROM {t}")
    conn.executemany("INSERT INTO products VALUES (?,?,?,?,?,?,?,?)", products)
    conn.executemany("INSERT INTO stock_tiers VALUES (?,?,?)", stock)
    conn.executemany("INSERT INTO length_tiers VALUES (?,?,?)", length)
    conn.executemany("INSERT OR REPLACE INTO extras VALUES (?,?,?)", extras)
    conn.execute("INSERT OR REPLACE INTO meta VALUES ('last_import', ?)",
                 (datetime.now(timezone.utc).isoformat(timespec="seconds"),))
    conn.execute("INSERT OR REPLACE INTO meta VALUES ('source_workbook', ?)", (os.path.basename(path),))
    conn.commit()
    conn.close()
    return {"products": len(products), "stock_tiers": len(stock),
            "length_tiers": len(length), "extras": len(extras)}


# ---- freight ------------------------------------------------------------
def round_005(x):
    """Round a $/cwt value to the nearest $0.05."""
    return round(round(float(x) / 0.05) * 0.05, 2)


def ensure_freight():
    conn = connect()
    conn.executescript(SCHEMA)
    reseed = False
    try:
        current = {r[0] for r in conn.execute("SELECT destination FROM freight WHERE freight_cwt IS NOT NULL")}
        if current != {r["region"] for r in REGIONS}:
            reseed = True   # regions added/removed
    except sqlite3.OperationalError:
        reseed = True       # old schema (no freight_cwt column)
    if reseed:
        conn.execute("DROP TABLE IF EXISTS freight")
        conn.executescript(SCHEMA)
        conn.executemany("INSERT OR REPLACE INTO freight VALUES (?,?,?)",
                         [(r["region"], round_005(r["freight_cwt"]), round_005(r["comp_cwt"])) for r in REGIONS])
        conn.commit()
    conn.close()


def get_freight(destination):
    ensure_freight()
    conn = connect()
    row = conn.execute("SELECT * FROM freight WHERE destination=?", (destination,)).fetchone()
    conn.close()
    return dict(row) if row else None


# ---- read helpers -------------------------------------------------------
def bootstrap():
    ensure_freight()
    """Everything the UI needs to render the form."""
    conn = connect()
    prods = [dict(r) for r in conn.execute(
        "SELECT key, thickness, thickness_in, grade, width, width_num, price_cwt, price_note FROM products "
        "ORDER BY thickness, grade, width")]
    extras = [dict(r) for r in conn.execute(
        "SELECT name, adder_cwt FROM extras WHERE selectable=1 ORDER BY name")]
    meta = {r["k"]: r["v"] for r in conn.execute("SELECT k, v FROM meta")}
    freight = [dict(r) for r in conn.execute("SELECT destination, freight_cwt, comp_cwt FROM freight ORDER BY destination")]
    conn.close()
    return {"products": prods, "extras": extras, "meta": meta, "freight": freight}


def get_product(key):
    conn = connect()
    row = conn.execute("SELECT * FROM products WHERE key=?", (key,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_tiers():
    conn = connect()
    stock = [dict(r) for r in conn.execute("SELECT * FROM stock_tiers")]
    length = [dict(r) for r in conn.execute("SELECT * FROM length_tiers")]
    extras = {r["name"]: r["adder_cwt"] for r in conn.execute("SELECT name, adder_cwt FROM extras")}
    conn.close()
    return stock, length, extras


def update_price(key, price_cwt):
    conn = connect()
    cur = conn.execute("UPDATE products SET price_cwt=?, price_note=NULL WHERE key=?",
                       (float(price_cwt), key))
    conn.commit()
    changed = cur.rowcount
    conn.close()
    return changed


def save_quote(customer, quote_no, quote, grand_total):
    conn = connect()
    conn.executescript(SCHEMA)
    cur = conn.execute(
        "INSERT INTO quotes (created_at, customer, quote_no, destination, adder_basis, use_inventory, grand_total, payload) "
        "VALUES (?,?,?,?,?,?,?,?)",
        (datetime.now().isoformat(timespec="seconds"), (customer or "").strip(), (quote_no or "").strip(),
         quote.get("destination", ""), quote.get("adder_basis", "line"),
         1 if quote.get("use_inventory") else 0, float(grand_total or 0), json.dumps(quote)))
    conn.commit()
    qid = cur.lastrowid
    conn.close()
    return qid


def list_quotes(limit=50):
    conn = connect()
    conn.executescript(SCHEMA)
    out = []
    for r in conn.execute("SELECT id, created_at, customer, quote_no, destination, grand_total, payload "
                          "FROM quotes ORDER BY id DESC LIMIT ?", (limit,)):
        d = dict(r)
        try:
            d["lines"] = len(json.loads(d.pop("payload")).get("lines", []))
        except Exception:
            d.pop("payload", None); d["lines"] = 0
        out.append(d)
    conn.close()
    return out


def get_quote(qid):
    conn = connect()
    conn.executescript(SCHEMA)
    row = conn.execute("SELECT * FROM quotes WHERE id=?", (qid,)).fetchone()
    conn.close()
    if not row:
        return None
    d = dict(row)
    try:
        d["payload"] = json.loads(d["payload"])
    except Exception:
        d["payload"] = {}
    return d


def delete_quote(qid):
    conn = connect()
    conn.executescript(SCHEMA)
    n = conn.execute("DELETE FROM quotes WHERE id=?", (qid,)).rowcount
    conn.commit()
    conn.close()
    return n


# ---- customers ----------------------------------------------------------
def get_cust_config():
    conn = connect()
    conn.executescript(SCHEMA)
    have = {r["k"]: r["v"] for r in conn.execute("SELECT k, v FROM cust_config")}
    for k, v in CUSTOMER_CONFIG.items():
        if k not in have:
            conn.execute("INSERT INTO cust_config VALUES (?,?)", (k, v))
    conn.commit()
    cfg = {r["k"]: r["v"] for r in conn.execute("SELECT k, v FROM cust_config")}
    conn.close()
    return {**CUSTOMER_CONFIG, **cfg}


def set_cust_config(updates):
    conn = connect()
    conn.executescript(SCHEMA)
    for k in CUSTOMER_CONFIG:
        if k in updates and updates[k] is not None:
            conn.execute("INSERT OR REPLACE INTO cust_config VALUES (?,?)", (k, float(updates[k])))
    conn.commit()
    conn.close()
    return get_cust_config()


def list_customers():
    conn = connect()
    conn.executescript(SCHEMA)
    rows = [dict(r) for r in conn.execute("SELECT name, segment, playbook, credit FROM customers ORDER BY name")]
    conn.close()
    return rows


def get_customer(name):
    if not name:
        return None
    conn = connect()
    conn.executescript(SCHEMA)
    row = conn.execute("SELECT * FROM customers WHERE name=?", (name.strip(),)).fetchone()
    conn.close()
    return dict(row) if row else None


def upsert_customer(name, segment=None, playbook=2, credit="low"):
    name = (name or "").strip()
    if not name:
        raise ValueError("customer name required")
    pb = int(playbook) if playbook in (1, 2, 3, "1", "2", "3") else 2
    credit = "high" if str(credit).lower().startswith("h") else "low"
    conn = connect()
    conn.executescript(SCHEMA)
    conn.execute("INSERT OR REPLACE INTO customers VALUES (?,?,?,?)",
                 (name, (segment or "").strip(), pb, credit))
    conn.commit()
    conn.close()
    return get_customer(name)


def delete_customer(name):
    conn = connect()
    conn.executescript(SCHEMA)
    n = conn.execute("DELETE FROM customers WHERE name=?", ((name or "").strip(),)).rowcount
    conn.commit()
    conn.close()
    return n


def customer_adjustments(name):
    """List of {label, cwt} $/cwt layers for a customer's playbook + credit (stacks)."""
    c = get_customer(name)
    if not c:
        return []
    cfg = get_cust_config()
    adjs = []
    pb_amt = cfg.get("pb%d_cwt" % (c["playbook"] or 2), 0.0)
    if pb_amt:
        adjs.append({"label": "Playbook %d" % c["playbook"], "cwt": pb_amt})
    if (c["credit"] or "low") == "high" and cfg.get("credit_high_cwt"):
        adjs.append({"label": "Credit: High Risk", "cwt": cfg["credit_high_cwt"]})
    return adjs


def import_customers_csv(path=None):
    """Bulk-load customers from a CSV with columns name, segment, playbook, credit."""
    import csv
    if path is None:
        path = os.path.join(HERE, "Customers.csv")
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    n = 0
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            row = {(k or "").strip().lower(): v for k, v in row.items()}
            if not row.get("name"):
                continue
            upsert_customer(row.get("name"), row.get("segment"),
                            row.get("playbook", 2), row.get("credit", "low"))
            n += 1
    return {"imported": n}


if __name__ == "__main__":
    stats = import_from_workbook()
    print("Imported:", stats)
