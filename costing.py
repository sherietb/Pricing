"""
Material-cost + spread layer for the Pricing App.

Reconstructs an estimated MATERIAL cost $/cwt per product from the workbook `Cost`
sheet, so the quote screen can show the **spread over material cost** (matches the
methodology of the workbook's `Gr65` "Spread" column) and softly warn when a line
falls below a target spread %.

IMPORTANT: this is *material* cost only (no conversion/freight-in/overhead), derived
from the Cost sheet - so "spread" here is spread-over-material, not full gross margin.
Costs are editable per product (Admin), and the material bases are tunable; grades with
no material basis (e.g. AR400) return no cost, so no spread is shown or flagged for them.
"""
import os
import sqlite3

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pricing.db")

# Seeded from Pricing 7.14.26.xlsx `Cost` sheet. Movable bases live in cost_config.
DEFAULT_BASES = {"hrc": 38.05, "plate_narrow": 35.65, "plate_wide": 26.45, "plate_120": 29.45}
DEFAULT_MIN_SPREAD = 35.0  # % - flag lines below this spread-over-material

GAUGES = {"16G", "14G", "12G", "11G", "10G"}
THK_EXTRA = {"316": 4.0, "14": 3.0, "516": 1.5}          # others 0
GRADE_HRC = {"786": 1.5, "871-65": 6.0, "572-65": 3.0}   # extras on HRC base
GRADE_PLATE = {"572-65": 3.25, "871-65": 6.5, "572-50": 3.0, "516-70": 3.0, "786": 1.5}
NO_COST_GRADES = {"AR400"}                                # no material basis -> unknown
SLIT, TEMPER, G16 = 1.75, 1.5, 1.5

SCHEMA = """
CREATE TABLE IF NOT EXISTS product_cost (key TEXT PRIMARY KEY, cost_cwt REAL, source TEXT);
CREATE TABLE IF NOT EXISTS cost_config (k TEXT PRIMARY KEY, v REAL);
"""


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _wnum(w):
    s = str(w or "")
    num = ""
    for ch in s:
        if ch.isdigit():
            num += ch
        elif num:
            break
    return int(num) if num else None


def init_config():
    conn = connect()
    conn.executescript(SCHEMA)
    have = {r["k"] for r in conn.execute("SELECT k FROM cost_config")}
    for k, v in DEFAULT_BASES.items():
        if k not in have:
            conn.execute("INSERT INTO cost_config VALUES (?,?)", (k, v))
    if "min_spread_pct" not in have:
        conn.execute("INSERT INTO cost_config VALUES ('min_spread_pct', ?)", (DEFAULT_MIN_SPREAD,))
    conn.commit()
    conn.close()


def get_config():
    init_config()
    conn = connect()
    cfg = {r["k"]: r["v"] for r in conn.execute("SELECT k, v FROM cost_config")}
    conn.close()
    base = dict(DEFAULT_BASES); base["min_spread_pct"] = DEFAULT_MIN_SPREAD
    base.update(cfg)
    return base


def set_config(updates):
    init_config()
    conn = connect()
    for k in list(DEFAULT_BASES) + ["min_spread_pct"]:
        if k in updates and updates[k] is not None:
            conn.execute("INSERT OR REPLACE INTO cost_config VALUES (?,?)", (k, float(updates[k])))
    conn.commit()
    conn.close()
    return get_config()


def material_cost(thickness, grade, width, bases=None):
    """Estimated material cost $/cwt for a product, or None if no material basis."""
    if grade in NO_COST_GRADES:
        return None
    b = bases or get_config()
    slit = SLIT if "slit" in str(width).lower() else 0.0
    if thickness in GAUGES:
        c = b["hrc"] + GRADE_HRC.get(grade, 0.0)
        if thickness == "16G":
            c += G16
        if thickness in ("14G", "16G"):
            c += TEMPER
        return round(c + slit, 2)
    wn = _wnum(width)
    if wn is None:
        return None
    if wn >= 120:
        wbase = b["plate_120"]
    elif wn >= 84:
        wbase = b["plate_wide"]
    else:
        wbase = b["plate_narrow"]
    return round(wbase + THK_EXTRA.get(thickness, 0.0) + GRADE_PLATE.get(grade, 0.0) + slit, 2)


def derive_costs():
    """(Re)compute derived costs for every product; never overwrite a manual override."""
    init_config()
    bases = get_config()
    conn = connect()
    prods = list(conn.execute("SELECT key, thickness, grade, width FROM products"))
    manual = {r["key"] for r in conn.execute("SELECT key FROM product_cost WHERE source='manual'")}
    n = 0
    for p in prods:
        if p["key"] in manual:
            continue
        c = material_cost(p["thickness"], p["grade"], p["width"], bases)
        conn.execute("INSERT OR REPLACE INTO product_cost VALUES (?,?, 'derived')", (p["key"], c))
        n += 1
    conn.commit()
    conn.close()
    return n


def ensure_costs():
    conn = connect()
    conn.executescript(SCHEMA)
    empty = conn.execute("SELECT COUNT(*) FROM product_cost").fetchone()[0] == 0
    conn.close()
    if empty:
        derive_costs()


def get_cost(key):
    ensure_costs()
    conn = connect()
    row = conn.execute("SELECT cost_cwt FROM product_cost WHERE key=?", (key,)).fetchone()
    conn.close()
    return row["cost_cwt"] if row and row["cost_cwt"] is not None else None


def set_cost(key, cost_cwt):
    init_config()
    conn = connect()
    conn.executescript(SCHEMA)
    conn.execute("INSERT OR REPLACE INTO product_cost VALUES (?,?, 'manual')", (key, float(cost_cwt)))
    conn.commit()
    n = conn.total_changes
    conn.close()
    return n


def get_min_spread():
    return get_config()["min_spread_pct"]


if __name__ == "__main__":
    print("Derived costs for", derive_costs(), "products. Config:", get_config())
