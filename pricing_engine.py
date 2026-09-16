"""
Pricing engine for the Willbanks steel Pricing App.

Pure calculation logic -- no I/O. All pricing data (products, adder tiers,
extras) is passed in from the database layer.

Confirmed rules (see project notes):
  * Prices are $/cwt (per 100 lb):  $/cwt = $/lb x 100.
  * A line's rate = base $/cwt + all applicable adders, THEN multiplied by weight:
        line_total = (rate / 100) * weight_lb
  * Weight (carbon steel): width_in * length_in * thickness_in * 0.2836 * qty
  * Stock-weight adder tiers ("X and under" = X lb or less):
        - over 20,000 lb  -> no adder ($0)
        - a boundary weight rolls UP to the looser (lower) adder
          (exactly 10,000 lb -> the "20,000 and under" $3 tier; exactly 20,000 -> $3)
  * Stock and custom-length adders can stack on the same line.
  * Only P&O, Temper Pass, and the 16-ga thickness adder are selectable extras
    (grade and slit-edge are already baked into the base price -> no double charge).
"""

STEEL_DENSITY_LB_IN3 = 0.2836

# Human-readable thickness labels for display (raw code -> shown text).
# Gauges (16G, 14G, ...) already read clearly, so they pass through unchanged.
THICKNESS_DISPLAY = {
    "316": '3/16"', "14": '1/4"', "516": '5/16"', "38": '3/8"', "716": '7/16"',
    "12": '1/2"', "916": '9/16"', "58": '5/8"', "1116": '11/16"', "34": '3/4"',
    "78": '7/8"', "18": '1/8"', "732": '7/32"', "932": '9/32"', "1": '1"',
}


def thickness_label(code):
    """Show fraction codes as fractions (14 -> 1/4\"); leave gauges/others as-is."""
    c = str(code).strip()
    return THICKNESS_DISPLAY.get(c, c)
# Assumed average weight of one coil or one truckload (placeholder until refined).
COIL_TL_WEIGHT_LB = 46000
QUANTITY_UNITS = ("pieces", "lbs", "coil", "tl")

# Only these processing extras may be added on top of this price list.
ALLOWED_EXTRAS = ("P&O", "Temper Pass", "Thickness Adder for 16 ga")

# Standard US carbon-steel sheet gauge -> decimal inches.
GAUGE_TO_INCH = {
    "16G": 0.0598, "14G": 0.0747, "12G": 0.1046, "11G": 0.1196, "10G": 0.1345,
}
# Fraction-of-inch thickness codes (written without the slash in the sheet).
FRACTION_TO_INCH = {
    "316": 3 / 16, "14": 1 / 4, "516": 5 / 16, "38": 3 / 8, "12": 1 / 2,
    "58": 5 / 8, "34": 3 / 4, "1": 1.0, '1-1/4"': 1.25, '1-1/2"': 1.5, '2"': 2.0,
}


def thickness_to_inch(code):
    """Decimal inches for a thickness code, or None if unknown."""
    key = str(code).strip()
    if key in GAUGE_TO_INCH:
        return GAUGE_TO_INCH[key]
    return FRACTION_TO_INCH.get(key)


def width_to_number(width):
    """Leading numeric part of a width label: '60-slit edge' -> 60.0, '96PL' -> 96.0."""
    s = str(width).strip()
    num = ""
    for ch in s:
        if ch.isdigit() or (ch == "." and "." not in num):
            num += ch
        elif num:
            break
    return float(num) if num else None


def piece_weight_lb(thickness_in, width_in, length_in, qty):
    """Total weight in lb for `qty` rectangular plates."""
    return thickness_in * width_in * length_in * STEEL_DENSITY_LB_IN3 * qty


def tier_adder(weight_lb, tiers):
    """
    Return (tier_label, adder_cwt) for a weight given ascending-by-max_weight tiers.

    "X and under" = X lb or less. A boundary weight rolls UP to the looser/lower
    adder; anything strictly above the largest tier gets no adder.
    """
    ordered = sorted(tiers, key=lambda t: t["max_weight"])
    for t in ordered:
        if weight_lb < t["max_weight"]:
            return t["tier_label"], t["adder_cwt"]
    top = ordered[-1]
    if weight_lb == top["max_weight"]:
        return top["tier_label"], top["adder_cwt"]
    return "over %s (no adder)" % f"{top['max_weight']:,}", 0.0


def compute_line(product, length_in, qty, *, unit="pieces", custom_length=False, extras=(),
                 inv_adjustment=None, freight=None, extra_adjustments=None, tier_weight=None, stock_tiers, length_tiers, extras_catalog):
    """
    Price a single quote line.

    product        : dict with keys thickness, thickness_in, grade, width,
                     width_num, price_cwt (float or None), price_note.
    extras         : iterable of extra names selected by the rep.
    stock_tiers    : list of {tier_label, max_weight, adder_cwt}
    length_tiers   : same shape, for custom-length adders
    extras_catalog : dict {extra_name: adder_cwt}

    Returns a result dict; if the product has no numeric price, `manual_quote`
    is True and no total is produced.
    """
    result = {
        "key": product["key"],
        "description": _describe(product),
        "manual_quote": False,
        "note": None,
        "unit": None,
        "qty": None,
        "weight_basis": None,
        "weight_lb": None,
        "breakdown": [],
        "rate_cwt": None,
        "line_total": None,
    }

    if product.get("price_cwt") is None:
        result["manual_quote"] = True
        result["note"] = f"Price is '{product.get('price_note') or 'not listed'}' - quote manually."
        return result

    unit = (unit or "pieces").strip().lower()
    if unit not in QUANTITY_UNITS:
        raise ValueError(f"unknown quantity unit: {unit!r}")
    qty = float(qty)
    if qty <= 0:
        raise ValueError("quantity must be positive")

    if unit == "lbs":
        weight = qty
        basis = f"{qty:,.0f} lb"
    elif unit in ("coil", "tl"):
        weight = qty * COIL_TL_WEIGHT_LB
        label = "coil" if unit == "coil" else "TL"
        basis = f"{qty:g} {label} x {COIL_TL_WEIGHT_LB:,} lb"
    else:  # pieces
        thickness_in = product.get("thickness_in") or thickness_to_inch(product["thickness"])
        width_in = product.get("width_num")
        if width_in is None:
            width_in = width_to_number(product["width"])
        if thickness_in is None or width_in is None:
            result["manual_quote"] = True
            result["note"] = ("Cannot compute piece weight (unrecognized thickness or width) - "
                              "price it by lbs, coil, or TL instead.")
            return result
        length_in = float(length_in or 0)
        if length_in <= 0:
            raise ValueError("length (inches) is required to price by the piece")
        weight = piece_weight_lb(thickness_in, width_in, length_in, qty)
        basis = f"{qty:g} pc x {width_in:g}\" x {length_in:g}\""

    result["unit"] = unit
    result["qty"] = qty
    result["weight_basis"] = basis

    base = float(product["price_cwt"])
    breakdown = [{"label": "Base price", "cwt": round(base, 4)}]

    tw = weight if tier_weight is None else float(tier_weight)
    slabel, sadd = tier_adder(tw, stock_tiers)
    if sadd:
        breakdown.append({"label": f"Stock adder ({slabel})", "cwt": round(sadd, 4)})

    if custom_length:
        llabel, ladd = tier_adder(tw, length_tiers)
        if ladd:
            breakdown.append({"label": f"Custom-length adder ({llabel})", "cwt": round(ladd, 4)})

    for name in extras:
        if name in ALLOWED_EXTRAS and name in extras_catalog:
            breakdown.append({"label": f"Extra: {name}", "cwt": round(float(extras_catalog[name]), 4)})

    if inv_adjustment and inv_adjustment.get("cwt"):
        breakdown.append({"label": inv_adjustment.get("label", "Inventory adjustment"),
                          "cwt": round(float(inv_adjustment["cwt"]), 4)})

    if freight and freight.get("cwt"):
        breakdown.append({"label": freight.get("label", "Freight"),
                          "cwt": round(float(freight["cwt"]), 4)})

    for a in (extra_adjustments or []):        # region freight/competitive, customer layers, etc.
        if a and a.get("cwt"):
            breakdown.append({"label": a.get("label", "Adjustment"), "cwt": round(float(a["cwt"]), 4)})

    rate = sum(item["cwt"] for item in breakdown)
    total = rate / 100.0 * weight

    result["weight_lb"] = round(weight, 1)
    result["breakdown"] = breakdown
    result["rate_cwt"] = round(rate, 4)
    result["line_total"] = round(total, 2)
    return result


def _describe(p):
    return f'{thickness_label(p["thickness"])} thick, {p["width"]} wide, grade {p["grade"]}'
