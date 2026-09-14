# Willbanks Pricing App

Internal web app that turns a product spec (thickness + grade + width + length + quantity)
into an extended, quote-ready price straight from the pricing workbook.

## Requirements
- Python 3.9+ (tested on 3.14)
- `openpyxl` (only needed to import the Excel workbook): `pip install openpyxl`
- Everything else is Python standard library (no framework, no build step).

## Run
```bash
python server.py            # http://localhost:8000
python server.py 8080       # choose a port
```
On first run it seeds `pricing.db` from `Pricing 7.14.26.xlsx`. Open the URL in a browser.
To let other people on the network use it, they browse to `http://<this-machine-ip>:8000`.

## How pricing works
`line_total = (base $/cwt + adders) / 100 x weight_lb`

- **Quantity can be entered by**: Pieces, Pounds (lbs), Coils, or Truckloads (TL).
  - Pieces: weight = thickness_in x width_in x length_in x 0.2836 lb/in3 x qty
    (gauges convert to decimal inches; fraction codes like `316` = 3/16"; length required).
  - Lbs: the quantity *is* the weight. Coils / TL: each assumed **46,000 lb** (placeholder).
- **Adders basis** toggle: *per line* (each line's own weight picks its tier) or
  *per order* (combined weight of all lines picks one tier applied to every line).
- **Freight ("Ship to")**: pick a destination and its $/cwt freight folds into the rate,
  shown as its own build-up line. "Pickup" = no freight. Destinations, miles, rate/mile,
  all-in cost and $/cwt live in `FREIGHT` in `database.py`. NOTE: the $/cwt assumes a
  ~full truckload (all-in cost / 46,000 lb); a partial load ships for the same truck cost,
  so freight per cwt would be higher on light orders.
- **Stock weight adder** by total weight tier. "20,000 and under" = 20,000 lb or less;
  over 20,000 lb has no adder. Boundary weights roll up to the lower adder.
- **Custom-length adder** stacks on top when the line is custom-cut.
- **Extras** selectable on this price list: P&O, Temper Pass, 16-ga adder. Grade and
  slit-edge are already baked into the base price, so they are intentionally not offered.
- Items priced `inquire`/`NA` route to **manual quote** (no computed total).

## Keeping prices current (two ways)
1. **Re-import Excel** - drop an updated workbook in place and click *Re-import Excel*
   in the Admin panel (or `python database.py`).
2. **Admin edit** - change a single product price by key in the Admin panel.

## Inventory-aware pricing (optional)
Import `Inventory Coil Projections V1.xlsx` (Admin -> *Re-import inventory*) and the app
can nudge each price to be more or less competitive based on stock position.

- **Signal**: *projected* months of supply = (on-hand + coils on PO) / monthly consumption
  (6-mo avg), rolled up by grade, plus per-SKU available lbs shown alongside.
- **Positions**: `long` (overstocked -> discount), `balanced` (no change), `short`
  (tight -> premium). Grades with no coil inventory (e.g. 516-70, AR400) price normally.
- **Pricing stance** (switchable any time by rep or admin, on the quote screen):
  - *Conservative* - long >= 6 mo -> -$1.50/cwt (only the most overstocked grades)
  - *Moderate* (default) - long >= 5 mo -> -$2.50/cwt (~3% off, 4 grades)
  - *Aggressive* - long >= 4.5 mo -> -$4.00/cwt (~4.7% off)
  - *Custom* - whatever thresholds/amounts you enter in the Admin rule editor.
  Short-side premium is on standby (nothing is that tight on projected basis today).
- The position is always **shown** on the quote line; the price only moves when the rep
  ticks **"Adjust price for inventory position"** (show-and-apply, rep stays in control).
- Joins to ~125 of 141 products.

### Keeping inventory fresh (auto-detect)
No manual clicking needed for routine refreshes:
- **Auto-import**: the app watches `Inventory Coil Projections V1.xlsx`. Whenever a new
  export overwrites that file, the app re-imports it automatically - on the next page
  load, and via a background check every ~10 minutes. Just save the export in place.
- **Staleness badge**: the quote screen shows *"as of <date> (N days old)"*; if the data
  is older than `STALE_DAYS` (default 2), it shows a red **STALE** badge so reps never
  quote off old positions. Change the threshold via `STALE_DAYS` in `inventory.py`.
- Manual **Re-import inventory** (Admin) still works for an on-demand refresh.
- Cadence is therefore just *how often someone saves a fresh export* - a daily save keeps
  everything current with zero clicks.

## Files
| File | Purpose |
|------|---------|
| `server.py` | Stdlib HTTP server + JSON API |
| `index.html` | Sales-rep pricing screen (single page) |
| `pricing_engine.py` | Pure pricing math (weight, adders, extended total) |
| `database.py` | SQLite store + Excel import |
| `pricing.db` | Generated database (safe to delete; re-seeds from Excel) |
| `Pricing 7.14.26.xlsx` | Source price list (sheet `1011-36`) + costs |

## Roadmap / not yet built
- **Smart intake**: paste an email or drop a PDF/Excel and have specs auto-extracted
  into quote lines (needs an AI key; the UI placeholder is in place and wires into the
  same engine).
- Refine the coil / TL assumed weight (currently a flat 46,000 lb) if you want it to
  vary by gauge/width. Change `COIL_TL_WEIGHT_LB` in `pricing_engine.py`.
- Save quotes to the database (Print / Save-as-PDF is available now via the browser).
- The workbook's `Gr65` sheet is excluded (its formulas are broken with `#REF!`).

## Margin guardrail (spread over material cost)
Each priced line shows its **spread over material cost** ($/cwt and %), and warns (red
"⚠ below floor") when it drops under the target — so inventory discounts + freight can't
quietly push a line below your margin threshold.

- **Material cost** is estimated from the workbook `Cost` sheet (CRU HRC/plate base +
  grade/thickness/width extras) in `costing.py`, and re-derives automatically on price
  re-import. It is **material cost only** (no conversion/freight-in/overhead) — so this is
  *spread over material*, matching the `Gr65` sheet's "Spread", not full gross margin.
- **Freight is excluded** from the spread (it's a pass-through cost, not margin).
- Grades with no material basis (e.g. AR400) show no cost and are never flagged.
- **Admin** exposes the **material bases** (HRC, plate 48&ndash;72", 84&ndash;96", 120")
  and the **min spread %** floor (default 35%, soft warning) — so a weekly cost update is
  just a couple of numbers; saving re-derives every non-overridden cost. Admin can also
  **override the material cost** for any product key (overrides survive re-derivation).
- ⚠ The derived costs are estimates — validate them before relying on the floor.

## Saved quotes
Build a quote, fill Customer / Quote #, and click **Save quote**. Saved quotes appear in
the *Saved quotes* list (date, customer, quote #, ship-to, lines, total). **Load** repopulates
the whole quote for editing or re-quoting a repeat customer (it recomputes at *current*
prices/rules, so a reloaded quote reflects today's pricing). Stored in the `quotes` table.

## Print / PDF
Add lines, fill the optional Customer / Quote # fields, then **Print / Save as PDF**.
The print view hides the form/admin and shows just the quote with a header, customer,
date, and adder basis.
