# Pricing Logic Changelog

Version-stamped history of the **pricing logic** in the Willbanks Pricing App.
Every generated quote is journaled (see the Quote Journal in Admin) with the
`LOGIC_VERSION` and a full config snapshot in effect at that moment, so any
historical quote can be traced back to exactly how it was priced.

Bump `LOGIC_VERSION` in `database.py` and add an entry here whenever the pricing
logic changes (adders, region rates, inventory rules, credit rules, engine math).

---

## 2026.10.20
- **Quote journal added.** Every saved/exported quote is logged with a timestamp,
  the logic version, and a config snapshot (region rates, customer adders, credit
  file, inventory rules, cost config).
- **High-risk credit = Top-20 Credit Risk Accounts file** (interim for the Invex
  "Statement Cycle D" field, which isn't in the SQL extracts). Newest matching
  credit file in the folder wins; every listed account gets the **+$2/cwt** credit
  adder. Replaced the old `credit_status = 'H'` rule.
- **Inventory-position adjustment now applies automatically** (no opt-in button).
- **Region auto-identified from the customer's ship-to** (dominant delivery ZIP ->
  TX/OK ZIP3 region map, city/state fallback). Rep can still override.
- **Internal Excel export** of a priced quote (product, weight lb/tons, inventory
  position, suggested price, each adder in its own column).

## 2026.09.18
- **Confirmed regional competitive rates (CCO):** West TX +$2, DFW 0 (baseline),
  OKC +$2, Tulsa +$3, Houston/South TX -$2, East TX -$1 per cwt.
  *(Freight $/cwt per region still placeholder pending confirmation.)*

## 2026.09.16  — CCO roadmap Phases 1-4
- **Phase 1 — Region:** freight-only / competitive / both modes that stack;
  inventory-position thresholds (<1mo +$2 / 1-2mo baseline / >2mo -$2);
  CTL vs discrete-plate split with separate rules.
- **Phase 2 — Customer layer:** playbook 1/2/3 = -$2 / mid / +$2; credit High/Low;
  live customer + credit pull from the ERP. Layers stack into the quote.
- **Phase 4 — Dynamic margin:** per-customer margin that ratchets up per win
  (capped) and eases back per loss; stacks as its own layer.

## Base pricing (ongoing)
- Prices are **$/cwt** ($/lb x 100). Line rate = base + all applicable adders,
  then `line_total = (rate / 100) x weight_lb`.
- Stock/volume adders by total weight ("X and under"); boundary weight rolls up to
  the looser tier; over 20,000 lb = no adder.
- Selectable extras: P&O, Temper Pass, 16-ga thickness adder (grade and slit-edge
  are baked into the base price).
- Base price book loaded from the newest `Pricing *.xlsx` in the folder.
