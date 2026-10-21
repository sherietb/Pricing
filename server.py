"""
Willbanks Pricing App -- internal web server (Python standard library only).

Run:
    python server.py            # serves http://localhost:8000
    python server.py 8080       # custom port

Endpoints:
    GET  /                      -> the rep pricing screen (index.html)
    GET  /api/bootstrap         -> products, selectable extras, inventory status/rule, meta
    POST /api/quote             -> {lines:[...], adder_basis, use_inventory}
    POST /api/admin/price       -> {key, price_cwt}       (admin edit)
    POST /api/reimport          -> reload prices from the Excel workbook
    POST /api/reimport-inventory-> reload inventory positions from the projections workbook
    POST /api/admin/inv-rule    -> {long_months, short_months, adj_long_cwt, adj_short_cwt}
"""
import json
import os
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import costing
import database as db
import inventory
import pricing_engine as pe
import quote_export

HERE = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(HERE, "index.html")


class Handler(BaseHTTPRequestHandler):
    server_version = "PricingApp/1.0"

    # -- helpers ----------------------------------------------------------
    def _send(self, code, body, ctype="application/json"):
        if isinstance(body, (dict, list)):
            body = json.dumps(body).encode("utf-8")
        elif isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype + ("; charset=utf-8" if "html" in ctype or "json" in ctype else ""))
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_download(self, data, filename, ctype):
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Disposition", 'attachment; filename="%s"' % filename)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _log_quote(self, event, payload, result, ref=None):
        """Journal a generated quote with the inventory/cost config the server owns."""
        extra = {}
        try:
            extra["inv_rules"] = inventory.get_rules()
        except Exception:
            pass
        try:
            extra["cost_config"] = costing.get_config()
        except Exception:
            pass
        if ref is not None:
            extra["saved_quote_id"] = ref
        db.log_quote(event, payload, result, extra_config=extra)

    def _read_json(self):
        length = int(self.headers.get("Content-Length", 0))
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def log_message(self, *args):
        pass  # quiet

    # -- routes -----------------------------------------------------------
    def do_GET(self):
        if self.path in ("/", "/index.html"):
            try:
                with open(INDEX, "rb") as f:
                    self._send(200, f.read(), "text/html")
            except FileNotFoundError:
                self._send(500, {"error": "index.html missing"})
        elif self.path == "/api/quotes":
            self._send(200, {"quotes": db.list_quotes()})
        elif self.path.startswith("/api/quote-log"):
            self._send(200, {"log": db.list_quote_log(), "logic_version": db.LOGIC_VERSION})
        elif self.path == "/api/export-log":
            data = quote_export.build_journal_workbook(db.list_quote_log(limit=5000, full=True))
            from datetime import datetime
            self._send_download(data, "Quote_Journal_%s.xlsx" % datetime.now().strftime("%Y%m%d"),
                                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        elif self.path.startswith("/api/dashboard"):
            import urllib.parse
            q = urllib.parse.parse_qs(self.path.split("?", 1)[1]) if "?" in self.path else {}
            def g(k):
                v = q.get(k, [""])[0].strip()
                return v or None
            days = int(q.get("days", ["90"])[0] or 90)
            self._send(200, db.dashboard(days, osr=g("osr"), isr=g("isr"),
                                         whs=g("whs"), form=g("form")))
        elif self.path.startswith("/api/quotes/"):
            qid = self.path.rsplit("/", 1)[-1]
            q = db.get_quote(int(qid)) if qid.isdigit() else None
            self._send(200, q or {"error": "not found"})
        elif self.path == "/api/bootstrap":
            try:
                inventory.maybe_auto_import()   # pick up a freshly-saved export
            except Exception:
                pass
            data = db.bootstrap()
            data["inventory_loaded"] = inventory.has_data()
            data["inv_rules"] = inventory.get_rules()   # {'ctl': {...}, 'plate': {...}}
            data["inventory_meta"] = inventory.get_meta()
            data["cost_config"] = costing.get_config()
            data["customers"] = db.list_customers()
            data["cust_config"] = db.get_cust_config()
            data["logic_version"] = db.LOGIC_VERSION
            self._send(200, data)
        else:
            self._send(404, {"error": "not found"})

    def do_POST(self):
        try:
            if self.path == "/api/quote":
                self._send(200, self._quote(self._read_json()))
            elif self.path == "/api/export-quote":
                payload = self._read_json()
                result = self._quote(payload)
                cust = db.get_customer(payload.get("customer_name"))
                pb = (cust or {}).get("playbook", 2) or 2
                pb_adder = db.get_cust_config().get("pb%d_cwt" % pb, 0.0)
                data = quote_export.build_workbook(payload, result, db.get_product,
                                                   playbook=pb, pb_adder=pb_adder)
                self._log_quote("exported", payload, result)
                from datetime import datetime
                fn = "Quote_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M")
                self._send_download(
                    data, fn,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            elif self.path == "/api/admin/price":
                payload = self._read_json()
                n = db.update_price(payload["key"], payload["price_cwt"])
                self._send(200, {"updated": n})
            elif self.path == "/api/reimport":
                res = db.import_from_workbook()
                costing.derive_costs()   # refresh material costs from the updated Cost sheet
                self._send(200, {"imported": res})
            elif self.path == "/api/admin/cost":
                pl = self._read_json()
                self._send(200, {"updated": costing.set_cost(pl["key"], pl["cost_cwt"])})
            elif self.path == "/api/admin/cost-config":
                cfg = costing.set_config(self._read_json())
                costing.derive_costs()
                self._send(200, {"cost_config": cfg})
            elif self.path == "/api/reimport-inventory":
                self._send(200, {"imported": inventory.import_inventory()})
            elif self.path == "/api/refresh-from-db":
                self._send(200, {"imported": inventory.import_inventory_from_db()})
            elif self.path == "/api/admin/inv-rule":
                pl = self._read_json()
                self._send(200, {"rules": inventory.set_rule(pl.get("product_type", "ctl"), pl)})
            elif self.path == "/api/admin/customer":
                pl = self._read_json()
                self._send(200, {"customer": db.upsert_customer(pl.get("name"), pl.get("segment"),
                                                                pl.get("playbook", 2), pl.get("credit", "low"))})
            elif self.path == "/api/admin/customer-delete":
                self._send(200, {"deleted": db.delete_customer(self._read_json().get("name"))})
            elif self.path == "/api/admin/customer-config":
                self._send(200, {"cust_config": db.set_cust_config(self._read_json())})
            elif self.path == "/api/admin/import-customers":
                self._send(200, db.import_customers_csv())
            elif self.path == "/api/admin/pull-customers-db":
                self._send(200, db.import_customers_from_db())
            elif self.path == "/api/dash-config":
                self._send(200, {"dash_config": db.set_dash_config(self._read_json())})
            elif self.path.startswith("/api/quotes/") and self.path.endswith("/status"):
                qid = self.path.split("/")[3]
                pl = self._read_json()
                self._send(200, {"updated": db.set_quote_status(
                    int(qid), pl.get("status"), pl.get("lost_reason"),
                    pl.get("competitor_cwt")) if qid.isdigit() else 0})
            elif self.path == "/api/quotes":
                pl = self._read_json()
                quote = pl.get("quote", {})
                result = None
                try:
                    result = self._quote(quote)
                    quote["_priced"] = [{"key": r.get("key"), "cwt": r.get("rate_cwt")}
                                        for r in result.get("lines", []) if r.get("key")]
                except Exception:
                    pass
                qid = db.save_quote(pl.get("customer"), pl.get("quote_no"),
                                    quote, pl.get("grand_total"))
                try:
                    self._log_quote("saved", quote, result or self._quote(quote), ref=qid)
                except Exception:
                    pass
                self._send(200, {"id": qid})
            elif self.path == "/api/item-history":
                pl = self._read_json()
                cust = pl.get("customer_name")
                out = [db.item_history(cust, it.get("key"), it.get("grade"), it.get("thickness"))
                       for it in pl.get("items", [])]
                self._send(200, {"history": out})
            elif self.path.startswith("/api/quotes/") and self.path.endswith("/delete"):
                qid = self.path.split("/")[3]
                self._send(200, {"deleted": db.delete_quote(int(qid)) if qid.isdigit() else 0})
            else:
                self._send(404, {"error": "not found"})
        except Exception as exc:  # surface a clean error to the UI
            self._send(400, {"error": str(exc)})

    # -- quote ------------------------------------------------------------
    def _quote(self, payload):
        stock, length, extras_catalog = db.get_tiers()
        lines = payload.get("lines", [])
        basis = (payload.get("adder_basis") or "line").lower()
        # Inventory-position adjustment is applied automatically whenever inventory data
        # is loaded (the old opt-in toggle defaulted off); pass use_inventory=false to suppress.
        use_inv = payload.get("use_inventory", True) is not False and inventory.has_data()
        rules = inventory.get_rules() if inventory.has_data() else None
        region = payload.get("destination")            # region name (kept key for back-compat)
        region_mode = (payload.get("region_mode") or "both").lower()   # freight | comp | both
        fr = db.get_freight(region) if region else None
        region_adjs = []
        if fr:
            if region_mode in ("freight", "both") and fr["freight_cwt"]:
                region_adjs.append({"label": "Freight: " + region, "cwt": fr["freight_cwt"]})
            if region_mode in ("comp", "both") and fr["comp_cwt"]:
                region_adjs.append({"label": "Region (competitive): " + region, "cwt": fr["comp_cwt"]})
        region_adjs += db.customer_adjustments(payload.get("customer_name"))   # playbook + credit (stack)
        min_spread = costing.get_min_spread()

        def price(line, tier_weight=None):
            if line.get("offsheet") or not line.get("key"):
                return _offsheet_result(line)
            product = db.get_product(line["key"])
            if not product:
                return _offsheet_result(line)   # key no longer in the price list
            inv = inventory.get_inventory_for(
                product["thickness"], product.get("width_num") or product.get("width"),
                product["grade"], line.get("form"), rules) if rules else None
            adj = None
            if use_inv and inv and inv.get("adjustment_cwt"):
                mos = inv.get("months_of_supply")
                tag = (inv.get("form") or inv.get("product_type") or "").upper()
                label = "Inventory" + (" (" + tag + ")" if tag else "") + ": " + inv["position"]
                if mos is not None:
                    label += f" ({mos} mo)"
                adj = {"label": label, "cwt": inv["adjustment_cwt"]}
            res = pe.compute_line(
                product,
                length_in=line.get("length_in"),
                qty=line["qty"],
                unit=line.get("unit", "pieces"),
                custom_length=bool(line.get("custom_length")),
                extras=line.get("extras", []),
                inv_adjustment=adj,
                extra_adjustments=region_adjs,
                tier_weight=tier_weight,
                stock_tiers=stock, length_tiers=length, extras_catalog=extras_catalog,
            )
            res["inventory"] = inv
            mat = costing.get_cost(line["key"])
            if res.get("rate_cwt") is not None and mat is not None:
                freight_c = next((b["cwt"] for b in res["breakdown"] if b["label"].startswith("Freight")), 0.0)
                sell = res["rate_cwt"] - freight_c   # freight is pass-through, not margin
                res["material_cost_cwt"] = round(mat, 2)
                res["spread_cwt"] = round(sell - mat, 2)
                res["spread_pct"] = round((sell - mat) / sell * 100, 1) if sell > 0 else None
                res["below_floor"] = bool(res["spread_pct"] is not None and res["spread_pct"] < min_spread)
            return res

        first = [price(l) for l in lines]
        order_weight = sum(r.get("weight_lb") or 0 for r in first)
        if basis == "order":
            results = []
            for i, l in enumerate(lines):
                r0 = first[i]
                if r0.get("error") or r0.get("manual_quote") or r0.get("off_sheet"):
                    results.append(r0)
                else:
                    results.append(price(l, tier_weight=order_weight))
        else:
            results = first

        grand = sum(r.get("line_total") or 0 for r in results)
        return {"lines": results, "grand_total": round(grand, 2),
                "adder_basis": basis, "order_weight_lb": round(order_weight, 1),
                "inventory_applied": use_inv}


def _offsheet_result(line):
    """A line for an item not on the price sheet: compute its weight, price it at $0,
    and flag it so the rep knows it needs a manual price."""
    th = (line.get("thickness") or "").strip()
    gr = (line.get("grade") or "").strip()
    wd = (line.get("width") or "").strip()
    desc = " / ".join(p for p in [th, gr, (wd + '" wide' if wd else "")] if p) or "custom item"
    unit = (line.get("unit") or "pieces").lower()
    try:
        qty = float(line.get("qty") or 0)
    except (TypeError, ValueError):
        qty = 0.0
    weight, wbasis = None, None
    if qty > 0:
        if unit == "lbs":
            weight, wbasis = qty, f"{qty:,.0f} lb"
        elif unit in ("coil", "tl"):
            weight = qty * pe.COIL_TL_WEIGHT_LB
            wbasis = f"{qty:g} {'coil' if unit == 'coil' else 'TL'} x {pe.COIL_TL_WEIGHT_LB:,} lb"
        else:  # pieces
            th_in = pe.thickness_to_inch(th)
            w_in = pe.width_to_number(wd) if wd else None
            try:
                L = float(line.get("length_in") or 0)
            except (TypeError, ValueError):
                L = 0.0
            if th_in and w_in and L > 0:
                weight = pe.piece_weight_lb(th_in, w_in, L, qty)
                wbasis = f"{qty:g} pc x {w_in:g}\" x {L:g}\""
    return {
        "key": line.get("key") or "(off-sheet)",
        "description": desc,
        "off_sheet": True,
        "manual_quote": False,
        "note": "Not on the price sheet — defaulted to $0, needs a manual price.",
        "unit": unit, "qty": qty,
        "weight_basis": wbasis,
        "weight_lb": round(weight) if weight else None,
        "breakdown": [{"label": "Base price", "cwt": 0.0}],
        "rate_cwt": 0.0, "line_total": 0.0, "inventory": None,
    }


def _inventory_watch(interval=600):
    """Background: re-import inventory whenever the export file is re-saved."""
    while True:
        time.sleep(interval)
        try:
            if inventory.maybe_auto_import():
                print("Inventory auto-refreshed from updated export.")
        except Exception as exc:
            print("Inventory auto-refresh skipped:", exc)


def _warmup():
    """Heavy startup work (live SQL inventory pull, cost derivation). Runs in the
    background so the HTTP server starts listening immediately -- otherwise the
    launcher opens the browser before the port is up and shows a connection error."""
    try:
        if inventory.db_available():
            print("Refreshing inventory live from SQL Server...")
            inventory.import_inventory_from_db()
            print("Inventory loaded from database.")
        elif inventory.maybe_auto_import():
            print("Inventory imported/refreshed from workbook.")
    except Exception as exc:
        print("Inventory not loaded:", exc)
    try:
        costing.ensure_costs()
    except Exception as exc:
        print("Costs not derived:", exc)


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
    if not os.path.exists(db.DB_PATH):
        print("No database found -- importing from workbook...")
        db.import_from_workbook()
    # Bind and start serving FIRST, then warm up inventory/costs in the background.
    httpd = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print("Server listening. Loading inventory in the background...")
    threading.Thread(target=_warmup, daemon=True).start()
    threading.Thread(target=_inventory_watch, daemon=True).start()
    httpd.serve_forever()


if __name__ == "__main__":
    print(f"Pricing App running -> http://localhost:{int(sys.argv[1]) if len(sys.argv) > 1 else 8000}")
    main()
