#!/usr/bin/env python3
"""
WAREHOUSE LOGISTICS TRACKING SYSTEM
-----------------------------------
- Manage products (SKU) and shelf locations
- Receive stock / Issue stock / Transfer between locations
- Manage shipments with status flow and tracking history
- Low-stock alerts, stock reports, movement history
- Import stock from CSV / Excel (.xlsx) files
- Storage via SQLite (Python standard library; Excel import needs `pip install openpyxl`)

Run:  python warehouse_tracker.py                        -> interactive menu
      python warehouse_tracker.py --demo                 -> load sample data, then open menu
      python warehouse_tracker.py --import stock.xlsx    -> import a file (preview + confirm)
      python warehouse_tracker.py --import count.csv --mode set   -> stock count (set quantities)
"""

import argparse
import csv
import os
import sqlite3
import sys
import uuid
from datetime import datetime
from contextlib import contextmanager

# Database lives next to this file, so the menu and the web app share the same data
DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "warehouse.db")

SHIPMENT_STATUSES = [
    "CREATED",      # Order created
    "PICKING",      # Items being picked
    "PACKED",       # Packed
    "SHIPPED",      # Left the warehouse
    "IN_TRANSIT",   # On the way
    "DELIVERED",    # Delivered
    "CANCELLED",    # Cancelled
]

# Allowed status transitions
VALID_TRANSITIONS = {
    "CREATED": {"PICKING", "CANCELLED"},
    "PICKING": {"PACKED", "CANCELLED"},
    "PACKED": {"SHIPPED", "CANCELLED"},
    "SHIPPED": {"IN_TRANSIT", "DELIVERED"},
    "IN_TRANSIT": {"DELIVERED"},
    "DELIVERED": set(),
    "CANCELLED": set(),
}


class WarehouseError(Exception):
    pass


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------- File import helpers
# Header names the importer recognises (case, spaces, "_" and "-" don't matter)
COLUMN_ALIASES = {
    "sku": {"sku", "item code", "code", "product code", "item id", "product id"},
    "name": {"name", "product", "product name", "item", "item name", "description"},
    "qty": {"qty", "quantity", "amount", "count", "stock", "units"},
    "location": {"location", "loc", "bin", "shelf", "location code"},
    "unit": {"unit", "uom"},
    "min_stock": {"min stock", "min", "minimum", "minimum stock", "reorder level", "reorder point"},
    "zone": {"zone", "area"},
    "capacity": {"capacity", "location capacity"},
    "ref": {"ref", "reference", "po", "po number", "note", "notes"},
}
REQUIRED_COLUMNS = {"sku", "qty", "location"}


def _normalize_header(h):
    return " ".join(str(h or "").strip().lower().replace("_", " ").replace("-", " ").split())


def _cell_str(v):
    """Turn a cell into clean text (Excel gives 1001.0 for the number 1001)."""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _to_int(v, field):
    if isinstance(v, bool):
        raise WarehouseError(f"{field} must be a number (got '{v}')")
    if isinstance(v, int):
        return v
    try:
        f = float(str(v).replace(",", "").strip())
    except ValueError:
        raise WarehouseError(f"{field} must be a number (got '{v}')")
    if not f.is_integer():
        raise WarehouseError(f"{field} must be a whole number (got '{v}')")
    return int(f)


def read_rows(path):
    """Read a .csv or .xlsx file.
    Returns (rows, ignored_headers) where rows = [(row_number_in_file, {field: value}), ...]."""
    if not os.path.isfile(path):
        raise WarehouseError(f"File not found: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as f:
            raw = list(csv.reader(f))
    elif ext in (".xlsx", ".xlsm"):
        try:
            from openpyxl import load_workbook
        except ImportError:
            raise WarehouseError("Reading Excel files needs openpyxl. Install it with: pip install openpyxl")
        wb = load_workbook(path, read_only=True, data_only=True)
        raw = [list(r) for r in wb.active.iter_rows(values_only=True)]  # first sheet
        wb.close()
    elif ext == ".xls":
        raise WarehouseError("The old .xls format isn't supported. In Excel, use Save As -> .xlsx or .csv")
    else:
        raise WarehouseError(f"Unsupported file type '{ext}'. Use .csv or .xlsx")

    # Header = first non-empty row
    start = next((i for i, r in enumerate(raw) if any(_cell_str(c) for c in r if c is not None)), None)
    if start is None:
        raise WarehouseError("The file is empty")
    mapping, ignored = {}, []
    for idx, h in enumerate(raw[start]):
        n = _normalize_header(h)
        if not n:
            continue
        key = next((k for k, al in COLUMN_ALIASES.items() if n in al and k not in mapping.values()), None)
        if key:
            mapping[idx] = key
        else:
            ignored.append(str(h))
    missing = REQUIRED_COLUMNS - set(mapping.values())
    if missing:
        found = ", ".join(str(h) for h in raw[start] if h not in (None, ""))
        raise WarehouseError(
            f"Missing required column(s): {', '.join(sorted(missing))}. Headers found: {found}"
        )

    rows = []
    for i, r in enumerate(raw[start + 1:], start=start + 2):  # +2 = 1-based row numbers like Excel
        rec = {}
        for idx, key in mapping.items():
            v = r[idx] if idx < len(r) else None
            if v is not None and _cell_str(v) != "":
                rec[key] = v
        if rec:  # skip blank rows
            rows.append((i, rec))
    return rows, ignored


class Warehouse:
    def __init__(self, db_file=DB_FILE):
        self.conn = sqlite3.connect(db_file)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self._depth = 0
        self._init_schema()

    # ------------------------------------------------------------------ DB
    @contextmanager
    def tx(self):
        """Transaction. Can be nested: only the outermost block commits / rolls back."""
        self._depth += 1
        try:
            yield self.conn
            if self._depth == 1:
                self.conn.commit()
        except Exception:
            if self._depth == 1:
                self.conn.rollback()
            raise
        finally:
            self._depth -= 1

    def _init_schema(self):
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS products (
                sku TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                unit TEXT DEFAULT 'pcs',
                min_stock INTEGER DEFAULT 0,
                created_at TEXT
            );
            CREATE TABLE IF NOT EXISTS locations (
                code TEXT PRIMARY KEY,          -- e.g. A-01-03 (zone-aisle-level)
                zone TEXT,
                capacity INTEGER DEFAULT 1000
            );
            CREATE TABLE IF NOT EXISTS inventory (
                sku TEXT REFERENCES products(sku),
                location TEXT REFERENCES locations(code),
                qty INTEGER NOT NULL DEFAULT 0 CHECK (qty >= 0),
                PRIMARY KEY (sku, location)
            );
            CREATE TABLE IF NOT EXISTS movements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ts TEXT,
                type TEXT,                      -- IN / OUT / TRANSFER
                sku TEXT,
                qty INTEGER,
                from_loc TEXT,
                to_loc TEXT,
                ref TEXT,                       -- order ref / note
                user TEXT
            );
            CREATE TABLE IF NOT EXISTS shipments (
                id TEXT PRIMARY KEY,
                customer TEXT,
                address TEXT,
                carrier TEXT,
                tracking_no TEXT,
                status TEXT,
                created_at TEXT,
                updated_at TEXT
            );
            CREATE TABLE IF NOT EXISTS shipment_items (
                shipment_id TEXT REFERENCES shipments(id),
                sku TEXT REFERENCES products(sku),
                qty INTEGER,
                PRIMARY KEY (shipment_id, sku)
            );
            CREATE TABLE IF NOT EXISTS shipment_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shipment_id TEXT REFERENCES shipments(id),
                ts TEXT,
                status TEXT,
                note TEXT
            );
            """
        )
        self.conn.commit()

    # ------------------------------------------------------- Master data
    def add_product(self, sku, name, unit="pcs", min_stock=0):
        with self.tx() as c:
            c.execute(
                "INSERT INTO products VALUES (?,?,?,?,?)",
                (sku.upper(), name, unit, int(min_stock), now()),
            )

    def add_location(self, code, zone="", capacity=1000):
        with self.tx() as c:
            c.execute(
                "INSERT INTO locations VALUES (?,?,?)",
                (code.upper(), zone, int(capacity)),
            )

    def _check_product(self, sku):
        if not self.conn.execute("SELECT 1 FROM products WHERE sku=?", (sku,)).fetchone():
            raise WarehouseError(f"SKU {sku} does not exist")

    def _check_location(self, loc):
        row = self.conn.execute("SELECT capacity FROM locations WHERE code=?", (loc,)).fetchone()
        if not row:
            raise WarehouseError(f"Location {loc} does not exist")
        return row["capacity"]

    def _loc_used(self, loc):
        return self.conn.execute(
            "SELECT COALESCE(SUM(qty),0) FROM inventory WHERE location=?", (loc,)
        ).fetchone()[0]

    def _qty_at(self, sku, loc):
        row = self.conn.execute(
            "SELECT qty FROM inventory WHERE sku=? AND location=?", (sku, loc)
        ).fetchone()
        return row["qty"] if row else 0

    def _log(self, mtype, sku, qty, from_loc=None, to_loc=None, ref="", user="system"):
        self.conn.execute(
            "INSERT INTO movements (ts,type,sku,qty,from_loc,to_loc,ref,user) VALUES (?,?,?,?,?,?,?,?)",
            (now(), mtype, sku, qty, from_loc, to_loc, ref, user),
        )

    # ------------------------------------------------------- Stock moves
    def receive(self, sku, qty, location, ref="", user="system"):
        """Receive stock into a location."""
        sku, location, qty = sku.upper(), location.upper(), int(qty)
        if qty <= 0:
            raise WarehouseError("Quantity must be > 0")
        self._check_product(sku)
        cap = self._check_location(location)
        if self._loc_used(location) + qty > cap:
            raise WarehouseError(f"Location {location} would exceed capacity ({cap})")
        with self.tx() as c:
            c.execute(
                """INSERT INTO inventory (sku,location,qty) VALUES (?,?,?)
                   ON CONFLICT(sku,location) DO UPDATE SET qty = qty + excluded.qty""",
                (sku, location, qty),
            )
            self._log("IN", sku, qty, to_loc=location, ref=ref, user=user)

    def issue(self, sku, qty, location=None, ref="", user="system"):
        """Issue stock. If no location is given, pick from locations with the most stock first."""
        sku, qty = sku.upper(), int(qty)
        if qty <= 0:
            raise WarehouseError("Quantity must be > 0")
        self._check_product(sku)
        if location:
            plan = [(location.upper(), qty)]
            if self._qty_at(sku, location.upper()) < qty:
                raise WarehouseError(f"Not enough {sku} at {location}")
        else:
            if self.total_stock(sku) < qty:
                raise WarehouseError(f"Not enough stock of {sku} ({self.total_stock(sku)} left)")
            plan, remaining = [], qty
            for r in self.conn.execute(
                "SELECT location, qty FROM inventory WHERE sku=? AND qty>0 ORDER BY qty DESC", (sku,)
            ):
                take = min(remaining, r["qty"])
                plan.append((r["location"], take))
                remaining -= take
                if remaining == 0:
                    break
        with self.tx() as c:
            for loc, q in plan:
                c.execute(
                    "UPDATE inventory SET qty = qty - ? WHERE sku=? AND location=?", (q, sku, loc)
                )
                self._log("OUT", sku, q, from_loc=loc, ref=ref, user=user)
            c.execute("DELETE FROM inventory WHERE qty = 0")
        return plan

    def transfer(self, sku, qty, from_loc, to_loc, user="system"):
        """Move stock between two locations."""
        sku, from_loc, to_loc, qty = sku.upper(), from_loc.upper(), to_loc.upper(), int(qty)
        cap = self._check_location(to_loc)
        self._check_location(from_loc)
        if self._qty_at(sku, from_loc) < qty:
            raise WarehouseError(f"Not enough {sku} at {from_loc}")
        if self._loc_used(to_loc) + qty > cap:
            raise WarehouseError(f"Location {to_loc} would exceed capacity ({cap})")
        with self.tx() as c:
            c.execute("UPDATE inventory SET qty = qty - ? WHERE sku=? AND location=?", (qty, sku, from_loc))
            c.execute(
                """INSERT INTO inventory (sku,location,qty) VALUES (?,?,?)
                   ON CONFLICT(sku,location) DO UPDATE SET qty = qty + excluded.qty""",
                (sku, to_loc, qty),
            )
            c.execute("DELETE FROM inventory WHERE qty = 0")
            self._log("TRANSFER", sku, qty, from_loc, to_loc, user=user)

    def set_stock(self, sku, location, qty, ref="", user="system"):
        """Stock count: set the quantity at a location to the counted number.
        The difference is logged as an ADJUST movement."""
        sku, location, qty = sku.upper(), location.upper(), int(qty)
        if qty < 0:
            raise WarehouseError("Counted quantity can't be negative")
        self._check_product(sku)
        cap = self._check_location(location)
        diff = qty - self._qty_at(sku, location)
        if diff == 0:
            return 0
        if self._loc_used(location) + diff > cap:
            raise WarehouseError(f"Location {location} would exceed capacity ({cap})")
        with self.tx() as c:
            c.execute(
                """INSERT INTO inventory (sku,location,qty) VALUES (?,?,?)
                   ON CONFLICT(sku,location) DO UPDATE SET qty = excluded.qty""",
                (sku, location, qty),
            )
            c.execute("DELETE FROM inventory WHERE qty = 0")
            self._log("ADJUST", sku, diff, to_loc=location, ref=ref, user=user)
        return diff

    # ------------------------------------------------------- File import
    def import_rows(self, rows, mode="receive", user="import"):
        """Apply rows from read_rows().
        mode='receive' -> add quantities to stock (goods arriving)
        mode='set'     -> set quantities to the counted number (stock count)
        Each row is all-or-nothing; bad rows are reported and skipped."""
        res = {"ok": 0, "units": 0, "new_products": [], "new_locations": [], "errors": []}
        for rownum, rec in rows:
            row_new = {"new_products": [], "new_locations": []}
            try:
                with self.tx():
                    units = self._import_one(rec, mode, user, row_new)
                res["ok"] += 1
                res["units"] += units
                res["new_products"] += row_new["new_products"]    # only count if the row succeeded
                res["new_locations"] += row_new["new_locations"]
            except (WarehouseError, sqlite3.Error) as e:
                res["errors"].append((rownum, str(e)))
        return res

    def _import_one(self, rec, mode, user, res):
        sku = _cell_str(rec.get("sku", "")).upper()
        loc = _cell_str(rec.get("location", "")).upper()
        if not sku:
            raise WarehouseError("SKU is empty")
        if not loc:
            raise WarehouseError("Location is empty")
        if "qty" not in rec:
            raise WarehouseError("Quantity is empty")
        qty = _to_int(rec["qty"], "Quantity")
        if mode == "receive" and qty <= 0:
            raise WarehouseError("Quantity must be > 0")
        min_stock = _to_int(rec["min_stock"], "Min stock") if "min_stock" in rec else None

        # Product: create if new, otherwise update any details given in the file
        if not self.conn.execute("SELECT 1 FROM products WHERE sku=?", (sku,)).fetchone():
            if "name" not in rec:
                raise WarehouseError(f"New SKU {sku} needs a name")
            self.add_product(sku, _cell_str(rec["name"]), _cell_str(rec.get("unit", "pcs")), min_stock or 0)
            res["new_products"].append(sku)
        else:
            for field, val in (("name", rec.get("name")), ("unit", rec.get("unit"))):
                if val is not None:
                    self.conn.execute(f"UPDATE products SET {field}=? WHERE sku=?", (_cell_str(val), sku))
            if min_stock is not None:
                self.conn.execute("UPDATE products SET min_stock=? WHERE sku=?", (min_stock, sku))

        # Location: create if new
        if not self.conn.execute("SELECT 1 FROM locations WHERE code=?", (loc,)).fetchone():
            cap = _to_int(rec["capacity"], "Capacity") if "capacity" in rec else 1000
            self.add_location(loc, _cell_str(rec.get("zone", "")), cap)
            res["new_locations"].append(loc)

        ref = _cell_str(rec.get("ref", "")) or "IMPORT"
        if mode == "set":
            return abs(self.set_stock(sku, loc, qty, ref, user))
        self.receive(sku, qty, loc, ref, user)
        return qty

    def preview_import(self, rows, mode="receive"):
        """Dry run on an in-memory copy of the database. Nothing real is changed."""
        tmp = Warehouse(":memory:")
        self.conn.backup(tmp.conn)
        try:
            return tmp.import_rows(rows, mode)
        finally:
            tmp.conn.close()

    # ------------------------------------------------------- Shipments
    def create_shipment(self, customer, address, items, carrier=""):
        """items: list[(sku, qty)]. Checks availability before creating the order."""
        for sku, qty in items:
            self._check_product(sku.upper())
            if self.available_stock(sku.upper()) < int(qty):
                raise WarehouseError(f"Not enough available stock for {sku}")
        sid = "SHP-" + uuid.uuid4().hex[:8].upper()
        with self.tx() as c:
            c.execute(
                "INSERT INTO shipments VALUES (?,?,?,?,?,?,?,?)",
                (sid, customer, address, carrier, "", "CREATED", now(), now()),
            )
            for sku, qty in items:
                c.execute("INSERT INTO shipment_items VALUES (?,?,?)", (sid, sku.upper(), int(qty)))
            c.execute(
                "INSERT INTO shipment_events (shipment_id,ts,status,note) VALUES (?,?,?,?)",
                (sid, now(), "CREATED", "Order created"),
            )
        return sid

    def update_shipment(self, sid, new_status, note="", tracking_no=None):
        sid, new_status = sid.upper(), new_status.upper()
        row = self.conn.execute("SELECT status FROM shipments WHERE id=?", (sid,)).fetchone()
        if not row:
            raise WarehouseError(f"Shipment {sid} not found")
        cur = row["status"]
        if new_status not in VALID_TRANSITIONS[cur]:
            raise WarehouseError(
                f"Cannot change {cur} -> {new_status}. Allowed: {sorted(VALID_TRANSITIONS[cur]) or 'none'}"
            )
        # When the goods physically leave the warehouse -> deduct stock
        if new_status == "SHIPPED":
            for it in self.conn.execute("SELECT sku, qty FROM shipment_items WHERE shipment_id=?", (sid,)).fetchall():
                self.issue(it["sku"], it["qty"], ref=sid)
        with self.tx() as c:
            c.execute(
                "UPDATE shipments SET status=?, updated_at=?, tracking_no=COALESCE(?, tracking_no) WHERE id=?",
                (new_status, now(), tracking_no, sid),
            )
            c.execute(
                "INSERT INTO shipment_events (shipment_id,ts,status,note) VALUES (?,?,?,?)",
                (sid, now(), new_status, note),
            )

    def track(self, sid):
        sid = sid.upper()
        s = self.conn.execute("SELECT * FROM shipments WHERE id=?", (sid,)).fetchone()
        if not s:
            raise WarehouseError(f"Shipment {sid} not found")
        items = self.conn.execute(
            """SELECT si.sku, p.name, si.qty FROM shipment_items si
               JOIN products p ON p.sku = si.sku WHERE shipment_id=?""", (sid,)
        ).fetchall()
        events = self.conn.execute(
            "SELECT ts, status, note FROM shipment_events WHERE shipment_id=? ORDER BY id", (sid,)
        ).fetchall()
        return s, items, events

    # ------------------------------------------------------- Reports
    def total_stock(self, sku):
        return self.conn.execute(
            "SELECT COALESCE(SUM(qty),0) FROM inventory WHERE sku=?", (sku,)
        ).fetchone()[0]

    def reserved_stock(self, sku):
        """Stock allocated to open shipments but not yet shipped."""
        return self.conn.execute(
            """SELECT COALESCE(SUM(si.qty),0) FROM shipment_items si
               JOIN shipments s ON s.id = si.shipment_id
               WHERE si.sku=? AND s.status IN ('CREATED','PICKING','PACKED')""", (sku,)
        ).fetchone()[0]

    def available_stock(self, sku):
        return self.total_stock(sku) - self.reserved_stock(sku)

    def stock_report(self):
        return self.conn.execute(
            """SELECT p.sku, p.name, p.unit, p.min_stock,
                      COALESCE(SUM(i.qty),0) AS total
               FROM products p LEFT JOIN inventory i ON i.sku = p.sku
               GROUP BY p.sku ORDER BY p.sku"""
        ).fetchall()

    def stock_by_location(self, sku=None):
        q = "SELECT sku, location, qty FROM inventory"
        args = ()
        if sku:
            q += " WHERE sku=?"
            args = (sku.upper(),)
        return self.conn.execute(q + " ORDER BY location, sku", args).fetchall()

    def low_stock(self):
        return [r for r in self.stock_report() if r["total"] <= r["min_stock"]]

    def location_usage(self):
        return self.conn.execute(
            """SELECT l.code, l.zone, l.capacity, COALESCE(SUM(i.qty),0) AS used
               FROM locations l LEFT JOIN inventory i ON i.location = l.code
               GROUP BY l.code ORDER BY l.code"""
        ).fetchall()

    def movement_history(self, sku=None, limit=20):
        q, args = "SELECT * FROM movements", []
        if sku:
            q += " WHERE sku=?"
            args.append(sku.upper())
        q += " ORDER BY id DESC LIMIT ?"
        args.append(limit)
        return self.conn.execute(q, args).fetchall()

    def list_shipments(self, status=None):
        q, args = "SELECT * FROM shipments", ()
        if status:
            q += " WHERE status=?"
            args = (status.upper(),)
        return self.conn.execute(q + " ORDER BY created_at DESC", args).fetchall()


# ====================================================================== CLI
def print_table(rows, headers):
    rows = [[("" if v is None else str(v)) for v in r] for r in rows]
    widths = [max(len(h), *(len(r[i]) for r in rows)) if rows else len(h) for i, h in enumerate(headers)]
    line = "+".join("-" * (w + 2) for w in widths)
    print(line)
    print(" | ".join(h.ljust(w) for h, w in zip(headers, widths)))
    print(line)
    for r in rows:
        print(" | ".join(v.ljust(w) for v, w in zip(r, widths)))
    print(line)
    if not rows:
        print("(empty)")


def ask(prompt, default=None):
    v = input(f"{prompt}{f' [{default}]' if default is not None else ''}: ").strip()
    return v if v else (default if default is not None else "")


def seed_demo(wh: Warehouse):
    if wh.stock_report():
        return
    for code, zone in [("A-01-01", "A"), ("A-01-02", "A"), ("B-02-01", "B"), ("C-01-01", "Cold")]:
        wh.add_location(code, zone, 500)
    wh.add_product("SKU001", "Dell Laptop", "pcs", 10)
    wh.add_product("SKU002", "Logitech Mouse", "pcs", 50)
    wh.add_product("SKU003", "Milk Carton", "case", 20)
    wh.receive("SKU001", 40, "A-01-01", "PO-1001")
    wh.receive("SKU002", 120, "A-01-02", "PO-1001")
    wh.receive("SKU003", 15, "C-01-01", "PO-1002")
    wh.transfer("SKU002", 30, "A-01-02", "B-02-01")
    sid = wh.create_shipment("ABC Company", "123 Main St, San Diego", [("SKU001", 5), ("SKU002", 10)], "UPS")
    wh.update_shipment(sid, "PICKING", "Started picking")
    wh.update_shipment(sid, "PACKED", "Packing done")
    wh.update_shipment(sid, "SHIPPED", "Handed to UPS", tracking_no="1Z999AA10123456784")
    print(f"✅ Sample data loaded. Sample shipment: {sid}")


def show_import_summary(res, title):
    print(f"\n{title}: {res['ok']} row(s) OK, {len(res['errors'])} row(s) with errors, "
          f"{res['units']} unit(s) changed")
    if res["new_products"]:
        print("  New products :", ", ".join(dict.fromkeys(res["new_products"])))
    if res["new_locations"]:
        print("  New locations:", ", ".join(dict.fromkeys(res["new_locations"])))
    if res["errors"]:
        print_table(res["errors"], ["Row", "Problem"])


def run_import(wh, path, mode="receive", assume_yes=False):
    """Read file -> preview -> confirm -> apply."""
    rows, ignored = read_rows(path)
    label = "Stock count (set quantities)" if mode == "set" else "Receive (add quantities)"
    print(f"\nRead {len(rows)} row(s) from {os.path.basename(path)}  |  Mode: {label}")
    if ignored:
        print("  Ignored columns:", ", ".join(ignored))
    preview = wh.preview_import(rows, mode)
    show_import_summary(preview, "PREVIEW")
    if preview["ok"] == 0:
        print("Nothing to import.")
        return None
    if not assume_yes:
        extra = " (rows with errors will be skipped)" if preview["errors"] else ""
        if ask(f"Import {preview['ok']} row(s){extra}? y/n", "n").lower() != "y":
            print("Cancelled. Nothing was changed.")
            return None
    result = wh.import_rows(rows, mode)
    show_import_summary(result, "✅ IMPORTED")
    return result


MENU = """
======== WAREHOUSE SYSTEM ========
 1. Add product            2. Add location
 3. Receive stock          4. Issue stock
 5. Transfer stock         6. Stock report
 7. Stock by location      8. Low-stock alerts
 9. Location capacity     10. Movement history
11. Create shipment       12. Update shipment status
13. Track shipment        14. List shipments
15. Import from CSV / Excel file
 0. Exit
=================================="""


def main():
    parser = argparse.ArgumentParser(description="Warehouse logistics tracking system")
    parser.add_argument("--demo", action="store_true", help="load sample data")
    parser.add_argument("--import", dest="import_file", metavar="FILE", help="import a .csv or .xlsx file")
    parser.add_argument("--mode", choices=["receive", "set"], default="receive",
                        help="receive = add quantities (default), set = stock count")
    parser.add_argument("--yes", action="store_true", help="import without asking for confirmation")
    args = parser.parse_args()

    wh = Warehouse()
    if args.demo:
        seed_demo(wh)
    if args.import_file:
        try:
            run_import(wh, args.import_file, args.mode, args.yes)
        except WarehouseError as e:
            print(f"❌ Error: {e}")
            sys.exit(1)
        return

    actions = {}

    def a1():
        wh.add_product(ask("SKU"), ask("Name"), ask("Unit", "pcs"), ask("Minimum stock", "0"))
        print("✅ Product added")

    def a2():
        wh.add_location(ask("Location code (e.g. A-01-01)"), ask("Zone"), ask("Capacity", "1000"))
        print("✅ Location added")

    def a3():
        wh.receive(ask("SKU"), ask("Quantity"), ask("Location"), ask("PO / note", ""))
        print("✅ Stock received")

    def a4():
        plan = wh.issue(ask("SKU"), ask("Quantity"), ask("Location (blank = auto-pick)", "") or None,
                        ask("Reference", ""))
        print("✅ Issued:", ", ".join(f"{q} from {l}" for l, q in plan))

    def a5():
        wh.transfer(ask("SKU"), ask("Quantity"), ask("From location"), ask("To location"))
        print("✅ Transferred")

    def a6():
        rows = [(r["sku"], r["name"], r["total"], wh.reserved_stock(r["sku"]),
                 wh.available_stock(r["sku"]), r["unit"], r["min_stock"]) for r in wh.stock_report()]
        print_table(rows, ["SKU", "Name", "On hand", "Reserved", "Available", "Unit", "Min"])

    def a7():
        print_table([tuple(r) for r in wh.stock_by_location(ask("SKU (blank = all)", "") or None)],
                    ["SKU", "Location", "Qty"])

    def a8():
        print_table([(r["sku"], r["name"], r["total"], r["min_stock"]) for r in wh.low_stock()],
                    ["SKU", "Name", "On hand", "Min"])

    def a9():
        print_table([(r["code"], r["zone"], r["used"], r["capacity"],
                      f"{100 * r['used'] / r['capacity']:.0f}%") for r in wh.location_usage()],
                    ["Location", "Zone", "Used", "Capacity", "%"])

    def a10():
        rows = wh.movement_history(ask("SKU (blank = all)", "") or None)
        print_table([(r["ts"], r["type"], r["sku"], r["qty"], r["from_loc"], r["to_loc"], r["ref"]) for r in rows],
                    ["Time", "Type", "SKU", "Qty", "From", "To", "Ref"])

    def a11():
        items = []
        print("Enter items (leave SKU blank to finish):")
        while True:
            sku = ask("  SKU", "")
            if not sku:
                break
            items.append((sku, int(ask("  Quantity"))))
        if not items:
            print("No items entered")
            return
        sid = wh.create_shipment(ask("Customer"), ask("Address"), items, ask("Carrier", ""))
        print(f"✅ Shipment {sid} created")

    def a12():
        sid = ask("Shipment ID")
        print("Statuses:", ", ".join(SHIPMENT_STATUSES))
        st = ask("New status").upper()
        tn = ask("Tracking number (optional)", "") or None
        wh.update_shipment(sid, st, ask("Note", ""), tn)
        print("✅ Updated")

    def a13():
        s, items, events = wh.track(ask("Shipment ID"))
        print(f"\n📦 {s['id']} | {s['customer']} | {s['address']}")
        print(f"   Carrier: {s['carrier']} {s['tracking_no'] or ''} | Status: {s['status']}")
        print_table([tuple(i) for i in items], ["SKU", "Name", "Qty"])
        print_table([tuple(e) for e in events], ["Time", "Status", "Note"])

    def a14():
        rows = wh.list_shipments(ask("Filter by status (blank = all)", "") or None)
        print_table([(r["id"], r["customer"], r["carrier"], r["tracking_no"], r["status"], r["updated_at"])
                     for r in rows], ["Shipment", "Customer", "Carrier", "Tracking #", "Status", "Updated"])

    def a15():
        path = ask("File path (.csv or .xlsx)").strip().strip('"').strip("'")  # handles drag-and-drop quotes
        print("Mode: 1 = Receive (add quantities to stock)   2 = Stock count (set to counted quantities)")
        run_import(wh, path, "set" if ask("Mode", "1") == "2" else "receive")

    for i, f in enumerate([a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15], 1):
        actions[str(i)] = f

    while True:
        print(MENU)
        choice = input("Choose: ").strip()
        if choice == "0":
            break
        fn = actions.get(choice)
        if not fn:
            print("❌ Invalid choice")
            continue
        try:
            fn()
        except (WarehouseError, sqlite3.IntegrityError, ValueError, OSError) as e:
            print(f"❌ Error: {e}")


if __name__ == "__main__":
    main()
