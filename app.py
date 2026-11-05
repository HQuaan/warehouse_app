"""
Warehouse Tracker — local web interface (Flask)

Run:  python app.py      (or press F5 in VS Code)
Then open http://127.0.0.1:5050

All business logic lives in warehouse_tracker.py; this file only turns
web requests into calls on the Warehouse class and renders the pages.
"""

import os
import re
import sqlite3
import time
import uuid
import webbrowser

from flask import Flask, abort, flash, g, redirect, render_template, request, url_for

import warehouse_tracker as wt
from warehouse_tracker import SHIPMENT_STATUSES, VALID_TRANSITIONS, Warehouse, WarehouseError

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
UPLOAD_NAME = re.compile(r"^[0-9a-f]{32}\.(csv|xlsx|xlsm)$")
PORT = int(os.environ.get("PORT", 5050))

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "local-dev-only")
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MB upload limit


# ------------------------------------------------------------------ helpers
def wh() -> Warehouse:
    """One database connection per request."""
    if "wh" not in g:
        g.wh = Warehouse(wt.DB_FILE)
    return g.wh


@app.teardown_appcontext
def close_db(_exc):
    w = g.pop("wh", None)
    if w is not None:
        w.conn.close()


@app.context_processor
def inject_globals():
    endpoint = request.endpoint or ""
    return {"section": endpoint.split("_")[0], "statuses": SHIPMENT_STATUSES}


def attempt(fn, ok=None):
    """Run a warehouse action (a zero-argument function), flash a success or error message,
    and return the result (or None if it failed)."""
    try:
        result = fn()
        if ok:
            flash(ok, "success")
        return result if result is not None else True
    except WarehouseError as e:
        flash(str(e), "error")
    except (ValueError, TypeError):
        flash("Please enter whole numbers for quantities and capacities.", "error")
    except sqlite3.IntegrityError:
        flash("That code already exists.", "error")
    return None


def form(name, default=""):
    return request.form.get(name, default).strip()


def num(name, default=None):
    """Read a whole number from the form (raises ValueError if it isn't one)."""
    v = form(name)
    return int(v) if v else default


def products_with_availability():
    w = wh()
    return [
        {"sku": r["sku"], "name": r["name"], "unit": r["unit"], "total": r["total"],
         "available": w.available_stock(r["sku"])}
        for r in w.stock_report()
    ]


# ---------------------------------------------------------------- dashboard
@app.route("/")
def dashboard():
    w = wh()
    report = w.stock_report()
    open_shipments = [s for s in w.list_shipments() if s["status"] not in ("DELIVERED", "CANCELLED")]
    return render_template(
        "dashboard.html",
        empty=not report,
        n_products=len(report),
        total_units=sum(r["total"] for r in report),
        low=w.low_stock(),
        locations=w.location_usage(),
        open_shipments=open_shipments,
        recent=w.movement_history(limit=8),
    )


@app.post("/demo")
def demo():
    if wh().stock_report():
        flash("Sample data is only loaded into an empty warehouse.", "error")
    else:
        attempt(lambda: wt.seed_demo(wh()), ok="Sample data loaded.")
    return redirect(url_for("dashboard"))


# ---------------------------------------------------------------- inventory
@app.route("/inventory")
def inventory():
    q = request.args.get("q", "").strip().lower()
    sku = request.args.get("sku", "").strip().upper() or None
    rows = [p for p in products_with_availability()
            if not q or q in p["sku"].lower() or q in p["name"].lower()]
    mins = {r["sku"]: r["min_stock"] for r in wh().stock_report()}
    return render_template("inventory.html", rows=rows, mins=mins, q=q,
                           by_location=wh().stock_by_location(sku), sku=sku)


# ---------------------------------------------------------------- stock moves
@app.route("/moves")
def moves():
    return render_template("moves.html", products=products_with_availability(),
                           locations=wh().location_usage())


@app.post("/moves/receive")
def moves_receive():
    attempt(lambda: wh().receive(form("sku"), num("qty"), form("location"), form("ref"), user="web"),
            ok=f"Received {form('qty')} × {form('sku')} into {form('location')}.")
    return redirect(url_for("moves"))


@app.post("/moves/issue")
def moves_issue():
    plan = attempt(lambda: wh().issue(form("sku"), num("qty"), form("location") or None,
                                      form("ref"), user="web"))
    if plan:
        flash("Issued " + ", ".join(f"{q} from {loc}" for loc, q in plan) + f" ({form('sku')}).", "success")
    return redirect(url_for("moves"))


@app.post("/moves/transfer")
def moves_transfer():
    attempt(lambda: wh().transfer(form("sku"), num("qty"), form("from_loc"), form("to_loc"), user="web"),
            ok=f"Moved {form('qty')} × {form('sku')} from {form('from_loc')} to {form('to_loc')}.")
    return redirect(url_for("moves"))


@app.post("/moves/product")
def moves_product():
    attempt(lambda: wh().add_product(form("sku"), form("name"), form("unit") or "pcs", num("min_stock", 0)),
            ok=f"Product {form('sku').upper()} added.")
    return redirect(url_for("moves"))


@app.post("/moves/location")
def moves_location():
    attempt(lambda: wh().add_location(form("code"), form("zone"), num("capacity", 1000)),
            ok=f"Location {form('code').upper()} added.")
    return redirect(url_for("moves"))


# ---------------------------------------------------------------- shipments
@app.route("/shipments")
def shipments():
    status = request.args.get("status") or None
    return render_template("shipments.html", shipments=wh().list_shipments(status), status=status,
                           products=products_with_availability())


@app.post("/shipments/create")
def shipments_create():
    items = [(sku, qty) for sku, qty in zip(request.form.getlist("sku"), request.form.getlist("qty"))
             if sku.strip() and qty.strip()]
    if not items:
        flash("Add at least one item.", "error")
        return redirect(url_for("shipments"))
    sid = attempt(lambda: wh().create_shipment(form("customer"), form("address"),
                                               [(sku, int(q)) for sku, q in items], form("carrier")))
    if sid:
        flash(f"Shipment {sid} created.", "success")
        return redirect(url_for("shipments_detail", sid=sid))
    return redirect(url_for("shipments"))


@app.route("/shipments/<sid>")
def shipments_detail(sid):
    try:
        s, items, events = wh().track(sid)
    except WarehouseError:
        abort(404)
    return render_template("shipment_detail.html", s=s, items=items, events=events,
                           next_statuses=sorted(VALID_TRANSITIONS[s["status"]],
                                                key=SHIPMENT_STATUSES.index))


@app.post("/shipments/<sid>/status")
def shipments_status(sid):
    new = form("status")
    attempt(lambda: wh().update_shipment(sid, new, form("note"), form("tracking_no") or None),
            ok=f"Status changed to {new}.")
    return redirect(url_for("shipments_detail", sid=sid))


# ---------------------------------------------------------------- file import
def clean_old_uploads(max_age_seconds=3600):
    """Delete previewed files that were never confirmed."""
    cutoff = time.time() - max_age_seconds
    for name in os.listdir(UPLOAD_DIR):
        p = os.path.join(UPLOAD_DIR, name)
        if os.path.isfile(p) and os.path.getmtime(p) < cutoff:
            os.remove(p)


@app.route("/import")
def imports():
    return render_template("import.html")


@app.post("/import/preview")
def imports_preview():
    f = request.files.get("file")
    mode = "set" if request.form.get("mode") == "set" else "receive"
    if not f or not f.filename:
        flash("Choose a file first.", "error")
        return redirect(url_for("imports"))
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".csv", ".xlsx", ".xlsm"):
        flash("Please upload a .csv or .xlsx file. (For old .xls files, use Save As → .xlsx in Excel.)", "error")
        return redirect(url_for("imports"))
    clean_old_uploads()
    token = uuid.uuid4().hex + ext            # saved under a random name, never the user's filename
    path = os.path.join(UPLOAD_DIR, token)
    f.save(path)
    try:
        rows, ignored = wt.read_rows(path)
        preview = wh().preview_import(rows, mode)
    except Exception as e:                    # bad format, corrupt Excel file, wrong encoding...
        os.remove(path)
        msg = str(e) if isinstance(e, WarehouseError) else \
            "Couldn't read that file. Is it a valid CSV (UTF-8) or Excel .xlsx file?"
        flash(msg, "error")
        return redirect(url_for("imports"))
    errors = dict(preview["errors"])
    return render_template("import.html", preview=preview, rows=rows[:300], n_rows=len(rows),
                           errors=errors, ignored=ignored, token=token, mode=mode, filename=f.filename)


@app.post("/import/apply")
def imports_apply():
    token, mode = form("token"), ("set" if form("mode") == "set" else "receive")
    if not UPLOAD_NAME.match(token):
        abort(400)
    path = os.path.join(UPLOAD_DIR, token)
    if not os.path.exists(path):
        flash("That upload has expired. Please upload the file again.", "error")
        return redirect(url_for("imports"))
    try:
        rows, _ = wt.read_rows(path)
        result = wh().import_rows(rows, mode, user="web-import")
    finally:
        os.remove(path)
    return render_template("import.html", result=result, filename=form("filename"), mode=mode)


# ---------------------------------------------------------------- history
@app.route("/history")
def history():
    sku = request.args.get("sku", "").strip() or None
    mtype = request.args.get("type", "")
    rows = wh().movement_history(sku, limit=500)
    if mtype:
        rows = [r for r in rows if r["type"] == mtype]
    return render_template("history.html", rows=rows, sku=sku or "", mtype=mtype)


@app.errorhandler(404)
def not_found(_e):
    return render_template("base.html", not_found=True), 404


if __name__ == "__main__":
    url = f"http://127.0.0.1:{PORT}"
    # Open the browser once (not again every time the auto-reloader restarts)
    if os.environ.get("WERKZEUG_RUN_MAIN") != "true" and not os.environ.get("NO_BROWSER"):
        webbrowser.open(url)
    print(f"\n  Warehouse Tracker running at {url}  (Ctrl+C to stop)\n")
    app.run(host="127.0.0.1", port=PORT, debug=True)
