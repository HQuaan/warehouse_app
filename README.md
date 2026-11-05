# Warehouse Tracker

Track stock, shelf locations and shipments for a warehouse, either in a web browser (local) or in a terminal menu. Both use the same database (`warehouse.db`).

## Project layout

```
warehouse_app/
├── app.py                 ← web app (Flask): routes and pages
├── warehouse_tracker.py   ← all the warehouse logic + terminal menu
├── templates/             ← HTML pages (Jinja templates)
├── static/style.css       ← styling
├── samples/               ← example files for the Import page
├── requirements.txt       ← Python packages to install
└── .vscode/launch.json    ← lets you press F5 to run
```

## First-time setup (in VS Code)

1. **File → Open Folder…** and choose the `warehouse_app` folder.
2. Install the **Python** extension from Microsoft if VS Code asks.
3. Open a terminal: **Terminal → New Terminal**, then create a virtual environment and install the packages:

   **Windows**
   ```powershell
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```
   **macOS / Linux**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```
4. Press **Ctrl+Shift+P** (Cmd+Shift+P on Mac), choose **Python: Select Interpreter**, and pick the one that shows `.venv`.

## Run it

- **Press F5** (or open the Run and Debug panel and pick **Warehouse: Web app**). Your browser opens at **http://127.0.0.1:5050**.
- Or run it from the terminal: `python app.py`
- Stop it with **Ctrl+C** in the terminal, or the red ■ button in VS Code.

On first launch the warehouse is empty. Click **Load sample data** on the dashboard to explore.

While it's running, saving a `.py` or `.html` file reloads the app automatically. Refresh the browser to see changes.

The **Warehouse: Terminal menu** option in the Run panel starts the original text menu instead.

## Pages

| Page | What it's for |
|---|---|
| Dashboard | Totals, low-stock list, how full each shelf is, open shipments, latest movements |
| Inventory | Stock per product (on hand / reserved / available) and per location; search |
| Stock moves | Receive, issue, transfer; add products and locations |
| Shipments | Create orders, filter by status, open one to move it along its status flow |
| Import file | Upload a CSV / Excel file, preview it, then confirm |
| History | Every stock movement, filterable by SKU and type |

## Notes

- The app only listens on `127.0.0.1`, so it's reachable from your own computer only.
- To start over, stop the app and delete `warehouse.db`.
- Port 5050 is busy? Set a different one: `PORT=5060 python app.py` (macOS/Linux) or `set PORT=5060 && python app.py` (Windows cmd).
