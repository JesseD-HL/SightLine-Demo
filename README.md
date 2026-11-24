# SightLine (portfolio copy)

A portfolio copy of a data-operations portal for apparel buying: it checks buying files (open-stock buysheets, hold orders, DPCI lists) against the
item master, finds what is missing or wrong, shows stock on hand, and produces the files that load the fixes.

**Everything here is invented.** The company (Northwind Apparel), brand, customers, styles, prices and stock are made up, and the "SAP" it reads is a
small in-memory SQLite database (`sapportal/demo_db.py`) that mimics the real views. There is no database login, network address, file share or
credential anywhere in this repository, and the original project's history is not part of it.

## What it shows

* **Validation**: every buysheet row checked for existence, color code and name, ZSLS price, ZVP1 MSRP, description, season, pieces per carton and
  rounding, with the numbers behind each check on hover.
* **Hold master**: stacked hold-order files reduced to two lists, color codes the file gets wrong and stock on hand by storage location, with an Excel export.
* **DPCI list**: a retailer item list checked against the item master and turned into creation sheets.
* **Structure**: a layered Python package (`sapportal/`, one module per job, import direction enforced by `tests/test_structure.py`), a single-file React
  page (`portal.html`), and regression tests for every parser layout.

## Run it

```powershell
pip install -r requirements.txt
python demo/make_samples.py     # writes demo/samples/*.xlsx
python server.py                # http://localhost:5050
```

Drop a file from `demo/samples/` on the Validation tab. Tests: `python tests/test_structure.py` (and the other `tests/test_*.py`).

## Safe by construction

`sapportal/demo_guard.py` answers only an allow-list of read-only endpoints, only on localhost (Host header checked), refuses cross-site requests, caps
uploads at 8 MB, sends no-store and anti-framing headers, and never returns a stack trace. The server binds to 127.0.0.1. Nothing in this copy can start a
process, open a browser session, or reach a network system.
