"""
Buysheet Portal - Flask backend
================================
Start it with:  python server.py        (portal.html and config.json sit in the same folder as server.py)

The code lives in the sapportal package, one module per job. A module only imports from the ones above it in this list:

  sapportal/config.py         settings and constants: where HANA is, what a size looks like, season codes, the build stamp
  sapportal/helpers.py        small pure functions: numbers, dates, seasons, ratios, PO / requisition numbers
  sapportal/hana.py           the HANA connection and the read-only queries
  sapportal/parsers.py        every file becomes row dicts: buysheets, legacy prebuys, the prebuy template
  sapportal/outputs.py        buysheet DATA + HEADER OUT, the .xlsm template fill, creation templates
  sapportal/validation.py     every row checked against what SAP holds
  sapportal/stock.py          the SAP Cut and Sold stock, validation report, PR CSV
  sapportal/diagnostics.py    what SAP really holds for anything the checks flagged, and the SAP explorer
  sapportal/reqcheck.py       PO / requisition check: SAP vs SAP, the margin file, and the buysheet
  sapportal/webapp.py         the Flask app object
  sapportal/routes.py         the HTTP endpoints, ItemHub, the PPK creation runner
  sapportal/sap_runner.py     the SAP tab: runs the MM41 creation scripts in tools/ from the creation template
  sapportal/delivery_flow.py  split the validated buysheet along a delivery flow or delivery chart

Endpoints:
  GET  /                    serve portal.html
  GET  /health              status + build stamp
  GET  /diagnostics         plain-text JSON for pasting into a chat
  POST /validate            upload a prebuy or buysheet, validate, cache the result
  POST /revalidate          check the loaded file against SAP again (no re-drop)
  GET  /generate            the buysheet: the .xlsm template filled (templates\\Buysheet_Template.xlsm), else a plain DATA workbook
  GET  /creation_templates  GENERIC / PPK / COLOR ADD creation workbook from the cache (?refresh=1 checks SAP again first)
  GET  /validation_report   SUMMARY / GENERIC VALIDATION / PPK VALIDATION / PPK COMPONENTS workbook from the cache
  POST /stock               the SAP Cut and Sold export (Excel / CSV): what is already in stock, joined by article number
  GET  /stock, POST /stock/clear, GET /validation
  GET  /req_csv             purchase-requisition CSV from the cache
  GET  /lookup              a PO or requisition from SAP, checked (and against the buysheet when one is loaded)
  POST /delivery_split      split a delivery file proportionally

Setup on a new machine: copy config.example.json to config.json and fill in your HANA host, user and password.
Nothing secret is kept in the code; the settings can also come from environment variables (HANA_HOST, HANA_PORT,
HANA_USER, HANA_PASSWORD, TOOLS_DIR, STOCK_VIEW) or from a file named by PORTAL_CONFIG. STOCK_VIEW (default
SAPDEMO.ZSD_CUTSOLD_STK) is where live stock is read; OFF switches live stock off.

This file also re-exports every name from the package, so `import server; server.parse_prebuy(...)` keeps working
for the tests and for the small scripts in tools/ that start with `from server import connect_hana`.
"""
import os

# Import order matters: the endpoint modules register their routes on the shared app when they are imported.
from sapportal import (config, helpers, hana, parsers, outputs, validation, stock, diagnostics, reqcheck, webapp, routes, delivery_flow, sap_runner, buysheets, retailportal, dpci_list, hold_check, demo_guard)

for _mod in (config, helpers, hana, parsers, outputs, validation, stock, diagnostics, reqcheck, webapp, routes, delivery_flow, sap_runner):
    globals().update({k: v for k, v in vars(_mod).items() if not k.startswith('__')})
del _mod

if __name__ == '__main__':
    print("SightLine demo: invented data only, answering on localhost only.")
    print("Open in browser:  http://localhost:5050")
    print(f"Build: {BUILD}")
    print("Press Ctrl+C to stop\n")
    app.run(host='127.0.0.1', port=5050, debug=False)
