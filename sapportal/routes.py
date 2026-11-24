"""The HTTP endpoints the portal calls, ItemHub, and the PPK creation runner."""

import io
import json
import os
import re
import sys
import traceback
from collections import defaultdict
from datetime import datetime
from flask import Response, jsonify, request, send_file, send_from_directory
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from .config import (
    BUILD,
    HANA_HOST,
    HANA_PORT,
    HANA_USER,
    HERE,
    KNOWN_SIZES,
    SETTINGS_SOURCE,
    STOCK_VIEW,
    TOOLS_DIR,
    _cache)
from .diagnostics import build_diagnostics, explore_sap
from .hana import connect_hana, hana_configured, query_color_codes_by_name, query_variant_articles
from .helpers import NO_SEASON_CODE, NO_SEASON_YEAR, no_season_style, normalize_color, pad10
from .outputs import BFONT, CTR, DFONT, NAVY, WBOLD, XLSM, add_filters, make_buysheet, make_creation_templates
from .parsers import parse_prebuy
from .reqcheck import build_lookup, buysheet_req_check, fetch_lookup_data, parse_margin
from .stock import apply_stock, live_stock, make_req_csv, make_validation_report, parse_cut_sold, stock_info
from .validation import SapView, validate_all
from .webapp import app


# ═══════════════════════════════════════════════════════════════════════════════
# 8. ROUTES
# ═══════════════════════════════════════════════════════════════════════════════
XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

def _send_xlsx(buf, stem):
    return send_file(buf, as_attachment=True, download_name=f"{stem}_{datetime.now():%m-%d-%Y}.xlsx", mimetype=XLSX)

def _from_cache(build, stem):
    """Shared body of the three download routes: need a validation in the cache, build the file, send it."""
    if not _cache.get('rows'):
        return jsonify({'error': 'No data — validate a file first'}), 400
    try:
        return _send_xlsx(build(), stem)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/')
def portal():
    if not os.path.exists(os.path.join(TOOLS_DIR, 'portal.html')):
        return Response(f"portal.html was not found in {TOOLS_DIR}. Put portal.html in the same folder as server.py.", status=404, mimetype='text/plain')
    return send_from_directory(TOOLS_DIR, 'portal.html')

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'build': BUILD, 'cached': bool(_cache.get('validation')), 'time': datetime.now().isoformat(), 'stock_view': STOCK_VIEW or None,
                    'hana': {'configured': hana_configured()}})

@app.route('/diagnostics')
def diagnostics():
    return Response(json.dumps(build_diagnostics(_cache), indent=2, default=str), mimetype='text/plain')

@app.route('/explore')
def explore():
    """Read-only look into HANA.  /explore (the six transactions) | /explore?q=ZPOP,ZPOCOND | /explore?view=<name>  (that object only) | /explore?view=<name>&col=<column>&like=<starts with>  (rows for one article etc.)"""
    terms = [t.upper() for t in re.split(r'[,\s]+', request.args.get('q', '')) if t] or None
    view = request.args.get('view', '').strip()
    col = request.args.get('col', '').strip()
    like = request.args.get('like', '').strip()
    if (col or like) and not (view and col and like) or (col and not re.fullmatch(r'[A-Za-z0-9_]{1,60}', col)) or (like and not re.fullmatch(r'[A-Za-z0-9_.\- ]{1,60}', like)):
        return Response('Filtering needs all of view, col and like, e.g. /explore?view=SOMEVIEW&col=MARA_ARTICLE&like=RDM271SS06F  (like is a prefix; letters, digits, _ . - only).', status=400, mimetype='text/plain')
    if any(not re.fullmatch(r'[A-Z0-9_]{2,40}', t) for t in terms or []) or (view and not re.fullmatch(r'[A-Za-z0-9_.\-/]{2,120}', view)):
        return Response('q takes transaction names like ZPOP,ZPOCOND; view takes a view name from the first section.', status=400, mimetype='text/plain')
    try:
        conn = connect_hana()
    except Exception as exc:
        return Response(f'Could not connect to HANA: {exc}', status=502, mimetype='text/plain')
    try:
        return Response(json.dumps(explore_sap(conn, terms, view or None, col or None, like or None), indent=2, default=str), mimetype='text/plain')
    finally:
        conn.close()

@app.route('/lookup', methods=['GET'])
def lookup():
    number, kind = request.args.get('q', '').strip(), request.args.get('kind', 'auto')
    if kind not in ('auto', 'po', 'req'): return jsonify({'error': 'kind must be auto, po or req'}), 400
    if not pad10(number): return jsonify({'error': 'Enter a PO or requisition number.'}), 400
    try:
        conn = connect_hana()
    except Exception as e:
        return jsonify({'error': f'Could not connect to HANA: {e}'}), 502
    try:
        k, po_rows, req_rows, fiber_rows, set_rows, notes = fetch_lookup_data(conn, number, kind)
    except LookupError as e:
        return jsonify({'error': str(e)}), 404
    except Exception as e:
        return jsonify({'error': f'SAP query failed: {e}'}), 502
    try:
        result = build_lookup(k, pad10(number), po_rows, req_rows, fiber_rows, set_rows, _cache.get('margin'), notes)
        result['buysheet'] = lookup_buysheet_check(conn, req_rows)
    finally:
        try: conn.close()
        except Exception: pass
    return jsonify(result)

def lookup_buysheet_check(conn, req_rows):
    """The requisition(s) of the lookup against the buysheet last dropped on the PREBUY tab (None when the lookup has no requisition)."""
    if not req_rows: return None
    v, rows = _cache.get('validation') or {}, _cache.get('rows') or []
    if v.get('file_type') != 'buysheet': rows = []
    try:
        variant_map = query_variant_articles(conn, list({r['style'] for r in rows if r.get('type') == 'OS'})) if rows else {}
        return buysheet_req_check(req_rows, rows, v, variant_map, v.get('filename') or '')
    except Exception as e:
        return {'file': v.get('filename') or '', 'reqs': [], 'issues': [], 'notes': [f'The buysheet check could not run: {e}']}

def margin_info(m):
    if not m: return {'loaded': False}
    rows = m['rows']
    return {'loaded': True, 'filename': m.get('filename'), 'sheet': m['sheet'], 'rows': len(rows), 'styles': len({r['key'] for r in rows}),
            'with_req': sum(1 for r in rows if r['req']), 'with_po': sum(1 for r in rows if r['po']),
            'terms': {t: sum(1 for r in rows if r['term'] == t) for t in sorted({r['term'] for r in rows})}}

@app.route('/margin', methods=['GET'])
def margin_current():
    return jsonify(margin_info(_cache.get('margin')))

@app.route('/margin', methods=['POST'])
def margin_upload():
    if 'file' not in request.files: return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    try:
        m = parse_margin(f.read()); m['filename'] = f.filename
    except Exception as e:
        return jsonify({'error': str(e)}), 400
    _cache['margin'] = m
    return jsonify(margin_info(m))

@app.route('/margin/clear', methods=['POST'])
def margin_clear():
    _cache.pop('margin', None)
    return jsonify(margin_info(None))

@app.route('/validate', methods=['POST'])
def validate():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    try:
        validation = validate_bytes(f.read(), f.filename)
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    return jsonify(validation_payload(validation))

def fill_color_codes(rows):
    """A DPCI list has color names only, so the color check has no code to test. A code is filled in ONLY from an exact spelling
    (case included: SAP keeps BLACK and Black apart), most specific first: the style's own colors, then every SAP color. The first
    step that leaves exactly one code decides it.  A spelling that differs (other case) is never applied: it is listed in the note
    as a suggestion for a person to confirm.  A new color (no code in SAP) or an exact name that still leaves several codes is left
    blank.  The existing color check then judges every code that was filled in.  Each row says what happened in color_code_note."""
    sap = SapView(rows)
    conn = connect_hana()
    try: anywhere = query_color_codes_by_name(conn, [r['color'] for r in rows])
    finally: conn.close()
    for r in rows:
        exact, loose = r['color_exact'], r['color'].replace(' ', '')
        own = [(name, code) for (style, code), name in sap.style_color.items() if style == r['style']]
        steps = [sorted({c for n, c in own if n == exact}), anywhere.get(exact, [])]
        pick = next((s_[0] for s_ in steps if len(s_) == 1), '')
        r['color_code'] = pick
        if pick:
            r['color_code_note'] = ''
            continue
        several = next((s_ for s_ in steps if s_), [])
        similar = sorted({f"{c} ({n})" for n, c in own if n != exact and n.upper().replace(' ', '') == loose} |
                         {f"{c} ({n})" for n, cs in anywhere.items() if n != exact and n.upper().replace(' ', '') == loose for c in cs})
        if several:
            r['color_code_note'] = f"{exact}: several color codes use exactly this name ({', '.join(several)}): pick one"
        elif similar:
            r['color_code_note'] = f"{exact}: no exact match in SAP. Spelled differently there: {', '.join(similar[:6])}. Confirm which one, or give a new code"
        else:
            r['color_code_note'] = f"{exact}: no color code in SAP uses this name: it needs a new code"

def validate_bytes(content, filename):
    """Parse a buysheet, check it against SAP and stock, and keep it as the loaded file. Shared by the upload and the BUYSHEETS tab."""
    rows, customers, file_type = parse_prebuy(content)
    if file_type == 'dpci_list':
        fill_color_codes(rows)
    validation = validate_all(rows)
    validation.update(customers=customers, filename=filename, file_type=file_type)
    apply_stock(validation, _cache.get('stock') or live_stock(validation))   # an export added earlier wins, else live from HANA
    _cache.update(validation=validation, rows=rows, customers=customers)
    return validation

def load_dpci_list(content, filename):
    """Read a DPCI list and check it against SAP. It is kept BESIDE the loaded buysheet (never in its place), so the creation templates can
    list what the buysheet needs and what the DPCI list needs in one export."""
    rows, customers, file_type = parse_prebuy(content)
    if file_type != 'dpci_list':
        raise ValueError('This is not a DPCI list: it needs a sheet with Generic Manufacturer Style, Component DPCI, Description, COLOR and SIZE columns.')
    fill_color_codes(rows)
    validation = validate_all(rows)
    validation.update(customers=customers, filename=filename, file_type=file_type)
    _cache['dpci'] = {'validation': validation, 'rows': rows}
    return validation

def refresh_dpci_list():
    """Check the loaded DPCI list against SAP again (articles created since the last check now exist)."""
    d = _cache.get('dpci')
    if not d: raise ValueError('No DPCI list is loaded')
    old = d['validation']
    new = validate_all(d['rows'])
    new.update(customers=old.get('customers', []), filename=old.get('filename'), file_type=old.get('file_type'))
    d['validation'] = new
    return new

def creation_source():
    """The validation the creation templates are built from: the loaded buysheet and the loaded DPCI list, added together (either alone is fine)."""
    main, dpci = _cache.get('validation'), (_cache.get('dpci') or {}).get('validation')
    if main and dpci:
        return {**main, 'rows': list(main.get('rows', [])) + list(dpci.get('rows', [])), 'generic': list(main.get('generic', [])) + list(dpci.get('generic', []))}
    return main or dpci

def refresh_validation():
    """Check the loaded file against SAP again (no file needed): a generic, color or size created since the last check now has its
    variant articles, so a prepack's components show them and a prepack that was blocked can become ready."""
    old, rows = _cache.get('validation'), _cache.get('rows')
    if not old or not rows: raise ValueError('No validation data - validate a file first')
    new = validate_all(rows)
    new.update(customers=old.get('customers', []), filename=old.get('filename'), file_type=old.get('file_type'))
    apply_stock(new, _cache.get('stock') or live_stock(new))
    _cache['validation'] = new
    return new

@app.route('/revalidate', methods=['POST'])
def revalidate():
    try:
        return jsonify(validation_payload(refresh_validation()))
    except ValueError as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        return jsonify({'error': f'SAP check failed: {e}', 'trace': traceback.format_exc()}), 500

def validation_payload(validation):
    return {'build': BUILD, 'filename': validation.get('filename'), 'file_type': validation.get('file_type'), 'customers': validation.get('customers', []),
            'summary': validation['summary'], 'generic': validation['generic'], 'ppk': validation['ppk'],
            'generic_unique': validation['generic_unique'], 'ppk_unique': validation['ppk_unique'], 'consistency': validation.get('consistency', [])}

@app.route('/validation', methods=['GET'])
def validation_current():
    if not _cache.get('validation'):
        return jsonify({'error': 'No validation data — validate a file first'}), 400
    return jsonify(validation_payload(_cache['validation']))

@app.route('/stock', methods=['GET'])
def stock_current():
    return jsonify({'loaded': bool(_cache.get('stock')), 'stock': stock_info(_cache.get('stock')), 'live': STOCK_VIEW or None})

@app.route('/stock', methods=['POST'])
def stock_upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    try:
        stock = parse_cut_sold(f.read(), f.filename)
    except Exception as e:
        return jsonify({'error': str(e)}), 400
    _cache['stock'] = stock
    refreshed = bool(_cache.get('validation'))
    if refreshed: apply_stock(_cache['validation'], stock)
    return jsonify({'stock': stock_info(stock), 'refreshed': refreshed})

# ═══════════════════════════════════════════════════════════════════════════════
# ITEMHUB - Target's item list against SAP: barcode -> article -> VD52 DPCI (reuses itemhub_dpci_compare.py)
# ═══════════════════════════════════════════════════════════════════════════════
ITEMHUB_CDP = os.environ.get('ITEMHUB_CDP', 'http://localhost:9222')     # the dedicated, signed-in Chrome
ITEMHUB_OUT = os.path.join(TOOLS_DIR, 'itemhub_dpci_out')
_lp_lock = __import__('threading').Lock()

def itemhub_module():
    """itemhub_dpci_compare.py from the portal folder (the same script that runs from the command line)."""
    import importlib
    if TOOLS_DIR not in sys.path: sys.path.insert(0, TOOLS_DIR)
    try:
        return importlib.import_module('itemhub_dpci_compare')
    except ImportError as exc:
        raise RuntimeError(f"itemhub_dpci_compare.py was not found next to server.py ({exc}).")

def itemhub_chrome_state():
    """Is the dedicated Chrome open on the debugging port, and is a ItemHub tab in it?"""
    import urllib.request
    try:
        with urllib.request.urlopen(ITEMHUB_CDP.rstrip('/') + '/json', timeout=2) as resp:
            tabs = json.loads(resp.read().decode())
    except Exception:
        return {'chrome': False, 'itemhub_tab': False, 'cdp': ITEMHUB_CDP}
    return {'chrome': True, 'itemhub_tab': any('itemhub.partnersonline.com' in (t.get('url') or '') for t in tabs), 'cdp': ITEMHUB_CDP}

def itemhub_run(source='chrome', items_file=None):
    """Pull the items (from Chrome, or a saved pull), check them against SAP, write the report, cache the results."""
    m = itemhub_module()
    os.makedirs(ITEMHUB_OUT, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M')
    if source == 'saved':
        path = items_file or max((os.path.join(ITEMHUB_OUT, f) for f in os.listdir(ITEMHUB_OUT) if f.startswith('itemhub_items_') and f.endswith('.csv')), key=os.path.getmtime, default=None)
        if not path: raise RuntimeError('No saved ItemHub pull yet: run the check from Chrome first.')
        rows, pulled = m.dedupe_items(m.read_lp_csv(path)), os.path.basename(path)
    else:
        state = itemhub_chrome_state()
        if not state['chrome']:
            raise RuntimeError(f"The ItemHub Chrome is not open on {ITEMHUB_CDP}. Start it with the ItemHub Chrome shortcut, sign in, open Item Search, then click again.")
        if not state['itemhub_tab']:
            raise RuntimeError('That Chrome has no ItemHub tab open. Open ItemHub (Item Search) in it, then click again.')
        args = type('A', (), {})(); args.cdp, args.page_size, args.no_prompt, args.profile = ITEMHUB_CDP, 500, True, None
        try:
            rows = m.fetch_itemhub(args)
        except SystemExit as exc:
            raise RuntimeError(str(exc) or 'ItemHub refused the request: sign in again in the ItemHub Chrome window, then click again.')
        path = os.path.join(ITEMHUB_OUT, f'itemhub_items_{stamp}.csv'); m.write_lp_csv(rows, path); pulled = os.path.basename(path)
    items = [m.prep_lp(r) for r in rows]
    cfg = m.load_settings(os.environ.get('PORTAL_CONFIG') or os.path.join(HERE, 'config.json'))
    sap = m.sap_from_hana(cfg, items)
    results = m.compare_all(items, sap)
    report = os.path.join(ITEMHUB_OUT, f'itemhub_vs_sap_{stamp}.xlsx'); m.write_report(results, report, sap)
    from collections import Counter
    n = Counter(r['Status'] for r in results); issues = Counter(i.strip() for r in results for i in r['Issues'].split(',') if i.strip())
    _cache['itemhub'] = {'when': datetime.now().strftime('%Y-%m-%d %H:%M'), 'source': source, 'pulled_from': pulled, 'report': report,
                           'summary': {'checked': len(results), 'ok': n.get('OK', 0), 'review': n.get('REVIEW', 0), 'error': n.get('ERROR', 0),
                                       'issues': [{'key': k, 'count': c} for k, c in issues.most_common()], 'dpci_source': getattr(sap, 'dpci_source', '')},
                           'rows': results}
    return _cache['itemhub']

@app.route('/itemhub', methods=['GET'])
def itemhub_current():
    lp = _cache.get('itemhub')
    return jsonify({'state': itemhub_chrome_state(), 'result': lp and {k: v for k, v in lp.items() if k != 'rows'}, 'rows': (lp or {}).get('rows', [])})

@app.route('/itemhub/run', methods=['POST'])
def itemhub_run_route():
    body = request.get_json(silent=True) or {}
    if not _lp_lock.acquire(blocking=False):
        return jsonify({'error': 'A ItemHub check is already running. Wait for it to finish.'}), 409
    try:
        lp = itemhub_run(body.get('source', 'chrome'), body.get('items_file'))
        return jsonify({'result': {k: v for k, v in lp.items() if k != 'rows'}, 'rows': lp['rows'], 'state': itemhub_chrome_state()})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        _lp_lock.release()

@app.route('/itemhub/report', methods=['GET'])
def itemhub_report():
    lp = _cache.get('itemhub')
    if not lp or not os.path.exists(lp['report']): return jsonify({'error': 'No ItemHub check yet'}), 400
    return send_file(lp['report'], as_attachment=True, download_name=os.path.basename(lp['report']))

@app.route('/stock/clear', methods=['POST'])
def stock_clear():
    _cache.pop('stock', None)
    refreshed = bool(_cache.get('validation'))
    if refreshed: apply_stock(_cache['validation'], live_stock(_cache['validation']))      # back to live stock
    return jsonify({'loaded': False, 'refreshed': refreshed, 'live': STOCK_VIEW or None})

@app.route('/generate', methods=['GET'])
def generate():
    if not _cache.get('rows'):
        return jsonify({'error': 'No data — validate a file first'}), 400
    try:
        buf, kind = make_buysheet(_cache['rows'], _cache.get('customers', []))
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    if kind == 'xlsm':
        return send_file(buf, as_attachment=True, download_name=f"Buysheet_{datetime.now():%m-%d-%Y}.xlsm", mimetype=XLSM)
    return _send_xlsx(buf, 'Buysheet_DATA')

@app.route('/creation_templates', methods=['GET'])
def creation_templates():
    if not creation_source():
        return jsonify({'error': 'No validation data — validate a file first'}), 400
    if request.args.get('refresh') == '1':                       # variants created in SAP since the last check show up in the file
        try:
            if _cache.get('validation'): refresh_validation()
            if _cache.get('dpci'): refresh_dpci_list()
        except Exception as e: return jsonify({'error': f'Could not refresh from SAP: {e}'}), 502
    try:
        return _send_xlsx(make_creation_templates(creation_source()), 'Creation_Templates')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ── PPK creation: runs sap_ppk_creation_playwright.py (MM41 in SAP WebGUI) on the machine the portal runs on ──
PPK_RUN_DIR    = os.path.join(TOOLS_DIR, 'ppk_creation_out')
PPK_RUN_SCRIPT = os.path.join(HERE, 'sap_ppk_creation_playwright.py')
_ppk_run = {'proc': None, 'log': '', 'cont': '', 'live': False, 'started': '', 'ready': [], 'skipped': [], 'file': ''}

def make_ppk_run_file(validation, path):
    """The PPK CREATION sheet alone, in the layout the script reads (sheet 1), for the prepacks that can be created NOW: every
    component already exists as a variant. Season / Year are written as values: the script reads with pandas, which sees no
    result for a formula the portal wrote. Returns (ready articles, skipped articles with the reason)."""
    wb = load_workbook(make_creation_templates(validation))
    del wb['GENERIC CREATION']; del wb['COLOR ADD']
    ws = wb['PPK CREATION']
    head = [c.value for c in ws[1]]
    col = {h: i for i, h in enumerate(head)}
    blocks = defaultdict(list)
    for row in ws.iter_rows(min_row=2):
        blocks[row[col['Material']].value].append(row)
    ready, skipped, drop = [], [], []
    for art, rows in blocks.items():
        bad = [r[col['VARIANT STATUS']].value for r in rows if not str(r[col['VARIANT STATUS']].value or '').startswith('OK')]
        if bad:
            skipped.append({'article': art, 'reason': str(bad[0])[:120]}); drop += [r[0].row for r in rows]
        else:
            ready.append(art)
            first = rows[0]
            s = str(art)[5:6]
            if no_season_style(art):                                    # CPU000FH01: no season in the style number
                first[col['Season']].value, first[col['Year']].value = NO_SEASON_CODE, NO_SEASON_YEAR
            else:
                first[col['Season']].value = ['SPR', 'SUM', 'FALL', 'HOL'][int(s) - 1] if s in '1234' and s else ''
                first[col['Year']].value = '20' + str(art)[3:5]
    for r in sorted(drop, reverse=True): ws.delete_rows(r)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return ready, skipped

def _ppk_running():
    p = _ppk_run['proc']
    return bool(p and p.poll() is None)

@app.route('/ppk_creation/start', methods=['POST'])
def ppk_creation_start():
    raise RuntimeError("Not available in the demo.")
    if _ppk_running(): return jsonify({'error': 'A creation run is already in progress.'}), 409
    if not _cache.get('validation'): return jsonify({'error': 'No validation data - validate a file first'}), 400
    if not os.path.exists(PPK_RUN_SCRIPT): return jsonify({'error': f'sap_ppk_creation_playwright.py was not found in {HERE}.'}), 500
    live = bool((request.get_json(silent=True) or {}).get('live'))
    try:                                   # the Chrome SAP_Chrome.bat starts (debug port, signed in): the run attaches to it
        import urllib.request
        with urllib.request.urlopen(ITEMHUB_CDP.rstrip('/') + '/json', timeout=2) as resp:
            tabs = json.loads(resp.read().decode('utf-8', 'replace'))
    except Exception:
        return jsonify({'error': f'The SAP Chrome is not open on {ITEMHUB_CDP}. Start SAP_Chrome.bat from the desktop, sign in to SAP, then try again.'}), 400
    if not any('prdap1-erp' in str(t.get('url', '')) or '/sap/bc/gui' in str(t.get('url', '')) for t in tabs):
        return jsonify({'error': 'That Chrome has no SAP tab open. Open SAP in it (SAP_Chrome.bat does), then try again.'}), 400
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    xlsx = os.path.join(PPK_RUN_DIR, f'ppk_creation_{stamp}.xlsx')
    try:
        refresh_validation()                                      # components created in SAP since the last check count as ready now
    except Exception as e:
        return jsonify({'error': f'Could not refresh from SAP before the run: {e}'}), 502
    try:
        ready, skipped = make_ppk_run_file(_cache['validation'], xlsx)
    except Exception as e:
        return jsonify({'error': f'Could not build the creation file: {e}'}), 500
    if not ready: return jsonify({'error': 'No prepack is ready to create (each needs its generic, colors and sizes in SAP first).', 'skipped': skipped}), 400
    cont, logp = os.path.join(PPK_RUN_DIR, f'continue_{stamp}.flag'), os.path.join(PPK_RUN_DIR, f'ppk_creation_{stamp}.log')
    logf = open(logp, 'w', encoding='utf-8')
    args = [sys.executable, '-u', PPK_RUN_SCRIPT, xlsx, f'--continue-file={cont}', f'--cdp={ITEMHUB_CDP}'] + (['--live'] if live else [])
    _ppk_run.update(proc=subprocess.Popen(args, cwd=TOOLS_DIR, stdout=logf, stderr=subprocess.STDOUT), log=logp, cont=cont, live=live,
                    started=datetime.now().strftime('%Y-%m-%d %H:%M:%S'), ready=ready, skipped=skipped, file=xlsx)
    return jsonify({'started': True, 'live': live, 'ready': ready, 'skipped': skipped})

@app.route('/ppk_creation/continue', methods=['POST'])
def ppk_creation_continue():
    if not _ppk_running() or not _ppk_run['cont']: return jsonify({'error': 'No creation run is waiting.'}), 400
    open(_ppk_run['cont'], 'w').close()
    return jsonify({'ok': True})

@app.route('/ppk_creation/stop', methods=['POST'])
def ppk_creation_stop():
    if _ppk_running(): _ppk_run['proc'].terminate()
    return jsonify({'ok': True})

@app.route('/ppk_creation/status', methods=['GET'])
def ppk_creation_status():
    tail = ''
    if _ppk_run['log'] and os.path.exists(_ppk_run['log']):
        with open(_ppk_run['log'], encoding='utf-8', errors='replace') as f: tail = ''.join(f.readlines()[-60:])
    running = _ppk_running()
    return jsonify({'running': running, 'exit_code': None if (running or not _ppk_run['proc']) else _ppk_run['proc'].returncode,
                    'live': _ppk_run['live'], 'started': _ppk_run['started'], 'ready': _ppk_run['ready'], 'skipped': _ppk_run['skipped'],
                    'waiting': running and 'WAITING:' in tail and not os.path.exists(_ppk_run['cont'] or '-'), 'log': tail,
                    'file': os.path.basename(_ppk_run['file'])})

@app.route('/validation_report', methods=['GET'])
def validation_report():
    if not _cache.get('validation'):
        return jsonify({'error': 'No validation data — validate a file first'}), 400
    return _from_cache(lambda: make_validation_report(_cache['validation']), 'Validation')

@app.route('/req_csv', methods=['GET'])
def req_csv():
    if not _cache.get('rows'):
        return jsonify({'error': 'No data — validate a file first'}), 400
    try:
        content = make_req_csv(_cache['rows'], _cache.get('customers', []), _cache.get('validation'))
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    return send_file(io.BytesIO(content.encode('utf-8')), as_attachment=True,
                     download_name=f"PR_Upload_{datetime.now():%m-%d-%Y}.csv", mimetype='text/csv')

@app.route('/delivery_split', methods=['POST'])
def delivery_split():
    if 'file' not in request.files: return jsonify({'error':'No file'}), 400
    f=request.files['file']
    del1=int(request.form.get('del1',0)); del2=int(request.form.get('del2',0))
    del1_date=request.form.get('del1_date',''); del2_date=request.form.get('del2_date','')
    if not del1 or not del2: return jsonify({'error':'Missing delivery quantities'}), 400
    total=del1+del2; prop=del1/total

    wb_in=load_workbook(io.BytesIO(f.read()),data_only=True); ws_in=wb_in.active
    header_ri=None; headers=[]
    for ri,row in enumerate(ws_in.iter_rows(values_only=True),1):
        cells=[str(v).strip().upper() if v else '' for v in row]
        if 'COLOR' in cells or 'STYLE' in cells:
            header_ri=ri; headers=cells; break
    if not header_ri: return jsonify({'error':'Could not find header row'}),400

    color_ci=next((headers.index(h) for h in ['COLOR DESCRIPTION','COLOR'] if h in headers),3)
    ratio_ci=next((headers.index(h) for h in ['RATIO','RATIOS'] if h in headers),None)
    style_ci=headers.index('STYLE') if 'STYLE' in headers else 1
    all_sizes=[h for h in headers if h in KNOWN_SIZES]
    size_cis={sz:headers.index(sz) for sz in all_sizes}

    wb_out=Workbook(); ws_out=wb_out.active; ws_out.title="DELIVERY SPLIT"
    out_hdrs=['STYLE','COLOR','DELIVERY']+all_sizes+['TOTAL']
    for ci,h in enumerate(out_hdrs,1):
        c=ws_out.cell(1,ci,value=h); c.font=WBOLD; c.fill=NAVY; c.alignment=CTR

    ri_out=2
    for row in ws_in.iter_rows(min_row=header_ri+1,values_only=True):
        if not any(v for v in row): continue
        style=str(row[style_ci]).strip() if row[style_ci] else ''
        color=str(row[color_ci]).strip() if row[color_ci] else ''
        if not color: continue
        sizes={sz:int(float(row[headers.index(sz)])) for sz in all_sizes
               if sz in headers and row[headers.index(sz)] and str(row[headers.index(sz)]).strip() not in ('','0','None')}
        total_qty=sum(sizes.values())
        if not total_qty: continue
        d1_sizes={sz:round(qty*prop) for sz,qty in sizes.items()}
        d2_sizes={sz:sizes[sz]-d1_sizes[sz] for sz in sizes}
        for label,sz_data,fill in [
            ('TOTAL',sizes,PatternFill('solid',start_color='D9E1F2')),
            (f'DEL 1{" ("+del1_date+")" if del1_date else ""}',d1_sizes,PatternFill('solid',start_color='E2EFDA')),
            (f'DEL 2{" ("+del2_date+")" if del2_date else ""}',d2_sizes,None)]:
            row_total=sum(sz_data.values())
            vals=[style,color,label]+[sz_data.get(sz,'') for sz in all_sizes]+[row_total]
            for ci,val in enumerate(vals,1):
                c=ws_out.cell(ri_out,ci,value=val)
                c.font=BFONT if label=='TOTAL' else DFONT; c.alignment=CTR
                if fill: c.fill=fill
            ri_out+=1
        ri_out+=1

    add_filters(wb_out)
    buf=io.BytesIO(); wb_out.save(buf); buf.seek(0)
    today=datetime.now().strftime('%m-%d-%Y')
    return send_file(buf,as_attachment=True,download_name=f'Delivery_Split_{today}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def lan_ip():
    """This machine's address on the network: what other computers type to reach the portal."""
    import socket
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.connect(('10.255.255.255', 1))        # no packet is sent; it just picks the outgoing interface
        ip = sock.getsockname()[0]; sock.close()
        return ip
    except Exception:
        return None
