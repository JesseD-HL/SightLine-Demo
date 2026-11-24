"""SAP tab: runs the MM41 creation scripts in tools/ on the machine the portal runs on, from the creation template of the loaded validation.
Registry: one entry per script. Add the pricing scripts (ZVP1 / ZSLS / PR00) here when they exist."""

import os
import sys
from datetime import datetime
from flask import jsonify, request
from openpyxl import load_workbook
from .config import HERE, TOOLS_DIR, _cache
from .helpers import NO_SEASON_CODE, NO_SEASON_YEAR, no_season_style
from .outputs import make_creation_templates
from .webapp import app

RUN_DIR = os.path.join(TOOLS_DIR, 'sap_run_out')
SCRIPTS = {      # id -> what the SAP tab shows and how it runs. 'desktop': drives the SAP GUI window by screen coordinates (PyAutoGUI)
    'generic': {'label': 'Generic creation (MM41)', 'script': os.path.join(HERE, 'tools', '_sap_generic_creation.py'), 'sheet': 'GENERIC CREATION', 'desktop': True,
                'about': 'Creates the new generics (styles SAP does not have) from the GENERIC CREATION tab: colors, sizes, weights, season. '
                         'It clicks by screen coordinates, so SAP must be open on MM41 on the screen the coordinates were set for, and nothing may be touched during the run.'},
}
_runs = {k: {'proc': None, 'log': '', 'live': False, 'started': '', 'file': '', 'rows': 0} for k in SCRIPTS}

def _running(k):
    p = _runs[k]['proc']
    return bool(p and p.poll() is None)

def _build_file(validation, sheet, path):
    """The creation workbook with Season / Year written as VALUES (the scripts read with pandas, which sees nothing for a formula
    the portal wrote). Returns the number of rows on the sheet the script reads."""
    wb = load_workbook(make_creation_templates(validation))
    ws = wb[sheet]
    head = [c.value for c in ws[1]]
    col = {h: i for i, h in enumerate(head)}
    n = 0
    for row in ws.iter_rows(min_row=2):
        m = str(row[col['Material']].value or '')
        if not m: continue
        n += 1
        if 'Season' in col and 'Year' in col and len(m) >= 6:
            if no_season_style(m): row[col['Season']].value, row[col['Year']].value = NO_SEASON_CODE, NO_SEASON_YEAR
            elif m[5] in '1234': row[col['Season']].value, row[col['Year']].value = ['SPR', 'SUM', 'FALL', 'HOL'][int(m[5]) - 1], '20' + m[3:5]
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return n

@app.route('/sap_run/scripts', methods=['GET'])
def sap_scripts():
    return jsonify([{'id': k, 'label': v['label'], 'about': v['about'], 'running': _running(k)} for k, v in SCRIPTS.items()])

@app.route('/sap_run/<sid>/start', methods=['POST'])
def sap_run_start(sid):
    raise RuntimeError("Not available in the demo.")
    if sid not in SCRIPTS: return jsonify({'error': 'Unknown script.'}), 404
    if any(_running(k) for k in SCRIPTS): return jsonify({'error': 'A script is already running. Stop it first.'}), 409
    if not _cache.get('validation'): return jsonify({'error': 'No validation data - validate a file first'}), 400
    sc = SCRIPTS[sid]
    if not os.path.exists(sc['script']): return jsonify({'error': f"{os.path.basename(sc['script'])} was not found in {os.path.dirname(sc['script'])}."}), 500
    live = bool((request.get_json(silent=True) or {}).get('live'))
    try:
        from .routes import refresh_validation
        refresh_validation()                                      # what was created in SAP since the last check is left out
    except Exception as e:
        return jsonify({'error': f'Could not refresh from SAP before the run: {e}'}), 502
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    xlsx, logp = os.path.join(RUN_DIR, f'{sid}_{stamp}.xlsx'), os.path.join(RUN_DIR, f'{sid}_{stamp}.log')
    try:
        rows = _build_file(_cache['validation'], sc['sheet'], xlsx)
    except Exception as e:
        return jsonify({'error': f'Could not build the creation file: {e}'}), 500
    if not rows: return jsonify({'error': f"Nothing to create: the {sc['sheet']} sheet has no rows."}), 400
    logf = open(logp, 'w', encoding='utf-8')
    args = [sys.executable, '-u', sc['script'], xlsx] + ([] if live else ['--test'])
    _runs[sid].update(proc=subprocess.Popen(args, cwd=TOOLS_DIR, stdout=logf, stderr=subprocess.STDOUT), log=logp, live=live,
                      started=datetime.now().strftime('%Y-%m-%d %H:%M:%S'), file=xlsx, rows=rows)
    return jsonify({'started': True, 'live': live, 'rows': rows})

@app.route('/sap_run/<sid>/stop', methods=['POST'])
def sap_run_stop(sid):
    if sid in SCRIPTS and _running(sid): _runs[sid]['proc'].terminate()
    return jsonify({'ok': True})

@app.route('/sap_run/<sid>/status', methods=['GET'])
def sap_run_status(sid):
    if sid not in SCRIPTS: return jsonify({'error': 'Unknown script.'}), 404
    r, tail = _runs[sid], ''
    if r['log'] and os.path.exists(r['log']):
        with open(r['log'], encoding='utf-8', errors='replace') as f: tail = ''.join(f.readlines()[-60:])
    running = _running(sid)
    return jsonify({'running': running, 'exit_code': None if (running or not r['proc']) else r['proc'].returncode, 'live': r['live'],
                    'started': r['started'], 'rows': r['rows'], 'log': tail, 'file': os.path.basename(r['file'])})
