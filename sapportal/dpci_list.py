"""The DPCI list on the ItemHub tab: one file, checked three ways, that also feeds creation.

1. SAP: every row's article is found in SAP (style + color + size) and the DPCI the list gives is compared with the DPCI kept in
   VD52 (the customer-material info table, KNMT).  SAP is only read.
2. ItemHub: the list's DPCI is looked up in a FULL ItemHub pull (every item, the page's filters ignored): its style and size must be
   the list's, and its barcode must be the UPC SAP holds.  ItemHub is only read.
3. Creation: what SAP does not have yet (style, color or size) is what the creation template export lists, on its existing sheets.
   That comes from the same validation of the same file, so this module only reports it.

The list is loaded here (/dpci_list/load) and kept BESIDE the loaded buysheet: the SAP tab's creation templates button builds from both,
so there is no creation button on this tab.  This module reads that loaded list."""

import io
import os
import re
from datetime import datetime
from flask import jsonify, request
from openpyxl import Workbook
from .config import _cache
from .hana import connect_hana
from .routes import ITEMHUB_CDP, ITEMHUB_OUT, _send_xlsx, itemhub_chrome_state, itemhub_module, load_dpci_list, refresh_dpci_list
from .webapp import app

ARTICLE_VIEW = '"_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS"'
KIND_TEXT = {'style': 'the style is not in SAP', 'color': 'the color is not on this style in SAP', 'size': 'the size is not on this style and color in SAP'}


def digits(v):
    return re.sub(r'\D', '', str(v or ''))


LP_PREFIX = 'itemhub_items_all_'          # a full pull made from this tab; the other saved pulls (a filtered search) are never used here
_lp_cache = {}


def lp_load():
    """The newest full ItemHub pull -> {'index': {DPCI digits: item}, 'items': n, 'when': text, 'file': name}, or None when there is none."""
    if not os.path.isdir(ITEMHUB_OUT):
        return None
    files = [os.path.join(ITEMHUB_OUT, f) for f in os.listdir(ITEMHUB_OUT) if f.startswith(LP_PREFIX) and f.endswith('.csv')]
    if not files:
        return None
    path = max(files, key=os.path.getmtime)
    stamp = os.path.getmtime(path)
    hit = _lp_cache.get(path)
    if not hit or hit['mtime'] != stamp:
        m = itemhub_module()
        items = [m.prep_lp(r) for r in m.dedupe_items(m.read_lp_csv(path))]
        hit = {'mtime': stamp, 'index': {i['dpci_key']: i for i in items if i['dpci_key']}, 'items': len(items),
               'when': datetime.fromtimestamp(stamp).strftime('%Y-%m-%d %H:%M'), 'file': os.path.basename(path)}
        _lp_cache[path] = hit
    return hit


def lp_problems(item, file_style, file_size, sap_upc):
    """What differs between one ItemHub item and the list row / SAP: [] when ItemHub agrees (style, size, and barcode = SAP's UPC)."""
    m = itemhub_module()
    out = []
    if m.key_style(file_style) not in item['styles']:
        out.append(f"ItemHub style is {', '.join(item['styles']) or 'blank'}, the list says {file_style}.")
    lp_size = item.get('lp_size') or ''
    if lp_size and m.canon_size(file_size)[0] != m.canon_size(lp_size)[0]:
        out.append(f"ItemHub size is {lp_size}, the list says {file_size}.")
    if sap_upc and item['gtins'] and sap_upc.lstrip('0') not in {g.lstrip('0') for g in item['gtins']}:
        out.append(f"ItemHub barcode {item['gtins'][0].lstrip('0')} is not the UPC in SAP ({sap_upc}).")
    return out


def sap_dpcis(conn, articles):
    """article (as SAP prints it) -> set of DPCI digit strings kept for it in VD52, any customer.  {} when the table cannot be found."""
    m = itemhub_module()
    cfg = m.load_settings()
    table = m.resolve_dpci_table(conn, cfg)
    if not table or not articles:
        return None
    forms = {f: a for a in articles for f in m.article_forms(a)}
    a_col, d_col = cfg['DPCI_ARTICLE_COL'], cfg['DPCI_COL']
    out = {}
    for art, dpci in m._select_in(conn, table, [a_col, d_col], a_col, list(forms)):
        key = forms.get(str(art).strip()) or forms.get(str(art).strip().lstrip('0'))
        if key and digits(dpci):
            out.setdefault(key, set()).add(digits(dpci))
    return out


def sap_upcs(conn, articles):
    """article -> the UPC SAP holds on it (MARA_UPC), as text."""
    out, cur = {}, conn.cursor()
    arts = sorted({a for a in articles if a})
    for i in range(0, len(arts), 500):
        part = arts[i:i + 500]
        cur.execute('SELECT "MARA_ARTICLE","MARA_UPC" FROM ' + ARTICLE_VIEW + ' WHERE "MARA_ARTICLE" IN (' + ','.join('?' * len(part)) + ')', tuple(part))
        for art, upc in cur.fetchall():
            out[str(art or '').strip()] = str(upc or '').strip()
    return out


def vd52_dpci(dpci):
    """The DPCI as VD52 keeps it: 043-23-2077 -> '043  23  2077' (two spaces where the dashes are)."""
    return '  '.join(p for p in re.split(r'\s*-\s*|\s+', str(dpci or '').strip()) if p)


def upload_rows(results):
    """The VD52 upload rows: [MATERIAL, DESCRIPTION (left empty), CUST MATERIAL (the DPCI, two spaces), UPC] for every row that has a SAP article and still needs its DPCI
    entered or corrected IN VD52 (the SAP check says so).  Rows already right in VD52 are left out, and so are rows SAP does not have yet
    (no article to enter it on).  When ItemHub was compared, a row ItemHub CONTRADICTS (it has the DPCI on another style, size or
    barcode) is held back too; a DPCI ItemHub does not have yet is not held back, and does not put a row in the file."""
    return [[r['SAP article'], None, vd52_dpci(r['List DPCI']), r.get('UPC', '')]
            for r in results if r['SAP article'] and r.get('VD52 need') and r['List DPCI'] and r.get('LP ok') is not False]


LIST_COLUMNS = ['Generic Manufacturer Style #', 'Component DPCI', 'Description', 'COLOR', 'SIZE', 'UPC', 'SAP article', 'Color code', 'Status', 'ItemHub']


def list_rows(results, rows):
    """EVERY row of the validated list, in the columns of the file that was loaded plus what the check found: the UPC and SAP article,
    the color code, the status, and ItemHub's answer (empty when ItemHub was not pulled).  Nothing is left out because it is already right."""
    desc = {(r['style'], r['color'], sz.upper()): d for r in rows for sz, d in (r.get('desc_by_size') or {}).items()}
    exact = {(r['style'], r['color']): r.get('color_exact') or r['color'] for r in rows}
    return [[x['Style'], x['List DPCI'], desc.get((x['Style'], x['Color'], x['Size'].upper()), ''), exact.get((x['Style'], x['Color']), x['Color']), x['Size'],
             x['UPC'], x['SAP article'], x['Color code'], x['Status'], x['ItemHub']] for x in results]


def check_rows(validation, rows, dpcis_for, upcs_for=None, lp_by_dpci=None):
    """One result row per style/color/size of the list.  dpcis_for(articles) -> {article: {dpci digits}} or None when VD52 is unreadable.
    lp_by_dpci: {DPCI digits: prepared ItemHub item} from a full pull, or None when ItemHub was not pulled (the ItemHub columns stay empty)."""
    want = {}
    for r in rows:
        for sz, d in (r.get('dpci_by_size') or {}).items():
            want[(r['style'], r['color'], sz.upper())] = d
    generic = [g for g in validation.get('generic', [])]
    arts = sorted({g['sap_article'] for g in generic if g.get('sap_article')})
    vd52 = dpcis_for(arts)
    upcs = upcs_for(arts) if upcs_for else {}
    out = []
    for g in generic:
        key = (g['style'], g.get('buy_color', ''), g['size'].upper())
        dpci = want.get(key, '')
        row = {'Style': g['style'], 'Color': g.get('buy_color', ''), 'SAP color': g.get('sap_color') or '', 'Color code': g.get('color_code', ''), 'Size': g['size'], 'List DPCI': dpci,
               'SAP article': g.get('sap_article') or '', 'UPC': upcs.get(g.get('sap_article') or '', ''), 'VD52 DPCI': '', 'ItemHub': '', 'LP barcode': '', 'LP ok': None, 'VD52 need': False, 'Status': 'OK', 'Action': ''}
        if g['status'] != '✓ EXISTS':
            row.update(Status='CREATE', Action='Create in SAP: ' + KIND_TEXT.get(g.get('missing_kind'), 'not in SAP') + '. It is on the creation templates; enter the DPCI in VD52 afterwards.')
        elif vd52 is None:
            row.update(Status='REVIEW', Action='The VD52 table (KNMT) could not be read: set DPCI_TABLE in config.json.')
        else:
            have = vd52.get(row['SAP article'], set())
            row['VD52 DPCI'] = ', '.join(sorted(have))
            if not have:
                row.update(Status='REVIEW', Action='Enter this DPCI in VD52 for the article.')
                row['VD52 need'] = True
            elif digits(dpci) not in have:
                row.update(Status='ERROR', Action='VD52 has a different DPCI: correct it in VD52 (or the list).')
                row['VD52 need'] = True
        if g.get('color_code') == '':
            row['Action'] = (row['Action'] + ' ' if row['Action'] else '') + 'No color code yet: a new color needs a code before it can be created.'
            if row['Status'] == 'OK': row['Status'] = 'REVIEW'
        if lp_by_dpci is not None:
            item = lp_by_dpci.get(digits(dpci))
            if item is None:
                row['ItemHub'] = '✗ not found'          # a warning only: ItemHub may not have the DPCI set up yet, so it cannot contradict the list
                row['Action'] = (row['Action'] + ' ' if row['Action'] else '') + 'ItemHub does not have this DPCI.'      # its own column and count: the row's status is about SAP
            else:
                probs = lp_problems(item, g['style'], g['size'], row['UPC'])
                row['ItemHub'], row['LP ok'] = ('✓' if not probs else '✗ differs'), not probs
                row['LP barcode'] = item['gtins'][0].lstrip('0') if item['gtins'] else ''
                if probs:
                    row['Action'] = (row['Action'] + ' ' if row['Action'] else '') + ' '.join(probs)
                    if row['Status'] != 'CREATE': row['Status'] = 'ERROR'
        out.append(row)
    return out


def loaded():
    d = _cache.get('dpci')
    return (d['validation'], d['rows']) if d else (None, None)


def run_check():
    """The loaded DPCI list checked against SAP -> (result rows, error response or None)."""
    validation, rows = loaded()
    if not validation:
        return None, (jsonify({'error': 'No DPCI list is loaded: drop the file on the ItemHub tab first.'}), 400)
    try:
        conn = connect_hana()
        try:
            lp = lp_load()
            res = check_rows(validation, rows, lambda arts: sap_dpcis(conn, arts), lambda arts: sap_upcs(conn, arts), lp['index'] if lp else None)
        finally:
            conn.close()
    except Exception as exc:
        return None, (jsonify({'error': str(exc)}), 502)
    return res, None


@app.route('/dpci_list/state', methods=['GET'])
def dpci_list_state():
    validation, _ = loaded()
    return jsonify({'loaded': bool(validation), 'filename': validation.get('filename') if validation else None})


@app.route('/dpci_list/check', methods=['GET'])
def dpci_list_check():
    res, err = run_check()
    if err: return err
    validation, rows = loaded()
    count = lambda s: sum(1 for r in res if r['Status'] == s)
    notes = sorted({r['color_code_note'] for r in rows if r.get('color_code_note')})
    lp = lp_load()
    lp_counts = ({'agree': sum(1 for r in res if r['ItemHub'] == '✓'), 'missing': sum(1 for r in res if r['ItemHub'] == '✗ not found'),
                  'differ': sum(1 for r in res if r['ItemHub'] == '✗ differs')} if lp else None)
    return jsonify({'rows': res, 'notes': notes, 'lp_counts': lp_counts, 'lp': ({k: lp[k] for k in ('items', 'when', 'file')} if lp else None), 'filename': validation.get('filename'), 'when': datetime.now().strftime('%Y-%m-%d %H:%M'),
                    'upload_rows': len(upload_rows(res)),
                    'summary': {'checked': len(res), 'ok': count('OK'), 'create': count('CREATE'), 'review': count('REVIEW'), 'error': count('ERROR')}})


@app.route('/dpci_list/export_list', methods=['GET'])
def dpci_list_export_list():
    """The validated DPCI list as a file, every row: the list's own columns, then UPC, SAP article, color code, status and ItemHub's answer."""
    res, err = run_check()
    if err: return err
    _, rows = loaded()
    wb = Workbook(); ws = wb.active; ws.title = 'DPCI LIST'
    ws.append(LIST_COLUMNS)
    for row in list_rows(res, rows):
        ws.append(row)
    for r in ws.iter_rows(min_row=2):
        for c in r: c.number_format = '@'                      # text: a UPC or DPCI must never turn into a number
    for col, width in zip('ABCDEFGHIJ', (24, 16, 50, 22, 8, 16, 16, 12, 10, 14)):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = 'A2'
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return _send_xlsx(buf, 'DPCI_List')


@app.route('/dpci_list/export', methods=['GET'])
def dpci_list_export():
    """The VD52 upload file from the loaded DPCI list, in the columns of the VD52 layout sheet (MATERIAL, DESCRIPTION, CUST MATERIAL): the
    DESCRIPTION is left empty and CUST MATERIAL holds the DPCI with two spaces where the dashes are.  The 4th column is the UPC."""
    res, err = run_check()
    if err: return err
    wb = Workbook(); ws = wb.active; ws.title = 'VD52'
    ws.append(['MATERIAL', 'DESCRIPTION', 'CUST MATERIAL', 'UPC'])
    for row in upload_rows(res):
        ws.append(row)
    for r in ws.iter_rows(min_row=2):
        for c in r: c.number_format = '@'                      # text: a UPC or DPCI must never turn into a number
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return _send_xlsx(buf, 'DPCI_Upload')


@app.route('/dpci_list/pull_lp', methods=['POST'])
def dpci_list_pull_lp():
    """Pull EVERY ItemHub item (the page's filters ignored) through the signed-in Chrome and keep it as the pull this tab compares with."""
    try:
        state = itemhub_chrome_state()
        if not state['chrome']:
            raise RuntimeError(f"The ItemHub Chrome is not open on {ITEMHUB_CDP}. Start it with the SightLine shortcut, sign in, open Item Search, then click again.")
        if not state['itemhub_tab']:
            raise RuntimeError('That Chrome has no ItemHub tab open. Open ItemHub (Item Search) in it, then click again.')
        m = itemhub_module()
        args = type('A', (), {})()
        args.cdp, args.page_size, args.no_prompt, args.profile, args.all = ITEMHUB_CDP, 500, True, None, True
        try:
            rows = m.dedupe_items(m.fetch_itemhub(args))
        except SystemExit as exc:
            raise RuntimeError(str(exc))
        os.makedirs(ITEMHUB_OUT, exist_ok=True)
        m.write_lp_csv(rows, os.path.join(ITEMHUB_OUT, f"{LP_PREFIX}{datetime.now():%Y%m%d_%H%M}.csv"))
        lp = lp_load()
        return jsonify({'lp': {k: lp[k] for k in ('items', 'when', 'file')}})
    except Exception as exc:
        return jsonify({'error': str(exc)}), 400


@app.route('/dpci_list/load', methods=['POST'])
def dpci_list_load():
    """Load a DPCI list (kept beside the loaded buysheet, which stays as it is)."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    try:
        v = load_dpci_list(f.read(), f.filename)
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400
    except Exception as exc:
        return jsonify({'error': str(exc)}), 500
    return jsonify({'filename': v.get('filename'), 'rows': len(v.get('generic', []))})


@app.route('/dpci_list/refresh', methods=['POST'])
def dpci_list_refresh():
    """Check the loaded DPCI list against SAP again (after articles or DPCIs were created in SAP)."""
    try:
        refresh_dpci_list()
    except ValueError as exc:
        return jsonify({'error': str(exc)}), 400
    except Exception as exc:
        return jsonify({'error': str(exc)}), 502
    return jsonify({'ok': True})
