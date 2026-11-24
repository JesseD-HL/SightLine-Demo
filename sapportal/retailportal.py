"""The RetailPortal tab: the items RetailPortal's catalog table shows (GTIN/UPC, color) checked against SAP.

READ ONLY on both sides.  RetailPortal is read through the signed-in Chrome (the same one ItemHub uses): the rows of the table
that is open on screen are copied, nothing is clicked, edited or exported.  SAP (HANA) is only read.
Checks: the UPC is on a SAP article, and the color is the same.  (Prices are not checked.)"""

import json
import re
import urllib.request
from datetime import datetime
from flask import jsonify
from .hana import connect_hana
from .routes import ITEMHUB_CDP
from .webapp import app

SO_HOST = 'retailportal.bigbox.com'
ARTICLE_VIEW = '"_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS"'

READ_TABLE_JS = """() => {
  const heads = [...document.querySelectorAll('table thead th')].map(h => h.innerText.trim());
  const rows = [...document.querySelectorAll('table tbody tr')].map(r => [...r.cells].map(c => c.innerText.trim()));
  return { heads, rows };
}"""


def gtin_forms(g):
    """The ways SAP may hold a 14-digit GTIN from RetailPortal: as shown, without the leading zeros, the 12-digit UPC, the 13-digit EAN."""
    g = re.sub(r'\D', '', str(g or ''))
    return [f for f in dict.fromkeys([g, g.lstrip('0'), g[2:], g[1:]]) if f]


def parse_table(heads, rows):
    """The page's table -> [{gtin, name, color}].  Cells are matched to headers from the right (a leading checkbox cell is skipped)."""
    out = []
    for r in rows:
        off = len(r) - len(heads)
        if off < 0 or not heads:
            continue
        cell = lambda h: r[off + heads.index(h)] if h in heads else ''
        g = re.sub(r'\D', '', cell('GTIN/UPC'))
        if not g:
            continue
        out.append({'gtin': g, 'name': cell('Product name'), 'color': cell('Color')})
    return out


def compare(items, by_upc):
    """items: the RetailPortal rows.  by_upc: SAP UPC -> [(article, style, color name, size)].  Returns one result row per item."""
    out = []
    for it in items:
        hits = next((by_upc[f] for f in gtin_forms(it['gtin']) if f in by_upc), None)
        row = {'GTIN': it['gtin'], 'Product': it['name'], 'SO color': it['color'],
               'Article': '', 'Style': '', 'Size': '', 'SAP color': '', 'Status': 'OK', 'Details': ''}
        if not hits:
            row.update(Status='ERROR', Details='This UPC is not on any SAP article.')
        else:
            art, style, color, size = hits[0]
            row.update(**{'Article': art, 'Style': style, 'Size': size, 'SAP color': color})
            notes = []
            if len({h[0] for h in hits}) > 1:
                row['Status'] = 'REVIEW'; notes.append('The UPC is on several SAP articles: ' + ', '.join(sorted({h[0] for h in hits})) + '.')
            if (it['color'] or '').strip().upper() != (color or '').strip().upper():
                row['Status'] = 'ERROR'; notes.append(f"Color differs: RetailPortal {it['color'] or 'blank'}, SAP {color or 'blank'}.")
            row['Details'] = ' '.join(notes)
        out.append(row)
    return out


def sap_lookup(conn, items):
    forms = sorted({f for it in items for f in gtin_forms(it['gtin'])})
    by_upc, cur = {}, conn.cursor()
    for i in range(0, len(forms), 500):
        part = forms[i:i + 500]
        cur.execute('SELECT "MARA_ARTICLE","F_GENERIC","CHARVAL_COLOR_CODE_DESCRIPTION","MARA_SIZE_CODE","MARA_UPC" FROM ' + ARTICLE_VIEW +
                    ' WHERE "MARA_UPC" IN (' + ','.join('?' * len(part)) + ') AND "F_ARTICLE_TYPE" <> ?', tuple(part) + ('C',))
        for art, style, color, size, upc in cur.fetchall():
            by_upc.setdefault(str(upc or '').strip(), []).append((str(art or '').strip(), str(style or '').strip(), str(color or '').strip(), str(size or '').strip()))
    return by_upc


def chrome_state():
    try:
        with urllib.request.urlopen(ITEMHUB_CDP.rstrip('/') + '/json', timeout=2) as resp:
            tabs = json.load(resp)
    except Exception:
        return {'chrome': False, 'tab': False}
    return {'chrome': True, 'tab': any(SO_HOST in (t.get('url') or '') for t in tabs)}


def read_retailportal():
    """The rows of the RetailPortal table on screen, through the signed-in Chrome.  Reads the page; clicks nothing."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(ITEMHUB_CDP)
        page = next((pg for c in browser.contexts for pg in c.pages if SO_HOST in pg.url), None)
        if page is None:
            raise RuntimeError('No RetailPortal tab is open in the signed-in Chrome. Open the catalog (Products) in it, then click again.')
        res = page.evaluate(READ_TABLE_JS)
    items = parse_table(res['heads'], res['rows'])
    if not items:
        raise RuntimeError('The RetailPortal tab shows no items with a GTIN/UPC. Open the catalog (Products) table, then click again.')
    return items


@app.route('/retailportal/state', methods=['GET'])
def retailportal_state():
    return jsonify(chrome_state())


@app.route('/retailportal/run', methods=['POST'])
def retailportal_run():
    try:
        items = read_retailportal()
        conn = connect_hana()
        try:
            by_upc = sap_lookup(conn, items)
        finally:
            conn.close()
        rows = compare(items, by_upc)
        count = lambda s: sum(1 for r in rows if r['Status'] == s)
        return jsonify({'rows': rows, 'summary': {'checked': len(rows), 'ok': count('OK'), 'review': count('REVIEW'), 'error': count('ERROR')},
                        'when': datetime.now().strftime('%Y-%m-%d %H:%M')})
    except Exception as exc:
        return jsonify({'error': str(exc)}), 400
