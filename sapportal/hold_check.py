"""The hold master on the Validation tab, reduced to the two things asked of it: which color codes the file gets wrong, and what is in stock.

Both lists are built from the same check results the Validation tab already produced (nothing new is asked of SAP), one line per style and
color code, and exported as one workbook."""

import io
from flask import jsonify
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from .config import SIZE_ORDER, _cache
from .routes import _send_xlsx
from .webapp import app

COLOR_KINDS = (('✗', 'WRONG NAME'), ('⚠ NEW ON STYLE', 'NEW ON STYLE'), ('⚠ NOT ON ANY ARTICLE', 'NOT IN SAP'), ('⚠ NO COLOR CODE', 'NO CODE'), ('⚠ NO BUY COLOR', 'NO COLOR NAME'))


def _kind(check):
    return next((label for start, label in COLOR_KINDS if str(check or '').startswith(start)), '')


def _problem(kind, r):
    code, sap = r.get('color_code') or '', r.get('sap_color') or ''
    return {'WRONG NAME': f'SAP calls {code} "{sap}" on this style; the file says "{r.get("buy_color") or ""}".',
            'NEW ON STYLE': f'{code} ("{sap}") is used on other SAP articles but not on this style yet.',
            'NOT IN SAP': f'No SAP article uses color code {code}.',
            'NO CODE': 'The file has no color code on this row.',
            'NO COLOR NAME': f'The file has no color name; SAP calls {code} "{sap}".'}.get(kind, '')


def color_mismatches(lines):
    """One entry per style / color code / color name typed in the file whose color check did not pass, with the SAP name for that code."""
    out = {}
    for r in lines:
        kind = _kind(r.get('color_check'))
        if not kind:
            continue
        key = (r['style'], r.get('color_code') or '', r.get('buy_color') or '')
        e = out.setdefault(key, {'style': key[0], 'code': key[1], 'file_color': key[2], 'sap_color': r.get('sap_color') or '', 'kind': kind,
                                 'problem': _problem(kind, r), 'customers': set(), 'units': 0})
        e['customers'].add(r.get('customer') or '')
        e['units'] += int(r.get('qty') or 0)
    rows = [dict(e, customers=sorted(c for c in e['customers'] if c)) for e in out.values()]
    return sorted(rows, key=lambda e: (e['kind'], e['style'], e['code']))


def in_stock(lines):
    """One entry per style / color code with stock on hand: the units on the hold orders, what is on hand, what is on hand in each storage
    location (segments are kept for the hover), and the sizes. Stock belongs to an article, so a size on several customers' lines is counted once."""
    groups, seen = {}, set()
    for r in lines:
        st = r.get('stock')
        key = (r['style'], r.get('color_code') or '')
        g = groups.setdefault(key, {'style': key[0], 'code': key[1], 'sap_color': r.get('sap_color') or '', 'hold': 0, 'inv': 0, 'sizes': {}, 'locations': {}, 'alt': {}, 'customers': set()})
        g['hold'] += int(r.get('qty') or 0)
        g['customers'].add(r.get('customer') or '')
        sz = g['sizes'].setdefault(r.get('size') or '', {'size': r.get('size') or '', 'hold': 0, 'inv': 0, 'article': r.get('sap_article') or ''})
        sz['hold'] += int(r.get('qty') or 0)
        art = r.get('sap_article') or ''
        if art in seen:
            continue
        seen.add(art)
        if st and st.get('inv', 0) > 0:
            g['inv'] += st['inv']; sz['inv'] += st['inv']
            for b in st.get('by') or []:
                if b.get('inv'):
                    loc = g['locations'].setdefault(b.get('sloc') or '—', {'inv': 0, 'segments': {}})
                    seg = b.get('segment') or '—'
                    loc['inv'] += b['inv']; loc['segments'][seg] = loc['segments'].get(seg, 0) + b['inv']
        for a in r.get('alt_stock') or []:
            if a.get('inv'):
                g['alt'][a['code']] = g['alt'].get(a['code'], 0) + a['inv']
    order = lambda z: SIZE_ORDER.index(z) if z in SIZE_ORDER else 99
    rows = []
    for g in groups.values():
        if g['inv'] <= 0 and not g['alt']:
            continue
        rows.append({'style': g['style'], 'code': g['code'], 'sap_color': g['sap_color'], 'hold': g['hold'], 'inv': g['inv'],
                     'locations': {k: {'inv': v['inv'], 'segments': v['segments']} for k, v in sorted(g['locations'].items())},
                     'alt': [{'code': c, 'inv': n} for c, n in sorted(g['alt'].items())],
                     'sizes': sorted(g['sizes'].values(), key=lambda s: order(s['size'])), 'customers': sorted(c for c in g['customers'] if c)})
    return sorted(rows, key=lambda e: (e['style'], e['code']))


def hold_view(validation):
    lines = validation.get('generic') or []
    stock = in_stock(lines)
    return {'colors': color_mismatches(lines), 'stock': stock, 'locations': sorted({loc for e in stock for loc in e['locations']}),
            'file': validation.get('filename'), 'lines': len(lines)}


def _loaded_hold():
    v = _cache.get('validation')
    if not v or v.get('file_type') != 'hold_master':
        return None, (jsonify({'error': 'Load a hold master on the Validation tab first.'}), 400)
    return v, None


@app.route('/hold_check', methods=['GET'])
def hold_check_view():
    v, err = _loaded_hold()
    return err if err else jsonify(hold_view(v))


@app.route('/hold_check/export', methods=['GET'])
def hold_check_export():
    v, err = _loaded_hold()
    if err:
        return err
    view = hold_view(v)
    locs = view['locations']
    wb = Workbook()
    ws = wb.active; ws.title = 'COLOR MISMATCHES'
    ws.append(['Style', 'Color code (file)', 'Color name (file)', 'What SAP has for that code', 'Problem', 'Type', 'Customers', 'Hold units'])
    for e in view['colors']:
        ws.append([e['style'], e['code'], e['file_color'], e['sap_color'], e['problem'], e['kind'], ', '.join(e['customers']), e['units']])
    st = wb.create_sheet('IN STOCK')
    st.append(['Style', 'Color code', 'SAP color', 'Hold units', 'On hand'] + locs + ['Same color under other code(s)'])
    for e in view['stock']:
        st.append([e['style'], e['code'], e['sap_color'], e['hold'], e['inv']] + [e['locations'].get(loc, {}).get('inv', 0) for loc in locs]
                  + ['; '.join(f"{a['code']}: {a['inv']}" for a in e['alt'])])
    sz = wb.create_sheet('BY SIZE')
    sz.append(['Style', 'Color code', 'Size', 'Article', 'Hold units', 'On hand'])
    for e in view['stock']:
        for s in e['sizes']:
            sz.append([e['style'], e['code'], s['size'], s['article'], s['hold'], s['inv']])
    for sheet, widths in ((ws, (14, 12, 26, 30, 70, 14, 40, 10)), (st, (14, 11, 30, 11, 11) + (10,) * len(locs) + (30,)), (sz, (14, 11, 7, 16, 11, 10))):
        for c in sheet[1]:
            c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864')
        sheet.freeze_panes = 'A2'; sheet.auto_filter.ref = sheet.dimensions
        for i, w in enumerate(widths):
            sheet.column_dimensions[chr(65 + i)].width = w
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return _send_xlsx(buf, 'Hold_Check')
