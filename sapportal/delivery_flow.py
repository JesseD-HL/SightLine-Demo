"""Delivery flow: split the validated buysheet along a delivery-flow file or delivery chart."""

import io
import re
import traceback
from datetime import datetime
from flask import jsonify, request, send_file
from openpyxl import load_workbook
from .config import _cache
from .helpers import base_style, num
from .outputs import XLSM, make_buysheet
from .routes import _send_xlsx
from .webapp import app


# ═══════════════════════════════════════════════════════════════════════════════
# DELIVERY FLOW: split the validated buysheet along a delivery-flow file (quantity per style per CRD column)
# ═══════════════════════════════════════════════════════════════════════════════
def _as_date(v):
    if v is None or v == '': return None
    if hasattr(v, 'date') and callable(v.date): return v.date()
    if hasattr(v, 'year') and hasattr(v, 'month') and hasattr(v, 'day'): return v
    t = str(v).strip()
    for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m/%d/%y', '%m-%d-%Y'):
        try: return datetime.strptime(t[:10] if fmt == '%Y-%m-%d' else t, fmt).date()
        except ValueError: pass
    return None

def parse_flow_file(file_bytes):
    """The delivery-flow sheet: blocks that each start with a CRD row (the dates over the delivery columns), then a header row
    holding Style# and QTY, then one row per style with its quantity under each CRD column. A style written as
    'SAM271HD04 (CPU000FH02)' is known by both numbers. Returns {style number -> entry}, notes."""
    ws = load_workbook(io.BytesIO(file_bytes), data_only=True).active
    flow, notes = {}, []
    style_ci = qty_ci = None
    crd = {}                                           # column -> the CRD date over it, for the block being read
    for ri, row in enumerate(ws.iter_rows(values_only=True), 1):
        cells = [str(v).strip().upper() if v is not None else '' for v in row]
        if 'CRD' in cells:
            label = cells.index('CRD')
            crd = {ci: _as_date(v) for ci, v in enumerate(row) if ci > label and _as_date(v)}
            continue
        if 'STYLE#' in cells:
            style_ci = cells.index('STYLE#'); qty_ci = cells.index('QTY') if 'QTY' in cells else style_ci + 1
            continue
        if style_ci is None or not crd or style_ci >= len(row): continue
        text = str(row[style_ci] or '').strip()
        if not text or text.upper().startswith('NORTHWIND TOTAL'): continue
        tokens = re.findall(r'[A-Z0-9]{6,}', text.upper())
        if not tokens: continue
        per = {}
        for ci, d in crd.items():
            q = num(row[ci]) if ci < len(row) else None
            if q: per[d] = per.get(d, 0) + int(round(q))
        if not per: notes.append(f'Row {ri}: {text} has no delivery quantities'); continue
        total = int(round(num(row[qty_ci]) or 0)) if qty_ci < len(row) else 0
        if total and total != sum(per.values()):
            notes.append(f'Row {ri}: {text} QTY {total:,} but its deliveries add to {sum(per.values()):,}')
        entry = next((flow[t] for t in tokens if t in flow), None)
        if entry:
            for d, q in per.items(): entry['per'][d] = entry['per'].get(d, 0) + q          # the same style listed twice: add them
            entry['qty'] += total
        else:
            entry = {'label': text, 'tokens': tokens, 'per': per, 'qty': total}
        for t in tokens: flow[t] = entry
    if not flow:
        return parse_flow_by_customer(ws)
    return flow, notes

def flow_customer_key(name):
    """A customer name for matching: upper case, spaces tidied, nothing else. The flow's customer must equal the buysheet's exactly;
    there is no fuzzy matching and no alias list (a flow that does not match is sent back to be fixed, see flow_customer_mismatches)."""
    return re.sub(r'\s+', ' ', str(name or '').strip().upper())

def flow_customer_mismatches(rows, flow):
    """The customers a delivery chart names that are not, exactly, a customer of the buysheet. Returns [(customer as written on the chart,
    the buysheet's customers of the same retailer)]. A flow with no customers (the Target layout) has nothing to check."""
    have = {}
    for r in rows:
        have.setdefault(flow_customer_key(r.get('customer')), str(r.get('customer') or '').strip())
    out, seen = [], set()
    for k, e in flow.items():
        if not isinstance(k, tuple): return []
        ck = k[1]
        if ck in have or ck in seen: continue
        seen.add(ck)
        retailer = ck.split(' ')[0]
        out.append((e.get('cust') or ck, sorted(v for kk, v in have.items() if kk.split(' ')[0] == retailer)))
    return out

def parse_flow_by_customer(ws):
    """The delivery-chart layout: a header row with STYLE, CUSTOMER and QTY, the CRD dates (PROPOSE SALES CRD) across that same row, then
    one row per style and customer with its quantity under each date. The EST IN WHSE rows under the dates are not used. Returns
    {(style number, customer key) -> entry}, notes."""
    flow, notes = {}, []
    style_ci = cust_ci = qty_ci = None
    crd = {}
    for ri, row in enumerate(ws.iter_rows(values_only=True), 1):
        cells = [re.sub(r'\s+', ' ', str(v).strip().upper()) if v is not None else '' for v in row]
        if 'STYLE' in cells and 'CUSTOMER' in cells:
            style_ci, cust_ci = cells.index('STYLE'), cells.index('CUSTOMER')
            qty_ci = cells.index('QTY') if 'QTY' in cells else None
            crd = {ci: _as_date(v) for ci, v in enumerate(row) if ci > cust_ci and _as_date(v)}
            continue
        if style_ci is None or not crd or max(style_ci, cust_ci) >= len(row): continue
        text, cust = str(row[style_ci] or '').strip(), str(row[cust_ci] or '').strip()
        if not text or not cust: continue
        tokens = re.findall(r'[A-Z0-9]{6,}', text.upper())
        if not tokens: continue
        per = {d: int(round(num(row[ci]))) for ci, d in crd.items() if ci < len(row) and num(row[ci])}
        if not per: continue
        total = int(round(num(row[qty_ci]) or 0)) if qty_ci is not None and qty_ci < len(row) else 0
        if total and total != sum(per.values()):
            notes.append(f'Row {ri}: {text} {cust} QTY {total:,} but its deliveries add to {sum(per.values()):,}')
        ck = flow_customer_key(cust)
        entry = next((flow[(t, ck)] for t in tokens if (t, ck) in flow), None)
        if entry:
            for d, q in per.items(): entry['per'][d] = entry['per'].get(d, 0) + q
            entry['qty'] += total
        else:
            entry = {'label': f'{text} · {cust}', 'cust': cust, 'tokens': tokens, 'per': per, 'qty': total}
        for t in tokens: flow[(t, ck)] = entry
    return flow, notes

def _alloc(n, weights, per_unit, carry):
    """Spread n whole units over the deliveries in proportion to weights, exactly (the sum is always n). Leftover units go to the
    deliveries furthest behind their share (running over the whole style), ties to the LAST delivery."""
    total_w = sum(weights)
    ideal = [n * w / total_w for w in weights]
    out = [int(x) for x in ideal]
    left = n - sum(out)
    cand = sorted((j for j, w in enumerate(weights) if w > 0), key=lambda j: ((ideal[j] - out[j]) * per_unit + carry[j], j), reverse=True)
    for j in cand[:left]: out[j] += 1
    for j in range(len(weights)): carry[j] += (ideal[j] - out[j]) * per_unit
    return out

def _alloc_cartons(n, carton, weights, per_unit, carry):
    """n units in whole cartons of `carton` (1 = no carton size given). Every delivery is a multiple of the carton; a part carton
    (n not a multiple) goes to the last delivery that has a share, so the total never changes."""
    if not carton or carton <= 1: return _alloc(n, weights, per_unit, carry), 0
    cartons, rem = divmod(n, carton)
    got = [c * carton for c in _alloc(cartons, weights, per_unit * carton, carry)]
    if rem:
        last = max(j for j, w in enumerate(weights) if w > 0)
        got[last] += rem
    return got, rem

def split_buysheet(rows, flow):
    """The validated buysheet rows split along the flow. Every quantity stays exact: per size for open stock, in whole packs for
    prepacks (a complex prepack in whole complex packs, the same pack count for every color). Returns (rows, report)."""
    notes, report, out_of = [], [], {}
    entries = {id(e): e for e in flow.values()}
    seen_styles = set()

    by_customer = any(isinstance(k, tuple) for k in flow)   # a delivery chart has a line per style AND customer; the Target flow one per style

    def entry_for(r):
        ck = flow_customer_key(r.get('customer')) if by_customer else None
        for t in (str(r.get('style') or '').upper(), base_style(str(r.get('raw_style') or r.get('style') or '')).upper()):
            m = re.match(r'^(.+\d)[A-Z]$', t)                  # CPU000FH01A: a lettered variant counts under CPU000FH01 on the flow
            for cand in [t] + ([m.group(1)] if m else []):
                k = (cand, ck) if by_customer else cand
                if k in flow: return flow[k]
        return None

    groups = []                                        # (kind, [rows]) in file order; a complex prepack's color rows are ONE group
    cx = {}
    for r in rows:
        if r.get('complex'):
            key = (r.get('raw_style'), r.get('customer'), str(r.get('delivery')))
            if key not in cx:
                cx[key] = ('C', []); groups.append(cx[key])
            cx[key][1].append(r)
        else:
            groups.append(('O' if r.get('type') == 'OS' else 'P', [r]))

    state = {}                                         # id(entry) -> {'dates', 'weights', 'carry', 'split'}
    result = {}                                        # id(group) -> list of new rows
    for g in groups:
        kind, grp = g
        e = entry_for(grp[0])
        if not e: result[id(g)] = grp; seen_styles.add(str(grp[0].get('style') or '')); continue
        st = state.setdefault(id(e), {'dates': sorted(e['per']), 'carry': None, 'split': {}, 'buy': 0, 'rows': 0})
        dates = st['dates']; weights = [e['per'][d] for d in dates]
        if st['carry'] is None: st['carry'] = [0.0] * len(dates)
        carry = st['carry']
        new = []

        def emit(base, j, **chg):
            n = dict(base); n.update(chg); n['split_crd'] = dates[j]; new.append(n)

        if kind == 'O':
            r = grp[0]; carton = num(r.get('pc_ctn_os') or r.get('pc_ctn')) or 0
            carton = int(carton) if float(carton).is_integer() else 0
            per_delivery = [dict() for _ in dates]
            for sz, qty in (r.get('sizes') or {}).items():
                carton_sz = (r.get('pc_by_size') or {}).get(sz) or carton
                got, rem = _alloc_cartons(int(qty), int(carton_sz), weights, 1, carry)
                if rem: notes.append(f"{r.get('style')} {r.get('color')} {sz}: {qty:,} is not whole cartons of {carton_sz}; the {rem} left over went to the last delivery")
                for j, q in enumerate(got):
                    if q: per_delivery[j][sz] = q
                st['buy'] += int(qty)
            for j, sizes in enumerate(per_delivery):
                if sizes:
                    emit(r, j, sizes=sizes); st['split'][dates[j]] = st['split'].get(dates[j], 0) + sum(sizes.values())
        else:
            packs, per_pack = [], []
            ok = True
            for r in grp:
                pp = num(r.get('row_pc') if kind == 'C' else r.get('pc_per_pac')) or 0
                pieces = num(r.get('ppk_qty')) or 0
                if not pp or pieces % pp:
                    ok = False; notes.append(f"{r.get('raw_style')} {r.get('color_code') or ''}: {pieces:,.0f} is not a whole number of packs of {pp:g}; left unsplit")
                    break
                packs.append(int(pieces // pp)); per_pack.append(pp)
            if ok and len(set(packs)) > 1:
                ok = False; notes.append(f"{grp[0].get('raw_style')}: its colors give different pack counts ({', '.join(f'{x:,}' for x in packs)}); left unsplit")
            if not ok:
                result[id(g)] = grp; continue
            carton = num(grp[0].get('pac_per_ctn')) or 0
            carton = int(carton) if float(carton).is_integer() else 0
            whole = sum(per_pack)                                      # pieces in one whole (complex) pack
            got, rem = _alloc_cartons(packs[0], carton, weights, whole, carry)
            if rem: notes.append(f"{grp[0].get('raw_style')}: {packs[0]:,} packs is not whole cartons of {carton}; the {rem} left over went to the last delivery")
            for j, pk in enumerate(got):
                if not pk: continue
                for r, pp in zip(grp, per_pack):
                    emit(r, j, ppk_qty=int(pk * pp))
                st['split'][dates[j]] = st['split'].get(dates[j], 0) + int(pk * whole)
            st['buy'] += int(packs[0] * whole)
        result[id(g)] = new
        st['rows'] += len(grp)
        seen_styles.add(str(grp[0].get('style') or ''))

    new_rows = [r for g in groups for r in result.get(id(g), g[1])]
    for eid, st in state.items():
        e = entries[eid]
        report.append({'style': e['label'], 'flow_total': sum(e['per'].values()), 'buy_total': st['buy'], 'rows': st['rows'],
                       'deliveries': [{'crd': d.isoformat(), 'flow': e['per'][d], 'split': st['split'].get(d, 0)} for d in st['dates']]})
        if st['buy'] != sum(e['per'].values()):
            notes.append(f"{e['label']}: the buysheet has {st['buy']:,} pieces but the flow has {sum(e['per'].values()):,}")
    def listed(items, n=25): return ', '.join(items[:n]) + (f' and {len(items) - n} more' if len(items) > n else '')
    unused = [e['label'] for eid, e in entries.items() if eid not in state]
    if unused: notes.append(f'In the flow file but not found in the buysheet ({len(unused)}): ' + listed(unused))
    kept = sorted({str(g[1][0].get('style') or '') + (f" ({g[1][0].get('customer')})" if by_customer else '') for g in groups if entry_for(g[1][0]) is None})
    if kept: notes.append(f'In the buysheet but not in the flow file, kept as they are ({len(kept)}): ' + listed(kept))
    return new_rows, {'styles': report, 'notes': notes, 'rows_in': len(rows), 'rows_out': len(new_rows)}

@app.route('/delivery_flow/preview', methods=['POST'])
def delivery_flow_preview():
    if not _cache.get('rows'): return jsonify({'error': 'Validate the buysheet first (PREBUY tab).'}), 400
    if (_cache.get('validation') or {}).get('file_type') != 'buysheet':
        return jsonify({'error': 'The delivery split needs a buysheet (a file with a DATA sheet) validated in the portal; the loaded file is a prebuy.'}), 400
    if 'file' not in request.files: return jsonify({'error': 'No file uploaded'}), 400
    try:
        flow, fnotes = parse_flow_file(request.files['file'].read())
        if not flow: return jsonify({'error': 'No styles found in the delivery flow file (looking for a CRD row, a Style# / QTY header, then style rows).'}), 400
        bad = flow_customer_mismatches(_cache['rows'], flow)
        if bad:                                                 # the flow must name the buysheet's customers exactly: send it back, do not guess
            lines = [f"• {c}  ({'the buysheet has: ' + ', '.join(have) if have else 'no customer of this retailer in the buysheet'})" for c, have in bad]
            return jsonify({'error': "The delivery chart's customer names must match the buysheet's exactly. Fix these on the chart and upload it again:\n" + '\n'.join(lines),
                            'mismatches': [{'chart_customer': c, 'buysheet_customers': have} for c, have in bad]}), 400
        new_rows, report = split_buysheet(_cache['rows'], flow)
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    report['notes'] = fnotes + report['notes']
    _cache['split'] = {'rows': new_rows}
    return jsonify(report)

@app.route('/delivery_flow/download', methods=['GET'])
def delivery_flow_download():
    sp = _cache.get('split')
    if not sp: return jsonify({'error': 'Split a buysheet first.'}), 400
    try:
        buf, kind = make_buysheet(sp['rows'], _cache.get('customers', []))
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    if kind == 'xlsm':
        return send_file(buf, as_attachment=True, download_name=f"Buysheet_SPLIT_{datetime.now():%m-%d-%Y}.xlsm", mimetype=XLSM)
    return _send_xlsx(buf, 'Buysheet_SPLIT')
