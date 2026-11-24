"""PO / requisition check: pulls a PO or requisition from SAP, compares the two, the margin file, and the buysheet."""

import io
import re
from collections import defaultdict
from openpyxl import load_workbook
from .config import BUILD
from .helpers import iso_date, num, pad10, up


# ═══════════════════════════════════════════════════════════════════════════════
# 7c. PO / REQ CHECK  (read-only: pulls a PO or requisition straight from SAP, compares the two, and checks the margin file)
# ═══════════════════════════════════════════════════════════════════════════════
PO_VIEW    = '"_SYS_BIC"."Acme.Data_Universe.Raw_Complete/P2P_RAW_PURCHASE_ORDERS"'
REQ_VIEW   = '"_SYS_BIC"."Acme.Data_Universe.Raw_Complete/P2P_RAW_REQUISITIONS"'
FIBER_VIEW = '"_SYS_BIC"."Acme.Data_Universe.Raw_Supporting/P2P_SUP_PO_FIBER_CONTENTS"'
SET_TABLE  = '"SAPDEMO"."ZSETCOMM"'

PO_COLS = ['EKKO_PO_NUMBER', 'EKPO_PO_ITEM', 'EKKO_PO_TYPE_ID', 'T161T_PO_TYPE_TEXT', 'EKKO_VENDOR_ID', 'EKKO_PURCH_ORG_ID', 'EKKO_RELEASE_INDICATOR',
           'EKKO_PO_DATE', 'EKPA_FACTORY_ID', 'EKPA_TRADING_CO_ID', 'EKPO_ARTICLE', 'EKPO_ARTICLE_SHORT_TEXT', 'EKPO_STORAGE_LOCATION_ID',
           'EKPO_PREQ_NUMBER', 'EKPO_PREQ_ITEM_NO', 'EKPO_DELETION_INDICATOR', 'EBAN_DELIVERY_DATE', 'EKPO_STAT_EXFTY_DATE', 'EKPO_STAT_DLVY_DATE',
           'EKET_SCHEDULE_LINE', 'EKET_DELIVERY_DATE', 'DLHD_HANDOVER_DATE', 'DLIT_01_DATE', 'DLIT_02_DATE', 'DLIT_03_DATE', 'DLIT_05_DATE',
           'F_EKET_PO_QTY_PC', 'F_GENERIC', 'MARM_PC_PER_PAC', 'T006A_PO_COMMERCIAL_UNIT', 'MAW1_COMMODITY_CODE', 'EIPO_COMMODITY_CODE',
           'EIPO_DISPATCH_COUNTRY', 'EIPO_TRADING_COUNTRY', 'EKPO_PO_QTY_PUOM', 'F_PO_QTY_PC', 'F_REQ_QTY_PC', 'F_NET_PRICE_PER_PC', 'F_PB00_PER_PC',
           'F_ZFS1_PER_PC', 'F_ZCST_RATE', 'F_ZCSR_RATE', 'F_ZTRF_RATE']
REQ_COLS = ['EBAN_REQUISITION_NO', 'EBAN_REQUISITION_ITEM', 'EBAN_ARTICLE', 'EBAN_DELETION_INDICATOR', 'EBAN_DELIVERY_DATE', 'EBAN_STORAGE_LOCATION',
            'EBAN_PURCHASE_ORDER_NO', 'EBAN_PO_ITEM_NO', 'F_REQ_QTY_PC', 'EBAN_QTY_DUOM', 'T006A_UNIT_COMMERCIAL', 'EBAN_RELEASE_INDICATOR', 'MAW1_COMMODITY_CODE', 'F_GENERIC', 'EBAN_DOC_TYPE']


def _s(v):
    return '' if v is None else str(v).strip()

def blank_date(v):
    s = _s(v)
    return '' if s in ('', '0', '00000000') else s

def _chunks(items, n=500):
    items = list(items)
    for i in range(0, len(items), n):
        yield items[i:i + n]

def _select(conn, view, cols, where_col, values):
    """SELECT the named columns where one column is IN a list of bound values. Names are fixed constants above; values are parameters."""
    out = []
    cur = conn.cursor()
    for part in _chunks(sorted(set(values))):
        marks = ', '.join('?' for _ in part)
        cur.execute(f'SELECT {", ".join(chr(34) + c + chr(34) for c in cols)} FROM {view} WHERE "{where_col}" IN ({marks})', tuple(part))
        names = [d[0] for d in cur.description]
        out += [dict(zip(names, r)) for r in cur.fetchall()]
    return out

def fetch_lookup_data(conn, number, kind='auto'):
    """Everything SAP holds for one PO or requisition number. Returns (kind, po_rows, req_rows, fiber_rows, set_rows, notes)."""
    n = pad10(number)
    if not n: raise ValueError('Enter a PO or requisition number.')
    notes = []
    po_rows, req_rows = [], []
    if kind in ('auto', 'po'):
        po_rows = _select(conn, PO_VIEW, PO_COLS, 'EKKO_PO_NUMBER', [n])
        if po_rows: kind = 'po'
    if not po_rows and kind in ('auto', 'req'):
        kind = 'req'
        req_rows = _select(conn, REQ_VIEW, REQ_COLS, 'EBAN_REQUISITION_NO', [n])
        if not req_rows: raise LookupError(f'{n} was not found as a PO or as a requisition.')
    if kind == 'po' and not po_rows: raise LookupError(f'PO {n} was not found.')
    if kind == 'po':
        reqs = {pad10(r['EKPO_PREQ_NUMBER']) for r in po_rows if pad10(r['EKPO_PREQ_NUMBER'])}
        if reqs: req_rows = _select(conn, REQ_VIEW, REQ_COLS, 'EBAN_REQUISITION_NO', reqs)
        else: notes.append('This PO has no requisition on any line.')
        if reqs:                                                  # every PO that draws on those requisitions, to see the split across POs
            known = {(r['EKKO_PO_NUMBER'], r['EKPO_PO_ITEM'], r['EKET_SCHEDULE_LINE']) for r in po_rows}
            po_rows += [r for r in _select(conn, PO_VIEW, PO_COLS, 'EKPO_PREQ_NUMBER', reqs) if (r['EKKO_PO_NUMBER'], r['EKPO_PO_ITEM'], r['EKET_SCHEDULE_LINE']) not in known]
    else:
        po_rows = _select(conn, PO_VIEW, PO_COLS, 'EKPO_PREQ_NUMBER', [n])
    generics = sorted({_s(r.get('F_GENERIC')) for r in po_rows + req_rows if _s(r.get('F_GENERIC'))})
    articles = sorted({_s(r.get('EKPO_ARTICLE')) for r in po_rows if _s(r.get('EKPO_ARTICLE'))} | {_s(r.get('EBAN_ARTICLE')) for r in req_rows if _s(r.get('EBAN_ARTICLE'))})
    fiber_rows, set_rows = [], []
    try:
        fiber_rows = _select(conn, FIBER_VIEW, ['ARTICLE', 'VENDOR', 'PURCHORG', 'FIBERCONT', 'FROMDELIVERY', 'TODELIVERY'], 'ARTICLE', generics) if generics else []
    except Exception as e:
        notes.append(f'Fiber lookup unavailable: {e}')
    try:
        set_rows = _select(conn, SET_TABLE, ['MATNR', 'EKORG', 'COMP_COMM_CODE', 'COMP_COMM_DESC', 'FOB_PRICE', 'Z1ST_SALE_PRICE', 'QUANTITY', 'VALID_FROM', 'VALID_TO'], 'MATNR', articles) if articles else []
    except Exception as e:
        notes.append(f'Set code lookup unavailable: {e}')
    return kind, po_rows, req_rows, fiber_rows, set_rows, notes

# ── requisition vs buysheet ─────────────────────────────────────────────────────
def _ymd(d):
    """A date as YYYYMMDD whatever it came as (date, '2027-04-09', '20270409'); '' when blank."""
    if not d: return ''
    if hasattr(d, 'strftime'): return d.strftime('%Y%m%d')
    t = re.sub(r'\D', '', str(d)[:10])
    return t if len(t) == 8 else ''

def _show_ymd(t):
    return f'{t[:4]}-{t[4:6]}-{t[6:]}' if len(t) == 8 else (t or '(blank)')

def buysheet_req_check(req_rows, rows, validation, variant_map, filename=''):
    """Check the requisition(s) in req_rows against the buysheet. A buysheet customer's requisition number is on the HEADER sheet
    (row['req']). Per requisition: every article, with its pieces, delivery date and storage location. Pure: no database, no files."""
    reqs = sorted({pad10(r['EBAN_REQUISITION_NO']) for r in req_rows if pad10(r['EBAN_REQUISITION_NO'])})
    out = {'file': filename, 'reqs': [], 'issues': [], 'notes': []}
    if not rows:
        out['notes'].append('No buysheet loaded: drop it on the PREBUY tab to check the requisition against it.')
        return out
    ppk_rows = [r for r in rows if r.get('type') == 'PPK']
    ppk_res = (validation or {}).get('ppk') or []
    ppk_of = {id(r): res for r, res in zip(ppk_rows, ppk_res)} if len(ppk_rows) == len(ppk_res) else {}
    new = lambda: {'qty': 0.0, 'dates': set(), 'stor': set(), 'customers': set()}
    expected = defaultdict(lambda: defaultdict(new))                  # req -> article -> what the buysheet says
    seen_complex, counts = set(), {}
    for r in rows:
        req = r.get('req')
        if not req or req not in reqs: continue
        lines = []
        if r.get('type') == 'OS':
            for sz, q in (r.get('sizes') or {}).items():
                art = variant_map.get((str(r['style']).upper(), str(r.get('color_code') or '').upper(), sz.upper())) or f"{r['style']}-{r.get('color_code', '')} {sz} (not in SAP)"
                lines.append((up(art), q))
        else:
            res = ppk_of.get(id(r))
            art = up(res['ppk_article'] if res else r['raw_style'])
            pieces = sum((r.get('sizes') or {}).values()) or num(r.get('ppk_qty')) or 0
            per_pack = num(r.get('row_pc') if r.get('complex') else r.get('pc_per_pac'))      # a prepack requisition is in PACKS, open stock in pieces
            packs = round(pieces / per_pack) if per_pack else 0
            if r.get('complex'):                                      # a complex prepack's color rows are ONE order: its packs count once
                key = (art, str(r.get('customer') or ''), str(r.get('delivery') or ''))
                counts.setdefault((req, art), []).append(packs)       # every color row of one order should give the same pack count
                if key in seen_complex: packs = 0
                seen_complex.add(key)
            lines.append((art, packs))
        for art, q in lines:
            e = expected[req][art]
            e['qty'] += q or 0; e['dates'].add(_ymd(r.get('delivery'))); e['stor'].add(up(str(r.get('warehouse') or '')))
            e['customers'].add(str(r.get('customer') or ''))
    os_articles = set(variant_map.values())
    actual = defaultdict(lambda: defaultdict(new))                    # req -> article -> what SAP says (lines not flagged deleted)
    for r in req_rows:
        if _s(r.get('EBAN_DELETION_INDICATOR')): continue
        e = actual[pad10(r['EBAN_REQUISITION_NO'])][up(_s(r['EBAN_ARTICLE']))]
        packs = _s(r.get('T006A_UNIT_COMMERCIAL')).upper() == 'PAC' and r.get('EBAN_QTY_DUOM') is not None      # a prepack line is in packs (F_REQ_QTY_PC is the pieces in them)
        e['qty'] += (num(r.get('EBAN_QTY_DUOM')) if packs else num(r.get('F_REQ_QTY_PC'))) or 0; e['dates'].add(_ymd(blank_date(r.get('EBAN_DELIVERY_DATE')))); e['stor'].add(up(_s(r.get('EBAN_STORAGE_LOCATION'))))

    unit = lambda art: 'pc' if art in os_articles else 'packs'      # open stock is in pieces, a prepack in packs
    def issue(level, req, art, msg): out['issues'].append({'level': level, 'req': req, 'article': art, 'msg': msg})
    for req in reqs:
        who = sorted({c for e in expected[req].values() for c in e['customers']})
        if not expected[req]:
            out['reqs'].append({'req': req, 'customers': [], 'buysheet_lines': 0, 'sap_lines': len(actual[req])})
            issue('error', req, '', f'Requisition {req} is not on the buysheet HEADER sheet, so it cannot be checked against the buysheet.')
            continue
        out['reqs'].append({'req': req, 'customers': who, 'buysheet_lines': len(expected[req]), 'sap_lines': len(actual[req])})
        for art in sorted(set(expected[req]) | set(actual[req])):
            b, s = expected[req].get(art), actual[req].get(art)
            if b and not s: issue('error', req, art, f"On the buysheet ({', '.join(sorted(b['customers']))}, {b['qty']:g} {unit(art)}) but not on the requisition")
            elif s and not b: issue('error', req, art, f"On the requisition ({s['qty']:g} {unit(art)}) but not on the buysheet for {', '.join(who)}")
            else:
                if abs(b['qty'] - s['qty']) > 0.5: issue('error', req, art, f"Quantity: buysheet {b['qty']:g} {unit(art)} vs requisition {s['qty']:g} {unit(art)}")
                bd, sd = b['dates'] - {''}, s['dates'] - {''}
                if bd and sd and bd != sd: issue('warn', req, art, f"Delivery date: buysheet {', '.join(map(_show_ymd, sorted(bd)))} vs requisition {', '.join(map(_show_ymd, sorted(sd)))}")
                bs, ss = b['stor'] - {''}, s['stor'] - {''}
                if bs and ss and bs != ss: issue('warn', req, art, f"Storage location: buysheet {', '.join(sorted(bs))} vs requisition {', '.join(sorted(ss))}")
    for (req, art), c in sorted(counts.items()):
        if max(c) - min(c) > 0.5: issue('warn', req, art, f"The color rows of this complex prepack on the buysheet give different pack counts ({', '.join(f'{x:g}' for x in c)}); the first is used")
    out['errors'] = sum(1 for i in out['issues'] if i['level'] == 'error'); out['warnings'] = sum(1 for i in out['issues'] if i['level'] == 'warn')
    return out

# ── margin file ─────────────────────────────────────────────────────────────────
MARGIN_FIELDS = {'CUSTOMER': 'customer', 'DELIVERY': 'delivery', 'REQ#': 'req', 'PO#': 'po', 'SELL PRICE': 'sell', 'QTY': 'qty', 'GENERIC': 'generic',
                 'COLOR DESC': 'color', 'SPLIT BY COLOR': 'split_color', 'VENDOR': 'vendor', 'FACTORY': 'factory', 'FOB': 'fob', 'TERM': 'term',
                 'DUTY RATE': 'duty_rate', 'SIZE RANGE': 'size_range', 'FIRST SALE VARIABLE': 'first_sale_var', 'COO': 'coo',
                 'MAIN FABRICATION': 'fabrication', 'ADDITIONAL TARIFF': 'add_tariff', 'FINAL DUTY': 'final_duty', 'FINAL PO FOB': 'final_fob',
                 'FIRST SALE': 'first_sale'}

def generic_key(v):
    """The style code in a margin GENERIC cell ('RDM271SS06F -TABLE PROGRAM' -> 'RDM271SS06F')."""
    t = _s(v).upper().split()
    return t[0] if t else ''

def parse_margin(file_bytes):
    """Rows of the 'Margin data recap' sheet as dicts (values only; the file is read with its saved formula results)."""
    if file_bytes[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
        raise ValueError('This file is password-protected (or an old .xls). In Excel: File > Info > Protect Workbook > Encrypt with Password, clear the password, and save as .xlsx. Then upload it again.')
    if file_bytes[:2] != b'PK':
        raise ValueError('This is not an Excel .xlsx/.xlsm file. Open it in Excel and save it as .xlsx.')
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    name = next((s for s in wb.sheetnames if 'MARGIN' in s.upper() and 'RECAP' in s.upper()), None) or wb.sheetnames[0]
    rows = list(wb[name].iter_rows(values_only=True))
    hi = next((i for i, r in enumerate(rows[:15]) if r and 'GENERIC' in [_s(c).upper() for c in r] and 'FOB' in [_s(c).upper() for c in r]), None)
    if hi is None: raise ValueError(f"No header row with GENERIC and FOB found on sheet '{name}'.")
    cols = {}
    for i, h in enumerate(rows[hi]):
        key = MARGIN_FIELDS.get(_s(h).upper())
        if key and key not in cols: cols[key] = i
    for need in ('generic', 'fob'):
        if need not in cols: raise ValueError(f'Margin sheet has no {need.upper()} column.')
    out = []
    for ri, r in enumerate(rows[hi + 1:], hi + 2):
        g = generic_key(r[cols['generic']] if cols['generic'] < len(r) else '')
        if not g: continue
        d = {k: (r[i] if i < len(r) else None) for k, i in cols.items()}
        d.update(row=ri, key=g, req=pad10(d.get('req')), po=pad10(d.get('po')))
        for k in ('qty', 'fob', 'duty_rate', 'add_tariff', 'final_duty', 'final_fob', 'first_sale', 'sell'):
            d[k] = num(d.get(k))
        d['delivery'] = iso_date(d.get('delivery')) if d.get('delivery') else ''
        d['term'] = _s(d.get('term')); d['vendor'] = _s(d.get('vendor')).upper(); d['factory'] = _s(d.get('factory')).upper()
        d['color'] = _s(d.get('color') or d.get('split_color')); d['coo'] = _s(d.get('coo')).upper(); d['fabrication'] = _s(d.get('fabrication'))
        d['size_range'] = _s(d.get('size_range')); d['customer'] = _s(d.get('customer')); d.pop('split_color', None)
        out.append(d)
    return {'sheet': name, 'rows': out}

# ── fiber text ──────────────────────────────────────────────────────────────────
_FIBER_ALIAS = {'POLY': 'POLYESTER', 'POLYESTER': 'POLYESTER', 'ELASTANE': 'SPANDEX', 'LYCRA': 'SPANDEX', 'SPANDEX': 'SPANDEX', 'ELASTANE.': 'SPANDEX',
                'POLYAMIDE': 'NYLON', 'NYLON': 'NYLON', 'COTTON': 'COTTON', 'CTN': 'COTTON', 'VISCOSE': 'RAYON', 'RAYON': 'RAYON'}

def fiber_pairs(text):
    """'80% POLYESTER / 20% SPANDEX' -> {'POLYESTER': 80.0, 'SPANDEX': 20.0}. Percent then fiber name; anything after the name is ignored."""
    pairs = {}
    for pct, word in re.findall(r'(\d+(?:\.\d+)?)\s*%\s*(?:RECYCLED\s+|ORGANIC\s+)?([A-Z]+)', _s(text).upper()):
        name = _FIBER_ALIAS.get(word, word)
        pairs[name] = pairs.get(name, 0) + float(pct)
    return pairs

def fiber_compare(sap_text, margin_text):
    """('ok'|'diff'|'manual', note). More than one set of percentages (size-specific fibers) goes to a person."""
    a, b = fiber_pairs(sap_text), fiber_pairs(margin_text)
    if not a: return 'diff', 'SAP fiber text has no percentages: ' + _s(sap_text)
    if not b: return 'diff', 'Margin file fabrication has no percentages: ' + _s(margin_text)
    if abs(sum(a.values()) - 100) > 0.5 or abs(sum(b.values()) - 100) > 0.5:
        return 'manual', f'Fiber percentages do not add to 100 (SAP: {_s(sap_text)} | margin: {_s(margin_text)}), check by eye'
    return ('ok', '') if a == b else ('diff', f'SAP {_s(sap_text)} vs margin {_s(margin_text)}')

# ── the comparison ──────────────────────────────────────────────────────────────
TOL_PRICE = 0.0101      # one cent
TOL_RATE  = 0.051       # rates are compared as percent: 0.32 in the margin file = 32 in SAP

def build_lookup(kind, number, po_rows, req_rows, fiber_rows, set_rows, margin=None, notes=None):
    """Join PO lines to requisition lines and run every check. Pure: no database, no files. Returns the JSON the portal shows."""
    n = pad10(number)
    notes = list(notes or [])
    # PO item rows carry quantity and price; schedule-line rows carry delivery dates. Merge them per PO item.
    items = {}
    for r in po_rows:
        key = (_s(r['EKKO_PO_NUMBER']), _s(r['EKPO_PO_ITEM']))
        it = items.setdefault(key, {'item_row': None, 'sched': []})
        if _s(r.get('EKET_SCHEDULE_LINE')):
            it['sched'].append(r)
        else:
            it['item_row'] = r
    req_by_key = {(pad10(r['EBAN_REQUISITION_NO']), _s(r['EBAN_REQUISITION_ITEM'])): r for r in req_rows}
    # pieces drawn on each requisition item, across every PO (the one asked about and any others)
    drawn = defaultdict(float)
    for (po, item), it in items.items():
        base = it['item_row'] or (it['sched'][0] if it['sched'] else None)
        if base is None: continue
        q = num(it['item_row'].get('F_PO_QTY_PC')) if it['item_row'] else None
        if q is None: q = sum(num(s.get('F_EKET_PO_QTY_PC')) or 0 for s in it['sched'])
        it['qty_pc'] = q
        if pad10(base.get('EKPO_PREQ_NUMBER')) and _s(base.get('EKPO_DELETION_INDICATOR')) != 'L':
            drawn[(pad10(base['EKPO_PREQ_NUMBER']), _s(base['EKPO_PREQ_ITEM_NO']))] += q or 0

    fiber_by = defaultdict(list)
    for f in fiber_rows: fiber_by[_s(f['ARTICLE']).upper()].append(f)
    set_by = defaultdict(list)
    for s in set_rows: set_by[_s(s['MATNR']).upper()].append(s)

    mrows = (margin or {}).get('rows') or []
    my_reqs = {pad10(r['EBAN_REQUISITION_NO']) for r in req_rows} | ({n} if kind == 'req' else set())
    tagged = [m for m in mrows if (kind == 'po' and m['po'] == n) or (m['req'] and m['req'] in my_reqs)]
    pool = tagged if tagged else mrows
    margin_note = ''
    if not mrows: margin_note = 'No margin file loaded.'
    elif tagged: margin_note = f'{len(tagged)} margin rows carry this REQ#/PO#; only those are used.'
    else: margin_note = 'Margin rows are not tagged with this REQ#/PO#, so they are matched by style only and quantity is not compared.'

    out_items = []
    focus_po = n if kind == 'po' else None
    for (po, item), it in sorted(items.items()):
        base = it['item_row'] or (it['sched'][0] if it['sched'] else None)
        if base is None: continue
        if focus_po and po != focus_po: continue                        # other POs only feed the "other POs" quantity below
        out_items.append(_check_item(po, item, it, base, req_by_key, drawn, fiber_by, set_by, pool, bool(tagged)))
    # requisition lines nobody ordered (requisition lookups, or items of the requisition not on this PO)
    used = {(i['req'], i['req_item']) for i in out_items if i['req']}
    for (rq, ri), r in sorted(req_by_key.items()):
        if (rq, ri) in used: continue
        if kind == 'po' and drawn.get((rq, ri)): continue               # ordered on another PO: shown there
        if kind == 'req' and drawn.get((rq, ri)): continue
        out_items.append(_req_only_item(rq, ri, r))
    margin_only = []
    if tagged:                                                          # quantity per style: margin file vs everything the PO carries for that style
        po_sum, first_at = defaultdict(float), {}
        for idx, i in enumerate(out_items):
            po_sum[i['generic']] += i['po_qty_pc'] or 0
            first_at.setdefault(i['generic'], idx)
        for g, idx in first_at.items():
            mrs = [m for m in tagged if _style_match(m['key'], g)]
            if mrs:
                mq = sum(m['qty'] or 0 for m in mrs)
                out_items[idx]['margin']['qty'] = mq
                if abs(mq - po_sum[g]) > 0.5: _flag(out_items[idx]['flags'], 'warn', f'Style quantity: margin file {mq:g} pc vs {po_sum[g]:g} pc on the PO (all colors and sizes)')
        margin_only = sorted({m['key'] for m in tagged if not any(_style_match(m['key'], g) for g in first_at)})
    summary = {'items': len(out_items), 'error': sum(1 for i in out_items for f in i['flags'] if f['level'] == 'error'),
               'warn': sum(1 for i in out_items for f in i['flags'] if f['level'] == 'warn'),
               'clean': sum(1 for i in out_items if not any(f['level'] in ('error', 'warn') for f in i['flags']))}
    return {'build': BUILD, 'kind': kind, 'number': n, 'items': out_items, 'summary': summary, 'notes': notes, 'margin_note': margin_note,
            'margin_only': margin_only, 'pos': sorted({i['po'] for i in out_items if i['po']}), 'reqs': sorted({i['req'] for i in out_items if i['req']})}

def _flag(flags, level, msg):
    flags.append({'level': level, 'msg': msg})

def _style_match(margin_key, sap_generic):
    """Same style code; a margin code one or two characters longer (RDM271SS06F vs RDM271SS06) still counts."""
    a, b = margin_key.upper(), sap_generic.upper()
    return bool(a and b) and (a == b or (a.startswith(b) and len(a) - len(b) <= 2) or (b.startswith(a) and len(b) - len(a) <= 2))

def _same(a, b, tol=TOL_PRICE):
    return a is not None and b is not None and abs(a - b) <= tol

def _check_item(po, item, it, base, req_by_key, drawn, fiber_by, set_by, pool, tagged):
    flags = []
    ir = it['item_row'] or base
    article = _s(base.get('EKPO_ARTICLE')); generic = _s(base.get('F_GENERIC')).upper() or article.split('-')[0].upper()
    rq, ri = pad10(base.get('EKPO_PREQ_NUMBER')), _s(base.get('EKPO_PREQ_ITEM_NO'))
    req = req_by_key.get((rq, ri)) if rq else None
    qty = it.get('qty_pc')
    fob = num(ir.get('F_NET_PRICE_PER_PC'))
    if fob is None: fob = num(base.get('F_PB00_PER_PC'))
    zfs1, zcsr, zcst, ztrf = (num(base.get(k)) for k in ('F_ZFS1_PER_PC', 'F_ZCSR_RATE', 'F_ZCST_RATE', 'F_ZTRF_RATE'))
    comm_m, comm_p = _s(base.get('MAW1_COMMODITY_CODE')), _s(base.get('EIPO_COMMODITY_CODE'))
    sched = [{'line': _s(s.get('EKET_SCHEDULE_LINE')), 'delivery': blank_date(s.get('EKET_DELIVERY_DATE')), 'ex_factory': blank_date(s.get('DLHD_HANDOVER_DATE') or s.get('DLIT_01_DATE')),
              'etd': blank_date(s.get('DLIT_02_DATE')), 'eta': blank_date(s.get('DLIT_03_DATE')), 'qty_pc': num(s.get('F_EKET_PO_QTY_PC'))} for s in it['sched']]
    ex_fact = next((s['ex_factory'] for s in sched if s['ex_factory']), '') or blank_date(base.get('EKPO_STAT_EXFTY_DATE'))
    po_deliv = next((s['delivery'] for s in sched if s['delivery']), '') or blank_date(base.get('EKPO_STAT_DLVY_DATE'))
    is_fs = _s(base.get('EKKO_PO_TYPE_ID')) == 'Z4'

    # ── SAP: nothing blank, HTS agrees ──
    for label, val in (('Article master commodity code (HTS)', comm_m), ('PO commodity code (HTS)', comm_p), ('Vendor', _s(base.get('EKKO_VENDOR_ID'))),
                       ('Factory', _s(base.get('EKPA_FACTORY_ID'))), ('Storage location', _s(base.get('EKPO_STORAGE_LOCATION_ID')))):
        if not val: _flag(flags, 'error', f'{label} is blank in SAP')
    if comm_m and comm_p and comm_m != comm_p: _flag(flags, 'error', f'HTS differs: article master {comm_m} vs PO {comm_p}')
    if not fob: _flag(flags, 'error', 'PO price is blank or zero')
    if is_fs and not zfs1: _flag(flags, 'error', 'First sale PO (Z4) but first sale price (ZFS1) is blank or zero')
    if zcsr is None or zcst is None: _flag(flags, 'error', 'Duty rate is blank in SAP (ZCSR base duty or ZCST duty)')
    if ztrf is None: _flag(flags, 'warn', 'US import tariff rate (ZTRF) is blank in SAP')
    if _s(base.get('EKKO_RELEASE_INDICATOR')) not in ('Y', ''): _flag(flags, 'warn', f"PO release indicator is '{_s(base.get('EKKO_RELEASE_INDICATOR'))}' (not released)")
    if _s(base.get('EKPO_DELETION_INDICATOR')): _flag(flags, 'error', f"PO line deletion indicator is '{_s(base.get('EKPO_DELETION_INDICATOR'))}'")
    if not sched: _flag(flags, 'warn', 'PO line has no delivery schedule line')

    # ── PO vs requisition ──
    req_qty = num(req.get('F_REQ_QTY_PC')) if req else num(ir.get('F_REQ_QTY_PC'))
    other = None
    if not rq:
        _flag(flags, 'warn', 'PO line is not tied to a requisition')
    elif not req:
        _flag(flags, 'warn', f'Requisition {rq}/{ri} was not found in the requisition view')
    else:
        if _s(req.get('EBAN_DELETION_INDICATOR')): _flag(flags, 'error', f"Requisition line {rq}/{ri} is flagged deleted ('{_s(req.get('EBAN_DELETION_INDICATOR'))}') but a PO draws on it")
        if _s(req.get('EBAN_ARTICLE')) and article and _s(req.get('EBAN_ARTICLE')) != article: _flag(flags, 'error', f"Article differs: requisition {_s(req.get('EBAN_ARTICLE'))} vs PO {article}")
        if _s(req.get('EBAN_STORAGE_LOCATION')) != _s(base.get('EKPO_STORAGE_LOCATION_ID')): _flag(flags, 'warn', f"Storage location differs: requisition {_s(req.get('EBAN_STORAGE_LOCATION'))} vs PO {_s(base.get('EKPO_STORAGE_LOCATION_ID'))}")
        total = drawn.get((rq, ri), 0)
        other = total - (qty or 0)
        if req_qty is not None and qty is not None:
            if abs(total - req_qty) > 0.5:
                _flag(flags, 'error' if total > req_qty else 'warn', f'Quantity: requisition {req_qty:g} pc vs {total:g} pc on POs' + (f' ({qty:g} on this PO, {other:g} on others)' if other else ''))
    # ── fiber ──
    fib = [f for f in fiber_by.get(generic, [])]
    d = po_deliv or blank_date(base.get('EBAN_DELIVERY_DATE'))
    vend = _s(base.get('EKKO_VENDOR_ID')).upper()
    vends = {vend, _s(base.get('EKPA_FACTORY_ID')).upper(), _s(base.get('EKPA_TRADING_CO_ID')).upper()} - {''}
    org = _s(base.get('EKKO_PURCH_ORG_ID')).upper()
    in_org = [f for f in fib if not org or not _s(f['PURCHORG']) or _s(f['PURCHORG']).upper() == org]      # fiber rows are per purchasing org
    def window(f): return not d or (_s(f['FROMDELIVERY']) <= d <= _s(f['TODELIVERY']))
    live = [f for f in in_org if _s(f['VENDOR']).upper() in vends and window(f)]            # this vendor's own row wins
    if not live: live = [f for f in in_org if not _s(f['VENDOR']) and window(f)]            # a row with no vendor applies to every vendor
    sap_fiber = next((_s(f['FIBERCONT']) for f in live), '')
    if not sap_fiber:
        mine = [f for f in in_org if _s(f['VENDOR']).upper() in vends or not _s(f['VENDOR'])]
        if mine: _flag(flags, 'error', f"Fiber content for {generic} / {vend} does not cover delivery {d or '(no date)'}: valid " + '; '.join(f"{_s(f['FROMDELIVERY'])}-{_s(f['TODELIVERY'])}" for f in mine))
        else: _flag(flags, 'error', f'No fiber content in SAP for {generic} / {vend}' + (f' (other vendors have it: {sorted({_s(f["VENDOR"]) for f in fib})})' if fib else ''))
    elif len({_s(f['FIBERCONT']) for f in live}) > 1:
        _flag(flags, 'warn', 'More than one fiber content is valid for this style, vendor and date')
    # ── set components (ZSETCOMM) ──
    comps = set_by.get(article.upper(), [])
    for c in comps:
        if not _s(c['COMP_COMM_CODE']): _flag(flags, 'error', f"Set component of {article} has no commodity code in ZSETCOMM")
    # ── margin file ──
    mr = [m for m in pool if _style_match(m['key'], generic)]
    if mr and any(m['key'] != generic for m in mr): _flag(flags, 'info', f"Margin file style {mr[0]['key']} matched to SAP style {generic} by prefix")
    msum = {'rows': len(mr)}
    if not pool:
        pass
    elif not mr:
        _flag(flags, 'warn', f'Style {generic} is not in the margin file')
    else:
        fobs = sorted({m['final_fob'] for m in mr if m['final_fob'] is not None})
        if fob is not None and fobs and not any(_same(fob, x) for x in fobs): _flag(flags, 'error', f'PO price {fob:g} is not a margin FINAL PO FOB ({", ".join(f"{x:g}" for x in fobs)})')
        terms = {m['term'] for m in mr if m['term']}
        fs_rows = [m for m in mr if m['term'].upper().startswith('FIRST')]
        if is_fs and terms and not fs_rows: _flag(flags, 'warn', f'SAP PO is first sale (Z4) but margin Term is {", ".join(sorted(terms))}')
        if not is_fs and fs_rows and zfs1: _flag(flags, 'info', 'Margin Term is First Sale; SAP PO type is not Z4')
        if fs_rows and zfs1 is not None:
            fsv = sorted({m['first_sale'] for m in fs_rows if m['first_sale'] is not None})
            if fsv and not any(_same(zfs1, x) for x in fsv): _flag(flags, 'error', f'First sale: SAP ZFS1 {zfs1:g} vs margin {", ".join(f"{x:g}" for x in fsv)}')
        for label, sap_v, key in (('Base duty (ZCSR)', zcsr, 'duty_rate'), ('US tariff (ZTRF) vs margin Additional Tariff', ztrf, 'add_tariff')):
            vals = sorted({round(m[key] * 100, 3) for m in mr if m[key] is not None})
            if sap_v is not None and vals and not any(_same(sap_v, x, TOL_RATE) for x in vals):
                _flag(flags, 'warn', f'{label}: SAP {sap_v:g}% vs margin {", ".join(f"{x:g}%" for x in vals)}')
        fin = sorted({round(m['final_duty'] * 100, 3) for m in mr if m['final_duty'] is not None})
        sap_fin = [v for v in (zcst, (zcsr or 0) + (ztrf or 0) if zcsr is not None and ztrf is not None else None, (zcst or 0) + (ztrf or 0) if zcst is not None and ztrf is not None else None) if v is not None]
        if fin and sap_fin and not any(_same(a, b, TOL_RATE) for a in sap_fin for b in fin):
            _flag(flags, 'warn', f'Final duty: margin {", ".join(f"{x:g}%" for x in fin)} matches none of SAP Duty ZCST {zcst:g}%, base+tariff {(zcsr or 0) + (ztrf or 0):g}%, Duty+tariff {(zcst or 0) + (ztrf or 0):g}%')
        mv = {m['vendor'] for m in mr if m['vendor']}
        if mv and vend and not any(vend in x or x in vend for x in mv): _flag(flags, 'warn', f'Vendor: SAP {vend} vs margin {", ".join(sorted(mv))}')
        fabs = sorted({m['fabrication'] for m in mr if m['fabrication']})
        if sap_fiber and fabs:
            res = [fiber_compare(sap_fiber, f) for f in fabs]
            if not any(r[0] == 'ok' for r in res):
                lvl = 'info' if all(r[0] == 'manual' for r in res) else 'warn'
                _flag(flags, lvl, 'Fiber: ' + res[0][1])
        msum.update(fob=fobs, term=sorted(terms), coo=sorted({m['coo'] for m in mr if m['coo']}), first_sale=sorted({m['first_sale'] for m in mr if m['first_sale'] is not None}),
                    duty_rate=sorted({m['duty_rate'] for m in mr if m['duty_rate'] is not None}), fabrication=fabs, rows_at=[m['row'] for m in mr][:20])
    return {'po': po, 'po_item': item, 'req': rq, 'req_item': ri, 'article': article, 'generic': generic, 'desc': _s(base.get('EKPO_ARTICLE_SHORT_TEXT')),
            'vendor': vend, 'factory': _s(base.get('EKPA_FACTORY_ID')), 'po_type': _s(base.get('EKKO_PO_TYPE_ID')), 'po_type_text': _s(base.get('T161T_PO_TYPE_TEXT')),
            'storage_po': _s(base.get('EKPO_STORAGE_LOCATION_ID')), 'storage_req': _s(req.get('EBAN_STORAGE_LOCATION')) if req else '',
            'req_qty_pc': req_qty, 'po_qty_pc': qty, 'other_po_qty_pc': other, 'unit': _s(base.get('T006A_PO_COMMERCIAL_UNIT')),
            'pc_per_pac': num(base.get('MARM_PC_PER_PAC')), 'fob_pc': fob, 'zfs1': zfs1, 'zcsr': zcsr, 'zcst': zcst, 'ztrf': ztrf,
            'hts_master': comm_m, 'hts_po': comm_p, 'dispatch': _s(base.get('EIPO_DISPATCH_COUNTRY')), 'trading': _s(base.get('EIPO_TRADING_COUNTRY')),
            'req_date': blank_date(req.get('EBAN_DELIVERY_DATE')) if req else blank_date(base.get('EBAN_DELIVERY_DATE')),
            'po_delivery': po_deliv, 'ex_factory': ex_fact, 'sched': sched, 'fiber': sap_fiber,
            'set_components': [{'code': _s(c['COMP_COMM_CODE']), 'desc': _s(c['COMP_COMM_DESC']), 'fob': num(c['FOB_PRICE']), 'first_sale': num(c['Z1ST_SALE_PRICE'])} for c in comps],
            'margin': msum, 'flags': flags}

def _req_only_item(rq, ri, r):
    flags = []
    if _s(r.get('EBAN_DELETION_INDICATOR')): _flag(flags, 'info', 'Requisition line is deleted')
    else: _flag(flags, 'info', 'No PO draws on this requisition line yet')
    if not _s(r.get('MAW1_COMMODITY_CODE')): _flag(flags, 'error', 'Article master commodity code (HTS) is blank in SAP')
    po = pad10(r.get('EBAN_PURCHASE_ORDER_NO'))
    if po and not _s(r.get('EBAN_DELETION_INDICATOR')): _flag(flags, 'warn', f'Requisition view says PO {po} item {_s(r.get("EBAN_PO_ITEM_NO"))} but no PO line was found')
    return {'po': '', 'po_item': '', 'req': rq, 'req_item': ri, 'article': _s(r.get('EBAN_ARTICLE')), 'generic': _s(r.get('F_GENERIC')).upper(), 'desc': '',
            'vendor': '', 'factory': '', 'po_type': '', 'po_type_text': '', 'storage_po': '', 'storage_req': _s(r.get('EBAN_STORAGE_LOCATION')),
            'req_qty_pc': num(r.get('F_REQ_QTY_PC')), 'po_qty_pc': None, 'other_po_qty_pc': None, 'unit': '', 'pc_per_pac': None, 'fob_pc': None, 'zfs1': None,
            'zcsr': None, 'zcst': None, 'ztrf': None, 'hts_master': _s(r.get('MAW1_COMMODITY_CODE')), 'hts_po': '', 'dispatch': '', 'trading': '',
            'req_date': blank_date(r.get('EBAN_DELIVERY_DATE')), 'po_delivery': '', 'ex_factory': '', 'sched': [], 'fiber': '', 'set_components': [],
            'margin': {'rows': 0}, 'flags': flags}
