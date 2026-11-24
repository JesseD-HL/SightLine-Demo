"""What the portal writes: buysheet DATA + HEADER OUT, the .xlsm template fill, creation templates, the PR CSV."""

import io
import os
import re
import zipfile
from collections import defaultdict
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from .config import KNOWN_SIZES, SIZE_ORDER, TOOLS_DIR, get_weights
from .hana import is_canadian_style
from .helpers import (
    NO_SEASON_CODE,
    NO_SEASON_YEAR,
    canonical_size,
    dash_list,
    no_season_style,
    num,
    season_from_style,
    up)


# ═══════════════════════════════════════════════════════════════════════════════
# 6. OUTPUTS - buysheet DATA + HEADER OUT, creation templates, PR CSV
# ═══════════════════════════════════════════════════════════════════════════════
NAVY  = PatternFill('solid', start_color='1F4E79')
WBOLD = Font(bold=True, color='FFFFFF', name='Arial', size=10)
DFONT = Font(name='Arial', size=10)
BFONT = Font(bold=True, name='Arial', size=10)
THIN  = Side(style='thin', color='BFBFBF')
BDR   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CTR   = Alignment(horizontal='center', vertical='center')
LEFT  = Alignment(horizontal='left',   vertical='center', indent=1)

# the buysheet's DATA sheet: these columns, in this order, with the prebuy's size columns between them
DATA_HEAD = ['CUSTOMER','STYLE','SEGMENT','COLOR CODE','COLOR','DESCRIPTION','CRD','DELIVERY','PRICE','MSRP',
             'SIZE RATIO','SIZES','PC PER PAC','PAC PER CTN','PPK QTY']
DATA_TAIL = ['PC PER CTN OS','DATE UPDATED','COMMENTS']
HEADER_OUT = ['CUSTOMER','Tickets','Hang/Fold','PER BAG','Warehouse','Pre-Buy Rcvd','Buysheet Created','REQ#','TERMS','TERMS DESC']

def buysheet_customer(r):
    """The buysheet code reads OS / PPK / COMPLEX out of the customer name, so every customer carries one of them."""
    kind = 'OS' if r.get('type') == 'OS' else ('COMPLEX' if r.get('complex') or r.get('pack_kind') == 'COMPLEX' else 'PPK')
    name = str(r.get('customer') or '').strip().upper()
    if re.search(rf'\b{kind}\b', name):
        return name
    return f"{name} - {kind}" if name else kind

def data_size_columns(rows):
    """Size columns of the buysheet: only sizes the PREBUY has quantities for, in the prebuy's own column order.
    Nothing is assumed: no fixed list, no renaming."""
    used = {sz for r in rows if r.get('type') == 'OS' for sz in (r.get('sizes') or {})}
    order = []
    for r in rows:
        for sz in r.get('size_order') or []:
            if sz in used and sz not in order:
                order.append(sz)
    return order + sorted(used - set(order))

def add_filters(wb, header_rows=None):
    """Every export is a filterable table: a filter on the header row of each sheet and the header frozen. header_rows: {sheet name: row} (default 1)."""
    for ws in wb.worksheets:
        hr = (header_rows or {}).get(ws.title, 1)
        if ws.max_row >= hr and ws.max_column >= 1:
            ws.auto_filter.ref = f"A{hr}:{get_column_letter(ws.max_column)}{max(ws.max_row, hr + 1)}"
            ws.freeze_panes = f"A{hr + 1}"
    return wb

def generic_pricing_table(ws, first_col=11):
    """A second table on the PRICING sheet, to the right of the variant table (which is left as it is): one row per generic with its
    sell price (ZSLS) and MSRP (ZVP1), from the variant rows.  A generic with the same price on every row is one line; one the buysheet prices two ways is one line per price."""
    seen, rows = set(), []
    for r in range(2, ws.max_row + 1):
        if not ws.cell(r, 1).value: continue
        mat, buy, msrp = ws.cell(r, 2).value, ws.cell(r, 5).value, ws.cell(r, 6).value      # Material, BUY PRICE, MSRP
        if (mat, buy, msrp) in seen: continue
        seen.add((mat, buy, msrp)); rows.append((mat, buy, msrp))
    labels = ['Material', 'SELL (ZSLS)', 'MSRP (ZVP1)']                 # the generic and both of its prices, from the buysheet
    for i, h in enumerate(labels):
        c = ws.cell(1, first_col + i, value=h); c.font = WBOLD; c.fill = NAVY; c.alignment = CTR; c.border = BDR
        ws.column_dimensions[get_column_letter(first_col + i)].width = max(ws.column_dimensions[get_column_letter(first_col + i)].width or 0, 14)
    for k, (mat, buy, msrp) in enumerate(rows, 2):
        for i, v in enumerate([mat, buy, msrp]):
            c = ws.cell(k, first_col + i, value=v); c.font = DFONT; c.alignment = CTR; c.border = BDR
    if rows:                                              # a real Excel table: it has its own filter, beside the variant table's
        tab = Table(displayName='GenericPricing', ref=f"{get_column_letter(first_col)}1:{get_column_letter(first_col + len(labels) - 1)}{len(rows) + 1}")
        tab.tableStyleInfo = TableStyleInfo(name='TableStyleLight1', showRowStripes=False)
        ws.add_table(tab)
    return rows


def make_data_sheet(rows, customers):
    sizes   = data_size_columns(rows)
    headers = DATA_HEAD + sizes + DATA_TAIL
    col     = {h: i for i, h in enumerate(headers, 1)}
    wb = Workbook(); ws = wb.active; ws.title = "DATA"
    if rows:
        style = next((r['style'] for r in rows if not no_season_style(r['style'])), rows[0]['style'])      # a 000 style (CPU000FH01) has no season to read
        ws.cell(1, 1).value = up(rows[0].get('brand') or '')
        label = next((r.get('season_label') for r in rows if r.get('season_label')), '')
        ws.cell(1, 2).value = label.replace(' ', '-') if label else season_from_style(style)
    for ci, h in enumerate(headers, 1):
        c = ws.cell(3, ci, value=h); c.font = WBOLD; c.fill = NAVY; c.alignment = CTR; c.border = BDR
    ws.row_dimensions[3].height = 22
    today = datetime.now().date()
    for ri, r in enumerate(rows, 4):
        is_os = r.get('type') == 'OS'
        d = r.get('delivery')
        crd, dlv = (r['split_crd'], None) if 'split_crd' in r else (d, d)         # a split row: CRD from the delivery flow, DELIVERY left for the user
        vals = {
            'CUSTOMER': buysheet_customer(r), 'STYLE': up(r['style'] if is_os else r['raw_style']),
            'SEGMENT': up(r.get('segment') or ''), 'COLOR CODE': up(r.get('color_code') or ''), 'COLOR': up(r.get('color') or ''),
            'DESCRIPTION': up(r.get('desc') or ''), 'CRD': crd, 'DELIVERY': dlv, 'PRICE': num(r.get('price')), 'MSRP': num(r.get('msrp')),
            'SIZE RATIO': None if is_os else up(r.get('ratio') or ''), 'SIZES': None if is_os else up(r.get('sizes_str') or ''),
            'PC PER PAC': None if is_os else r.get('pc_per_pac'), 'PAC PER CTN': None if is_os else (num(r.get('pac_per_ctn')) or 1),       # a prepack with no PAC PER CTN typed: 1
            'PPK QTY': None if is_os else r.get('ppk_qty'),
            'PC PER CTN OS': num(r.get('pc_ctn_os') or r.get('pc_ctn')) if is_os else None,
            'DATE UPDATED': today, 'COMMENTS': up(r.get('notes') or ''),
        }
        if is_os:
            for sz, qty in (r.get('sizes') or {}).items():
                vals[sz] = qty
        for h, v in vals.items():
            c = ws.cell(ri, col[h], value=v if v != '' else None)
            if h in ('CRD', 'DELIVERY', 'DATE UPDATED'): c.number_format = 'mm/dd/yyyy'
        for ci in range(1, len(headers) + 1):
            c = ws.cell(ri, ci); c.font = DFONT; c.border = BDR
            c.alignment = LEFT if headers[ci - 1] in ('CUSTOMER', 'COLOR', 'DESCRIPTION', 'COMMENTS') else CTR
        ws.row_dimensions[ri].height = 18
    wide = {'CUSTOMER': 26, 'STYLE': 18, 'SEGMENT': 11, 'COLOR CODE': 11, 'COLOR': 22, 'DESCRIPTION': 36, 'CRD': 11, 'DELIVERY': 11,
            'PRICE': 8, 'MSRP': 8, 'SIZE RATIO': 14, 'SIZES': 16, 'PC PER PAC': 10, 'PAC PER CTN': 10, 'PPK QTY': 9,
            'PC PER CTN OS': 12, 'DATE UPDATED': 12, 'COMMENTS': 24}
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = wide.get(h, 7)

    # HEADER OUT: one line per customer exactly as it appears on DATA; everything comes from the prebuy header
    wh = wb.create_sheet("HEADER OUT")
    for ci, h in enumerate(HEADER_OUT, 1):
        c = wh.cell(1, ci, value=h); c.font = WBOLD; c.fill = NAVY; c.alignment = CTR; c.border = BDR
        wh.column_dimensions[get_column_letter(ci)].width = 30 if ci == 1 else 16
    seen = {}
    for r in rows:
        cust = buysheet_customer(r)
        if cust not in seen:
            terms = r.get('terms')
            seen[cust] = [cust, up(r.get('tickets') or ''), up(r.get('hang_fold') or ''), None, up(r.get('warehouse') or ''), None, today, None,
                          num(terms) if num(terms) is not None else up(terms or ''), up(r.get('terms_desc') or '')]
    for ri, vals in enumerate(seen.values(), 2):
        for ci, v in enumerate(vals, 1):
            c = wh.cell(ri, ci, value=v if v != '' else None); c.font = DFONT; c.border = BDR
            c.alignment = LEFT if ci == 1 else CTR
            if ci in (6, 7): c.number_format = 'mm/dd/yyyy'
    add_filters(wb, {'DATA': 3})
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def ppk_components(r):
    """[(size, units per pack)] for a prepack row, from the size grid (legacy prebuy / buysheet) or from
    the SIZES + SIZE RATIO text (new template). The creation template needs this per size."""
    if r.get('sizes'):
        return list(r['sizes'].items())
    sizes, ratio = dash_list(r.get('sizes_str')), dash_list(r.get('ratio'))
    if sizes and len(sizes) == len(ratio):
        try: return [(s, int(q)) for s, q in zip(sizes, ratio)]
        except ValueError: pass
    return []

PPK_UNIT = 'H5'      # the unit of measure on every PPK CREATION row (H4 is used up)

def make_creation_templates(validation):
    """GENERIC CREATION, PPK CREATION (one row per size of each missing prepack) and COLOR ADD, in the layout the MM41 automation scripts read.
    Only what has to be created is listed. GENERIC CREATION: styles SAP does not have, every color and size of the buysheet. COLOR ADD
    (third tab): the colors / sizes missing on styles SAP already has (colors missing the same sizes share a row)."""
    rows = validation.get('rows', [])
    missing_ppks     = [r for r in validation.get('ppk', []) if r['exists'] == '✗ MISSING']
    os_rows_by_style = defaultdict(list)
    for r in rows:
        if r['type'] == 'OS': os_rows_by_style[r['style']].append(r)
    ppk_rows_by_art = defaultdict(list)
    for r in rows:
        if r['type'] == 'PPK': ppk_rows_by_art[r['raw_style']].append(r)
    # generics that only a PREPACK needs (no open-stock line for them): created with the colors and sizes its components use
    prepack_generic = {}
    adds = defaultdict(lambda: defaultdict(set))       # a style SAP already has -> {color code: the sizes it lacks}
    for r in validation.get('generic', []):
        if r['status'] == '✗ MISSING' and r.get('missing_kind') != 'style':
            adds[r['style']][r['color_code']].add(r['size'])
    for res in missing_ppks:
        for n in (res.get('readiness') or {}).get('needs', []):
            if n['kind'] != 'generic':                   # a color / sizes a prepack needs on a style SAP already has
                if n['color']: adds[n['style']][n['color']].update(n['sizes'])
                continue
            e = prepack_generic.setdefault(n['style'], {'codes': [], 'sizes': []})
            if n['color'] and n['color'] not in e['codes']: e['codes'].append(n['color'])
            e['sizes'] += [z for z in n['sizes'] if z not in e['sizes']]
    new_styles = {r['style'] for r in validation.get('generic', []) if r['status'] == '✗ MISSING' and r.get('missing_kind') == 'style'} | set(prepack_generic)
    missing_generics = sorted(new_styles | set(adds))

    def header(ws, labels):
        for ci, h in enumerate(labels, 1):
            c = ws.cell(1, ci, value=h); c.font = WBOLD; c.fill = NAVY; c.alignment = CTR; c.border = BDR
    def write(ws, ri, vals):
        for ci, v in enumerate(vals, 1):
            c = ws.cell(ri, ci, value=v); c.font = DFONT; c.alignment = CTR; c.border = BDR
    season_formulas = lambda ri: (f'=IF(MID(A{ri},4,3)="000","{NO_SEASON_YEAR}","20"&MID(A{ri},4,2))',
                                  f'=IF(MID(A{ri},4,3)="000","{NO_SEASON_CODE}",CHOOSE(MID(A{ri},6,1),"SPR","SUM","FALL","HOL"))')       # (Year, Season) - the column order SAP uses; a 000 style carries no season 0000 / year 9999

    wb = Workbook()
    ws_g = wb.active; ws_g.title = "GENERIC CREATION"
    header(ws_g, ['Material','COLOR CODES','DESCRIPTION','Pcs per Carton','SIZES','WT2','WT','LB','1','2','3','4','VOL','FT3','Year','Season'])
    ri = 2
    color_adds, variant_jobs = [], []
    in_order =lambda sizes: ';'.join(sorted(sizes, key=lambda s_: SIZE_ORDER.index(s_) if s_ in SIZE_ORDER else 99))
    for style in missing_generics:
        style_rows, pg = os_rows_by_style.get(style, []), prepack_generic.get(style)
        ppk_style_rows = [r for r in rows if r['type'] == 'PPK' and r.get('style') == style]
        r0 = style_rows[0] if style_rows else (ppk_style_rows[0] if ppk_style_rows else None)
        if r0 is None: continue
        if style in new_styles:                          # SAP has no such style: all of it
            combos = [(sorted({r['color_code'] for r in style_rows} | set(pg['codes'] if pg else [])),
                       {sz for r in style_rows for sz in r['sizes']} | set(pg['sizes'] if pg else []))]
        else:                                            # SAP has the style: only the colors / sizes it lacks
            by_sizes = defaultdict(list)
            for color, sizes in adds[style].items(): by_sizes[frozenset(sizes)].append(color)
            combos = [(sorted(colors), set(sizes)) for sizes, colors in by_sizes.items()]
        wt, vol = get_weights(style, r0.get('desc', ''))
        for colors, sizes in combos:
            variant_jobs.append((style, colors, sizes, style in new_styles))
            if any(x.get('pc_by_size') for x in style_rows):
                cartons = {c for x in style_rows for c in (x.get('pc_by_size') or {}).values()}
                style_pc = next(iter(cartons)) if len(cartons) == 1 else ''
            else:
                style_pc = r0.get('pc_ctn_os', '') if style_rows else ''
            vals = [style, ';'.join(colors), r0.get('desc', ''), style_pc, in_order(sizes), wt, wt, 'LB', None, None, None, None, vol, 'FT3']
            if style in new_styles:
                write(ws_g, ri, vals + list(season_formulas(ri))); ri += 1
            else:
                color_adds.append(vals)                  # the style exists: goes on the COLOR ADD tab, not on GENERIC CREATION

    ws_p = wb.create_sheet("PPK CREATION")
    header(ws_p, ['Material','Description','Batch','Qty','Unit','WT2','WT','LB','1','2','3','4','VOL','FT3','Year','Season','PC PER PPK','PPK PER CTN','Mdse Catgry','COMPONENT','VARIANT STATUS'])
    FLAG = PatternFill('solid', start_color='FCE4E4')
    ri = 2
    by_article = defaultdict(list)
    for res in missing_ppks: by_article[res['ppk_article']].append(res)          # one block per prepack, however many orders / color rows it has
    for art, results in by_article.items():
        art_rows = ppk_rows_by_art.get(art, [])
        r = art_rows[0] if art_rows else {}
        comps, seen, conflicts = [], {}, []
        for res in results:                                   # the prepack's components, each (color code, size) once
            for c in (res.get('readiness') or {}).get('components', []):
                key = (c['color'], canonical_size(c['size']))
                if key in seen:
                    if seen[key]['qty'] != c['qty'] and f"{c['color']} {c['size']}" not in conflicts: conflicts.append(f"{c['color']} {c['size']}")
                    continue
                seen[key] = c; comps.append(c)
        is_complex = any(x.get('complex') for x in art_rows)
        mdse = next(((res.get('readiness') or {}).get('mdse_cat') for res in results if (res.get('readiness') or {}).get('mdse_cat')), '')
        color_name = str(r.get('color') or '').strip()
        desc = up(f"{r.get('desc', '')} - {color_name}") if (color_name and not is_complex) else up(r.get('desc', ''))     # SAP names a prepack "DESCRIPTION - COLOR"
        pc_ppk = sum(c['qty'] for c in comps) or None
        per_ctn = num(r.get('pac_per_ctn')) or 1
        per_ctn = int(per_ctn) if float(per_ctn).is_integer() else per_ctn
        wt, vol = get_weights(r.get('style', art), r.get('desc', ''))
        ppk_wt  = round(wt * pc_ppk, 4) if wt and pc_ppk else None
        ppk_vol = round(vol * pc_ppk, 4) if vol and pc_ppk else None
        style = r.get('style') or art.split('-')[0]
        say = {'generic': f"MISSING: generic {style} is not in SAP", 'color': None, 'size': None}
        for j, c in enumerate(comps):
            first = j == 0
            y_f, s_f = season_formulas(ri)
            if c['status'] == 'ok': note = 'OK'
            elif c['status'] == 'generic': note = say['generic']
            elif c['status'] == 'color': note = f"MISSING: color {c['color'] or '(no color code)'} is not on {style}"
            else: note = f"MISSING: size {c['size']} is not on {style} {c['color']}"
            if c.get('code_from_name'): note += ' (the COLOR CODE cell was blank: color taken from the article name)'
            if f"{c['color']} {c['size']}" in conflicts: note += ' (the ratio differs between orders: the first is used)'
            write(ws_p, ri, [art, desc if first else '', c['article'] or None, c['qty'], PPK_UNIT,        # Material on EVERY row: the MM41 script groups a prepack's rows by it
                             ppk_wt if first else '', ppk_wt if first else '', 'LB', None, None, None, None,
                             ppk_vol if first else '', 'FT3', y_f if first else '', s_f if first else '',
                             pc_ppk if first else '', per_ctn if first else '', mdse if first else '', f"{c['color']} {c['size']}".strip(), note])
            if c['status'] != 'ok':
                ws_p.cell(ri, 3).fill = FLAG; ws_p.cell(ri, 21).fill = FLAG
            ri += 1

    ws_c = wb.create_sheet("COLOR ADD")                  # third tab: new colors / sizes on styles SAP already has
    header(ws_c, ['Material','COLOR CODES','DESCRIPTION','Pcs per Carton','SIZES','WT2','WT','LB','1','2','3','4','VOL','FT3','Year','Season'])
    for k, vals in enumerate(color_adds, 2):
        write(ws_c, k, vals + list(season_formulas(k)))

    # fourth tab: pieces per carton, ONE SIZE PER ROW (the cartons differ by size, so one number per style cannot hold them)
    ws_k = wb.create_sheet("PC PER CTN")
    header(ws_k, ['Material', 'COLOR CODE', 'SIZE', 'Pcs per Carton', 'Article'])     # the first four keep their order; Article = the variant     # the first four columns keep their order; Article = the variant to set it on

    # fifth tab: the VARIANT ARTICLES to be created, one row per color x size. A style SAP does not have starts at 001; a style it
    # has continues after its last variant (any color / size). Colors in code order, sizes in size order, numbered in that order.
    ws_v = wb.create_sheet("VARIANTS")
    header(ws_v, ['Article', 'Unit', 'H5', 'Material', 'COLOR CODE', 'SIZE', 'NOTE'])
    last, note_all = {}, ''
    old_styles = sorted({j[0] for j in variant_jobs if not j[3]})
    if old_styles:
        try:
            from .hana import connect_hana, query_variant_articles
            conn = connect_hana()
            try: existing = query_variant_articles(conn, old_styles)
            finally: conn.close()
            for (gen, _c, _s), art in existing.items():
                m_ = re.fullmatch(re.escape(gen) + r'(\d+)', art or '')
                if m_: last[gen] = max(last.get(gen, 0), int(m_.group(1)))
        except Exception as exc:
            note_all = f"the last variant of an existing style could not be read from HANA ({str(exc).splitlines()[0][:80]}): numbers start at 001 - check them"
    ws_pr = wb.create_sheet("PRICING")        # sixth tab: one row per variant article; the condition columns are filled in later
    header(ws_pr, ['Article', 'Material', 'COLOR CODE', 'SIZE', 'BUY PRICE', 'MSRP', 'ZVP1', 'ZSLS', 'PR00'])
    ref = {}                                  # (style, color code) -> the buysheet's price and MSRP, for reference
    for style_, rws in os_rows_by_style.items():
        for x in rws: ref.setdefault((style_, x['color_code']), (x.get('price'), x.get('msrp')))
    counter, ri, rp = {}, 2, 2
    for style, colors, sizes, is_new in variant_jobs:
        n = counter.setdefault(style, 0 if is_new else last.get(style, 0))
        for color in colors:
            for sz in sorted(sizes, key=lambda s_: SIZE_ORDER.index(s_) if s_ in SIZE_ORDER else 99):
                n += 1
                write(ws_v, ri, [f"{style}{n:03d}", 'PC', 'H5', style, color, sz,
                                 'new style: starts at 001' if is_new else (f"after the last variant {last[style]:03d}" if style in last else note_all)])
                bp, mp = ref.get((style, color), (None, None))
                write(ws_pr, rp, [f"{style}{n:03d}", style, color, sz, bp, mp, None, None, None])
                ri += 1; rp += 1
        counter[style] = n

    # PRICING also lists articles SAP ALREADY has whose price is blank or differs from the buysheet's (the style existing is no reason to leave them out)
    has_row = {ws_pr.cell(r_, 1).value for r_ in range(2, ws_pr.max_row + 1)}
    needs_price = lambda g_: (not g_.get('sap_price') or str(g_.get('price_check') or '').startswith('✗')
                              or not g_.get('sap_msrp') or str(g_.get('msrp_check') or '').startswith('✗'))     # SELL or MSRP blank in SAP, or different from the buysheet's
    for g in validation.get('generic', []):
        a_ = g.get('sap_article')
        if g.get('status') == '✓ EXISTS' and a_ and a_ not in has_row and needs_price(g):
            write(ws_pr, rp, [a_, g['style'], g['color_code'], g['size'], g.get('buy_price'), g.get('buy_msrp'), None, None, None]); rp += 1; has_row.add(a_)
    for g in validation.get('ppk', []):
        a_ = g.get('ppk_article')
        if g.get('exists') == '✓ EXISTS' and a_ and a_ not in has_row and needs_price(g):
            write(ws_pr, rp, [a_, str(a_).split('-')[0], g.get('color_code'), '', g.get('buy_price'), None, None, None, None]); rp += 1; has_row.add(a_)

    # PC PER CTN rows (filled here, after the variants are numbered): the variants being created (their planned article number) and the
    # variants SAP already has whose carton is missing or differs from the buysheet's (their SAP article).
    planned = {}
    for _r in range(2, ws_v.max_row + 1):
        planned[(ws_v.cell(_r, 4).value, ws_v.cell(_r, 5).value, ws_v.cell(_r, 6).value)] = ws_v.cell(_r, 1).value
    sap_art = {}
    for g in validation.get('generic', []):
        if g['status'] == '✓ EXISTS' and (g.get('pc_ctn_check', '').startswith('✗') or g.get('pc_ctn_check', '').startswith('⚠ NO SAP')):
            m_ = re.search(r'SAP=(\d+)', g['pc_ctn_check'])
            sap_art[(g['style'], g['color_code'], g['size'])] = (g.get('sap_article'), g.get('pc_ctn_note') or g['pc_ctn_check'], int(m_.group(1)) if m_ else None,
                                                                'MISMATCH' if m_ else 'NO SAP VALUE')
    pcs, art_of, note_of, sap_pc_of, status_of = {}, {}, {}, {}, {}
    for style, rws in os_rows_by_style.items():
        for r in rws:
            for sz in r['sizes']:
                key = (style, r['color_code'], sz)
                new_variant = style in missing_generics and (style in new_styles or sz in adds[style].get(r['color_code'], ()))
                if not new_variant and key not in sap_art: continue
                pc = (r.get('pc_by_size') or {}).get(sz) or num(r.get('pc_ctn_os') or r.get('pc_ctn'))
                if pc: pcs.setdefault(key, []).append(int(pc) if float(pc).is_integer() else pc)
                if new_variant: art_of[key], note_of[key], sap_pc_of[key], status_of[key] = planned.get(key, ''), 'to be created', None, 'NEW'
                else: art_of[key], note_of[key], sap_pc_of[key], status_of[key] = sap_art[key]
    order_key = lambda k_: (k_[0], k_[1], SIZE_ORDER.index(k_[2]) if k_[2] in SIZE_ORDER else 99)
    for k, key in enumerate(sorted(pcs, key=order_key), 2):
        vals_ = pcs[key]; same = len(set(vals_)) == 1
        extra = '' if same else f"the buysheet has {', '.join(str(v) for v in sorted(set(vals_)))} for this size: the first is shown"
        write(ws_k, k, [key[0], key[1], key[2], vals_[0], art_of.get(key, '')])

    add_filters(wb)
    generic_pricing_table(ws_pr)                          # after add_filters: the variant table above keeps its own columns and filter
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf


# ── 6b. the buysheet TEMPLATE (.xlsm): fill its DATA and HEADER tables, keep everything else byte for byte ──
BUYSHEET_TEMPLATE = os.environ.get('BUYSHEET_TEMPLATE') or os.path.join(TOOLS_DIR, 'templates', 'Buysheet_Template.xlsm')
XLSM = 'application/vnd.ms-excel.sheet.macroEnabled.12'

def _xml_text(v):
    return (str(v).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;'))

def _excel_serial(d):
    if hasattr(d, 'date'): d = d.date()
    return (d - datetime(1899, 12, 30).date()).days

def centric_style(style):
    """'RDZ272ZS01C' -> 'RDZ272ZS01': the style without its trailing variant letter(s)."""
    return str(style or '')[:10]

def template_cell_value(col, r, i, rows_in_customer, merged=False):
    """What goes in DATA column `col` for prebuy row r (None = leave blank). Everything text is upper case."""
    ppk = r['type'] == 'PPK'
    style = r.get('raw_style') or r.get('style') or ''
    today = datetime.now()
    size_qty = r.get('sizes') or {}
    simple = {
        'CUSTOMER': buysheet_customer(r), 'STYLE': style, 'CENTRIC STYLE': centric_style(r.get('style') or style),
        'SEGMENT': r.get('segment'), 'COLOR CODE': r.get('color_code'), 'COLOR': r.get('color'), 'DESCRIPTION': r.get('desc'),
        'CRD': r['split_crd'] if 'split_crd' in r else r.get('delivery'), 'DELIVERY': None if 'split_crd' in r else r.get('delivery'),
        'PRICE': num(r.get('price')), 'MSRP': num(r.get('msrp')),
        'DELIVERY SALES': None if 'split_crd' in r else r.get('delivery'),    # merged template: the prebuy's date is what sales wants; DELIVERY PROD is production's to fill
        'SIZE RATIO': r.get('ratio') if ppk else None, 'SIZES': r.get('sizes_str') if ppk else None,
        'PC PER PAC': r.get('pc_per_pac') if ppk else None, 'PAC PER CTN': (r.get('pac_per_ctn') or 1) if ppk else None,
        'PPK QTY': r.get('ppk_qty') if ppk else None, 'PC PER CTN OS': None if ppk else r.get('pc_ctn_os'),
        'DATE UPDATED': today, 'COMMENTS': r.get('notes'), 'ROW_LINE': i + 1,
    }
    if 'split_crd' in r and col in ('CRD SALES', 'CRD PROD'):                  # a split row (merged template): the flow's CRD on both CRD columns
        return r['split_crd']
    if merged and col in ('PRICE (CAD)', 'PRICE (USD)'):                      # a generic ending in C is Canadian: its price is CAD, the rest USD
        canadian = (r.get('sap_currency') == 'CAD') if r.get('sap_currency') else is_canadian_style(r.get('raw_style') or r.get('style') or style)
        return num(r.get('price')) if canadian == (col == 'PRICE (CAD)') else None
    if col in simple:
        v = simple[col]
        return up(v) if isinstance(v, str) else v
    if col in KNOWN_SIZES or col in size_qty:
        return size_qty.get(col) if not ppk else None
    return None

def make_buysheet_from_template(rows, customers, template_bytes):
    """A copy of the buysheet template with DATA and HEADER filled from the prebuy rows. Only the two sheet parts,
    their two table parts and the calc settings are touched: VBA, Power Query, pivots and slicers stay as they are."""
    zin = zipfile.ZipFile(io.BytesIO(template_bytes))
    parts = {n: zin.read(n) for n in zin.namelist()}
    wbx = parts['xl/workbook.xml'].decode('utf-8')
    rels = parts['xl/_rels/workbook.xml.rels'].decode('utf-8')
    rid_target = {m.group(2): m.group(1) for m in re.finditer(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)}
    rid_target.update({m.group(1): m.group(2) for m in re.finditer(r'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', rels)})
    def sheet_part(name):
        m = re.search(rf'<sheet [^>]*name="{re.escape(name)}"[^>]*r:id="([^"]+)"', wbx) or re.search(rf'<sheet [^>]*r:id="([^"]+)"[^>]*name="{re.escape(name)}"', wbx)
        if not m: raise ValueError(f'the buysheet template has no sheet named {name}')
        return 'xl/' + rid_target[m.group(1)].lstrip('/').replace('xl/', '', 1) if not rid_target[m.group(1)].startswith('xl/') else rid_target[m.group(1)]
    def table_part(sheet):
        srel = parts.get(sheet.replace('worksheets/', 'worksheets/_rels/') + '.rels', b'').decode('utf-8')
        m = re.search(r'Target="\.\./tables/([^"]+)"', srel)
        if not m: raise ValueError(f'{sheet} has no table')
        return 'xl/tables/' + m.group(1)

    # ── DATA ──
    dsheet = sheet_part('DATA'); dtable = table_part(dsheet)
    sx, tx = parts[dsheet].decode('utf-8'), parts[dtable].decode('utf-8')
    tref = re.search(r' ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', tx)
    first_col, head_row, last_col = tref.group(1), int(tref.group(2)), tref.group(3)
    cols = re.findall(r'<tableColumn [^>]*name="([^"]+)"', tx)
    cols = [c.replace('&amp;', '&') for c in cols]
    # size columns the prebuy has that the table does not: insert after the table's last size column
    want_sizes = []
    for r in rows:
        if r['type'] == 'OS':
            for sz in (r.get('size_order') or list(r.get('sizes') or {})):
                if sz in (r.get('sizes') or {}) and sz not in cols and sz not in want_sizes: want_sizes.append(sz)
    if want_sizes:
        last_size_i = max(i for i, c in enumerate(cols) if c in KNOWN_SIZES)
        cols = cols[:last_size_i + 1] + want_sizes + cols[last_size_i + 1:]
    ncols = len(cols)
    rows_xml = re.findall(r'<row r="(\d+)"[^>]*>.*?</row>|<row r="(\d+)"[^>]*/>', sx, re.S)
    def row_xml(n): 
        m = re.search(rf'<row r="{n}"[^>]*?(?:/>|>.*?</row>)', sx, re.S); return m.group(0) if m else ''
    head_xml, ex_xml = row_xml(head_row), row_xml(head_row + 1)
    style_of = lambda xml, letter: (re.search(rf'<c r="{letter}\d+"[^>]*?\ss="(\d+)"', xml) or [None, None])[1] if xml else None
    head_style = style_of(head_xml, first_col) or '1'
    orig_cols = [c.replace('&amp;', '&') for c in re.findall(r'<tableColumn [^>]*name="([^"]+)"', tx)]
    orig_letter = {c: get_column_letter(i + 1) for i, c in enumerate(orig_cols)}
    ex_style = {c: style_of(ex_xml, orig_letter[c]) for c in orig_cols}
    ex_formula = {}
    for c in orig_cols:
        m = re.search(rf'<c r="{orig_letter[c]}{head_row + 1}"[^>]*>(?:<f[^>]*>(.*?)</f>)', ex_xml, re.S)
        if m: ex_formula[c] = m.group(1)
    size_style = next((ex_style[c] for c in orig_cols if c in KNOWN_SIZES and ex_style.get(c)), None)

    def cell(letter, rown, value, style, formula=None):
        s_attr = f' s="{style}"' if style else ''
        if formula is not None:
            return f'<c r="{letter}{rown}"{s_attr}><f>{formula}</f></c>'
        if value is None or value == '':
            return f'<c r="{letter}{rown}"{s_attr}/>'
        if hasattr(value, 'year'):
            return f'<c r="{letter}{rown}"{s_attr}><v>{_excel_serial(value)}</v></c>'
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            v = int(value) if float(value).is_integer() else value
            return f'<c r="{letter}{rown}"{s_attr}><v>{v}</v></c>'
        return f'<c r="{letter}{rown}"{s_attr} t="inlineStr"><is><t>{_xml_text(value)}</t></is></c>'

    new_rows = ['<row r="%d" spans="1:%d">%s</row>' % (head_row, ncols, ''.join(cell(get_column_letter(i + 1), head_row, c, head_style) for i, c in enumerate(cols)))]
    by_customer = defaultdict(int)
    for i, r in enumerate(rows):
        rown = head_row + 1 + i
        cells = []
        for ci, c in enumerate(cols):
            letter = get_column_letter(ci + 1)
            if c in ex_formula:
                cells.append(cell(letter, rown, None, ex_style.get(c), ex_formula[c]))
            else:
                cells.append(cell(letter, rown, template_cell_value(c, r, i, by_customer, merged='PRICE (CAD)' in cols), ex_style.get(c) if c in ex_style else size_style))
        new_rows.append('<row r="%d" spans="1:%d">%s</row>' % (rown, ncols, ''.join(cells)))
    last_row = head_row + max(len(rows), 1)
    last_letter = get_column_letter(ncols)
    # brand + season in A1 / B1
    label = next((r.get('season_label') for r in rows if r.get('season_label')), '')
    brand = next((r.get('brand') for r in rows if r.get('brand')), '')
    sx = re.sub(r'<c r="A1"[^>]*?(?:/>|>.*?</c>)', cell('A', 1, up(brand), style_of(row_xml(1), 'A')), sx, count=1, flags=re.S)
    sx = re.sub(r'<c r="B1"[^>]*?(?:/>|>.*?</c>)', cell('B', 1, label.replace('-', ' ') if label else '', style_of(row_xml(1), 'B')), sx, count=1, flags=re.S)
    # replace every row from the header row down
    keep = re.search(r'<sheetData>(.*?)</sheetData>', sx, re.S).group(1)
    kept = ''.join(m.group(0) for m in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', keep, re.S) if int(m.group(1)) < head_row)
    sx = re.sub(r'<sheetData>.*?</sheetData>', lambda m: '<sheetData>' + kept + ''.join(new_rows) + '</sheetData>', sx, count=1, flags=re.S)
    sx = re.sub(r'<dimension ref="[^"]+"/>', f'<dimension ref="A1:{last_letter}{last_row}"/>', sx, count=1)
    sx = re.sub(r'<selection [^>]*/>', '<selection activeCell="A4" sqref="A4"/>', sx, count=1)
    new_ref = f'{first_col}{head_row}:{last_letter}{last_row}'
    tx = re.sub(r' ref="[A-Z]+\d+:[A-Z]+\d+"', f' ref="{new_ref}"', tx)
    if want_sizes:
        max_id = max(int(x) for x in re.findall(r'<tableColumn [^>]*id="(\d+)"', tx))
        col_xml = {m.group(1).replace('&amp;', '&'): m.group(0) for m in re.finditer(r'<tableColumn [^>]*name="([^"]+)"[^>]*?(?:/>|>.*?</tableColumn>)', tx, re.S)}
        parts_cols = []
        for c in cols:
            if c in col_xml: parts_cols.append(col_xml[c])
            else:
                max_id += 1; parts_cols.append(f'<tableColumn id="{max_id}" name="{_xml_text(c)}"/>')
        tx = re.sub(r'<tableColumns [^>]*>.*?</tableColumns>', lambda m: f'<tableColumns count="{ncols}">' + ''.join(parts_cols) + '</tableColumns>', tx, count=1, flags=re.S)
    parts[dsheet], parts[dtable] = sx.encode('utf-8'), tx.encode('utf-8')

    # ── HEADER ──
    hsheet = sheet_part('HEADER'); htable = table_part(hsheet)
    hx, htx = parts[hsheet].decode('utf-8'), parts[htable].decode('utf-8')
    href = re.search(r' ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', htx)
    hhead = int(href.group(2))
    hcols = [c.replace('&amp;', '&') for c in re.findall(r'<tableColumn [^>]*name="([^"]+)"', htx)]
    def hrow(n):
        m = re.search(rf'<row r="{n}"[^>]*?(?:/>|>.*?</row>)', hx, re.S); return m.group(0) if m else ''
    hex_xml = hrow(hhead + 1)
    hstyle = {c: style_of(hex_xml, get_column_letter(i + 1)) for i, c in enumerate(hcols)}
    seen = {}
    for r in rows:
        cust = buysheet_customer(r)
        if cust not in seen:
            seen[cust] = {'CUSTOMER': cust, 'TICKETS': up(r.get('tickets') or ''), 'HANG_FOLD': up(r.get('hang_fold') or ''), 'PER BAG': None,
                          'WHSE': up(r.get('warehouse') or ''), 'REQ#': None, 'TERMS': r.get('terms'), 'TERMS DESC': up(r.get('terms_desc') or '')}
    hrows = []
    for i, vals in enumerate(seen.values()):
        rown = hhead + 1 + i
        hrows.append('<row r="%d" spans="1:%d">%s</row>' % (rown, len(hcols), ''.join(cell(get_column_letter(ci + 1), rown, vals.get(c), hstyle.get(c)) for ci, c in enumerate(hcols))))
    hlast = hhead + max(len(seen), 1)
    hkeep = ''.join(m.group(0) for m in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', re.search(r'<sheetData>(.*?)</sheetData>', hx, re.S).group(1), re.S) if int(m.group(1)) <= hhead)
    hx = re.sub(r'<sheetData>.*?</sheetData>', lambda m: '<sheetData>' + hkeep + ''.join(hrows) + '</sheetData>', hx, count=1, flags=re.S)
    hx = re.sub(r'<dimension ref="[^"]+"/>', f'<dimension ref="A1:{get_column_letter(len(hcols))}{hlast}"/>', hx, count=1)
    htx = re.sub(r' ref="[A-Z]+\d+:[A-Z]+\d+"', f' ref="A{hhead}:{get_column_letter(len(hcols))}{hlast}"', htx)
    parts[hsheet], parts[htable] = hx.encode('utf-8'), htx.encode('utf-8')

    # ── recalc on open; drop the stale calc chain ──
    wbx = re.sub(r'<calcPr([^>]*)/>', lambda m: '<calcPr' + re.sub(r' fullCalcOnLoad="\d"', '', m.group(1)) + ' fullCalcOnLoad="1"/>', wbx, count=1)
    parts['xl/workbook.xml'] = wbx.encode('utf-8')
    if 'xl/calcChain.xml' in parts:
        del parts['xl/calcChain.xml']
        parts['xl/_rels/workbook.xml.rels'] = re.sub(r'<Relationship [^>]*calcChain\.xml"[^>]*/>', '', rels).encode('utf-8')
        parts['[Content_Types].xml'] = re.sub(r'<Override [^>]*calcChain\.xml"[^>]*/>', '', parts['[Content_Types].xml'].decode('utf-8')).encode('utf-8')

    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename in parts:
                zout.writestr(info.filename, parts[info.filename], compress_type=zipfile.ZIP_DEFLATED)
    out.seek(0)
    return out

def make_buysheet(rows, customers):
    """The buysheet download: the .xlsm template filled when it is installed, else the plain DATA + HEADER OUT workbook."""
    if os.path.exists(BUYSHEET_TEMPLATE):
        with open(BUYSHEET_TEMPLATE, 'rb') as f:
            return make_buysheet_from_template(rows, customers, f.read()), 'xlsm'
    return make_data_sheet(rows, customers), 'xlsx'
