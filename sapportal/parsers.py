"""Every file becomes a list of row dicts (type OS or PPK) plus the customers in it: buysheets, legacy prebuys, the prebuy template."""

import io
import math
import re
from openpyxl import load_workbook
from .config import KNOWN_SIZES, SIZE_ORDER
from .hana import is_canadian_style, norm_code
from .helpers import base_style, canonical_size, dash_list, is_ppk, num, pad10, parse_season


# ═══════════════════════════════════════════════════════════════════════════════
# 4. PARSERS - every file becomes a list of row dicts (type OS or PPK) + the customers in it
# ═══════════════════════════════════════════════════════════════════════════════
def detect_file_type(wb):
    """A buysheet has a DATA sheet; anything else is a prebuy (legacy layout or the new template)."""
    return 'buysheet' if 'DATA' in wb.sheetnames else 'prebuy'

# ── 4a. buysheet (DATA sheet) ─────────────────────────────────────────────────
def read_storage_locations(wb):
    """customer -> SAP storage location, from the buysheet's HEADER sheet (HEADER, HEADER & FACTORY, ...).
    Columns are found by their heading, so a different column order does not matter."""
    out = {}
    for name in wb.sheetnames:
        if not name.strip().upper().startswith('HEADER'):
            continue
        ws, hdr_row = wb[name], None
        for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
            cells = [str(v).strip().upper() if v is not None else '' for v in row]
            is_loc = lambda c: 'STORAGE' in c or c in ('WAREHOUSE', 'WHSE')          # the buysheet template's own heading is WHSE
            if 'CUSTOMER' in cells and any(is_loc(c) for c in cells):
                hdr_row, cust_i = ri, cells.index('CUSTOMER')
                loc_i = next(i for i, c in enumerate(cells) if is_loc(c))
                break
        if hdr_row is None:
            continue
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            if len(row) > max(cust_i, loc_i) and row[cust_i] and row[loc_i] and str(row[loc_i]).strip():
                out.setdefault(str(row[cust_i]).strip().upper(), str(row[loc_i]).strip().upper())
    return out

def read_header_table(wb):
    """customer -> requisition number ('' if the customer's line has none yet), from the REQ# column of the HEADER sheet's table
    (one line per customer).  A line with a customer and a real 10-digit number means the requisition is created."""
    out = {}
    for name in wb.sheetnames:
        if not name.strip().upper().startswith('HEADER'):
            continue
        ws, hdr_row = wb[name], None
        for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
            cells = [str(v).strip().upper() if v is not None else '' for v in row]
            if 'CUSTOMER' in cells and any(c.startswith('REQ') for c in cells):
                hdr_row, cust_i = ri, cells.index('CUSTOMER')
                req_i = next(i for i, c in enumerate(cells) if c.startswith('REQ'))
                break
        if hdr_row is None:
            continue
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            if len(row) > max(cust_i, req_i) and row[cust_i]:
                req = pad10(row[req_i]) if row[req_i] else ''
                cust = str(row[cust_i]).strip().upper()
                if req.strip('0'):
                    out[cust] = out[cust] if out.get(cust) else req
                else:
                    out.setdefault(cust, '')
    return out

def read_header_reqs(wb):
    """customer -> requisition number (10 digits), only the customers that have one."""
    return {c: r for c, r in read_header_table(wb).items() if r}

def read_season_label(all_rows):
    """'Q2 2027' from a label such as Q2-2027 near the top of the DATA sheet ('' if there is none)."""
    for r in all_rows[:3]:
        for v in r:
            label = parse_season(v)
            if label:
                return label
    return ''

def parse_buysheet(wb):
    """Parse buysheet DATA sheet — same logic as validator."""
    ws = wb['DATA'] if 'DATA' in wb.sheetnames else wb.active
    all_rows = list(ws.iter_rows(values_only=True))

    header_idx = None
    for i, r in enumerate(all_rows):
        if r[0] and str(r[0]).upper().strip() == 'CUSTOMER':
            header_idx = i; break
    if header_idx is None:
        header_idx = next((i for i, r in enumerate(all_rows) if any(v for v in r)), 0)

    header    = list(all_rows[header_idx])
    data_rows = [r for r in all_rows[header_idx+1:] if any(v for v in r)]
    header_upper = [str(h).upper().strip().replace('\n',' ').replace('\r',' ') if h else '' for h in header]

    def find_col(name):
        for i, h in enumerate(header_upper):
            if name in h: return i
        return None

    idx = {
        'style':        find_col('STYLE'),
        'color_code':   find_col('COLOR CODE'),
        'color':        next((i for i,h in enumerate(header_upper) if h == 'COLOR'), None),
        'desc':         find_col('DESCRIPTION'),
        'pc_ctn':       find_col('PC PER CTN'),
        'ratio':        find_col('SIZE RATIO'),
        'sizes_col':    find_col('SIZES'),
        'pc_per_pac':   find_col('PC PER PAC'),
        'pac_per_ctn':  find_col('PAC PER CTN'),
        'ppk_qty':      find_col('PPK QTY'),
        'pc_ctn_os':    find_col('PC PER CTN OS'),
        'price':        find_col('STORE COST (CAD)') or find_col('PRICE'),                 # older templates
        'price_cad':    next((i for i, h in enumerate(header_upper) if h == 'PRICE (CAD)'), None),   # the merged template: CAD for a Canadian generic (ends in C),
        'price_usd':    next((i for i, h in enumerate(header_upper) if h == 'PRICE (USD)'), None),   # USD for the others
        'fineline':     next((i for i, h in enumerate(header_upper) if h == 'FINELINE'), None),
        'msrp':         find_col('MSRP'),
        # the date a requisition uses: DELIVERY PROD on the merged template (what production can do; DELIVERY SALES is only what sales
        # would like), else DELIVERY on older ones. A blank stays blank: another date is never substituted.
        'delivery':     next((i for i, h in enumerate(header_upper) if h == 'DELIVERY PROD'), None) if 'DELIVERY PROD' in header_upper else find_col('DELIVERY'),
        'segment':      find_col('SEGMENT'),
        'customer':     find_col('CUSTOMER'),
        'warehouse':    next((i for i, h in enumerate(header_upper) if h in ('WAREHOUSE', 'WHSE', 'STORAGE LOCATION', 'STORAGE LOC')), None),   # exact heading only: 'IN WHSE DATE' is a date
    }
    size_indices = [(i, str(header[i]).strip().upper())
                    for i, h in enumerate(header_upper) if h in KNOWN_SIZES]
    # a LONG buysheet: one row per size, with a SIZE column and a QTY / QUANTITY column instead of one column per size
    size_col = next((i for i, h in enumerate(header_upper) if h == 'SIZE'), None)
    qty_col  = next((i for i, h in enumerate(header_upper) if h in ('QTY', 'QUANTITY')), None)
    long_fmt = size_col is not None and qty_col is not None and not size_indices

    storage_by_customer = read_storage_locations(wb)
    req_by_customer = read_header_reqs(wb)
    season_label = read_season_label(all_rows)

    def gv_cust(r):
        i = idx.get('customer')
        return r[i] if i is not None and i < len(r) else None

    rows = []
    for r in data_rows:
        si = idx['style']
        if si is None or not r[si]: continue
        raw_style = str(r[si]).strip()
        if not raw_style: continue

        cust_kind = re.search(r'\b(OS|PPK|COMPLEX)\b', str(gv_cust(r) or '').upper())
        cust_kind = cust_kind.group(1) if cust_kind else ''
        ppk = is_ppk(raw_style) or cust_kind in ('PPK', 'COMPLEX')      # a complex prepack's article is the bare style: the customer name says COMPLEX
        sizes = {}
        long_ppk_qty = None
        pc_by_size = {}
        for si2, sz_label in size_indices:
            if si2 < len(r) and r[si2]:
                try:
                    qty = int(float(r[si2]))
                    if qty > 0: sizes[sz_label] = qty
                except: pass
        if long_fmt:
            lab = str(r[size_col] or '').strip().upper() if size_col < len(r) else ''
            try: qty = int(float(r[qty_col])) if qty_col < len(r) and r[qty_col] not in (None, '') else 0
            except: qty = 0
            if lab == 'PPK QTY': long_ppk_qty = qty or None           # a prepack line carries its pack count, not a size
            elif lab and qty > 0:
                sizes[lab] = qty
                pc_i = idx.get('pc_ctn_os') if idx.get('pc_ctn_os') is not None else idx.get('pc_ctn')
                try: size_pc = int(float(r[pc_i])) if pc_i is not None and pc_i < len(r) and r[pc_i] not in (None, '') else None
                except: size_pc = None
                if size_pc: pc_by_size[lab] = size_pc

        def gv(key):
            i = idx.get(key)
            return r[i] if i is not None and i < len(r) else None

        entry = {
            'type':        'PPK' if ppk else 'OS',
            'pack_kind':   'COMPLEX' if cust_kind == 'COMPLEX' else ('PPK' if ppk else 'OS'),
            'complex':     cust_kind == 'COMPLEX',
            'raw_style':   raw_style,
            'style':       base_style(raw_style) if ppk else raw_style,
            'color_code':  str(gv('color_code') or '').strip(),
            'color':       str(gv('color') or '').strip().upper(),
            'desc':        str(gv('desc') or '').strip(),
            'price':       (gv('price_cad') if (idx.get('price_cad') is not None and is_canadian_style(raw_style)) else gv('price_usd') if idx.get('price_cad') is not None else gv('price')),
            'price_cad':   gv('price_cad') if idx.get('price_cad') is not None else None,     # both kept: SAP's price currency picks the right one at validation
            'price_usd':   gv('price_usd') if idx.get('price_cad') is not None else None,
            'fineline':    norm_code(gv('fineline')),
            'msrp':        gv('msrp'),
            'delivery':    gv('delivery'),
            'segment':     str(gv('segment') or '').strip(),
            'customer':    str(gv('customer') or '').strip(),
            'warehouse':   str(gv('warehouse') or '').strip() or storage_by_customer.get(str(gv('customer') or '').strip().upper(), ''),
            'season_label': season_label,
            'req':         req_by_customer.get(str(gv('customer') or '').strip().upper(), ''),     # the requisition on the HEADER sheet for this customer
            'pc_ctn':      gv('pc_ctn'),
            'pc_ctn_os':   gv('pc_ctn_os'),
            'ratio':       str(gv('ratio') or '').strip(),
            'row_pc':      sum(int(x) for x in dash_list(gv('ratio')) if x.isdigit()),      # this color's pieces per pack (complex prepacks)
            'sizes_str':   str(gv('sizes_col') or '').strip(),
            'pc_per_pac':  gv('pc_per_pac'),
            'pac_per_ctn': gv('pac_per_ctn'),
            'ppk_qty':     long_ppk_qty if long_ppk_qty is not None else gv('ppk_qty'),
            'sizes':       sizes,
            'size_order':  [lab for _, lab in size_indices],
            'pack_kind':   'COMPLEX' if re.search(r'\bCOMPLEX\b', str(gv('customer') or '').upper()) else ('PPK' if ppk else 'OS'),
        }
        if long_fmt: entry['pc_by_size'] = pc_by_size      # long layout only: the carton of each size (they differ)
        rows.append(entry)

    if long_fmt:                          # fold the size rows of one order back into one row with its sizes, as the wide layout gives
        merged, seen = [], {}
        for e in rows:
            key = (e['customer'], e['raw_style'], e['color_code'], e['color'], str(e['delivery']), e['desc'], str(e['price']), str(e['msrp']), e['type'])
            if key in seen:
                m = seen[key]
                for z, q in e['sizes'].items(): m['sizes'][z] = m['sizes'].get(z, 0) + q
                for z, c in e['pc_by_size'].items(): m['pc_by_size'].setdefault(z, c)
                if e['ppk_qty'] and not m['ppk_qty']: m['ppk_qty'] = e['ppk_qty']
            else:
                seen[key] = e; merged.append(e)
        for e in merged:
            e['size_order'] = sorted(e['sizes'], key=lambda z: SIZE_ORDER.index(z) if z in SIZE_ORDER else 99)
        rows = merged
    customers = list(set(r['customer'] for r in rows if r['customer']))
    return rows, customers

# ── 4b. legacy prebuys: OPEN STOCK / PREPACK tabs with a header block above the grid ──
def get_header_info(ws):
    """Header block of a legacy prebuy: labels sit in one row and their values directly BELOW them
    (D1 'SELECT WAREHOUSE:' over D2 'SJCA'). A value to the right is accepted only when it is not itself a label."""
    info = {'customer': None, 'warehouse': None, 'segment': None, 'hang_fold': None, 'tickets': None, 'brand': None, 'terms': None, 'terms_desc': None}
    keys = {'CUSTOMER:': 'customer', 'CUSTOMER': 'customer', 'BRAND:': 'brand',
            'SELECT WAREHOUSE:': 'warehouse', 'WAREHOUSE:': 'warehouse', 'WAREHOUSE': 'warehouse',
            'SEGMENT:': 'segment', 'SEGMENT': 'segment', 'STOCK SEGMENT:': 'segment',
            'HANG/FOLD:': 'hang_fold', 'HANG/FOLD': 'hang_fold', 'PRICE TICKETS:': 'tickets', 'PRICE TICKETS': 'tickets',
            'TERMS/ALLOWANCES': 'terms', 'DESCRIPTION OF TERMS': 'terms_desc'}
    other_labels = {'SAMPLES NEEDED:', 'PACKS PER CTN'}
    grid = [list(r) for r in ws.iter_rows(min_row=1, max_row=6, values_only=True)]

    def is_label(v):
        t = str(v).strip().upper()
        return t.endswith(':') or t in keys or t in other_labels

    def value_for(r, i):
        below = grid[r + 1][i] if r + 1 < len(grid) and i < len(grid[r + 1]) else None
        right = grid[r][i + 1] if i + 1 < len(grid[r]) else None
        for cand in (below, right):
            if cand not in (None, '') and not is_label(cand):
                return str(cand).strip()
        return None

    for r, row in enumerate(grid):
        for i, v in enumerate(row):
            key = keys.get(str(v).strip().upper()) if v else None
            if key and info[key] is None:
                info[key] = value_for(r, i)
    return info

def find_header_row(ws, required_cols):
    for ri, row in enumerate(ws.iter_rows(values_only=True), 1):
        cells = [str(v).strip().upper() if v else '' for v in row]
        if all(c in cells for c in required_cols):
            return ri, cells
    return None, None

def has_ppk_column(ws):
    for row in ws.iter_rows(min_row=1, max_row=6, values_only=True):
        cells = [str(v).strip().upper() if v else '' for v in row]
        if 'PPK' in cells or 'PPK QTY' in cells:
            return True
    return False

def parse_os_tab(ws):
    info      = get_header_info(ws)
    header_ri, header_row = find_header_row(ws, ['STYLE'])
    if not header_row: return info['customer'], []

    def ci(name): return header_row.index(name) if name in header_row else None
    idx = {
        'style':   ci('STYLE'), 'code': ci('COLOR CODE'),
        'color':   ci('COLOR DESCRIPTION'), 'desc': ci('DESCRIPTION'),
        'delivery':ci('DELIVERY'), 'price': ci('PRICE'), 'msrp': ci('MSRP'),
        'segment': ci('SEGMENT'), 'pc_ctn': ci('PC PER CTN OS'),
    }
    size_idx = {canonical_size(h): ci_n for ci_n, h in enumerate(header_row)
                if canonical_size(h) in KNOWN_SIZES}

    rows = []
    for row in ws.iter_rows(min_row=header_ri+1, values_only=True):
        si = idx['style']
        if si is None or not row[si]: continue
        style = str(row[si]).strip()
        if not re.match(r'^[A-Z]{3}\d', style.upper()): continue
        sizes = {}
        for sz, ci_n in size_idx.items():
            v = row[ci_n]
            if v and str(v).strip() not in ('','0','None'):
                try: sizes[sz] = int(float(v))
                except: pass
        if not sizes: continue
        row_seg = str(row[idx['segment']]).strip() if idx['segment'] is not None and row[idx['segment']] else info.get('segment','')
        rows.append({
            'type':'OS', 'raw_style':style, 'style':style,
            'customer': info['customer'], 'warehouse': info.get('warehouse',''),
            'color_code': str(row[idx['code']]).strip()          if idx['code']  is not None and row[idx['code']]  else '',
            'color':      str(row[idx['color']]).strip().upper()  if idx['color'] is not None and row[idx['color']] else '',
            'desc':       str(row[idx['desc']]).strip()           if idx['desc']  is not None and row[idx['desc']]  else '',
            'delivery':   row[idx['delivery']]                    if idx['delivery'] is not None                    else None,
            'price':      row[idx['price']]                       if idx['price'] is not None                       else None,
            'msrp':       row[idx['msrp']]                        if idx['msrp'] is not None                        else None,
            'segment':    row_seg,
            'pc_ctn_os':  row[idx['pc_ctn']]                      if idx['pc_ctn'] is not None and row[idx['pc_ctn']] else None,
            'sizes':      sizes,
            'size_order': list(size_idx), 'pack_kind': 'OS',
            'ratio':'', 'sizes_str':'', 'pc_per_pac':None, 'pac_per_ctn':None, 'ppk_qty':None,
        })
    return info['customer'], rows

def is_complex_header(style):
    """SFM273CR01-A, SFM271SS05-A1, SFW273QZ08-ROLLOUT, SFW273QZ08-PACKXSS: a style with a letter suffix (a name, not a color code) and no
    color code is a complex prepack's own row."""
    t = str(style or '').strip().upper()
    return bool(re.match(r'^[A-Z]{3}\d', t) and re.search(r'-[A-Z]+\d*$', t))

def parse_ppk_tab(ws):
    info = get_header_info(ws)
    header_ri, header_row = find_header_row(ws, ['STYLE','COLOR CODE'])
    if not header_row: return info['customer'], []

    def ci(name): return header_row.index(name) if name in header_row else None
    idx = {'style':ci('STYLE'),'code':ci('COLOR CODE'),'color':ci('COLOR DESCRIPTION'),'price':ci('PRICE'),'msrp':ci('MSRP')}
    desc_ci  = ci('DESCRIPTION') or ci('VENDOR DESCRIPTION')
    size_idx = {canonical_size(h): ci_n for ci_n, h in enumerate(header_row) if canonical_size(h) in KNOWN_SIZES}

    rows = []
    current_style = None; current_price = None; current_msrp = None
    cx_base = None                                   # base style of a complex block being skipped: parse_complex_tab reads those rows
    for row in ws.iter_rows(min_row=header_ri+1, values_only=True):
        if not any(v for v in row): continue
        style_val = row[idx['style']] if idx['style'] is not None else None
        code_here = row[idx['code']] if idx['code'] is not None else None
        if style_val and not code_here and is_complex_header(style_val):
            cx_base = base_style(str(style_val).strip()).upper(); continue
        if cx_base:
            if code_here and (not style_val or base_style(str(style_val).strip()).upper() == cx_base): continue
            cx_base = None
        if style_val:
            sv = str(style_val).strip()
            if re.match(r'^[A-Z]{3}\d', sv.upper()):
                current_style = sv
                if idx['price'] is not None and row[idx['price']]: current_price = row[idx['price']]
                if idx['msrp']  is not None and row[idx['msrp']]:  current_msrp  = row[idx['msrp']]
        code_val = row[idx['code']] if idx['code'] is not None else None
        if code_val and current_style:
            sizes = {}
            for sz, ci_n in size_idx.items():
                v = row[ci_n]
                if v and str(v).strip() not in ('','0','None'):
                    try: sizes[sz] = int(float(v))
                    except: pass
            if not sizes: continue
            vals = list(sizes.values())
            g = math.gcd(*vals) if len(vals) > 1 else vals[0]
            ratio_vals = [v//g for v in vals]
            raw_style  = f"{current_style}-{str(code_val).strip()}"
            rows.append({
                'type':'PPK', 'raw_style':raw_style, 'style':current_style,
                'customer': info['customer'], 'warehouse': info.get('warehouse',''),
                'color_code': str(code_val).strip(),
                'color':      str(row[idx['color']]).strip().upper() if idx['color'] is not None and row[idx['color']] else '',
                'desc':       str(row[desc_ci]).strip() if desc_ci is not None and row[desc_ci] else '',
                'price':      current_price,
                'msrp':       current_msrp,
                'segment':    info.get('segment',''),
                'delivery':   None,
                'sizes':      sizes,
                'sizes_str':  '-'.join(sizes.keys()),
                'ratio':      '-'.join(str(r) for r in ratio_vals),
                'pc_per_pac': sum(ratio_vals),
                'ppk_qty':    g,
                'pac_per_ctn':None,
                'pc_ctn_os':  None,
                'size_order': list(size_idx), 'pack_kind': 'COMPLEX' if 'COMPLEX' in ws.title.upper() else 'PPK',
            })
    return info['customer'], rows

RATIO_HDR_ALIASES = {'SM': 'S', 'MD': 'M', 'MED': 'M', 'LG': 'L', 'LRG': 'L', 'XS2': 'XS'}      # the prebuy form heads its XS ratio column XS2

def parse_complex_tab(ws, require_suffix=False):
    """COMPLEX tab of a legacy prebuy. A row with a STYLE but no COLOR CODE (SFM273CR01-A) is the whole prepack: its
    PREPACK RATIO is the pack. The colored rows under it are its components, until the next such row. Ratios are read
    from the PREPACK RATIO columns; units are only a cross-check (a mismatch warns, it never blocks)."""
    info = get_header_info(ws)
    header_ri, header_row = find_header_row(ws, ['STYLE', 'COLOR CODE'])
    if not header_row: return info['customer'], []
    def ci(name): return header_row.index(name) if name in header_row else None
    idx = {'style': ci('STYLE'), 'code': ci('COLOR CODE'), 'color': ci('COLOR DESCRIPTION'), 'price': ci('PRICE'), 'msrp': ci('MSRP'), 'units': ci('UNITS')}
    desc_ci = ci('DESCRIPTION') or ci('VENDOR DESCRIPTION')

    # where the PREPACK RATIO band starts (merged band above the header); else the first size that repeats
    ratio_start = None
    for r in ws.iter_rows(min_row=max(1, header_ri - 3), max_row=header_ri, values_only=True):
        for i, v in enumerate(r):
            if v and 'PREPACK RATIO' in str(v).upper():
                ratio_start = i
    ratio_end = None                  # a later TOTAL / NET band (the form's total-pieces column is headed 3XL under it) ends the ratio area
    if ratio_start is not None:
        for r in ws.iter_rows(min_row=max(1, header_ri - 3), max_row=header_ri, values_only=True):
            for i, v in enumerate(r):
                if i > ratio_start and v and str(v).strip().upper() in ('TOTAL', 'NET') and (ratio_end is None or i < ratio_end):
                    ratio_end = i
    unit_cols, ratio_cols = {}, {}
    for i, h in enumerate(header_row):
        sz = RATIO_HDR_ALIASES.get(h, canonical_size(h)) if h else ''
        if sz not in KNOWN_SIZES: continue
        if ratio_end is not None and i >= ratio_end: continue
        in_ratio = (ratio_start is not None and i >= ratio_start) or (ratio_start is None and sz in unit_cols)
        (ratio_cols if in_ratio else unit_cols)[sz] = i

    def read(row, cols):
        out = {}
        for sz, i in cols.items():
            v = row[i] if i < len(row) else None
            n = num(v)
            if n and n > 0: out[sz] = int(round(n))
        return out

    rows, head, comps, n_pack = [], None, [], 0
    def close():
        nonlocal head, comps
        if head and comps:
            whole = sum(head['ratio'].values()) or sum(sum(c['_ratio'].values()) for c in comps)
            for c in comps:
                warn = []
                for sz, q in head['ratio'].items():
                    if sum(x['_ratio'].get(sz, 0) for x in comps) != q:
                        warn.append(f"{sz}: colors sum to {sum(x['_ratio'].get(sz, 0) for x in comps)}, pack ratio says {q}")
                packs = {x['ppk_qty'] / sum(x['_ratio'].values()) for x in comps if x['_ratio'] and x['ppk_qty']}
                if len(packs) > 1: warn.append('colors imply different pack counts: ' + ', '.join(f'{p:g}' for p in sorted(packs)))
                c.update(pc_per_pac=whole, qty_conflict=bool(warn), notes='; '.join(warn[:3]))
            units = head.get('units')
            if units and whole and not any(c['ppk_qty'] for c in comps):      # units only on the pack's own row: packs = units / pieces per pack, the same for every color
                for c in comps: c['ppk_qty'] = round(units / whole * c['row_pc'])
                if units % whole:
                    for c in comps: c['qty_conflict'] = True; c['notes'] = (c.get('notes') + '; ' if c.get('notes') else '') + f"{units:g} units is not a whole number of {whole}-piece packs"
        rows.extend(comps); head, comps = None, []

    for row in ws.iter_rows(min_row=header_ri + 1, values_only=True):
        if not any(v for v in row): continue
        style = str(row[idx['style']] or '').strip() if idx['style'] is not None else ''
        code = str(row[idx['code']] or '').strip() if idx['code'] is not None else ''
        if style and not code and re.match(r'^[A-Z]{3}\d', style.upper()):
            if require_suffix and not is_complex_header(style):
                close(); head = None; continue                   # a color-less row that is not a complex header: not ours
            close(); n_pack += 1
            head = {'style': style, 'ratio': read(row, ratio_cols),
                    'price': row[idx['price']] if idx['price'] is not None else None,
                    'msrp': row[idx['msrp']] if idx['msrp'] is not None else None,
                    'units': num(row[idx['units']]) if idx['units'] is not None else None}
            continue
        if not (head and code): continue
        if style and base_style(style).upper() != base_style(head['style']).upper():
            close(); continue                                    # another style's row: not a component of this pack
        ratio = read(row, ratio_cols)
        units = read(row, unit_cols)
        if not ratio: continue
        article = head['style'] if is_complex_header(head['style']) else f"{head['style']}-A"
        if style and is_complex_header(style): article = style          # the name typed on the color rows (SFW273QZ08-XLXXL) is the pack's article      # -A, -A1: kept as typed
        comps.append({
            'type': 'PPK', 'pack_kind': 'COMPLEX', 'complex': True, 'complex_group': f"{article}#{n_pack}",
            'raw_style': article.upper(), 'style': base_style(head['style']),
            'customer': info['customer'], 'warehouse': info.get('warehouse', ''),
            'color_code': code.upper(), 'color': str(row[idx['color']] or '').strip().upper() if idx['color'] is not None else '',
            'desc': str(row[desc_ci]).strip() if desc_ci is not None and row[desc_ci] else '',
            'price': head['price'], 'msrp': head['msrp'], 'segment': info.get('segment', ''), 'delivery': None,
            'sizes': ratio, 'sizes_str': '-'.join(ratio), 'ratio': '-'.join(str(v) for v in ratio.values()),
            'row_pc': sum(ratio.values()), 'pc_per_pac': None,
            'ppk_qty': sum(units.values()) or None,        # this color's own pieces, as the rest of the pipeline expects
            '_ratio': ratio, 'pac_per_ctn': None, 'pc_ctn_os': None, 'size_order': list(ratio_cols),
        })
    close()
    for r in rows: r.pop('_ratio', None)
    return info['customer'], rows

def parse_simple_ppk_tab(ws):
    info = get_header_info(ws)
    header_ri, header_row = find_header_row(ws, ['STYLE'])
    if not header_row: return info['customer'], []
    def ci(name): return header_row.index(name) if name in header_row else None
    idx = {'style':ci('STYLE'),'code':ci('COLOR CODE'),'color':ci('COLOR DESCRIPTION'),
           'desc':ci('DESCRIPTION'),'ppk':ci('PPK'),'price':ci('PRICE'),'msrp':ci('MSRP')}
    rows = []
    for row in ws.iter_rows(min_row=header_ri+1, values_only=True):
        si = idx['style']
        if si is None or not row[si]: continue
        style = str(row[si]).strip()
        if not re.match(r'^[A-Z]{3}\d', style.upper()): continue
        ppk_qty = row[idx['ppk']] if idx['ppk'] is not None else None
        if not ppk_qty: continue
        rows.append({
            'type':'PPK', 'raw_style':style, 'style':style,
            'customer': info['customer'], 'warehouse': info.get('warehouse',''),
            'color_code': str(row[idx['code']]).strip()          if idx['code']  is not None and row[idx['code']]  else '',
            'color':      str(row[idx['color']]).strip().upper()  if idx['color'] is not None and row[idx['color']] else '',
            'desc':       str(row[idx['desc']]).strip()           if idx['desc']  is not None and row[idx['desc']]  else '',
            'price':      row[idx['price']]                       if idx['price'] is not None                       else None,
            'msrp':       row[idx['msrp']]                        if idx['msrp'] is not None                        else None,
            'ppk_qty':    int(float(ppk_qty)),
            'segment':    info.get('segment',''), 'delivery': None,
            'sizes':{},'sizes_str':'','ratio':'','pc_per_pac':None,'pac_per_ctn':None,'pc_ctn_os':None,
            'size_order': [], 'pack_kind': 'PPK',
        })
    return info['customer'], rows

# ── 4c. the new prebuy template: HEADER + OS grid (units by size) + PPK grid (sizes / ratio / qty as text) ──
TEMPLATE_MARK = 'HIGH LIFE PREBUY TEMPLATE'

def is_template(wb):
    """The new prebuy template announces itself in HEADER!A1."""
    return 'HEADER' in wb.sheetnames and str(wb['HEADER']['A1'].value or '').strip().upper().startswith(TEMPLATE_MARK)

def read_template_header(ws):
    """HEADER tab = label in column A, value in column B. Sales type segment and storage location ONCE here;
    they apply to every row and travel on to the buysheet and the PR file."""
    raw = {}
    for row in ws.iter_rows(min_row=1, max_row=40, values_only=True):
        if row and row[0] and len(row) > 1 and row[1] not in (None, ''):
            raw[str(row[0]).strip().upper()] = row[1]
    def text(key, upper=False):
        v = str(raw.get(key, '') or '').strip()
        return v.upper() if upper else v
    return {'customer': text('CUSTOMER'), 'brand': text('BRAND'), 'season_label': parse_season(raw.get('SEASON')), 'segment': text('SEGMENT', True),
            'warehouse': text('WAREHOUSE / STORAGE LOCATION', True), 'currency': text('CURRENCY', True),
            'hang_fold': text('HANG / FOLD', True), 'tickets': text('PRICE TICKETS'),
            'terms': raw.get('TERMS / ALLOWANCES'), 'terms_desc': text('DESCRIPTION OF TERMS')}

def template_columns(ws):
    """(header row number, {HEADING: column index}) - columns are found by heading, never by position."""
    for ri, row in enumerate(ws.iter_rows(min_row=1, max_row=12, values_only=True), 1):
        cells = [re.sub(r'\s+', ' ', str(v)).strip().upper() if v is not None else '' for v in row]
        if 'STYLE' in cells and 'COLOR CODE' in cells:
            return ri, {h: i for i, h in enumerate(cells) if h}
    return None, {}

def template_header_fields(header):
    """What every row takes from the HEADER tab."""
    return {'customer': header['customer'], 'segment': header['segment'], 'warehouse': header['warehouse'], 'hang_fold': header['hang_fold'],
            'tickets': header['tickets'], 'currency': header['currency'], 'brand': header['brand'], 'season_label': header['season_label'],
            'terms': header['terms'], 'terms_desc': header['terms_desc']}

def template_rows(ws, kind, header):
    """OS tab: a grid, one row per style + color, units in the size columns (only the sizes the prebuy has)."""
    hr, C = template_columns(ws)
    if hr is None:
        return []
    size_cols = [(h, i) for h, i in C.items() if h in KNOWN_SIZES]
    def cell(row, h):
        i = C.get(h)
        return row[i] if i is not None and i < len(row) else None
    out = []
    for ri, row in enumerate(ws.iter_rows(min_row=hr + 1, values_only=True), hr + 1):
        style = str(cell(row, 'STYLE') or '').strip().upper()
        if not style or str(cell(row, 'NOTES') or '').strip().upper().startswith('EXAMPLE'):
            continue
        units = {}
        for h, i in size_cols:
            q = num(row[i]) if i < len(row) else None
            if q and q > 0:
                units[h] = int(round(q))
        if not units:
            continue
        pc = cell(row, 'PC PER CTN OS')
        out.append({**template_header_fields(header), 'segment': str(cell(row, 'SEGMENT') or '').strip().upper() or header['segment'],
                    'style': style, 'color_code': str(cell(row, 'COLOR CODE') or '').strip().upper(),
                    'color': str(cell(row, 'COLOR DESCRIPTION') or '').strip().upper(), 'desc': str(cell(row, 'DESCRIPTION') or '').strip(),
                    'price': cell(row, 'SELL PRICE'), 'msrp': cell(row, 'MSRP'), 'delivery': cell(row, 'DELIVERY'),
                    'notes': str(cell(row, 'NOTES') or '').strip(), 'size_order': [h for h, _ in size_cols], 'source': f"{ws.title}!{ri}",
                    'type': 'OS', 'pack_kind': 'OS', 'raw_style': style, 'pc_ctn': pc, 'pc_ctn_os': pc, 'sizes': units, 'ratio': '',
                    'sizes_str': '', 'pc_per_pac': None, 'pac_per_ctn': None, 'ppk_qty': None})
    return out

def template_ppk_rows(ws, header, complex_mode=False):
    """PPK tab: a grid, one ROW per prepack. No size columns: SIZES, SIZE RATIO and PPK QTY are typed as text.
    The buysheet article is built from the two fields sales type: STYLE + '-' + COLOR CODE (SFW271SH01 + S470 -> SFW271SH01-S470).
    COMPLEX tab (complex_mode): one row per COLOR; rows with the same STYLE are one prepack. PPK QTY is that color's
    own pieces; those rows become one complex_group and the article STYLE-A (see apply_complex_split)."""
    hr, C = template_columns(ws)
    if hr is None:
        return []
    def cell(row, h):
        i = C.get(h)
        return row[i] if i is not None and i < len(row) else None
    out = []
    for ri, row in enumerate(ws.iter_rows(min_row=hr + 1, values_only=True), hr + 1):
        style = str(cell(row, 'STYLE') or '').strip().upper()
        code = str(cell(row, 'COLOR CODE') or '').strip().upper()
        if (not style and not code) or str(cell(row, 'NOTES') or '').strip().upper().startswith('EXAMPLE'):
            continue                  # an empty row or the EXAMPLE row; a half-filled one is kept so it shows up as incomplete
        qty = num(cell(row, 'QTY') if 'QTY' in C else cell(row, 'PPK QTY'))
        ratio = dash_list(cell(row, 'SIZE RATIO'))
        extra = ({'raw_style': style, 'complex_group': style,
                  'row_pc': sum(int(x) for x in ratio if x.isdigit())} if complex_mode else {'raw_style': f"{style}-{code}"})
        out.append({**template_header_fields(header), **extra, 'segment': str(cell(row, 'SEGMENT') or '').strip().upper() or header['segment'],
                    'style': style, 'color_code': code,
                    'color': str(cell(row, 'COLOR DESCRIPTION') or '').strip().upper(), 'desc': str(cell(row, 'DESCRIPTION') or '').strip(),
                    'price': cell(row, 'SELL PRICE'), 'msrp': cell(row, 'MSRP'), 'delivery': cell(row, 'DELIVERY'),
                    'ratio': '-'.join(ratio), 'sizes_str': '-'.join(dash_list(cell(row, 'SIZES'))),
                    'pc_per_pac': num(cell(row, 'PC PER PAC')) or (sum(int(x) for x in ratio if x.isdigit()) or None),   # no column on the template: pieces per pack = the ratio added up
                    'pac_per_ctn': num(cell(row, 'PAC PER CTN')),
                    'ppk_qty': int(round(qty)) if qty is not None else None,
                    'notes': str(cell(row, 'NOTES') or '').strip(), 'source': f"{ws.title}!{ri}",
                    'type': 'PPK', 'pack_kind': 'PPK', 'sizes': {}, 'pc_ctn': None, 'pc_ctn_os': None, 'size_order': []})
    return out

def apply_complex_split(ppk):
    """COMPLEX tab rows: the rows of one 'complex_group' (same style) are one prepack. PPK QTY on each row is that
    color's own pieces, as typed; the whole-prepack size is every color's ratio added together."""
    groups = {}
    for r in ppk:
        if r.get('complex_group'):
            groups.setdefault(r['complex_group'], []).append(r)
    for style, grp in groups.items():
        whole = sum(x['row_pc'] for x in grp)
        article = f"{style}-A"                     # a complex prepack is the style plus a letter suffix, never a color code
        for x in grp:
            x.update(complex=True, pack_kind='COMPLEX', pc_per_pac=whole, raw_style=article)

def parse_template(wb):
    header = read_template_header(wb['HEADER'])
    rows = template_rows(wb['OS'], 'OS', header) if 'OS' in wb.sheetnames else []
    ppk = template_ppk_rows(wb['PPK'], header) if 'PPK' in wb.sheetnames else []
    cx = template_ppk_rows(wb['COMPLEX'], header, complex_mode=True) if 'COMPLEX' in wb.sheetnames else []
    apply_complex_split(cx)
    ppk = ppk + cx
    return rows + ppk, ([header['customer']] if header['customer'] else [])

# ── 4d. entry point ──
def _dpci_sheet(wb):
    """The sheet of a DPCI list that holds the list: the first one whose top row has the list's columns. Other sheets (such as the VD52
    upload layout saved beside it) are ignored, and the sheet that happened to be active when the file was saved does not matter."""
    for ws in wb.worksheets:
        head = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        cells = [str(v or '').strip().upper() for v in head]
        if 'COMPONENT DPCI' in cells and 'COLOR' in cells and 'SIZE' in cells and any('STYLE' in c for c in cells):
            return ws
    return None

def is_dpci_list(wb):
    """A DPCI list: one flat sheet of style, component DPCI, description, color name and size (no quantities, no color codes)."""
    return 'DATA' not in wb.sheetnames and _dpci_sheet(wb) is not None

def parse_dpci_list(wb):
    """The list as open-stock rows (one per style and color, sizes across) so the normal checks and the creation sheets can use it.
    Each size carries a quantity of 1 only as a placeholder: the rows are marked creation_only so no requisition is written from them.
    Color codes are not in the file: the portal fills them from SAP by color name (see routes.validate_bytes).
    Stray spaces (leading, trailing, doubled) in the style, color, size and DPCI cells are trimmed. Every other difference is kept
    exactly as typed, case included: it is a real difference when the color is compared with SAP."""
    ws = _dpci_sheet(wb)
    rows_iter = ws.iter_rows(values_only=True)
    head = [str(v or '').strip().upper() for v in next(rows_iter)]
    col = lambda test: next(i for i, h in enumerate(head) if test(h))
    i_style, i_dpci, i_color, i_size = col(lambda h: 'STYLE' in h), head.index('COMPONENT DPCI'), head.index('COLOR'), head.index('SIZE')
    i_desc = head.index('DESCRIPTION') if 'DESCRIPTION' in head else None
    grouped = {}
    for r in rows_iter:
        trim = lambda v: ' '.join(str(v or '').split())
        cells = {'style': trim(r[i_style]), 'color': trim(r[i_color]), 'size': trim(r[i_size]), 'DPCI': trim(r[i_dpci])}
        if not cells['style'] and not cells['size']:
            continue
        style, size, exact = cells['style'].upper(), cells['size'].upper(), cells['color']
        g = grouped.setdefault((style, exact), {
            'type': 'OS', 'raw_style': style, 'style': style, 'customer': '', 'warehouse': '', 'color_code': '', 'color': exact.upper(), 'color_exact': exact,
            'desc': ' '.join(str(r[i_desc] or '').split()) if i_desc is not None else '', 'delivery': None, 'price': None, 'msrp': None,
            'segment': '', 'pc_ctn_os': None, 'sizes': {}, 'size_order': [], 'pack_kind': 'OS', 'ratio': '', 'sizes_str': '',
            'pc_per_pac': None, 'pac_per_ctn': None, 'ppk_qty': None, 'creation_only': True, 'dpci_by_size': {}, 'desc_by_size': {}})
        g['sizes'][size] = 1
        g['size_order'].append(size)
        g['dpci_by_size'][size] = cells['DPCI']
        g['desc_by_size'][size] = ' '.join(str(r[i_desc] or '').split()) if i_desc is not None else ''
    return list(grouped.values()), []

def _hold_sheet(wb):
    """The sheet of a hold master (every hold order entry file stacked into one): one line per customer / style / color / size, with the
    columns Generic, Color Code, Color Desc, Size, Qty and Customer."""
    for ws in wb.worksheets:
        head = [str(v or '').strip().upper() for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())]
        if {'GENERIC', 'COLOR CODE', 'COLOR DESC', 'SIZE', 'QTY', 'CUSTOMER'} <= set(head):
            return ws
    return None

def parse_hold_master(wb):
    """The hold master as open-stock rows (sizes across), one per customer / style / color code / delivery / price, so every normal check runs on it:
    exists, color code against SAP's name, ZSLS, ZVP1, description, season, and the stock under the same color's other codes.
    Prepack lines (size PPK, a ratio instead of a size) and sizes the portal does not know are counted, not checked: skipped = how many."""
    ws = _hold_sheet(wb)
    it = ws.iter_rows(values_only=True)
    head = [str(v or '').strip().upper() for v in next(it)]
    at = lambda *names: next((head.index(n) for n in names if n in head), None)
    ix = {k: at(*v) for k, v in {'customer': ('CUSTOMER',), 'style': ('GENERIC',), 'code': ('COLOR CODE',), 'color': ('COLOR DESC',), 'size': ('SIZE',), 'qty': ('QTY',),
                                  'desc': ('DESCRIPTION',), 'delivery': ('DELIVERY',), 'price': ('PRICE',), 'msrp': ('MSRP',), 'segment': ('SEGMENT',),
                                  'source': ('SOURCE FILE',), 'division': ('DIVISION',)}.items()}
    get = lambda r, k: r[ix[k]] if ix[k] is not None else None
    grouped, skipped = {}, 0
    for r in it:
        style = ' '.join(str(get(r, 'style') or '').split()).upper()
        if not style:
            continue
        size = canonical_size(str(get(r, 'size') or '').strip())
        qty = num(get(r, 'qty'))
        if size not in KNOWN_SIZES or not qty:
            skipped += 1
            continue
        cust, code = ' '.join(str(get(r, 'customer') or '').split()), ' '.join(str(get(r, 'code') or '').split()).upper()
        g = grouped.setdefault((cust, style, code, str(get(r, 'color') or '').strip().upper(), get(r, 'delivery'), get(r, 'price'), get(r, 'msrp')), {
            'type': 'OS', 'raw_style': style, 'style': style, 'customer': cust, 'warehouse': '', 'color_code': code,
            'color': str(get(r, 'color') or '').strip().upper(), 'desc': ' '.join(str(get(r, 'desc') or '').split()),
            'delivery': get(r, 'delivery'), 'price': get(r, 'price'), 'msrp': get(r, 'msrp'), 'segment': str(get(r, 'segment') or '').strip(),
            'pc_ctn_os': None, 'sizes': {}, 'size_order': [], 'pack_kind': 'OS', 'ratio': '', 'sizes_str': '', 'pc_per_pac': None, 'pac_per_ctn': None, 'ppk_qty': None,
            'source': str(get(r, 'source') or ''), 'division': str(get(r, 'division') or '')})
        g['sizes'][size] = g['sizes'].get(size, 0) + int(qty)
        if size not in g['size_order']: g['size_order'].append(size)
    rows = list(grouped.values())
    for g in rows:
        g['size_order'].sort(key=lambda z: SIZE_ORDER.index(z) if z in SIZE_ORDER else 99)
        g['sizes'] = {z: g['sizes'][z] for z in g['size_order']}
    return rows, sorted({g['customer'] for g in rows if g['customer']}), skipped

def parse_prebuy(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    if 'DATA' not in wb.sheetnames and _hold_sheet(wb) is not None:
        rows, customers, skipped = parse_hold_master(wb)
        return rows, customers, 'hold_master'
    if is_dpci_list(wb):
        rows, customers = parse_dpci_list(wb)
        return rows, customers, 'dpci_list'
    if detect_file_type(wb) == 'buysheet':
        rows, customers = parse_buysheet(wb)
        return rows, customers, 'buysheet'
    if is_template(wb):
        rows, customers = parse_template(wb)
        return rows, customers, 'template'
    all_rows=[]; customers=[]
    SKIP=['DO NOT','COLOR CODE','SETTINGS','REFERENCE','LOOKUP','TERMS']
    for sheet in wb.sheetnames:
        ws=wb[sheet]; snu=sheet.strip().upper()
        if any(x in snu for x in SKIP): continue
        if 'COMPLEX' in snu:
            cust,rows=parse_complex_tab(ws)
            if not rows: cust,rows=parse_ppk_tab(ws)
        elif 'PREPACK' in snu or 'PPK' in snu:
            cust,rows=parse_ppk_tab(ws)
            _, cx_rows = parse_complex_tab(ws, require_suffix=True)      # complex packs typed inline: a shaded -A row with its colors below
            rows = rows + cx_rows
        elif has_ppk_column(ws): cust,rows=parse_simple_ppk_tab(ws)
        else: cust,rows=parse_os_tab(ws)
        if rows:
            hf = get_header_info(ws)
            for r in rows:
                r.setdefault('hang_fold', hf.get('hang_fold') or '')
                r.setdefault('tickets', hf.get('tickets') or '')
                r.setdefault('brand', hf.get('brand') or '')
                r.setdefault('terms', hf.get('terms') or '')
                r.setdefault('terms_desc', hf.get('terms_desc') or '')
            if cust and cust not in customers: customers.append(cust)
            all_rows.extend(rows)
    return all_rows, customers, 'prebuy'
