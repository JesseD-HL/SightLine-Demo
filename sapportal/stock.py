"""The SAP "Cut and Sold" stock, from an export or live, joined to every article by article number."""

import csv
import io
import re
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from .config import STOCK_VIEW, _cache
from .hana import connect_hana, query_variant_articles
from .helpers import mmddyyyy, num, season_text, up
from .outputs import BDR, CTR, DFONT, LEFT, NAVY, WBOLD
from .validation import check_counts, gen_check_list, ppk_check_list


# ═══════════════════════════════════════════════════════════════════════════════
# 6c. STOCK - the SAP "Cut and Sold" report, exported, joined to every article by article number
# ═══════════════════════════════════════════════════════════════════════════════
STOCK_HEADERS = {   # key -> the report's column headings it can appear under (lower case, single spaces)
    'site': ('site', 'plant'), 'merch': ('merchcat', 'merch cat'), 'season': ('yr-seas', 'season'), 'article': ('article', 'material'),
    'segment': ('segment', 'stock segment'), 'sloc': ('s.loc', 'sloc', 'storage location', 'storage loc'), 'desc': ('description',),
    'color': ('color desc', 'color'), 'size': ('size',),
    'on_po': ('on po',), 'in_txt': ('in txt', 'in transit'), 'inv': ('inv qty',), 'dlv': ('dlv qty',), 'order': ('order qty',),
    'ata': ('ata inv',), 'ats': ('ats inv',), 'oversold': ('oversold inv',)}
STOCK_NUMBERS = ('on_po', 'in_txt', 'inv', 'dlv', 'order', 'ata', 'ats', 'oversold')
STOCK_CHECK = (('stock_check', 'In stock'),)

def _norm_header(v):
    return re.sub(r'\s+', ' ', str(v or '').strip().lower())

def sap_number(v):
    """SAP quantities: 12,428  /  1.234,5  /  5-  (a trailing minus is negative). Blank or text -> 0."""
    if v is None or v == '': return 0.0
    if isinstance(v, (int, float)): return float(v)
    t = str(v).strip().replace('\xa0', '').replace(' ', '')
    neg = t.endswith('-') or t.startswith('-') or (t.startswith('(') and t.endswith(')'))
    t = t.strip('-()')
    t = t.replace('.', '').replace(',', '.') if re.fullmatch(r'\d{1,3}(\.\d{3})+,\d+', t) else t.replace(',', '')
    try: n = float(t)
    except ValueError: return 0.0
    return -n if neg else n

def read_stock_table(file_bytes, filename=''):
    """Rows of cells from an Excel export or from a delimited text export (CSV, SAP 'text with tabs')."""
    if (filename or '').lower().endswith(('.xlsx', '.xlsm')) or file_bytes[:2] == b'PK':
        ws = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True).worksheets[0]
        return [list(r) for r in ws.iter_rows(values_only=True)]
    if file_bytes[:2] in (b'\xff\xfe', b'\xfe\xff'): text = file_bytes.decode('utf-16')
    else:
        try: text = file_bytes.decode('utf-8-sig')
        except UnicodeDecodeError: text = file_bytes.decode('cp1252', errors='replace')
    lines = [l for l in text.splitlines() if l.strip() and not re.fullmatch(r'[\s\-+=|_]+', l)]      # drop SAP's ruler lines
    sample = '\n'.join(lines[:40])
    delim = max(('\t', ';', '|', ','), key=lambda d: sample.count(d))
    out = []
    for r in csv.reader(lines, delimiter=delim):
        r = [c.strip() for c in r]
        if delim == '|':
            r = r[1:] if r and r[0] == '' else r
            r = r[:-1] if r and r[-1] == '' else r
        out.append(r)
    return out

def parse_cut_sold(file_bytes, filename=''):
    """The Cut and Sold export -> {'articles': {ARTICLE: totals + the split by site / segment / storage location}, ...}.
    The report's own total and sub-total rows (no article) are skipped, or every unit would be counted twice."""
    table = read_stock_table(file_bytes, filename)
    head = next((i for i, row in enumerate(table[:40])
                 if 'article' in [_norm_header(c) for c in row] and any(h in [_norm_header(c) for c in row] for h in ('inv qty', 'ats inv', 'ats total'))), None)
    if head is None:
        raise ValueError("This does not look like the Cut and Sold report: no row has an 'Article' column together with 'Inv Qty' or 'ATS Inv'. "
                         "Export the report to Excel or CSV with its columns, then try again.")
    cells = [_norm_header(c) for c in table[head]]
    col = {}
    for key, names in STOCK_HEADERS.items():                      # the first heading variant that is present wins
        for n in names:
            if n in cells:
                col[key] = cells.index(n); break
    articles, lines, skipped = {}, 0, 0
    for row in table[head + 1:]:
        get = lambda k: row[col[k]] if k in col and col[k] < len(row) else None
        art = str(get('article') or '').strip().upper()
        if not art:
            skipped += 1; continue
        lines += 1
        a = articles.setdefault(art, dict({k: 0.0 for k in STOCK_NUMBERS}, by={}, desc='', color='', size='', season=''))
        a['desc'] = a['desc'] or str(get('desc') or '').strip(); a['color'] = a['color'] or str(get('color') or '').strip()
        a['size'] = a['size'] or str(get('size') or '').strip(); a['season'] = a['season'] or str(get('season') or '').strip()
        where = (str(get('site') or '').strip(), str(get('segment') or '').strip(), str(get('sloc') or '').strip())
        b = a['by'].setdefault(where, {k: 0.0 for k in STOCK_NUMBERS})
        for k in STOCK_NUMBERS:
            n = sap_number(get(k)); a[k] += n; b[k] += n
    whole = lambda x: int(x) if float(x).is_integer() else round(x, 2)
    for a in articles.values():
        for k in STOCK_NUMBERS: a[k] = whole(a[k])
        a['by'] = sorted(({'site': w[0], 'segment': w[1], 'sloc': w[2], **{k: whole(v) for k, v in n.items()}} for w, n in a['by'].items()), key=lambda x: -x['inv'])
    m = re.search(r'(\d{4})(\d{2})(\d{2})[_\-](\d{2})(\d{2})(\d{2})', filename or '')
    return {'source': 'export', 'filename': filename, 'as_of': f"{m.group(1)}-{m.group(2)}-{m.group(3)} {m.group(4)}:{m.group(5)}" if m else '',
            'lines': lines, 'skipped_total_rows': skipped, 'articles': articles, 'article_count': len(articles),
            'total_inv': whole(sum(a['inv'] for a in articles.values())), 'total_ats': whole(sum(a['ats'] for a in articles.values())),
            'sites': sorted({b['site'] for a in articles.values() for b in a['by'] if b['site']}),
            'segments': sorted({b['segment'] for a in articles.values() for b in a['by'] if b['segment']}),
            'columns': sorted(col), 'missing_columns': [k for k in ('inv', 'ats', 'on_po', 'order') if k not in col]}

STOCK_TYPES = (('inv', 'CLABS', 'on hand'), ('transfer', 'CUMLM', 'in transfer'), ('quality', 'CINSM', 'in quality inspection'),
               ('restricted', 'CEINM', 'restricted'), ('blocked', 'CSPEM', 'blocked'), ('returns', 'CRETM', 'returns'))

def query_hana_stock(conn, articles):
    """On-hand stock straight from STOCK_VIEW (default SAPDEMO.ZSD_CUTSOLD_STK, the Cut and Sold report's own stock view:
    checked line for line against the report, identical). One line per article / site / segment / storage location.
    'On hand' is unrestricted stock (CLABS), as on the report; returns, blocked and the other stock types are kept
    beside it per storage location, because that stock is owned too."""
    schema, view = STOCK_VIEW.split('.', 1)
    ident = lambda x: '"' + str(x).replace('"', '""') + '"'
    arts = sorted({str(a).strip().upper() for a in articles if a})
    sums = ', '.join(f'SUM({ident(col)}) AS {ident(key)}' for key, col, _ in STOCK_TYPES)
    cur, rows = conn.cursor(), []
    for i in range(0, len(arts), 500):
        part = arts[i:i + 500]
        cur.execute(f'SELECT "MATNR", "WERKS", "SGT_SCAT", "LGORT", {sums} FROM {ident(schema)}.{ident(view)} '
                    f'WHERE "MATNR" IN ({", ".join("?" for _ in part)}) GROUP BY "MATNR", "WERKS", "SGT_SCAT", "LGORT"', tuple(part))
        rows += cur.fetchall()
    whole = lambda x: int(x) if float(x).is_integer() else round(float(x), 2)
    out, lines = {}, 0
    for r in rows:
        art = str(r[0] or '').strip().upper()
        nums = {key: float(r[4 + i] or 0) for i, (key, _, _) in enumerate(STOCK_TYPES)}
        if not art or not any(nums.values()):
            continue
        lines += 1
        a = out.setdefault(art, {'inv': 0.0, 'ats': None, 'on_po': None, 'in_txt': None, 'dlv': None, 'order': None, 'ata': None,
                                 'oversold': None, 'by': [], 'other': {}, 'other_total': 0.0, 'desc': '', 'color': '', 'size': '', 'season': ''})
        a['inv'] += nums['inv']
        other = {k: whole(v) for k, v in nums.items() if k != 'inv' and v}
        for k, v in other.items():
            a['other'][k] = whole(a['other'].get(k, 0) + v); a['other_total'] += v
        a['by'].append({'site': str(r[1] or '').strip(), 'segment': str(r[2] or '').strip(), 'sloc': str(r[3] or '').strip(),
                        'inv': whole(nums['inv']), 'ats': None, 'other': other})
    for a in out.values():
        a['inv'] = whole(a['inv']); a['other_total'] = whole(a['other_total'])
        a['by'].sort(key=lambda b: (-b['inv'], b['sloc'], b['segment']))
    return {'source': 'hana', 'view': STOCK_VIEW, 'filename': STOCK_VIEW, 'as_of': datetime.now().strftime('%Y-%m-%d %H:%M'),
            'lines': lines, 'skipped_total_rows': 0, 'articles': out, 'article_count': len(out),
            'total_inv': whole(sum(a['inv'] for a in out.values())), 'total_ats': None,
            'total_other': whole(sum(a['other_total'] for a in out.values())),
            'sites': sorted({b['site'] for a in out.values() for b in a['by'] if b['site']}),
            'segments': sorted({b['segment'] for a in out.values() for b in a['by'] if b['segment']}),
            'labels': {k: label for k, _, label in STOCK_TYPES}, 'columns': [], 'missing_columns': []}

def live_stock(validation):
    """Live stock for every article in this validation (each size's article, each prepack and its components).
    A failure never stops the validation: it is reported on the stock strip and an export can still be added."""
    S = validation.setdefault('summary', {})
    S.pop('stock_error', None)
    if not STOCK_VIEW:
        return None
    arts = ({r.get('sap_article') for r in validation.get('generic', [])} | {a['article'] for r in validation.get('generic', []) for a in r.get('alt_articles') or []}
            | {r.get('ppk_article') for r in validation.get('ppk', [])}
            | {c.get('article') for r in validation.get('ppk', []) for c in r.get('components') or []})
    arts = {str(a).strip().upper() for a in arts if a}
    if not arts:
        return None
    try:
        conn = connect_hana()
        try:
            stock = query_hana_stock(conn, arts)
        finally:
            conn.close()
    except Exception as exc:
        S['stock_error'] = f"Live stock could not be read from {STOCK_VIEW}: {str(exc).splitlines()[0][:200]}. You can still add a Cut and Sold export."
        return None
    _cache['live_stock'] = stock
    return stock

def stock_info(stock):
    """The stock report's summary for the portal (never the per-article detail)."""
    return {k: v for k, v in stock.items() if k != 'articles'} if stock else None

def _stock_numbers(agg, source='export'):
    return None if not agg else dict({k: agg[k] for k in ('inv', 'ats', 'on_po', 'in_txt', 'order', 'dlv', 'oversold')}, by=agg['by'],
                                     other=agg.get('other', {}), other_total=agg.get('other_total', 0), source=source)

def apply_stock(validation, stock):
    """Adds what is already in stock to every result: an open-stock line by its SAP variant article, a prepack by its own
    article and by each component article (and how many packs the components could make). No stock -> clears it again."""
    arts = (stock or {}).get('articles', {})
    src = (stock or {}).get('source', 'export')
    unknown = None if src == 'hana' else 0               # live stock has no ATS / PO figures (yet): unknown, not zero
    for lst in (validation.get('generic', []), validation.get('generic_unique', [])):
        for r in lst:
            if not stock:
                r.pop('stock', None); r.pop('stock_check', None); r.pop('alt_stock', None); continue
            agg = arts.get((r.get('sap_article') or '').upper())
            r['stock'] = _stock_numbers(agg, src)
            alt = []
            for a in r.get('alt_articles') or []:                      # same style + color name, other color code
                g = arts.get(str(a['article']).upper())
                if g and g['inv'] > 0: alt.append({**a, 'inv': g['inv']})
            r['alt_stock'] = alt
            alt_n = sum(a['inv'] for a in alt)
            own_n = int(agg['inv']) if agg and agg['inv'] > 0 else 0
            r['stock_check'] = (f"ℹ IN STOCK {own_n:,}" + (f" · OTHER CODE {int(alt_n):,}" if alt_n else '')) if own_n else \
                               (f"ℹ OTHER CODE {int(alt_n):,}" if alt_n else '✓ NONE')
    for lst in (validation.get('ppk', []), validation.get('ppk_unique', [])):
        for r in lst:
            if not stock:
                r.pop('stock', None); r.pop('stock_check', None); continue
            comps = r.get('components') or []
            for c in comps:
                a = arts.get((c.get('article') or '').upper())
                c['inv'], c['ats'], c['on_po'] = (a['inv'], a['ats'], a['on_po']) if a else (0, unknown, unknown)
            mine = [c for c in comps if not r.get('complex') or c.get('code') == r.get('color_code')]
            have = [c for c in mine if c['inv'] > 0]
            sized = [c for c in mine if (c.get('qty') or 0) > 0]
            own = arts.get((r.get('ppk_article') or '').upper())
            usable = lambda c: c['ats'] if c['ats'] is not None else c['inv']        # ATS when known, else on hand
            r['stock'] = {'own': _stock_numbers(own, src), 'components': len(mine), 'components_in_stock': len(have), 'source': src,
                          'packs': min(int(usable(c) // c['qty']) for c in sized) if sized else None, 'units': sum(c['inv'] for c in mine)}
            r['stock_check'] = (f"ℹ IN STOCK {int(own['inv']):,}" if own and own['inv'] > 0
                                else f"ℹ COMPONENTS {len(have)}/{len(mine)}" if have else '✓ NONE')
    S = validation.setdefault('summary', {})
    S['checks_gen'] = check_counts(validation.get('generic_unique', []), gen_check_list(validation.get('generic_unique', []), STOCK_CHECK if stock else ()))
    S['checks_ppk'] = check_counts(validation.get('ppk_unique', []), ppk_check_list(validation.get('ppk_unique', []), STOCK_CHECK if stock else ()))
    if stock: S['stock'] = stock_info(stock)
    else: S.pop('stock', None)

def make_validation_report(validation):
    """GENERIC VALIDATION, PPK VALIDATION, PPK COMPONENTS and SUMMARY - the layout _sap_creation_validator.py produces."""
    GREEN, RED, AMBER = PatternFill('solid', start_color='E2EFDA'), PatternFill('solid', start_color='FCE4E4'), PatternFill('solid', start_color='FFF2CC')
    GF, RF, AF = Font(name='Arial', size=10, color='375623', bold=True), Font(name='Arial', size=10, color='C00000', bold=True), Font(name='Arial', size=10, color='7F6000', bold=True)
    def write(ws, headers, lines, widths, left=(1, 2)):
        for ci, h in enumerate(headers, 1):
            c = ws.cell(1, ci, h); c.font = WBOLD; c.fill = NAVY; c.alignment = CTR; c.border = BDR
        for ri, line in enumerate(lines, 2):
            for ci, v in enumerate(line, 1):
                c = ws.cell(ri, ci, v); c.border = BDR; c.alignment = LEFT if ci in left else CTR; c.font = DFONT
                if isinstance(v, str) and v and v[0] in '✓✗⚠':
                    c.fill, c.font = (GREEN, GF) if v[0] == '✓' else (RED, RF) if v[0] == '✗' else (AMBER, AF)
        for ci, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(len(lines), 1) + 1}"
    wb = Workbook(); ws = wb.active; ws.title = 'SUMMARY'
    S = validation.get('summary', {})
    lines = [('CREATION VALIDATION SUMMARY', None), (validation.get('filename', ''), None), ('', None),
             ('GENERIC (OPEN STOCK)', None), (f"✓ Size rows already in SAP : {S.get('gen_exists', 0)}", 'g'), (f"✗ Size rows to create       : {S.get('gen_missing', 0)}", 'r')]
    count_text = lambda c: f"{c['label']:<12} ✓ {c['ok']}   ✗ {c['bad']}   ⚠ {c['warn']}" + (f"   ℹ {c['info']}" if c.get('info') else '')
    for c in S.get('checks_gen', []): lines.append((count_text(c), 'r' if c['bad'] else 'g'))
    lines += [('', None), ('PREPACKS', None), (f"✓ Already in SAP : {S.get('ppk_exists', 0)}", 'g'), (f"✗ To create       : {S.get('ppk_missing', 0)}", 'r')]
    for c in S.get('checks_ppk', []): lines.append((count_text(c), 'r' if c['bad'] else 'g'))
    for i, (t, kind) in enumerate(lines, 1):
        c = ws.cell(i, 1, t); c.font = Font(name='Consolas', size=14 if i == 1 else 10, bold=(i == 1 or kind is None), color={'g': '375623', 'r': 'C00000'}.get(kind, '000000'))
    ws.column_dimensions['A'].width = 60
    wg = wb.create_sheet('GENERIC VALIDATION')
    write(wg, ['STYLE','COLOR CODE','SIZE','STATUS','WHY','BUY PRICE','SAP PRICE','PRICE ✓','BUY COLOR','SAP COLOR','COLOR ✓','BUY DESC','SAP DESC','DESC ✓','PC PER CTN ✓','ROUNDING ✓','ROWS IN FILE','CUSTOMERS'],
          [[r['style'], r['color_code'], r['size'], r['status'], r.get('missing_reason', ''), r.get('buy_price'), r.get('sap_price'), r['price_check'], r.get('buy_color'), r.get('sap_color'),
            r['color_check'], r.get('buy_desc'), r.get('sap_desc'), r['desc_check'], r['pc_ctn_check'], r['round_check'], r.get('count', 1), ', '.join(r.get('customers', []))] for r in validation.get('generic_unique', validation.get('generic', []))],
          [16, 12, 8, 14, 50, 10, 10, 18, 22, 22, 24, 36, 36, 12, 24, 14, 10, 40])
    p = validation.get('ppk_unique', validation.get('ppk', []))
    wp = wb.create_sheet('PPK VALIDATION')
    write(wp, ['PPK ARTICLE','PREBUY ARTICLE','COLOR CODE','CC ✓','EXISTS','PC PER PPK ✓','BUY PRICE','SAP PRICE','PRICE ✓','BUY COLOR','SAP COLOR','COLOR ✓',
               'BUY SIZES','SAP SIZES','SIZES ✓','BUY RATIO','SAP RATIO','RATIO ✓','RATIO DIFFERENCES','BUY DESC','SAP DESC','ROUNDING ✓','ROWS IN FILE','CUSTOMERS'],
          [[r['ppk_article'], r.get('prebuy_article', ''), r['color_code'], r['cc_check'], r['exists'], r['pc_ppk_check'], r.get('buy_price'), r.get('sap_price'), r['price_check'],
            r.get('buy_color'), r.get('sap_color'), r['color_check'], r.get('buy_sizes'), r.get('sap_sizes'), r['sizes_check'], r.get('buy_ratio'), r.get('sap_ratio'), r['ratio_check'],
            r.get('ratio_note', ''), r.get('buy_desc'), r.get('sap_desc'), r['round_check'], r.get('count', 1), ', '.join(r.get('customers', []))] for r in p],
          [22, 20, 12, 10, 12, 30, 10, 10, 22, 22, 22, 14, 16, 16, 10, 14, 14, 12, 40, 36, 36, 30, 10, 40])
    wc = wb.create_sheet('PPK COMPONENTS')
    write(wc, ['PPK ARTICLE','COMPONENT ARTICLE','COLOR CODE','COLOR','SIZE','QTY PER PACK'],
          [[r['ppk_article'], c.get('article'), c.get('code'), c.get('color'), c.get('size'), c.get('qty')] for r in p if r['exists'] == '✓ EXISTS' for c in r.get('components', [])],
          [22, 24, 12, 24, 8, 12])
    if S.get('stock'):
        wst = wb.create_sheet('IN STOCK')
        by_text = lambda s_: '; '.join(f"{b['segment'] or '-'}/{b['sloc'] or '-'}: {b['inv']:,}" for b in (s_ or {}).get('by', []))
        lines_st = []
        for r in validation.get('generic_unique', validation.get('generic', [])):
            st = r.get('stock')
            if st and st['inv'] > 0:
                lines_st.append(['OPEN STOCK', r.get('sap_article'), r['style'], r['color_code'], r['size'], st['inv'], st['ats'], st['on_po'], st['in_txt'], st['order'], st['oversold'], by_text(st)])
            for a in r.get('alt_stock') or []:
                lines_st.append(['SAME COLOR, OTHER CODE', a['article'], r['style'], a['code'], r['size'], a['inv'], None, None, None, None, None, f"{a['color']} (buying code {r['color_code']})"])
        for r in p:
            st = r.get('stock') or {}
            if st.get('own') and st['own']['inv'] > 0:
                o = st['own']; lines_st.append(['PREPACK', r['ppk_article'], r['ppk_article'].split('-')[0], r['color_code'], '', o['inv'], o['ats'], o['on_po'], o['in_txt'], o['order'], o['oversold'], by_text(o)])
            for c in r.get('components', []):
                if c.get('inv', 0) > 0 and (not r.get('complex') or c.get('code') == r.get('color_code')):
                    lines_st.append(['PREPACK COMPONENT', c.get('article'), r['ppk_article'], c.get('code'), c.get('size'), c['inv'], c['ats'], c['on_po'], '', '', '', f"{c.get('qty')} per pack"])
        write(wst, ['TYPE','ARTICLE','STYLE / PREPACK','COLOR CODE','SIZE','ON HAND','ATS','ON PO','IN TRANSIT','ORDER QTY','OVERSOLD','WHERE'], lines_st, [18, 20, 22, 12, 8, 10, 10, 10, 10, 10, 10, 60])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

PR_HEADER = ['Group By', 'Article', 'Purch req quantity', 'Delivery Date(MMDDYYYY)',
             'Site', 'Storage Location', 'Purch Group', 'Stock Segment', 'SEASON']

def make_req_csv(rows, customers, validation=None):
    """The purchase-requisition upload: exactly the nine columns, in this order. Plain UTF-8 (no BOM), CRLF line endings.
    Open stock: one line per size variant. Prepacks: one line per buysheet row; only a complex prepack's color rows
    (one order, one article) collapse into a single line. Warehouse, delivery date and segment come from each buysheet row.
    Prepack quantity is in PACKS: the order's pieces divided by PC PER PPK (60 pc at 5 pc per ppk = 12).
    Anything SAP does not have yet (size variants or prepacks) is left out: it has to be created first.
    The prepack's article is the one the validation found in SAP, which can differ from the name the prebuy typed."""
    os_rows = [r for r in rows if r.get('type') == 'OS' and not r.get('creation_only')]     # a DPCI list has no quantities: nothing to requisition
    ppk_rows = [r for r in rows if r.get('type') == 'PPK']
    ppk_res = (validation or {}).get('ppk') or []
    ppk_pairs = list(zip(ppk_rows, ppk_res)) if len(ppk_rows) == len(ppk_res) else []     # results are in row order

    conn = connect_hana()
    try:
        variant_map = query_variant_articles(conn, list({r['style'] for r in os_rows}))
    finally:
        conn.close()

    lines = {}                                                # (every column but the quantity) -> quantity, in order of first appearance
    class _Merge:
        def writerow(self, row):
            key = tuple(row[:2]) + tuple(row[3:])
            lines[key] = lines.get(key, 0) + row[2]
    w = _Merge()
    groups = {}                                               # customer -> Group By number, in order of first appearance
    group_no = lambda r: groups.setdefault(str(r.get('customer') or '').strip(), len(groups) + 1)
    for r in os_rows:
        style, code = r['style'], str(r.get('color_code') or '')
        group = group_no(r)
        for sz, qty in (r.get('sizes') or {}).items():
            article = variant_map.get((style.upper(), code.upper(), sz.upper()), '')
            if not article:
                continue
            w.writerow([group, up(article), int(qty), mmddyyyy(r.get('delivery')), 'H001', up(r.get('warehouse') or ''),
                        'H01', up(r.get('segment') or ''), up(season_text(r.get('season_label'), style))])

    seen_complex = set()                                      # a complex prepack's color rows are ONE order: written once
    for r, res in ppk_pairs:
        if res.get('exists') != '✓ EXISTS':
            continue
        article = res['ppk_article']
        if r.get('complex'):
            key = (article, str(r.get('customer') or ''), str(r.get('delivery') or ''))
            if key in seen_complex:
                continue
            seen_complex.add(key)
        pieces = sum((r.get('sizes') or {}).values()) or num(r.get('ppk_qty'))      # the row in pieces (an old-layout prebuy types its grid in pieces)
        per_pack = num(r.get('row_pc') if r.get('complex') else r.get('pc_per_pac')) # complex: this color's pieces per pack; else pieces per pack
        if not pieces or not per_pack:
            continue                                          # no quantity or no pieces-per-pack: the pack count cannot be worked out
        qty = pieces / per_pack                               # PAC qty = pieces / PC PER PPK  (PAC PER CTN does not enter into it)
        w.writerow([group_no(r), up(article), int(round(qty)), mmddyyyy(r.get('delivery')), 'H001', up(r.get('warehouse') or ''),
                    'H01', up(r.get('segment') or ''), up(season_text(r.get('season_label'), r['style']))])
    # Same customer, article, delivery date, warehouse, segment and season = one requisition line: quantities are added up.
    out = io.StringIO()
    cw = csv.writer(out, lineterminator='\r\n')
    cw.writerow(PR_HEADER)
    for key, qty in lines.items():
        cw.writerow([key[0], key[1], qty, *key[2:]])
    return out.getvalue()
