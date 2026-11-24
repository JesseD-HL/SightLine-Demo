"""Checks for the delivery-flow split (server.split_buysheet / parse_flow_file).  Run:  python tests/test_split.py"""
import io, os, sys, datetime
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import openpyxl
import server

D = [datetime.date(2027, 4, 9), datetime.date(2027, 4, 16), datetime.date(2027, 4, 23)]


def flow_file(styles):
    """A flow sheet in the real layout: CRD row (label in col H), header row with Style# / QTY, style rows, NORTHWIND TOTAL."""
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append([None, 'TITLE'])
    ws.append([None])
    ws.append([None, None, None, None, None, None, None, 'CRD', *D])
    ws.append([None, None, None, None, None, None, None, 'IN WH DATE - PLUS 30 DAYS', *[d + datetime.timedelta(30) for d in D]])
    ws.append([None, 'Trading Company', 'Factory', 'COO', 'Brand/Division', 'Style#', 'QTY', 'IN WH DATE - PLUS 50 DAYS', *[d + datetime.timedelta(50) for d in D]])
    for text, qtys in styles:
        ws.append([None, 'T', 'F', 'C', 'B', text, sum(qtys), 'Production Request', *qtys])
    ws.append([None, None, None, None, None, 'NORTHWIND TOTAL', 1])
    b = io.BytesIO(); wb.save(b)
    return b.getvalue()


def os_row(style, color, sizes, ctn):
    return {'type': 'OS', 'style': style, 'raw_style': style, 'color': color, 'sizes': dict(sizes), 'pc_ctn_os': ctn, 'delivery': datetime.date(2027, 4, 1)}


def cx_row(color, pp, units, ctn):
    return {'type': 'PPK', 'complex': True, 'style': 'SFM273CR01', 'raw_style': 'SFM273CR01-A', 'customer': 'BULKCO - COMPLEX', 'color': color,
            'color_code': color[:2], 'row_pc': pp, 'pc_per_pac': 30, 'ppk_qty': units, 'pac_per_ctn': ctn, 'delivery': datetime.date(2027, 4, 1)}


def test_parse_flow_file():
    flow, notes = server.parse_flow_file(flow_file([('SAM271SH07 (SPM253SH01)', [5000, 5000, 4400])]))
    e = flow['SAM271SH07']
    assert flow['SPM253SH01'] is e and e['per'] == {D[0]: 5000, D[1]: 5000, D[2]: 4400} and not notes, (e, notes)


def test_open_stock_exact_by_size_and_whole_cartons():
    sizes = {'S': 2400, 'M': 4800, 'L': 4800, 'XL': 2400}
    flow, _ = server.parse_flow_file(flow_file([('SAM271SH07', [5000, 5000, 4400])]))
    rows, rep = server.split_buysheet([os_row('SAM271SH07', 'BLACK', sizes, 12)], flow)
    for sz, q in sizes.items():
        assert sum(r['sizes'].get(sz, 0) for r in rows) == q, (sz, rows)                  # exact down to the size
    assert all(q % 12 == 0 for r in rows for q in r['sizes'].values())                    # every delivery a whole number of cartons
    assert [r['split_crd'] for r in rows] == D, [r['split_crd'] for r in rows]
    assert rep['styles'][0]['buy_total'] == 14400 == rep['styles'][0]['flow_total']


def test_complex_example_from_the_manual_method():
    """BLACK 15/pack 150,000, GREY 9/pack 90,000, GREEN 6/pack 60,000 over 3 equal deliveries -> 3,333 / 3,333 / 3,334 packs."""
    flow, _ = server.parse_flow_file(flow_file([('SFM273CR01', [100000, 100000, 100000])]))
    grp = [cx_row('BLACK', 15, 150000, 1), cx_row('GREY', 9, 90000, 1), cx_row('GREEN', 6, 60000, 1)]
    rows, rep = server.split_buysheet(grp, flow)
    by = {d: [r['ppk_qty'] for r in rows if r['split_crd'] == d] for d in D}
    assert by[D[0]] == [49995, 29997, 19998] and by[D[1]] == [49995, 29997, 19998] and by[D[2]] == [50010, 30006, 20004], by
    assert [sum(v) for v in by.values()] == [99990, 99990, 100020]
    assert sum(sum(v) for v in by.values()) == 300000                                     # the total never changes


def test_complex_colors_with_different_pack_counts_are_flagged_not_split():
    flow, _ = server.parse_flow_file(flow_file([('SFM273CR01', [100000, 100000, 100000])]))
    grp = [cx_row('BLACK', 15, 150000, 1), cx_row('GREY', 9, 99000, 1)]
    rows, rep = server.split_buysheet(grp, flow)
    assert len(rows) == 2 and 'split_crd' not in rows[0] and any('different pack counts' in n for n in rep['notes']), rep['notes']


def test_prepack_cartons_part_carton_goes_to_last_delivery():
    flow, _ = server.parse_flow_file(flow_file([('SFM273CR01', [1000, 1000, 1000])]))
    r = cx_row('BLACK', 30, 30 * 100, 6)                                                  # 100 packs, 6 per carton: 16 cartons + 4 packs
    rows, rep = server.split_buysheet([r], flow)
    packs = [x['ppk_qty'] // 30 for x in rows]
    assert sum(packs) == 100 and all(p % 6 == 0 for p in packs[:-1]) and packs[-1] % 6 == 4, packs


def test_lettered_variant_is_split_with_its_base_style():
    """CPU000FH01A is counted inside CPU000FH01 on the flow: split along it, and never reported as missing from the flow."""
    flow, _ = server.parse_flow_file(flow_file([('SAM271HD03 (CPU000FH01)', [3000, 3000, 0])]))
    a = os_row('CPU000FH01', 'BLACK', {'S': 2400}, 12)
    b = os_row('CPU000FH01A', 'BLACK', {'S': 3600}, 12)
    rows, rep = server.split_buysheet([a, b], flow)
    assert all('split_crd' in r for r in rows), rows
    assert sum(r['sizes']['S'] for r in rows if r['style'] == 'CPU000FH01A') == 3600           # exact
    assert not any('CPU000FH01A' in n and 'not in the flow' in n for n in rep['notes']), rep['notes']
    assert rep['styles'][0]['buy_total'] == 6000 == rep['styles'][0]['flow_total']             # both variants counted against the one flow line


def test_flow_listing_both_the_base_style_and_the_lettered_one_keeps_them_apart():
    """When the flow has CPU000FH01 AND CPU000FH01A as their own lines, each is split by its own line (no pooling)."""
    flow, _ = server.parse_flow_file(flow_file([('SAM271HD03 (CPU000FH01)', [3000, 3000, 0]), ('CPU000FH01A', [2400, 2400, 1200])]))
    a = os_row('CPU000FH01', 'BLACK', {'S': 6000}, 12)
    b = os_row('CPU000FH01A', 'BLACK', {'S': 6000}, 12)
    rows, rep = server.split_buysheet([a, b], flow)
    base = {r['split_crd']: r['sizes']['S'] for r in rows if r['style'] == 'CPU000FH01'}
    var = {r['split_crd']: r['sizes']['S'] for r in rows if r['style'] == 'CPU000FH01A'}
    assert base == {D[0]: 3000, D[1]: 3000} and var == {D[0]: 2400, D[1]: 2400, D[2]: 1200}, (base, var)
    assert len(rep['styles']) == 2 and not rep['notes'], rep['notes']


def chart_file(lines):
    """A delivery chart in the Offprice layout: header row with STYLE / CUSTOMER / QTY and the CRD dates across it, EST IN WHSE rows
    under it, then one row per style and customer."""
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(['TITLE'])
    ws.append([None] * 8 + ['SALES VERSION'])
    ws.append(['TRADING COMPANY', 'FACTORY', 'COO/PORT', 'DEST', 'CRD (SI CUT-OFF)', 'STYLE ', 'CUSTOMER', 'QTY ', 'PROPOSE SALES CRD', *D, None, 'TOTAL', 'COMMENTS'])
    ws.append([None] * 8 + ['EST IN WHSE -   35 DAYS', *[d + datetime.timedelta(35) for d in D]])
    for style, cust, qtys in lines:
        ws.append(['T', 'F', 'C', 'LA', 'THUR', style, cust, sum(qtys), None, *qtys, None, sum(qtys), 'OK'])
    ws.append(['TOTAL', None, None, None, None, None, None, 1])
    b = io.BytesIO(); wb.save(b)
    return b.getvalue()


def test_delivery_chart_layout_matches_style_and_customer():
    flow, notes = server.parse_flow_file(chart_file([
        ('RVM272SH03 ', 'MARSHALLS SJPA PPK', [7200, 4800, 0]),
        ('RVM272SH03 ', 'LINE SJPA PPK', [0, 6000, 0]),
        ('RVM272SH03 ', 'MARSHALLS TBCA PPK', [0, 0, 1200])]))
    assert len({id(e) for e in flow.values()}) == 3 and not notes, notes
    def row(cust, qty):
        r = os_row('RVM272SH03', 'BLACK', {'S': qty}, 12); r['customer'] = cust; return r
    rows, rep = server.split_buysheet([row('MARSHALLS SJPA PPK', 12000), row('LINE SJPA PPK', 6000), row('WINNERS SDBC PPK', 2400)], flow)
    mar = {r['split_crd']: r['sizes']['S'] for r in rows if r['customer'] == 'MARSHALLS SJPA PPK'}
    line = {r['split_crd']: r['sizes']['S'] for r in rows if r['customer'] == 'LINE SJPA PPK'}
    assert mar == {D[0]: 7200, D[1]: 4800} and line == {D[1]: 6000}, (mar, line)         # each customer split by its own chart line
    win = [r for r in rows if r['customer'] == 'WINNERS SDBC PPK']
    assert len(win) == 1 and 'split_crd' not in win[0]                                     # no chart line: kept as it is
    text = ' '.join(rep['notes'])
    assert 'MARSHALLS TBCA PPK' in text and 'WINNERS SDBC PPK' in text, rep['notes']        # both unmatched sides are reported


def test_chart_customers_must_match_the_buysheet_exactly():
    """No aliases and no fuzzy matching: TBCA on the chart is NOT read as SJCA. The preview refuses and says what to fix."""
    flow, _ = server.parse_flow_file(chart_file([('RVM272SH03 ', 'MARSHALLS TBCA PPK', [0, 1200, 0]), ('RVM272SH03 ', 'MARSHALLS SJPA PPK', [0, 0, 600])]))
    def row(cust):
        r = os_row('RVM272SH03', 'BLACK', {'S': 600}, 12); r['customer'] = cust; return r
    rows = [row('MARSHALLS SJCA PPK'), row('MARSHALLS SJPA PPK')]
    bad = server.flow_customer_mismatches(rows, flow)
    assert bad == [('MARSHALLS TBCA PPK', ['MARSHALLS SJCA PPK', 'MARSHALLS SJPA PPK'])], bad
    assert server.flow_customer_mismatches([row('MARSHALLS TBCA PPK'), row('MARSHALLS SJPA PPK')], flow) == []       # exact names: nothing to fix
    assert server.flow_customer_mismatches([row('marshalls  tbca ppk'), row('MARSHALLS SJPA PPK')], flow) == []      # case and extra spaces are ignored, nothing else


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
