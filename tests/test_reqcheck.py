"""The requisition is checked against the buysheet (REQ# per customer from the HEADER sheet).  Run:  python tests/test_reqcheck.py"""
import datetime, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import server

REQ = '0010039354'
VARIANTS = {('X', 'BK', 'S'): 'X-BK-S', ('X', 'BK', 'M'): 'X-BK-M'}


def os_row(sizes, req=REQ, cust='TARGET HANG OS', deliv=datetime.date(2027, 4, 9), whse='RICA'):
    return {'type': 'OS', 'style': 'X', 'raw_style': 'X', 'color_code': 'BK', 'sizes': sizes, 'req': req, 'customer': cust, 'delivery': deliv, 'warehouse': whse}


def line(article, qty, item='10', deliv='20270409', stor='RICA', deleted=''):
    return {'EBAN_REQUISITION_NO': REQ, 'EBAN_REQUISITION_ITEM': item, 'EBAN_ARTICLE': article, 'F_REQ_QTY_PC': qty, 'EBAN_DELIVERY_DATE': deliv,
            'EBAN_STORAGE_LOCATION': stor, 'EBAN_DELETION_INDICATOR': deleted}


def pack_line(article, packs, per_pack=6, item='20'):
    d = line(article, packs * per_pack, item); d.update(EBAN_QTY_DUOM=packs, T006A_UNIT_COMMERCIAL='PAC'); return d


def check(rows, req_rows):
    return server.buysheet_req_check(req_rows, rows, {}, VARIANTS, 'bs.xlsm')


def test_everything_agrees():
    r = check([os_row({'S': 120, 'M': 240})], [line('X-BK-S', 120), line('X-BK-M', 240, '20')])
    assert r['issues'] == [] and r['reqs'][0]['customers'] == ['TARGET HANG OS'], r


def test_quantity_article_date_and_storage_differences_are_reported():
    r = check([os_row({'S': 120, 'M': 240})], [line('X-BK-S', 96, deliv='20270416', stor='ALCA'), line('X-BK-L', 12, '30')])
    msgs = sorted((i['level'], i['article'], i['msg'][:12]) for i in r['issues'])
    assert ('error', 'X-BK-S', 'Quantity: bu') in msgs and ('warn', 'X-BK-S', 'Delivery dat') in msgs and ('warn', 'X-BK-S', 'Storage loca') in msgs, msgs
    assert ('error', 'X-BK-M', 'On the buysh') in msgs and ('error', 'X-BK-L', 'On the requi') in msgs, msgs


def test_a_deleted_requisition_line_does_not_count():
    r = check([os_row({'S': 120})], [line('X-BK-S', 120), line('X-BK-M', 50, '20', deleted='X')])
    assert r['issues'] == [], r['issues']


def test_requisition_not_on_the_header_sheet():
    r = check([os_row({'S': 120}, req='')], [line('X-BK-S', 120)])
    assert r['errors'] == 1 and 'HEADER' in r['issues'][0]['msg'], r


def test_other_customers_requisitions_are_not_mixed_in():
    r = check([os_row({'S': 120}), os_row({'S': 999}, req='0010039999', cust='OTHER')], [line('X-BK-S', 120)])
    assert r['issues'] == [], r['issues']


def test_no_buysheet_loaded_says_so():
    r = check([], [line('X-BK-S', 120)])
    assert r['issues'] == [] and 'PREBUY' in r['notes'][0], r


def ppk_row(pieces, per_pack=6, complex_=False, color='A', row_pc=None):
    return {'type': 'PPK', 'style': 'P', 'raw_style': 'P-S1', 'sizes': {}, 'ppk_qty': pieces, 'pc_per_pac': per_pack, 'req': REQ, 'customer': 'TARGET HANG PPK',
            'delivery': datetime.date(2027, 4, 9), 'warehouse': 'RICA', 'complex': complex_, 'row_pc': row_pc, 'color': color}


def test_prepack_requisition_is_in_packs_open_stock_in_pieces():
    rows = [os_row({'S': 120}), ppk_row(600)]                                  # 600 pc at 6 per pack = 100 packs
    assert check(rows, [line('X-BK-S', 120), pack_line('P-S1', 100)])['issues'] == []
    r = check(rows, [line('X-BK-S', 120), pack_line('P-S1', 600, per_pack=1)])           # the pieces typed on a pack requisition is a mismatch
    assert any(i['article'] == 'P-S1' and 'buysheet 100 packs vs requisition 600 packs' in i['msg'] for i in r['issues']), r['issues']


def test_complex_prepack_colors_are_one_order():
    rows = [ppk_row(150, complex_=True, color='A', row_pc=15), ppk_row(90, complex_=True, color='B', row_pc=9)]     # 10 packs, both colors
    assert check(rows, [pack_line('P-S1', 10, 24)])['issues'] == []


def test_complex_colors_that_disagree_on_packs_are_warned():
    rows = [ppk_row(150, complex_=True, color='A', row_pc=15), ppk_row(99, complex_=True, color='B', row_pc=9)]      # 10 packs vs 11
    r = check(rows, [pack_line('P-S1', 10, 24)])
    assert [i['level'] for i in r['issues']] == ['warn'] and 'different pack counts (10, 11)' in r['issues'][0]['msg'], r['issues']


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
