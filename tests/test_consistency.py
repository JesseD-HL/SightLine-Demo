"""A style with more than one price, MSRP or fineline in the same file is reported.  Run:  python tests/test_consistency.py"""
import os, sys, unittest
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..'))
import server


def row(art, price=None, msrp=None, fineline=None, cust='A', typ='PPK'):
    return {'type': typ, 'raw_style': art, 'style': art.split('-')[0], 'customer': cust, 'price': price, 'msrp': msrp, 'fineline': fineline}


def test_nothing_reported_when_every_row_agrees():
    rows = [row('X-S1', 8.79, 19.97, '105'), row('X-S2', 8.79, 19.97, '105'), row('Y-S1', 9, 12, '3')]
    assert server.check_consistency(rows) == []


def test_two_prices_on_one_generic_are_reported_with_where():
    rows = [row('X-S1', 8.79, 19.97, '105', cust='M1'), row('X-S2', 9.5, 19.97, '105', cust='M2'), row('X-S3', 9.5, 19.97, '105', cust='M2')]
    out = server.check_consistency(rows)
    assert [(c['style'], c['kind']) for c in out] == [('X', 'price')], out
    vals = {v['value']: v for v in out[0]['values']}
    assert vals[8.79]['rows'] == 1 and vals[9.5]['rows'] == 2 and vals[9.5]['customers'] == ['M2'] and vals[9.5]['articles'] == ['X-S2', 'X-S3'], vals


def test_a_prepack_counts_under_its_generic_and_finelines_are_checked_too():
    rows = [row('X', 5, 10, '105', typ='OS'), row('X-S1', 5, 12, '106')]               # open stock and a prepack of the same generic
    assert sorted((c['kind'], [v['value'] for v in c['values']]) for c in server.check_consistency(rows)) == [('fineline', ['105', '106']), ('msrp', [10.0, 12.0])]


def test_cad_and_usd_columns_are_compared_separately():
    r1 = dict(row('X-S1', 9.5, 19.97, '105'), price_cad=9.5, price_usd=6.94)
    r2 = dict(row('X-S2', 9.5, 19.97, '105'), price_cad=9.5, price_usd=6.42)
    assert [(c['kind'], [v['value'] for v in c['values']]) for c in server.check_consistency([r1, r2])] == [('price_usd', [6.42, 6.94])]


def test_blank_values_are_not_a_second_value():
    rows = [row('X-S1', 5, None, ''), row('X-S2', 5, 10, '105')]
    assert server.check_consistency(rows) == []


def test_real_bigbox_men_buysheet_finds_rem231ts16cb():
    path = os.path.join(HERE, '..', 'samples', 'Bigbox CA DOMESTIC - Ridgeline Men - Q2 2027 Buysheet 11.13.xlsm')
    if not os.path.exists(path): raise unittest.SkipTest('real sample buysheets are not in git (samples/)')
    rows, _, _ = server.parse_prebuy(open(path, 'rb').read())
    found = {c['kind']: sorted(v['value'] for v in c['values']) for c in server.check_consistency(rows) if c['style'] == 'REM231TS16CB'}
    assert found.get('price_usd') == [6.42, 6.94] and 'price_cad' not in found, found       # CAD is 9.50 everywhere; the USD column is stale on some rows


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except unittest.SkipTest as e:
            print('skip  ', name, '-', e)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
