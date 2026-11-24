"""PR upload: lines for the same customer, article, delivery date (and site/segment/season) are added into one line."""
import csv, io, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from sapportal import stock


class _Conn:
    def close(self): pass


def _csv(rows):
    stock.connect_hana = lambda: _Conn()
    stock.query_variant_articles = lambda conn, styles: {('RVM273FZ14', '001', 'S'): 'RVM273FZ14001', ('RVM273FZ14', '001', 'M'): 'RVM273FZ14002'}
    return list(csv.reader(io.StringIO(stock.make_req_csv(rows, [], None))))[1:]


def _row(customer, delivery, sizes):
    return {'type': 'OS', 'style': 'RVM273FZ14', 'color_code': '001', 'customer': customer, 'delivery': delivery,
            'sizes': sizes, 'warehouse': '', 'segment': 'MCE_GLOBAL', 'season_label': ''}


def test_same_customer_variant_and_date_become_one_line():
    out = _csv([_row('A', '2027-09-01', {'S': 10, 'M': 5}), _row('A', '2027-09-01', {'S': 16})])
    assert [(r[1], r[2]) for r in out] == [('RVM273FZ14001', '26'), ('RVM273FZ14002', '5')], out


def test_other_date_or_customer_stays_separate():
    out = _csv([_row('A', '2027-09-01', {'S': 10}), _row('A', '2027-08-01', {'S': 16}), _row('B', '2027-09-01', {'S': 3})])
    assert len(out) == 3 and [r[0] for r in out] == ['1', '1', '2'], out


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
