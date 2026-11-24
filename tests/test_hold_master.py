"""Hold master: the stacked hold order entry files become open-stock rows the normal checks run on."""
import io, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from openpyxl import Workbook
from sapportal.parsers import parse_prebuy

HEAD = ['Division', 'Customer', 'Generic', 'Description', 'Color Code', 'Color Desc', 'Size', 'Qty', 'Segment', 'Delivery', 'Price', 'MSRP', 'Source file']


def _file(lines):
    wb = Workbook(); ws = wb.active; ws.title = 'MASTER'; ws.append(HEAD)
    for l in lines: ws.append(l)
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def test_sizes_go_across_one_row_per_customer_style_color():
    rows, customers, kind = parse_prebuy(_file([
        ['Mens', 'BIRCH', 'SLM273JK02', 'LETTERMAN', 'S00FR', 'Chocolate Chip', 'S', 10, 'OTH', '2027-07-01', 28, 80, 'a.xlsx'],
        ['Mens', 'BIRCH', 'SLM273JK02', 'LETTERMAN', 'S00FR', 'Chocolate Chip', 'M', 5, 'OTH', '2027-07-01', 28, 80, 'a.xlsx'],
        ['Mens', 'BASEEX', 'SLM273JK02', 'LETTERMAN', 'S00FR', 'Chocolate Chip', 'S', 7, 'OTH', '2027-07-01', 28, 80, 'b.xlsx']]))
    assert kind == 'hold_master' and customers == ['BASEEX', 'BIRCH']
    birch = next(r for r in rows if r['customer'] == 'BIRCH')
    assert birch['sizes'] == {'S': 10, 'M': 5} and birch['color_code'] == 'S00FR' and birch['type'] == 'OS' and len(rows) == 2


def test_prepack_lines_and_empty_quantities_are_skipped_not_checked():
    rows, _, _ = parse_prebuy(_file([
        ['Mens', "DALTON'S", 'SLM273PA01', 'PANT', 'SF23', 'Grey', 'PPK', 944, 'ECM', None, 15.75, 45, 'c.xlsx'],
        ['Mens', 'BIRCH', 'SLM273JK02', 'X', 'S00FR', 'Choc', 'S', 0, 'OTH', None, 28, 80, 'a.xlsx']]))
    assert rows == []


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
