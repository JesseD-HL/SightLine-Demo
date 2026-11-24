"""Regression checks for the prebuy parsers in server.py: one small workbook per layout, asserting what the parser returns.
Run from the SAP TOOLS folder:   python tests/test_parsers.py
Add a case whenever a prebuy layout is fixed or a new one appears, BEFORE changing the parsers."""
import io, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import openpyxl
import server


def parse(sheets):
    """sheets: {title: [row, row, ...]} -> parse_prebuy(...) rows."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for r, row in enumerate(rows, 1):
            for c, v in enumerate(row, 1):
                if v is not None:
                    ws.cell(r, c, v)
    buf = io.BytesIO()
    wb.save(buf)
    rows, customers, kind = server.parse_prebuy(buf.getvalue())
    return rows, customers, kind


HDR_BLOCK = [[None, 'CUSTOMER:'], [None, 'BULKCO'], [], []]


def test_complex_tab_is_one_prepack():
    rows, _, _ = parse({'COMPLEX': HDR_BLOCK[:1] + [[None, 'BULKCO']] + [[]] + [
        ['STYLE', 'COLOR CODE', 'COLOR DESCRIPTION', 'DESCRIPTION', 'S', 'M', 'L', 'TOTAL UNITS', 'PRICE', 'SM', 'MD', 'LG'],
        ['SFM273CR01-A', None, None, 'CREW', 740 * 4, 740 * 7, 740 * 10, None, 13.55, 4, 7, 10],
        ['SFM273CR01', 'S143', 'BLACK', 'CREW', 740 * 2, 740 * 3, 740 * 4, None, None, 2, 3, 4],
        ['SFM273CR01', 'RI39', 'OXFORD', 'CREW', 740 * 2, 740 * 4, 740 * 6, None, None, 2, 4, 6],
    ]})
    assert len(rows) == 2, rows
    assert {r['raw_style'] for r in rows} == {'SFM273CR01-A'}
    assert all(r['complex'] and r['pc_per_pac'] == 21 for r in rows)
    assert [r['row_pc'] for r in rows] == [9, 12]
    assert not any(r['qty_conflict'] for r in rows)          # color ratios add up to the pack's: no warning


def test_complex_tab_warns_when_ratios_disagree():
    rows, _, _ = parse({'COMPLEX': [[None, 'CUSTOMER:'], [None, 'BULKCO'], [], [],
        ['STYLE', 'COLOR CODE', 'COLOR DESCRIPTION', 'DESCRIPTION', 'S', 'M', 'PRICE', 'SM', 'MD'],
        ['SFM273CR01-A', None, None, 'CREW', 600, 900, 13.55, 4, 6],
        ['SFM273CR01', 'S143', 'BLACK', 'CREW', 200, 300, None, 2, 3],
        ['SFM273CR01', 'RI39', 'OXFORD', 'CREW', 200, 300, None, 1, 3],      # S: 2+1 != 4
    ]})
    assert len(rows) == 2 and all(r['qty_conflict'] for r in rows), rows


def test_prepack_tab_ratio_from_units():
    rows, customers, kind = parse({'PREPACK': HDR_BLOCK + [
        ['STYLE', 'COLOR CODE', 'COLOR DESCRIPTION', 'DESCRIPTION', 'S', 'M', 'L', 'PRICE'],
        ['SFM273CR01', 'S143', 'BLACK', 'CREW', 740, 1110, 1480, 13.55],
    ]})
    assert kind == 'prebuy' and len(rows) == 1
    r = rows[0]
    assert r['raw_style'] == 'SFM273CR01-S143' and r['ratio'] == '2-3-4' and r['ppk_qty'] == 370 and r['pc_per_pac'] == 9
    assert r['pack_kind'] == 'PPK'


def test_open_stock_tab():
    rows, _, _ = parse({'OPEN STOCK': HDR_BLOCK + [
        ['STYLE', 'COLOR CODE', 'COLOR DESCRIPTION', 'DESCRIPTION', 'DELIVERY', 'S', 'M', 'L'],
        ['SFM273CR01', 'S143', 'BLACK', 'CREW', None, 10, 20, 30],
    ]})
    assert len(rows) == 1 and rows[0]['type'] == 'OS' and rows[0]['sizes'] == {'S': 10, 'M': 20, 'L': 30}


def test_simple_ppk_column_tab():
    rows, _, _ = parse({'SHEET1': HDR_BLOCK + [
        ['STYLE', 'COLOR CODE', 'COLOR DESCRIPTION', 'DESCRIPTION', 'PPK', 'PRICE'],
        ['SFM273CR01-S143', 'S143', 'BLACK', 'CREW', 50, 13.55],
    ]})
    assert len(rows) == 1 and rows[0]['type'] == 'PPK' and rows[0]['ppk_qty'] == 50


def test_inline_complex_and_plain_pack_on_one_prepack_tab():
    """The real prebuy form: a shaded -A row with its colors below, on the PREPACK tab, next to an ordinary one-color pack."""
    rows, _, _ = parse({'PREPACK TEMPLATE': [[None, 'CUSTOMER:'], [None, 'BULKCO'], [None, None, 'PREPACK RATIO'], [],
        ['STYLE', 'COLOR CODE', 'COLOR DESCRIPTION', 'DESCRIPTION', 'S', 'M', 'PRICE', 'SM', 'MD'],
        ['SFM273CR01-A', None, None, 'CREW', 600, 900, 13.55, 2, 3],
        ['SFM273CR01', 'S143', 'BLACK', 'CREW', 400, 600, None, 4, 6],
        ['SFM273CR01', 'RI39', 'OXFORD', 'CREW', 200, 300, None, 2, 3],
        ['SFM273PA16', 'S143', 'BLACK', 'JOGGER', 100, 200, 9.0, None, None],
    ]})
    cx = [r for r in rows if r.get('complex')]
    plain = [r for r in rows if not r.get('complex')]
    assert len(cx) == 2 and {r['raw_style'] for r in cx} == {'SFM273CR01-A'}, rows
    assert len(plain) == 1 and plain[0]['raw_style'] == 'SFM273PA16-S143', plain


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:300])
    sys.exit(1 if failed else 0)
