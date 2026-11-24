"""The GENERIC CREATION sheet lists only what has to be created.  Run:  python tests/test_templates.py"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import openpyxl
import server


def os_rows(style, colors, sizes):
    return [{'type': 'OS', 'style': style, 'raw_style': style, 'color_code': c, 'color': c, 'sizes': {z: 100 for z in sizes}, 'desc': 'DESC ' + style,
             'pc_ctn_os': 12} for c in colors]


def gen(style, color, size, kind):
    return {'style': style, 'color_code': color, 'size': size, 'status': '✗ MISSING' if kind else '✓ EXISTS', 'missing_kind': kind}


def sheets(rows, generic, ppk=()):
    wb = openpyxl.load_workbook(server.make_creation_templates({'rows': rows, 'generic': generic, 'ppk': list(ppk)}))
    pick = lambda name: [(r[0], r[1], r[4]) for r in wb[name].iter_rows(min_row=2, values_only=True) if r[0]]
    assert wb.sheetnames == ['GENERIC CREATION', 'PPK CREATION', 'COLOR ADD', 'PC PER CTN', 'VARIANTS', 'PRICING'], wb.sheetnames       # colors added are the third tab
    return pick('GENERIC CREATION'), pick('COLOR ADD')


def test_existing_style_lists_only_the_new_color():
    rows = os_rows('X', ['A', 'B', 'C'], ['S', 'M', 'L'])
    generic = [gen('X', c, z, 'color' if c == 'C' else None) for c in 'ABC' for z in ('S', 'M', 'L')]
    assert sheets(rows, generic) == ([], [('X', 'C', 'S;M;L')])                       # on COLOR ADD, not on GENERIC CREATION, and not A;B;C


def test_new_style_lists_everything():
    rows = os_rows('Y', ['A', 'B'], ['S', 'M'])
    generic = [gen('Y', c, z, 'style') for c in 'AB' for z in ('S', 'M')]
    assert sheets(rows, generic) == ([('Y', 'A;B', 'S;M')], [])


def test_one_missing_size_lists_only_that_size_and_color():
    rows = os_rows('Z', ['A', 'B'], ['S', 'M', 'L'])
    generic = [gen('Z', c, z, 'size' if (c, z) == ('A', 'L') else None) for c in 'AB' for z in ('S', 'M', 'L')]
    assert sheets(rows, generic) == ([], [('Z', 'A', 'L')])


def test_colors_missing_different_sizes_are_not_crossed():
    rows = os_rows('W', ['A', 'B'], ['S', 'M', 'L'])
    generic = [gen('W', c, z, 'size' if (c, z) in (('A', 'L'), ('B', 'S')) else None) for c in 'AB' for z in ('S', 'M', 'L')]
    g, c = sheets(rows, generic)
    assert g == [] and sorted(c) == [('W', 'A', 'L'), ('W', 'B', 'S')]                 # not A;B x S;L


def test_a_color_a_prepack_needs_on_an_existing_style_is_listed():
    rows = [{'type': 'PPK', 'style': 'X2', 'raw_style': 'X2-Q', 'desc': 'PPK DESC', 'color': 'Q', 'color_code': 'Q'}]
    res = {'ppk_article': 'X2-Q', 'exists': '✗ MISSING', 'readiness': {'state': 'blocked', 'components': [],
           'needs': [{'kind': 'color', 'style': 'X2', 'color': 'Q', 'sizes': ['S', 'M']}]}}
    assert sheets(rows, [], [res]) == ([], [('X2', 'Q', 'S;M')])


def test_a_wrong_pc_per_ctn_is_listed_even_when_the_style_exists():
    import io
    rows = os_rows('V', ['A'], ['S', 'M'])
    generic = [dict(gen('V', 'A', z, None), sap_article='V' + z, pc_ctn_check=chk) for z, chk in (('S', '✗ SAP=10 BUY=12'), ('M', '✓'))]
    wb = openpyxl.load_workbook(server.make_creation_templates({'rows': rows, 'generic': generic, 'ppk': []}))
    pc = [r for r in wb['PC PER CTN'].iter_rows(min_row=2, values_only=True) if r[0]]
    assert [(r[0], r[1], r[2], r[3], r[4]) for r in pc] == [('V', 'A', 'S', 12, 'VS')], pc            # only the wrong size, with its SAP article
    assert wb['GENERIC CREATION'].max_row == 1 and wb['COLOR ADD'].max_row == 1                         # nothing to create


def test_a_blank_sap_price_is_listed_on_pricing_even_when_the_style_exists():
    rows = os_rows('U', ['A'], ['S', 'M'])
    generic = [dict(gen('U', 'A', z, None), sap_article='U' + z, price_check=chk, buy_price=9.5, sap_price=sp, sap_msrp=20, msrp_check='✓') for z, chk, sp in (('S', '⚠ NO SAP PRICE', None), ('M', '✓', 9.5))]
    ppk = [{'ppk_article': 'U-A', 'exists': '✓ EXISTS', 'price_check': '✗ SAP=5.00 BUY=6.00', 'sap_price': 5.0, 'sap_msrp': 20, 'msrp_check': '✓', 'color_code': 'A', 'buy_price': 6.0}]
    wb = openpyxl.load_workbook(server.make_creation_templates({'rows': rows, 'generic': generic, 'ppk': ppk}))
    pr = [r[:6] for r in wb['PRICING'].iter_rows(min_row=2, values_only=True) if r[0]]
    assert pr == [('US', 'U', 'A', 'S', 9.5, None), ('U-A', 'U', 'A', None, 6, None)], pr


def test_pricing_has_a_deduplicated_generic_table_beside_the_variant_table():
    rows = [dict(r, price=9.5, msrp=20) for r in os_rows('T', ['A', 'B'], ['S', 'M'])]
    generic = [gen('T', c, z, 'style') for c in 'AB' for z in ('S', 'M')]
    wb = openpyxl.load_workbook(server.make_creation_templates({'rows': rows, 'generic': generic, 'ppk': []}))
    ws = wb['PRICING']
    variants = [r[:9] for r in ws.iter_rows(min_row=2, max_col=9, values_only=True) if r[0]]
    assert len(variants) == 4 and ws.cell(1, 1).value == 'Article' and ws.cell(1, 9).value == 'PR00'             # the variant table: untouched
    assert ws.auto_filter.ref.startswith('A1:I'), ws.auto_filter.ref                                              # and its filter still covers A:I only
    head = [ws.cell(1, c).value for c in range(11, 14)]
    assert head == ['Material', 'SELL (ZSLS)', 'MSRP (ZVP1)'], head                                              # both prices
    body = [[ws.cell(r, c).value for c in range(11, 14)] for r in range(2, ws.max_row + 1) if ws.cell(r, 11).value]
    assert body == [['T', 9.5, 20]], body                                                                          # 4 variants, one generic line
    assert list(ws.tables) == ['GenericPricing']


def test_a_generic_priced_two_ways_gets_one_line_per_price():
    rows = os_rows('T2', ['A', 'B'], ['S'])
    rows[0].update(price=9.5, msrp=20); rows[1].update(price=10.5, msrp=20)
    generic = [gen('T2', c, 'S', 'style') for c in 'AB']
    ws = openpyxl.load_workbook(server.make_creation_templates({'rows': rows, 'generic': generic, 'ppk': []}))['PRICING']
    assert [[ws.cell(r, c).value for c in range(11, 14)] for r in range(2, 4)] == [['T2', 9.5, 20], ['T2', 10.5, 20]]


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
