"""Sam's prebuy: every block (ROLLOUT, PACKXSS, PACKM, PACKL, PACKXLXXL) is ONE complex prepack of four colors; packs = the block's
units / pieces per pack; each color's quantity = packs x its own ratio.  Run:  python tests/test_sams.py"""
import os, sys, unittest
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..'))
import server

SAMPLE = os.path.join(HERE, '..', 'samples', "Prebuy - Clubmart Q3 '27 Summit Women- Prepacks 11.17.xlsx")


def blocks():
    if not os.path.exists(SAMPLE): raise unittest.SkipTest('real sample buysheets are not in git (samples/)')
    rows, customers, kind = server.parse_prebuy(open(SAMPLE, 'rb').read())
    out = {}
    for r in rows: out.setdefault(r['raw_style'], []).append(r)
    return rows, out


def test_ten_complex_packs_of_four_colors():
    rows, b = blocks()
    assert len(b) == 10 and all(len(v) == 4 and all(x['complex'] for x in v) for v in b.values()), {k: len(v) for k, v in b.items()}
    assert sorted(b) == sorted(f'{s}-{k}' for s in ('SFW273QZ08', 'SFW273PA31') for k in ('ROLL', 'XSS', 'M', 'L', 'XLXXL')), sorted(b)   # the names typed on the color rows


def test_packs_are_the_block_units_over_the_pack_size():
    _, b = blocks()
    want = {'SFW273QZ08-ROLL': (36, 5587), 'SFW273QZ08-XSS': (36, 216), 'SFW273QZ08-M': (36, 236), 'SFW273QZ08-XLXXL': (36, 254),
            'SFW273PA31-ROLL': (30, 7314), 'SFW273PA31-XSS': (30, 601), 'SFW273PA31-L': (30, 516)}
    for art, (pc, packs) in want.items():
        v = b[art]
        assert v[0]['pc_per_pac'] == pc and {x['ppk_qty'] / x['row_pc'] for x in v} == {packs}, (art, v[0]['pc_per_pac'], [x['ppk_qty'] for x in v])


def test_color_quantity_follows_its_own_ratio():
    _, b = blocks()
    black = next(x for x in b['SFW273QZ08-M'] if x['color_code'] == 'S143')      # 9 of the 36 pieces in the M pack are black
    assert black['ppk_qty'] == 236 * 9 and black['sizes'] == {'M': 9}, black


def test_whole_buy_adds_up_to_the_file_total():
    rows, _ = blocks()
    assert sum(r['ppk_qty'] for r in rows) == 518934


def test_buysheet_defaults_pac_per_ctn_to_one_for_prepacks():
    import openpyxl
    rows, _ = blocks()
    buf = server.make_data_sheet(rows, ["SAM'S CLUB"])
    ws = openpyxl.load_workbook(buf)['DATA']
    hi = next(i for i, r in enumerate(ws.iter_rows(max_row=5, values_only=True), 1) if r[0] == 'CUSTOMER')
    head = [c.value for c in ws[hi]]
    vals = {r[head.index('PAC PER CTN')] for r in ws.iter_rows(min_row=hi + 1, values_only=True) if r[0]}
    assert vals == {1}, vals
    rows[0]['pac_per_ctn'] = 6                                                  # a value typed on the prebuy wins
    ws = openpyxl.load_workbook(server.make_data_sheet(rows, ["SAM'S CLUB"]))['DATA']
    assert next(r for r in ws.iter_rows(min_row=hi + 1, values_only=True) if r[0])[head.index('PAC PER CTN')] == 6


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
