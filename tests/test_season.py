"""A style with 000 where the year and quarter sit (CPU000FH01) carries no season: season 0000, year 9999.  Run:  python tests/test_season.py"""
import os, sys, io
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import openpyxl
import server


def test_season_text_from_style():
    assert server.season_text('', 'SFM273CR01') == '2027 FALL'                      # unchanged for an ordinary style
    assert server.season_text('', 'CPU000FH01') == '9999 0000'
    assert server.season_text('', 'CPU000FH01A') == '9999 0000'
    assert server.season_text('Q2 2027', 'CPU000FH01') == '2027 SUM'                 # an explicit header season still wins in the PR file


def test_buysheet_label_is_not_read_from_a_000_style():
    assert server.season_from_style('CPU000FH01') == ''
    assert server.season_from_style('SFM273CR01') == 'Q3-2027'


def _validation(article):
    comp = {'color': 'S143', 'size': 'S', 'qty': 2, 'status': 'ok', 'article': article + 'S'}
    return {'rows': [{'type': 'PPK', 'raw_style': article, 'style': article.split('-')[0], 'desc': 'X', 'color': '', 'pac_per_ctn': 1}],
            'generic': [], 'ppk': [{'ppk_article': article, 'exists': '✗ MISSING',
                                    'readiness': {'state': 'ready', 'needs': [], 'mdse_cat': 'M', 'components': [comp]}}]}


def test_creation_template_formulas_default_for_000_styles():
    wb = openpyxl.load_workbook(server.make_creation_templates(_validation('CPU000FH01A-S143')))
    ws = wb['PPK CREATION']
    head = [c.value for c in ws[1]]
    season, year = ws.cell(2, head.index('Season') + 1).value, ws.cell(2, head.index('Year') + 1).value
    assert 'MID(A2,4,3)="000"' in season and '"0000"' in season and 'CHOOSE(MID(A2,6,1)' in season, season
    assert 'MID(A2,4,3)="000"' in year and '"9999"' in year and '"20"&MID(A2,4,2)' in year, year


def test_ppk_run_file_writes_literal_defaults(tmp_path=None):
    import tempfile
    path = os.path.join(tempfile.mkdtemp(), 'run.xlsx')
    for art, want in (('CPU000FH01A-S143', ('0000', '9999')), ('SFM273CR01-S143', ('FALL', '2027'))):
        ready, _ = server.make_ppk_run_file(_validation(art), path)
        ws = openpyxl.load_workbook(path)['PPK CREATION']
        head = [c.value for c in ws[1]]
        got = (str(ws.cell(2, head.index('Season') + 1).value), str(ws.cell(2, head.index('Year') + 1).value))
        assert ready == [art] and got == want, (art, got)


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
