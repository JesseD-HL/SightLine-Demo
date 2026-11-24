"""DPCI list: one file, two jobs (VD52 DPCI check and what to create in SAP).  Run:  python tests/test_dpci_list.py"""
import io, os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import openpyxl
from sapportal import parsers, dpci_list, stock

HEAD = ['Generic Manufacturer Style', 'Component DPCI', 'Description', 'COLOR', 'SIZE']


def _wb(rows):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sheet1'
    ws.append(HEAD)
    for r in rows: ws.append(r)
    buf = io.BytesIO(); wb.save(buf)
    return openpyxl.load_workbook(io.BytesIO(buf.getvalue()), data_only=True)


def test_the_list_is_recognised_and_read_as_creation_only_rows():
    wb = _wb([['CPU000FH01', '043-23-2077', 'HOODIE ', ' DELICIOSO ', 'XS'], ['CPU000FH01', '043-23-2078', 'HOODIE', 'DELICIOSO', ' S ']])
    assert parsers.is_dpci_list(wb)
    rows, customers = parsers.parse_dpci_list(wb)
    assert len(rows) == 1 and rows[0]['creation_only'] and rows[0]['color'] == 'DELICIOSO' and list(rows[0]['sizes']) == ['XS', 'S']
    assert rows[0]['dpci_by_size'] == {'XS': '043-23-2077', 'S': '043-23-2078'} and rows[0]['desc'] == 'HOODIE'


def test_stray_spaces_are_trimmed_but_case_is_kept_as_a_real_difference():
    rows, _ = parsers.parse_dpci_list(_wb([['CPU000FH01', '043-23-2077', 'X', ' Black ', 'M'], ['CPU000FH01', '043-23-2078', 'X', 'BLACK', 'M']]))
    assert [r['color_exact'] for r in rows] == ['Black', 'BLACK']        # two colors: SAP keeps them apart too


def test_the_list_is_found_when_another_sheet_was_active_at_save_time():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sheet1'
    ws.append(HEAD); ws.append(['CPU000FH01', '043-23-2077', 'X', 'Black', 'S'])
    wb.create_sheet('Sheet2').append(['MATERIAL', 'DESCRIPTION', 'CUST MATERIAL', 'DPCI'])        # the VD52 layout saved beside the list
    wb.active = 1
    buf = io.BytesIO(); wb.save(buf)
    wb2 = openpyxl.load_workbook(io.BytesIO(buf.getvalue()), data_only=True)
    assert wb2.active.title == 'Sheet2' and parsers.is_dpci_list(wb2) and len(parsers.parse_dpci_list(wb2)[0]) == 1


def test_a_buysheet_is_not_mistaken_for_a_dpci_list():
    wb = openpyxl.Workbook(); wb.active.title = 'DATA'; wb.active.append(HEAD)
    assert not parsers.is_dpci_list(wb)


def _generic(size, exists=True, article='A1', kind='', code='001'):
    return {'style': 'S1', 'buy_color': 'BLACK', 'color_code': code, 'size': size, 'status': '✓ EXISTS' if exists else '✗ MISSING',
            'sap_article': article if exists else '', 'missing_kind': kind}


def test_each_row_is_ok_differs_missing_or_to_create():
    rows = [{'style': 'S1', 'color': 'BLACK', 'dpci_by_size': {'S': '043-23-2077', 'M': '043-23-2078', 'L': '043-23-2079', 'XL': '043-23-2080'}}]
    validation = {'generic': [_generic('S', article='A1'), _generic('M', article='A2'), _generic('L', article='A3'), _generic('XL', exists=False, kind='size')]}
    out = dpci_list.check_rows(validation, rows, lambda arts: {'A1': {'043232077'}, 'A2': {'043230001'}})
    assert [r['Status'] for r in out] == ['OK', 'ERROR', 'REVIEW', 'CREATE'], out
    assert 'Create in SAP' in out[3]['Action'] and 'size' in out[3]['Action']


def test_a_row_without_a_color_code_is_never_ok():
    rows = [{'style': 'S1', 'color': 'BLACK', 'dpci_by_size': {'S': '043-23-2077'}}]
    out = dpci_list.check_rows({'generic': [_generic('S', article='A1', code='')]}, rows, lambda arts: {'A1': {'043232077'}})
    assert out[0]['Status'] == 'REVIEW' and 'color code' in out[0]['Action']


def test_a_row_right_in_sap_stays_ok_and_is_not_uploaded_when_itemhub_lacks_the_dpci():
    r = _row_check({}, vd52={'A1': {'043232077'}})
    assert r['Status'] == 'OK' and r['ItemHub'] == '✗ not found' and r['VD52 need'] is False
    assert dpci_list.upload_rows([r]) == []


def test_the_vd52_upload_has_material_a_blank_column_dpci_with_two_spaces_and_the_upc():
    res = [{'SAP article': 'A1', 'List DPCI': '043-23-2077', 'UPC': '199071176575', 'Status': 'REVIEW', 'VD52 need': True},     # no DPCI in VD52: enter it
           {'SAP article': 'A2', 'List DPCI': '043-23-2078', 'UPC': '199071176582', 'Status': 'ERROR', 'VD52 need': True},      # a different DPCI: correct it
           {'SAP article': 'A3', 'List DPCI': '043-23-2079', 'UPC': '199071176599', 'Status': 'OK', 'VD52 need': False},        # already right: left out
           {'SAP article': '', 'List DPCI': '043-23-2080', 'UPC': '', 'Status': 'CREATE', 'VD52 need': False}]                  # no article yet: left out
    assert dpci_list.upload_rows(res) == [['A1', None, '043  23  2077', '199071176575'], ['A2', None, '043  23  2078', '199071176582']]
    assert dpci_list.vd52_dpci('043-23-2077') == '043  23  2077' and dpci_list.vd52_dpci('043 - 23 - 2077') == '043  23  2077'


def _lp_item(style='S1', size='S', gtin='00199071176575'):
    from sapportal.routes import itemhub_module
    return {'styles': [itemhub_module().key_style(style)], 'gtins': [gtin], 'lp_size': size}


def _row_check(lp, upc='199071176575', vd52=None):
    rows = [{'style': 'S1', 'color': 'BLACK', 'dpci_by_size': {'S': '043-23-2077'}}]
    validation = {'generic': [_generic('S', article='A1')]}
    return dpci_list.check_rows(validation, rows, lambda arts: vd52 if vd52 is not None else {}, lambda arts: {'A1': upc}, lp)[0]


def test_a_dpci_itemhub_does_not_have_yet_is_flagged_but_still_uploaded():
    r = _row_check({})
    assert r['ItemHub'] == '✗ not found' and r['LP ok'] is None and r['Status'] == 'REVIEW'
    assert len(dpci_list.upload_rows([r])) == 1             # ItemHub cannot contradict a DPCI it does not have


def test_without_a_itemhub_pull_the_itemhub_columns_stay_empty_and_the_row_is_still_uploaded():
    r = _row_check(None)
    assert r['ItemHub'] == '' and r['LP ok'] is None and len(dpci_list.upload_rows([r])) == 1


def test_the_creation_templates_are_built_from_the_buysheet_and_the_dpci_list_together():
    from sapportal import routes
    from sapportal.config import _cache
    saved = {k: _cache.get(k) for k in ('validation', 'dpci')}
    try:
        _cache['validation'] = {'rows': [{'style': 'BUY1'}], 'generic': [{'style': 'BUY1', 'size': 'S'}], 'ppk': [{'p': 1}], 'filename': 'buysheet.xlsm'}
        _cache['dpci'] = {'validation': {'rows': [{'style': 'DP1'}], 'generic': [{'style': 'DP1', 'size': 'M'}], 'ppk': [], 'filename': 'list.xlsx'}, 'rows': []}
        both = routes.creation_source()
        assert [r['style'] for r in both['rows']] == ['BUY1', 'DP1'] and [g['style'] for g in both['generic']] == ['BUY1', 'DP1'] and both['ppk'] == [{'p': 1}]
        assert _cache['validation']['filename'] == 'buysheet.xlsm' and len(_cache['validation']['rows']) == 1          # the buysheet itself is untouched
        _cache['validation'] = None
        assert routes.creation_source()['filename'] == 'list.xlsx'                                                       # the list alone is enough
        _cache['dpci'] = None
        assert routes.creation_source() is None
    finally:
        _cache.update(saved)


def test_a_file_that_is_not_a_dpci_list_is_refused_and_nothing_is_loaded():
    from sapportal import routes
    from sapportal.config import _cache
    wb = openpyxl.Workbook(); wb.active.append(['not', 'a', 'list'])
    buf = io.BytesIO(); wb.save(buf)
    before = _cache.get('dpci')
    try:
        routes.load_dpci_list(buf.getvalue(), 'other.xlsx')
        assert False, 'should have been refused'
    except ValueError as exc:
        assert 'not a DPCI list' in str(exc)
    assert _cache.get('dpci') is before


def test_the_exported_list_has_every_row_with_its_upc_even_when_all_are_right():
    wb = _wb([['CPU000FH01', '043-23-2077', 'HOODIE ONE', ' Black ', 'XS'], ['CPU000FH01', '043-23-2078', 'HOODIE TWO', ' Black ', 'S']])
    rows, _ = parsers.parse_dpci_list(wb)
    results = [{'Style': 'CPU000FH01', 'Color': 'BLACK', 'Size': 'XS', 'List DPCI': '043-23-2077', 'UPC': '709006100195', 'SAP article': 'A1', 'Color code': '001', 'Status': 'OK', 'ItemHub': ''},
               {'Style': 'CPU000FH01', 'Color': 'BLACK', 'Size': 'S', 'List DPCI': '043-23-2078', 'UPC': '709006100201', 'SAP article': 'A2', 'Color code': '001', 'Status': 'OK', 'ItemHub': '✓'}]
    out = dpci_list.list_rows(results, rows)
    assert out == [['CPU000FH01', '043-23-2077', 'HOODIE ONE', 'Black', 'XS', '709006100195', 'A1', '001', 'OK', ''],
                   ['CPU000FH01', '043-23-2078', 'HOODIE TWO', 'Black', 'S', '709006100201', 'A2', '001', 'OK', '✓']], out     # each row's own description, the color as typed


def test_no_requisition_is_written_from_a_dpci_list():
    stock.connect_hana = lambda: type('C', (), {'close': lambda self: None})()
    stock.query_variant_articles = lambda conn, styles: {('S1', '001', 'S'): 'S1001'}
    row = {'type': 'OS', 'style': 'S1', 'color_code': '001', 'customer': '', 'delivery': None, 'sizes': {'S': 1}, 'creation_only': True}
    assert stock.make_req_csv([row], [], None).strip().count('\n') == 0     # the header line only


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
