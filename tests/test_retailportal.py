"""RetailPortal tab: the page's table is parsed, and each UPC is checked against SAP (UPC on an article, same color)."""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from sapportal import retailportal as so

HEADS = ['GTIN/UPC', 'Product name', 'Available channels', 'Color', 'Is variant', 'Item ID']


def test_table_cells_are_matched_to_headers_even_with_a_leading_checkbox_cell():
    rows = [['', '00199071176575', 'Tee', '1P', 'Cabernet', 'Yes', '16453614386'], ['', '', '', '', '', '', '']]
    assert so.parse_table(HEADS, rows) == [{'gtin': '00199071176575', 'name': 'Tee', 'color': 'Cabernet'}]


def test_gtin_forms_cover_upc_and_ean():
    forms = so.gtin_forms('00199071176575')
    assert '199071176575' in forms and '0199071176575' in forms


def _item(gtin='00199071176575', color='Cabernet'):
    return {'gtin': gtin, 'name': 'Tee', 'color': color}


def test_matching_upc_and_color_is_ok():
    sap = {'199071176575': [('REM241TE14093', 'REM241TE14', 'CABERNET', 'S')]}
    r = so.compare([_item()], sap)[0]
    assert r['Status'] == 'OK' and r['Article'] == 'REM241TE14093'


def test_a_different_color_is_an_error():
    sap = {'199071176575': [('REM241TE14093', 'REM241TE14', 'COBALT', 'S')]}
    r = so.compare([_item()], sap)[0]
    assert r['Status'] == 'ERROR' and 'Color differs' in r['Details']


def test_an_upc_sap_does_not_have_is_an_error_and_two_articles_need_review():
    assert so.compare([_item()], {})[0]['Status'] == 'ERROR'
    sap = {'199071176575': [('A1', 'S', 'CABERNET', 'S'), ('A2', 'S', 'CABERNET', 'S')]}
    assert so.compare([_item()], sap)[0]['Status'] == 'REVIEW'


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
