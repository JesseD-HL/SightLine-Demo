"""The buysheet share is production: it is listed and copied from, never written to.  Run:  python tests/test_buysheets.py"""
import os, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import server
from sapportal import buysheets


def fake_share():
    root = tempfile.mkdtemp()
    os.makedirs(os.path.join(root, '2027', 'Q1 2027'))
    for name in ('A Q1 Buysheet.xlsm', '~$A Q1 Buysheet.xlsm', 'notes.txt'):
        with open(os.path.join(root, '2027', 'Q1 2027', name), 'wb') as f:
            f.write(b'data')
    return root


def snapshot(root):
    return sorted((os.path.join(d, n), os.path.getmtime(os.path.join(d, n)), os.path.getsize(os.path.join(d, n)))
                  for d, _, names in os.walk(root) for n in names)


def test_lists_only_excel_buysheets_and_skips_lock_files():
    buysheets.BUYSHEET_ROOT = fake_share()
    out = buysheets.list_buysheets()
    assert out['year'] == '2027' and [f['name'] for q in out['quarters'] for f in q['files']] == ['A Q1 Buysheet.xlsm']


def test_reading_leaves_the_share_untouched_and_saves_no_copy():
    buysheets.BUYSHEET_ROOT = root = fake_share()
    before, files_before = snapshot(root), set(os.listdir(os.path.join(buysheets.HERE, 'sapportal')))
    assert buysheets.read_from_share('2027/Q1 2027/A Q1 Buysheet.xlsm') == b'data'
    assert snapshot(root) == before and not hasattr(buysheets, 'COPY_DIR')


def test_a_path_outside_the_share_is_refused():
    buysheets.BUYSHEET_ROOT = fake_share()
    for bad in ('../../etc/passwd', '2027/../../x.xlsx', 'C:/Windows/win.ini'):
        try:
            buysheets.read_from_share(bad)
        except (ValueError, FileNotFoundError):
            continue
        raise AssertionError(bad)


def test_reqs_are_read_from_the_header_sheet_and_remembered():
    import openpyxl
    buysheets.BUYSHEET_ROOT = root = fake_share()
    buysheets.REQ_CACHE = os.path.join(tempfile.mkdtemp(), 'reqs.json')
    buysheets._req_cache.clear()
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = 'HEADER'
    ws.append(['CUSTOMER', 'REQ#']); ws.append(['ACME FOLD', 10040113]); ws.append(['ACME HANG', None]); ws.append(['BETA', '0010040099'])
    data = wb.create_sheet('DATA'); data.append(['STYLE']); data.append(['X'])
    wb.save(os.path.join(root, '2027', 'Q1 2027', 'B Q1 Buysheet.xlsm'))
    rel = '2027/Q1 2027/B Q1 Buysheet.xlsm'
    before = snapshot(root)
    got = buysheets.req_status(rel)
    assert got['customers'] == 3 and got['created'] == 2 and got['lines'] == {'ACME FOLD': '0010040113', 'ACME HANG': '', 'BETA': '0010040099'}, got
    assert buysheets.cached_req(rel) == got and os.path.exists(buysheets.REQ_CACHE)          # remembered, on our own file
    assert snapshot(root) == before                                                              # the share was only read


def test_only_the_chosen_quarter_is_listed():
    buysheets.BUYSHEET_ROOT = root = fake_share()
    os.makedirs(os.path.join(root, '2027', 'Q2 2027'))
    open(os.path.join(root, '2027', 'Q2 2027', 'C Q2 Buysheet.xlsm'), 'wb').write(b'x')
    out = buysheets.list_buysheets()                                   # no quarter named: the newest
    assert out['quarter_names'] == ['Q1 2027', 'Q2 2027'] and out['quarter'] == 'Q2 2027' and [q['quarter'] for q in out['quarters']] == ['Q2 2027']
    assert [q['quarter'] for q in buysheets.list_buysheets('2027', 'Q1 2027')['quarters']] == ['Q1 2027']


class FakeSap:
    """A stand-in for the HANA connection: requisition 1 has all its lines on POs, 2 some, 3 none."""
    def cursor(self): return self
    def execute(self, sql, params=()): self.sql = sql
    def fetchall(self):
        if 'GROUP BY' in self.sql: return [('0010000001', 2, 2), ('0010000002', 4, 1), ('0010000003', 1, 0)]
        return [('0010000001', '4500000001'), ('0010000001', '4500000002'), ('0010000002', '4500000003')]
    def close(self): pass


def test_the_module_never_opens_a_share_path_for_writing():
    src = open(os.path.join(os.path.dirname(buysheets.__file__), 'buysheets.py'), encoding='utf-8').read()
    for banned in ('os.remove', 'os.rename', 'shutil.move', 'os.unlink', 'os.rmdir', 'shutil.rmtree', "src, 'w", "src, 'a", 'os.makedirs', "'wb'"):
        assert banned not in src, banned


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
