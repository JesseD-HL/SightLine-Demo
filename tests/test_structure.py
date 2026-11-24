"""The sapportal package keeps its layers: a module only imports from modules BELOW it, so nothing depends on a circle, and server.py
still exposes the names the tests and the small scripts use.  Run:  python tests/test_structure.py"""
import ast, glob, os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, '..')
sys.path.insert(0, ROOT)

LAYERS = ['config', 'helpers', 'demo_db', 'hana', 'parsers', 'outputs', 'validation', 'stock', 'diagnostics', 'reqcheck', 'webapp', 'routes', 'delivery_flow', 'sap_runner', 'buysheets', 'retailportal', 'dpci_list', 'hold_check', 'demo_guard']


def imports_of(module):
    tree = ast.parse(open(os.path.join(ROOT, 'sapportal', module + '.py'), encoding='utf-8').read())
    return {m for n in ast.walk(tree) if isinstance(n, ast.ImportFrom) and n.level == 1 for m in ([n.module] if n.module else [a.name for a in n.names])}


def test_every_module_is_in_the_layer_list():
    on_disk = {os.path.basename(p)[:-3] for p in glob.glob(os.path.join(ROOT, 'sapportal', '*.py'))} - {'__init__'}
    assert on_disk == set(LAYERS), (on_disk ^ set(LAYERS))


def test_modules_only_import_from_below():
    for i, mod in enumerate(LAYERS):
        for dep in imports_of(mod):
            assert dep in LAYERS and LAYERS.index(dep) < i, f'{mod} imports {dep}, which is not below it'


def test_server_exposes_what_the_tests_and_scripts_use():
    import server
    for name in ('parse_prebuy', 'make_data_sheet', 'make_creation_templates', 'make_ppk_run_file', 'parse_flow_file', 'split_buysheet', 'flow_customer_mismatches',
                 'buysheet_req_check', 'season_text', 'season_from_style', 'connect_hana', 'validate_all', 'refresh_validation', 'app', 'BUILD', 'lan_ip'):
        assert hasattr(server, name), name


def test_the_routes_are_registered():
    import server
    rules = {r.rule for r in server.app.url_map.iter_rules()}
    for rule in ('/', '/validate', '/revalidate', '/creation_templates', '/lookup', '/delivery_flow/preview', '/ppk_creation/start', '/itemhub/run', '/sap_run/scripts', '/buysheets/list', '/retailportal/run', '/hold_check', '/hold_check/export'):
        assert rule in rules, rule
    assert len(list(server.app.url_map.iter_rules())) >= 29, len(list(server.app.url_map.iter_rules()))      # 28 routes + Flask's /static


def test_the_build_stamp_matches_the_page():
    import re, server
    page = open(os.path.join(ROOT, 'portal.html'), encoding='utf-8').read()
    assert re.search(r'PAGE_BUILD = "%s"' % re.escape(server.BUILD), page), 'portal.html PAGE_BUILD differs from the server BUILD'


if __name__ == '__main__':
    failed = 0
    for name, fn in sorted((n, f) for n, f in globals().items() if n.startswith('test_')):
        try:
            fn(); print('ok    ', name)
        except Exception as e:
            failed += 1; print('FAILED', name, '-', repr(e)[:400])
    sys.exit(1 if failed else 0)
