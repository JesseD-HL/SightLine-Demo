"""What SAP really holds for anything the checks flagged, and the read-only SAP explorer."""

import re
from datetime import datetime
from .config import BUILD
from .hana import connect_hana
from .helpers import base_style


# ═══════════════════════════════════════════════════════════════════════════════
# 7. DIAGNOSTICS
# ═══════════════════════════════════════════════════════════════════════════════
def build_diagnostics(cache, connect=None):
    """Everything needed to tell 'SAP has no price' from 'the lookup missed it' in one paste.
    Reads the last validation from the cache, then asks HANA directly about the articles that
    came back without a price, and lists which condition types SAP holds for the first of them."""
    out = {'build': BUILD, 'time': datetime.now().isoformat(timespec='seconds')}
    v = cache.get('validation')
    if not v:
        out['note'] = 'Nothing validated yet - upload a file on the BUYSHEET tab and click Validate first.'
        return out
    out.update({'file': v.get('filename'), 'file_type': v.get('file_type'),
                'summary': v.get('summary'), 'counts': v.get('diag')})

    d = v.get('diag') or {}
    tokens = list(d.get('os_styles_without_sap_price', [])[:3]) + list(d.get('ppk_without_sap_price', [])[:3])
    if not tokens:                       # everything priced - probe the first row so the page is still useful
        first = (v.get('generic') or v.get('ppk') or [{}])[0]
        t = first.get('style') or first.get('ppk_article')
        if t: tokens.append(t)
    out['probed'] = tokens

    try:
        conn = (connect or connect_hana)()
    except Exception as exc:
        out['hana_error'] = str(exc)
        return out
    try:
        cur = conn.cursor()
        probes = {}
        for tok in tokens:
            if not re.fullmatch(r'[A-Za-z0-9_\-]+', tok):
                continue
            try:
                cur.execute(
                    'SELECT "ARTICLE" AS "Article", "F_AMOUNT_PER_PC" AS "Price", "SEGMENT" AS "Segment", '
                    '"VALIDITY_START_DATE" AS "Valid From", "VALIDITY_END_DATE" AS "Valid To" '
                    'FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Supporting/MDA_SUP_ZSLS" '
                    'WHERE "ARTICLE" LIKE ? ORDER BY "ARTICLE" LIMIT 12', (tok + '%',))
                cols = [c[0] for c in cur.description]
                probes[tok] = [dict(zip(cols, [None if x is None else str(x).strip() for x in row]))
                               for row in cur.fetchall()]
            except Exception as exc:
                probes[tok] = f'ERROR: {exc}'
        out['zsls_rows_under_each_probed_article'] = probes

        missing = {}
        for r in v.get('generic') or []:
            if r.get('status') == '✗ MISSING':
                missing.setdefault(r['style'], []).append(f"{r['color_code']}/{r['size']}")
        held = {}
        for style, asked in list(missing.items())[:3]:
            if not re.fullmatch(r'[A-Za-z0-9_\-]+', style):
                continue
            try:
                cur.execute(
                    'SELECT "MARA_ARTICLE" AS "Article", "F_GENERIC" AS "Generic", "MARA_COLOR_CODE" AS "Color Code", '
                    '"MARA_SIZE_CODE" AS "Size", "F_ARTICLE_TYPE" AS "Article Type", "CHARVAL_COLOR_CODE_DESCRIPTION" AS "Color Name" '
                    'FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS" '
                    'WHERE "MARA_ARTICLE" LIKE ? OR "F_GENERIC" = ? ORDER BY "MARA_COLOR_CODE", "MARA_SIZE_CODE" LIMIT 150',
                    (style + '%', style))
                cols = [c[0] for c in cur.description]
                held[style] = {
                    'portal_asked_for_but_found_missing (color/size)': asked[:40],
                    'every_article_sap_has_under_this_style': [
                        dict(zip(cols, [None if x is None else str(x).strip() for x in row])) for row in cur.fetchall()],
                }
            except Exception as exc:
                held[style] = f'ERROR: {exc}'
        if held:
            out['sap_articles_under_styles_with_missing_rows'] = held

        # prepacks the portal called missing: everything SAP holds under that article, any article type
        held_ppk = {}
        for art in list({r['ppk_article'] for r in v.get('ppk') or [] if r.get('exists') == '✗ MISSING'})[:3]:
            if not re.fullmatch(r'[A-Za-z0-9_\-]+', art):
                continue
            try:
                cur.execute(
                    'SELECT "MAST_ARTICLE" AS "Prepack", "MARA_ARTICLE" AS "Component", "F_ARTICLE_TYPE" AS "Article Type", '
                    '"MARA_COLOR_CODE" AS "Color Code", "MARA_SIZE_CODE" AS "Size", "STPO_PPK_QTY" AS "PPK Qty", "MARM_PC_PER_PAC" AS "PC Per PAC" '
                    'FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS" '
                    'WHERE "MAST_ARTICLE" LIKE ? OR "MARA_ARTICLE" LIKE ? ORDER BY "MAST_ARTICLE", "MARA_SIZE_CODE" LIMIT 60',
                    (art + '%', art + '%'))
                cols = [c[0] for c in cur.description]
                held_ppk[art] = {'portal_queried': f'MAST_ARTICLE = {art!r} AND F_ARTICLE_TYPE = \'C\'',
                                 'every_row_sap_has_starting_with_this_article': [
                                     dict(zip(cols, [None if x is None else str(x).strip() for x in row])) for row in cur.fetchall()]}
            except Exception as exc:
                held_ppk[art] = f'ERROR: {exc}'
        if held_ppk:
            out['sap_rows_under_prepacks_called_missing'] = held_ppk

        if tokens and re.fullmatch(r'[A-Za-z0-9_\-]+', tokens[0]):
            style = base_style(tokens[0])
            try:
                cur.execute(
                    'SELECT "CONDITION_TYPE" AS "Type", "KONP_CURRENCY" AS "Currency", "SEGMENT" AS "Segment", '
                    'COUNT(*) AS "Rows", MIN("ARTICLE") AS "Sample article" '
                    'FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Supporting/MDA_SUP_ARTMAS_PRICING_CONDITIONS" '
                    'WHERE "ARTICLE" LIKE ? GROUP BY "CONDITION_TYPE", "KONP_CURRENCY", "SEGMENT" '
                    'ORDER BY "CONDITION_TYPE"', (style + '%',))
                cols = [c[0] for c in cur.description]
                out['condition_types_under_' + style] = [
                    dict(zip(cols, [None if x is None else str(x).strip() for x in row])) for row in cur.fetchall()]
            except Exception as exc:
                out['condition_types_error'] = str(exc)
    finally:
        conn.close()
    return out

# ═══════════════════════════════════════════════════════════════════════════════
# 7b. SAP EXPLORER  (read-only: finds where the PO / margin data lives in HANA; nothing is written)
# ═══════════════════════════════════════════════════════════════════════════════
EXPLORE_TERMS = ['ZPOP', 'ZPOCOND', 'ZPOWB', 'ZARTIMP', 'AFC', 'ZSETCOMM']       # SAP transactions the PO / margin tab has to match
EXPLORE_FIELDS = {                                                               # field -> column-name fragments that might carry it
    'HTS code':          ['HTS', 'STEUC', 'TARIFF'],
    'Comm code':         ['COMM%CODE', 'COMMOD'],
    'Set code':          ['SET%CODE'],
    'Chain code':        ['CHAIN'],
    'Fiber content':     ['FIBER', 'FIBRE', 'COMPOS'],
    'Duty rate':         ['DUTY', 'ZOLL'],
    'Country of origin': ['HERKL', 'ORIGIN'],
    'PO / req link':     ['BANFN', 'EBELN', 'PURCH%REQ', 'PO%NUMBER'],
}

def _ident(x):
    return '"' + str(x).replace('"', '""') + '"'

def explore_sap(conn, terms=None, view=None, col=None, like=None):
    """Read-only questions put to HANA, each in its own try/except so one refusal never hides the rest:
    1. which views and tables are NAMED after each transaction
    2. the dictionary route: transaction -> program (TSTC) -> every table that program uses (D010TAB), Z tables first
    3. which view/table columns look like HTS, comm code, set code, chain code, fiber content, duty rate, origin, req/PO number
    4. with view=NAME: the first rows of that view and how many of its columns are actually filled in"""
    sample_only = bool(view) and not terms          # ?view=X on its own: just that object, without the long catalog sections
    terms = terms or EXPLORE_TERMS
    out = {'build': BUILD, 'time': datetime.now().isoformat(timespec='seconds')}
    if not sample_only:
        out['asked_about'] = terms
    cur = conn.cursor()

    def run(sql, params=()):
        cur.execute(sql, params)
        cols = [c[0] for c in cur.description]
        return [dict(zip(cols, [None if x is None else str(x).strip() for x in row])) for row in cur.fetchall()]

    def safe(fn, *args):
        try:
            return fn(*args)
        except Exception as exc:
            return f'ERROR: {exc}'

    named = {}
    for t in ([] if sample_only else terms):
        like = '%' + t + '%'
        named[t] = {
            'views': safe(run, 'SELECT "SCHEMA_NAME", "VIEW_NAME" FROM "SYS"."VIEWS" WHERE "SCHEMA_NAME" <> \'SYS\' '
                               'AND UPPER("VIEW_NAME") LIKE ? ORDER BY "SCHEMA_NAME", "VIEW_NAME" LIMIT 40', (like,)),
            'tables': safe(run, 'SELECT "SCHEMA_NAME", "TABLE_NAME" FROM "SYS"."TABLES" WHERE "SCHEMA_NAME" <> \'SYS\' '
                                'AND UPPER("TABLE_NAME") LIKE ? ORDER BY "SCHEMA_NAME", "TABLE_NAME" LIMIT 40', (like,)),
        }
    if not sample_only:
        out['1_objects_named_like_each_transaction'] = named

    def dictionary_route():
        homes = run('SELECT DISTINCT "SCHEMA_NAME", "TABLE_NAME" FROM "SYS"."TABLES" WHERE "TABLE_NAME" IN (\'TSTC\', \'D010TAB\', \'DD02T\')')
        schema_of = {h['TABLE_NAME']: h['SCHEMA_NAME'] for h in homes}
        if not {'TSTC', 'D010TAB'} <= set(schema_of):
            return ('The SAP dictionary tables (TSTC / D010TAB) are not visible to this HANA user, so a transaction cannot be '
                    'traced to its tables from here. Use sections 1 and 3.')
        tab = lambda n: f'{_ident(schema_of[n])}.{_ident(n)}'
        res = {}
        for t in terms:
            tc = run(f'SELECT "TCODE", "PGMNA" AS "Program", "DYPNO" AS "Screen" FROM {tab("TSTC")} WHERE "TCODE" = ?', (t,))
            if not tc:
                res[t] = 'Not a transaction code in this SAP system (TSTC has no row for it).'
                continue
            join = (f'LEFT JOIN {tab("DD02T")} t ON t."TABNAME" = d."TABNAME" AND t."DDLANGUAGE" = \'E\'') if 'DD02T' in schema_of else ''
            text = 't."DDTEXT" AS "Description"' if join else '\'\' AS "Description"'
            used = run(f'SELECT DISTINCT d."TABNAME" AS "Table", {text} FROM {tab("D010TAB")} d {join} '
                       f'WHERE d."MASTER" = ? ORDER BY CASE WHEN d."TABNAME" LIKE \'Z%\' THEN 0 ELSE 1 END, d."TABNAME" LIMIT 250',
                       (tc[0]['Program'],))
            res[t] = {'transaction': tc[0], 'tables_the_program_uses_z_tables_first': used}
        return res
    if not sample_only:
        out['2_dictionary_route_transaction_to_tables'] = safe(dictionary_route)

    def field_search():
        found = {}
        for field, frags in EXPLORE_FIELDS.items():
            conds = ' OR '.join(['UPPER("COLUMN_NAME") LIKE ?'] * len(frags))
            params = tuple('%' + f + '%' for f in frags)
            in_views = safe(run, f'SELECT "SCHEMA_NAME", "VIEW_NAME" AS "OBJECT", "COLUMN_NAME" FROM "SYS"."VIEW_COLUMNS" '
                                 f'WHERE "SCHEMA_NAME" = \'_SYS_BIC\' AND ({conds}) ORDER BY "VIEW_NAME", "POSITION" LIMIT 40', params)
            in_tables = safe(run, f'SELECT "SCHEMA_NAME", "TABLE_NAME" AS "OBJECT", "COLUMN_NAME" FROM "SYS"."TABLE_COLUMNS" '
                                  f'WHERE "SCHEMA_NAME" NOT LIKE \'%SYS%\' AND ({conds}) ORDER BY "TABLE_NAME", "POSITION" LIMIT 40', params)
            found[field] = {'in_views': in_views, 'in_tables': in_tables}
        return found
    if not sample_only:
        out['3_columns_that_could_hold_each_field'] = safe(field_search)

    if view:
        def sample():
            # resolve the name through HANA's own catalog (view, then calculation view, then table): only a name the catalog lists reaches the SQL text
            hit = run('SELECT "SCHEMA_NAME", "VIEW_NAME" AS "OBJ", \'view\' AS "KIND" FROM "SYS"."VIEWS" WHERE "VIEW_NAME" = ? AND "SCHEMA_NAME" <> \'SYS\' LIMIT 1', (view,))
            if not hit:
                hit = run('SELECT DISTINCT "SCHEMA_NAME", "VIEW_NAME" AS "OBJ", \'view\' AS "KIND" FROM "SYS"."VIEW_COLUMNS" WHERE "VIEW_NAME" = ? AND "SCHEMA_NAME" <> \'SYS\' LIMIT 1', (view,))
            if not hit:
                hit = run('SELECT "SCHEMA_NAME", "TABLE_NAME" AS "OBJ", \'table\' AS "KIND" FROM "SYS"."TABLES" WHERE "TABLE_NAME" = ? AND "SCHEMA_NAME" <> \'SYS\' LIMIT 1', (view,))
            if not hit:
                return f'No view or table called {view!r} in the catalog. Use a name from section 1 or 3.'
            schema, name, kind = hit[0]['SCHEMA_NAME'], hit[0]['OBJ'], hit[0]['KIND']
            src = f'{_ident(schema)}.{_ident(name)}'
            cols_sql = ('SELECT "COLUMN_NAME" FROM "SYS"."VIEW_COLUMNS" WHERE "SCHEMA_NAME" = ? AND "VIEW_NAME" = ? ORDER BY "POSITION"' if kind == 'view'
                        else 'SELECT "COLUMN_NAME" FROM "SYS"."TABLE_COLUMNS" WHERE "SCHEMA_NAME" = ? AND "TABLE_NAME" = ? ORDER BY "POSITION"')
            where, params, shown_filter = '', (), None
            if col:
                catalog_cols = [r['COLUMN_NAME'] for r in run(cols_sql, (schema, name))]
                match = next((c for c in catalog_cols if c.upper() == col.upper()), None)
                if not match:
                    return f'{col!r} is not a column of {name}. Its columns are: {catalog_cols}'
                where, params = f' WHERE UPPER({_ident(match)}) LIKE ?', (str(like or '').upper() + '%',)   # column name from the catalog; the value is a bound parameter
                shown_filter = f'{match} starts with {str(like or "").upper()!r}'
            cols = [r['COLUMN_NAME'] for r in run(cols_sql, (schema, name))][:300]          # every column the catalog lists, not just the first 80
            rows = run(f'SELECT * FROM {src}{where} LIMIT {5 if col else 3}', params)
            cols = cols or (list(rows[0].keys()) if rows else [])
            fills = ', '.join(f'SUM(CASE WHEN {_ident(c)} IS NULL OR TRIM(TO_NVARCHAR({_ident(c)})) = \'\' THEN 0 ELSE 1 END) AS "F{i}"'
                              for i, c in enumerate(cols))
            stats = safe(run, f'SELECT COUNT(*) AS "N", {fills} FROM (SELECT * FROM {src}{where} LIMIT 5000) AS "T"', params) if cols else []
            res = {kind: f'{schema}.{name}'}
            if shown_filter:
                res['filter'] = shown_filter
            if isinstance(stats, list) and stats:
                counts = {c: int(stats[0][f'F{i}']) for i, c in enumerate(cols)}
                res['rows_in_sample'] = int(stats[0]['N'])                                   # capped at 5000; a 0 is counted as filled, only blank / NULL are not
                res['filled_counts'] = {c: v for c, v in counts.items() if v}
                res['always_empty'] = [c for c, v in counts.items() if not v]
            else:
                res.update({'columns': cols, 'stats': stats})
            res['sample_rows'] = [{k: v for k, v in r.items() if v not in ('', None)} for r in rows]
            return res
        out['4_view_sample'] = safe(sample)
    return out
