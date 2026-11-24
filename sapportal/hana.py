"""The HANA connection and the read-only queries."""

import re
from datetime import datetime
from . import demo_db
from .helpers import color_from_ppk, num


# ═══════════════════════════════════════════════════════════════════════════════
# 3. HANA - connection + read-only queries
# ═══════════════════════════════════════════════════════════════════════════════
def norm_code(v):
    """A comparable code: ' 3609 ', 3609, 3609.0 and '0003609' are all '3609'."""
    t = '' if v is None else str(v).strip()
    if re.fullmatch(r'-?\d+\.0+', t): t = t.split('.')[0]
    if re.fullmatch(r'\d+', t): t = t.lstrip('0') or '0'
    return t.upper()

def is_canadian_style(raw_style):
    """Fallback only (used when SAP's price currency is not known): a generic ending in C or CA is Canadian."""
    return str(raw_style or '').split('-')[0].strip().upper().endswith(('CA', 'C'))

def buy_price_for(r, currency):
    """The buysheet price in SAP's currency: PRICE (CAD) when SAP prices the article in CAD, PRICE (USD) when in USD.
    A buysheet without the two columns, or an article SAP has no currency for, keeps the price as read."""
    if r.get('price_cad') is None and r.get('price_usd') is None: return num(r.get('price'))
    if currency == 'CAD': return num(r.get('price_cad'))
    if currency == 'USD': return num(r.get('price_usd'))
    return num(r.get('price'))

def hana_configured():
    return True                      # the demo always has its invented database

def connect_hana():
    return demo_db.connect()

def query_hana_generics(conn, generic_styles):
    if not generic_styles: return {}
    style_list = ', '.join(f"'{s}'" for s in generic_styles)

    def run(with_fineline):
        # The Bigbox fineline is the article's external material group (EXTWG; MARA_EXT_MERCH_CAT in the BI views).
        extra = ', "MARA_EXT_MERCH_CAT" AS "Fineline", "MARA_MERCHANDISE_CATEGORY" AS "Mdse Cat"' if with_fineline else ''
        sql = f"""
        SELECT "F_GENERIC" AS "Generic", "MARA_ARTICLE" AS "Article",
               "MARA_COLOR_CODE" AS "Color Code", "MARA_SIZE_CODE" AS "Size Code",
               "F_GENERIC_DESCRIPTION" AS "Description", "WRF_SORT_POSITION" AS "Size Sort",
               "F_PC_PER_CAR_FIX" AS "PC Per Ctn",
               "CHARVAL_COLOR_CODE_DESCRIPTION" AS "Color Name", "F_SEASON" AS "Season", "F_YEAR" AS "Year"{extra}
        FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS"
        WHERE "F_GENERIC" IN ({style_list}) AND "F_ARTICLE_TYPE" <> 'C'
        """
        cursor = conn.cursor(); cursor.execute(sql)
        return cursor.fetchall(), [d[0] for d in cursor.description]
    try:
        rows, cols = run(True)
    except Exception as exc:                                  # a HANA without that column must not stop the validation
        if not any(c in str(exc).upper() for c in ('MARA_EXT_MERCH_CAT', 'MARA_MERCHANDISE_CATEGORY')): raise
        rows, cols = run(False)
    result = {}
    for row in rows:
        r = dict(zip(cols, row))
        generic    = str(r.get('Generic','') or '').strip()
        color_code = str(r.get('Color Code','') or '').strip()
        size_code  = str(r.get('Size Code','') or '').strip().upper()
        hana_desc  = str(r.get('Description','') or '').strip().upper()
        size_sort  = str(r.get('Size Sort','') or '99999').strip()
        pc_per_ctn = r.get('PC Per Ctn')
        color_name = str(r.get('Color Name','') or '').strip().upper()
        result[(generic, color_code, size_code)] = {
            'article':    str(r.get('Article','') or '').strip(),
            'desc':       hana_desc,
            'sort':       size_sort,
            'pc_per_ctn': pc_per_ctn,
            'color_name': color_name,
            'season':     (f"{str(r.get('Year') or '').strip()} {str(r.get('Season') or '').strip().upper()}").strip(),     # SAP's season assignment, e.g. '2027 FALL'
            'fineline':   norm_code(r.get('Fineline')),
            'mdse_cat':   str(r.get('Mdse Cat', '') or '').strip().upper(),        # the merchandise category, e.g. RIDGELINE-M: the MM41 script needs it
        }
    return result

def query_hana_ppks(conn, ppk_articles):
    if not ppk_articles: return {}, {}
    article_list = ', '.join(f"'{a}'" for a in ppk_articles)
    sql = f"""
        SELECT "MAST_ARTICLE" AS "PPK Article", "MARA_ARTICLE" AS "Component",
               "CHARVAL_COLOR_CODE_DESCRIPTION" AS "Color Name",
               "MARA_SIZE_CODE" AS "Size Code", "STPO_PPK_QTY" AS "PPK Qty",
               "F_GENERIC_DESCRIPTION" AS "Description", "WRF_SORT_POSITION" AS "Size Sort",
               "MARM_PC_PER_PAC" AS "PC Per PAC"
        FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS"
        WHERE "MAST_ARTICLE" IN ({article_list}) AND "F_ARTICLE_TYPE" = 'C'
    """
    cursor = conn.cursor(); cursor.execute(sql); rows = cursor.fetchall()
    cols = [d[0] for d in cursor.description]
    ppk_exists = {}; ppk_details = {}
    for row in rows:
        r          = dict(zip(cols, row))
        ppk_art    = str(r.get('PPK Article','') or '').strip()
        color_name = str(r.get('Color Name','') or '').strip().upper()
        size_code  = str(r.get('Size Code','') or '').strip().upper()
        ppk_qty    = r.get('PPK Qty')
        size_sort  = str(r.get('Size Sort','') or '99999').strip()
        hana_desc  = str(r.get('Description','') or '').strip().upper()
        pc_per_pac = r.get('PC Per PAC')
        ppk_exists[ppk_art] = True
        if ppk_art not in ppk_details:
            ppk_details[ppk_art] = {
                'color_name': color_name, 'hana_desc': hana_desc,
                'variants': [], 'seen': set(), 'pc_per_pac': pc_per_pac
            }
        if size_code not in ppk_details[ppk_art]['seen']:
            ppk_details[ppk_art]['seen'].add(size_code)
            ppk_details[ppk_art]['variants'].append((size_sort, size_code, ppk_qty))
            ppk_details[ppk_art].setdefault('components', []).append(
                {'article': str(r.get('Component') or '').strip(), 'code': color_from_ppk(ppk_art), 'color': color_name, 'size': size_code,
                 'qty': int(round(float(ppk_qty))) if ppk_qty not in (None, '') else None, 'sort': size_sort})
    return ppk_exists, ppk_details

def query_hana_ppks_for_generics(conn, generics):
    """Every prepack under these generics with its components -> {article: {'generic', 'comps': [(code, name, size, qty)], 'pc_per_pac', 'desc'}}.
    Used to find a prepack by CONTENT (color + sizes + ratio) when the typed article name is not found - the same
    search check_8_7.py does."""
    generics = sorted({str(g).strip() for g in generics if g})
    if not generics: return {}
    gen_list = ', '.join(f"'{g}'" for g in generics)
    sql = f"""
        SELECT "MAST_ARTICLE" AS "PPK", "MARA_ARTICLE" AS "Component", "F_GENERIC" AS "Generic", "MARA_COLOR_CODE" AS "Color Code",
               "CHARVAL_COLOR_CODE_DESCRIPTION" AS "Color Name", "MARA_SIZE_CODE" AS "Size", "STPO_PPK_QTY" AS "Qty",
               "WRF_SORT_POSITION" AS "Sort", "MARM_PC_PER_PAC" AS "PC Per PAC", "F_GENERIC_DESCRIPTION" AS "Description"
        FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS"
        WHERE "F_GENERIC" IN ({gen_list}) AND "F_ARTICLE_TYPE" = 'C'
        ORDER BY "MAST_ARTICLE", "WRF_SORT_POSITION"
    """
    cursor = conn.cursor(); cursor.execute(sql); rows = cursor.fetchall()
    cols = [d[0] for d in cursor.description]
    out = {}
    for row in rows:
        r = dict(zip(cols, row))
        art = str(r.get('PPK') or '').strip().upper()
        if not art: continue
        e = out.setdefault(art, {'generic': str(r.get('Generic') or '').strip().upper(), 'comps': [], 'pc_per_pac': r.get('PC Per PAC'),
                                 'desc': str(r.get('Description') or '').strip().upper(), 'seen': set()})
        size = str(r.get('Size') or '').strip().upper(); code = str(r.get('Color Code') or '').strip().upper()
        if (code, size) in e['seen']: continue
        e['seen'].add((code, size))
        try: qty = int(round(float(r.get('Qty') or 0)))
        except (TypeError, ValueError): qty = 0
        e['comps'].append((code, str(r.get('Color Name') or '').strip().upper(), size, qty, str(r.get('Sort') or '99999')))
        e.setdefault('components', []).append({'article': str(r.get('Component') or r.get('PPK') or '').strip(), 'code': code,
                                               'color': str(r.get('Color Name') or '').strip().upper(), 'size': size, 'qty': qty, 'sort': str(r.get('Sort') or '99999')})
    return out

def query_hana_ppk_descriptions(conn, ppk_articles):
    if not ppk_articles: return {}
    article_list = ', '.join(f"'{a}'" for a in ppk_articles)
    sql = f"""
        SELECT "MARA_ARTICLE" AS "Article", "MAKT_ARTICLE_DESCRIPTION" AS "Description"
        FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS"
        WHERE "MARA_ARTICLE" IN ({article_list})
    """
    cursor = conn.cursor(); cursor.execute(sql); rows = cursor.fetchall()
    cols = [d[0] for d in cursor.description]
    result = {}
    for row in rows:
        r = dict(zip(cols, row))
        article = str(r.get('Article','') or '').strip()
        desc    = str(r.get('Description','') or '').strip().upper()
        if article and desc: result[article] = desc
    return result

def query_hana_prices(conn, articles):
    """ZSLS price per article from MDA_SUP_ZSLS (already a per-piece amount).

    The view can hold more than one row per article (price history / future prices), so the
    choice is deterministic: a row valid today wins, otherwise the one with the latest start.
    Long IN lists are sent in chunks of 500.
    """
    articles = sorted({str(a).strip() for a in articles if a})
    if not articles: return {}
    today = datetime.now().strftime('%Y%m%d')
    norm  = lambda v, default: (str(v).replace('-', '')[:8] if v else default)
    best  = {}                                   # article -> (rank, price)
    cursor = conn.cursor()
    for i in range(0, len(articles), 500):
        article_list = ', '.join(f"'{a}'" for a in articles[i:i+500])
        sql = f"""
            SELECT "ARTICLE" AS "Article", "F_AMOUNT_PER_PC" AS "Price",
                   "VALIDITY_START_DATE" AS "Valid From", "VALIDITY_END_DATE" AS "Valid To"
            FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Supporting/MDA_SUP_ZSLS"
            WHERE "ARTICLE" IN ({article_list})
        """
        cursor.execute(sql)
        cols = [d[0] for d in cursor.description]
        for row in cursor.fetchall():
            r = dict(zip(cols, row))
            article = str(r.get('Article') or '').strip()
            price   = r.get('Price')
            if not article or price is None: continue
            start = norm(r.get('Valid From'), '00000000')
            end   = norm(r.get('Valid To'),   '99991231')
            rank  = (start <= today <= end, start)
            if article not in best or rank > best[article][0]:
                best[article] = (rank, float(price))
    query_hana_prices.currency = query_hana_price_currency(conn, articles)         # the view has the amount, not the currency
    return {art: price for art, (rank, price) in best.items()}

PRICE_CURRENCY_SOURCES = ('"SAPDEMO"."ZSAC_ZSLS_A504"', '"SAPDEMO"."ZSAC_ZSLS_A996"')    # the ZSLS condition records: MATNR + KONWA (what MEK3 shows)

def query_hana_price_currency(conn, articles):
    """article -> 'CAD' / 'USD' (...) from the ZSLS condition records. The pricing view carries the amount (KONP_AMOUNT) but not the
    currency, which SAP keeps on the condition record (KONWA). An article with two different currencies is left out (unknown), and a
    condition table that cannot be read is skipped: the buysheet check then falls back to the generic's currency, then to C / CA."""
    arts, out, cur = sorted({str(a) for a in articles if a}), {}, conn.cursor()
    for src in PRICE_CURRENCY_SOURCES:
        seen = {}
        for i in range(0, len(arts), 500):
            part = arts[i:i + 500]
            try:
                cur.execute(f'SELECT DISTINCT "MATNR", "KONWA" FROM {src} WHERE "MATNR" IN ({", ".join("?" for _ in part)})', tuple(part))
                for m, k in cur.fetchall():
                    seen.setdefault(str(m or '').strip(), set()).add(str(k or '').strip().upper())
            except Exception:
                break
        for a, ks in seen.items():
            ks.discard('')
            if a not in out and len(ks) == 1: out[a] = next(iter(ks))
    return out

def query_hana_color_codes(conn, codes):
    """Color code -> set of descriptions SAP uses for it, read from the articles in HANA
    (the same source the standalone color_code_validator uses)."""
    codes = sorted({str(c).strip() for c in codes if c})
    out = {}
    if not codes: return out
    cursor = conn.cursor()
    for i in range(0, len(codes), 500):
        code_list = ', '.join(f"'{c}'" for c in codes[i:i+500])
        cursor.execute(f"""
            SELECT DISTINCT "MARA_COLOR_CODE" AS "Code", "CHARVAL_COLOR_CODE_DESCRIPTION" AS "Description"
            FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS"
            WHERE "MARA_COLOR_CODE" IN ({code_list})
        """)
        for code, desc in cursor.fetchall():
            code = str(code or '').strip(); desc = str(desc or '').strip().upper()
            if code and desc: out.setdefault(code, set()).add(desc)
    return out

def query_color_codes_by_name(conn, names):
    """Color name exactly as SAP spells it (case kept) -> the sorted color codes that use that spelling.  Names are asked for without
    regard to case, so a caller can fall back from an exact spelling to a looser one."""
    names = sorted({' '.join(str(n).split()).upper() for n in names if n})
    out, cursor = {}, conn.cursor()
    for i in range(0, len(names), 500):
        part = names[i:i + 500]
        cursor.execute('SELECT DISTINCT "CHARVAL_COLOR_CODE_DESCRIPTION", "MARA_COLOR_CODE" FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS" '
                       'WHERE UPPER("CHARVAL_COLOR_CODE_DESCRIPTION") IN (' + ','.join('?' * len(part)) + ')', tuple(part))
        for desc, code in cursor.fetchall():
            name, code = ' '.join(str(desc or '').split()), str(code or '').strip()
            if name and code:
                out.setdefault(name, set()).add(code)
    return {n: sorted(c) for n, c in out.items()}

def query_variant_articles(conn, generics):
    """Get variant articles for PR CSV generation."""
    if not generics: return {}
    gen_list = ', '.join(f"'{g}'" for g in generics)
    sql = f"""
        SELECT "MARA_ARTICLE","F_GENERIC","MARA_COLOR_CODE","MARA_SIZE_CODE"
        FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS"
        WHERE "F_GENERIC" IN ({gen_list}) AND "F_ARTICLE_TYPE" != 'C'
    """
    cursor = conn.cursor(); cursor.execute(sql); rows = cursor.fetchall()
    result = {}
    for row in rows:
        article,gen,code,size = [str(x).strip().upper() if x else '' for x in row]
        result[(gen,code,size)] = article
    return result


def query_hana_msrp(conn, articles):
    """MSRP per article: the ZVP1 condition (per piece), from the article's pricing conditions. An article-level record wins over a
    customer-specific one; then a record valid today; then the latest start. Long IN lists are sent in chunks of 500."""
    articles = sorted({str(a).strip() for a in articles if a})
    if not articles: return {}
    today = datetime.now().strftime('%Y%m%d')
    norm  = lambda v, default: (str(v).replace('-', '')[:8] if v else default)
    best, cursor = {}, conn.cursor()
    for i in range(0, len(articles), 500):
        part = articles[i:i + 500]
        cursor.execute('SELECT "ARTICLE", "CUSTOMER", "KONP_AMOUNT", "KONP_PER", "VALID_FROM", "VALID_TO" '
                       'FROM "_SYS_BIC"."Acme.Data_Universe.Raw_Supporting/MDA_SUP_ARTMAS_PRICING_CONDITIONS" '
                       'WHERE "CONDITION_TYPE" = \'ZVP1\' AND "ARTICLE" IN (' + ", ".join("?" for _ in part) + ')', tuple(part))
        for art, cust, amount, per, vf, vt in cursor.fetchall():
            art = str(art or '').strip()
            if not art or amount is None: continue
            start, end = norm(vf, '00000000'), norm(vt, '99991231')
            rank = (not str(cust or '').strip(), start <= today <= end, start)
            if art not in best or rank > best[art][0]:
                best[art] = (rank, float(amount) / float(per or 1))
    return {a: v for a, (r, v) in best.items()}
