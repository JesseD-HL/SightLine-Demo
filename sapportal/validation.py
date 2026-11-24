"""Every row checked against what SAP holds."""

import json
import re
from collections import defaultdict
from .config import SIZE_ORDER
from .hana import (
    buy_price_for,
    connect_hana,
    norm_code,
    query_hana_color_codes,
    query_hana_generics,
    query_hana_ppk_descriptions,
    query_hana_ppks,
    query_hana_ppks_for_generics,
    query_hana_msrp,
    query_hana_prices)
from .helpers import (
    build_ratio_sizes,
    canonical_size,
    color_from_ppk,
    compare_ratio,
    dash_list,
    desc_diff_note,
    generic_from_ppk,
    iso_date,
    normalize_color,
    num,
    price_check_text,
    season_text,
    strip_color)
from .outputs import ppk_components


# ═══════════════════════════════════════════════════════════════════════════════
# 5. VALIDATION - every row against what SAP holds
# ═══════════════════════════════════════════════════════════════════════════════
class SapView:
    """Everything the checks need from HANA for one file, fetched once and indexed."""

    def __init__(self, rows):
        os_rows  = [r for r in rows if r['type'] == 'OS']
        ppk_rows = [r for r in rows if r['type'] == 'PPK']
        self.generic_styles = sorted({r['style'] for r in os_rows})
        self.ppk_articles   = sorted({r['raw_style'] for r in ppk_rows})
        ppk_styles          = sorted({generic_from_ppk(a) for a in self.ppk_articles})

        conn = connect_hana()
        try:
            self.generic_exists = query_hana_generics(conn, self.generic_styles)            # (style, code, size) -> variant
            self.ppk_exists, self.ppk_details = query_hana_ppks(conn, self.ppk_articles)
            self.ppk_desc = query_hana_ppk_descriptions(conn, self.ppk_articles)           # prepack article -> its own name
            self.ppks_under_generic = query_hana_ppks_for_generics(conn, ppk_styles)         # every prepack SAP has for these styles
            ppk_gen_exists = query_hana_generics(conn, ppk_styles)
            self.ppk_variants = ppk_gen_exists                                             # kept: a prepack is built from these variants
            self.generic_desc = query_hana_ppk_descriptions(conn, sorted(set(self.generic_styles) | set(ppk_styles)))   # the generic ARTICLE's own name
            variant_arts = {v['article'] for v in self.generic_exists.values() if v.get('article')}
            self.price_keys = set(self.ppk_articles) | variant_arts | set(self.generic_styles) | set(self.ppks_under_generic)   # found-by-content prepacks need a price too
            self.prices = query_hana_prices(conn, self.price_keys)
            try: self.msrp = query_hana_msrp(conn, self.price_keys)                  # ZVP1, per piece
            except Exception: self.msrp = {}                                          # a failed MSRP read must not stop the validation
            self.price_ccy = dict(getattr(query_hana_prices, 'currency', {}))     # article -> CAD / USD, from the price record
            self.code_names = query_hana_color_codes(conn, {r['color_code'] for r in os_rows})  # code -> names SAP uses for it anywhere
        finally:
            conn.close()

        # per-style facts, taken from the variants. A prepack row ('STYLE-COLOR') never supplies the style description.
        self.style_desc, self.style_pc_ctn, self.style_color = {}, {}, {}
        self.style_colors, self.color_sizes = defaultdict(set), defaultdict(set)
        self.style_pc_ctn_votes = defaultdict(lambda: defaultdict(int))
        for (g, c, s), v in self.generic_exists.items():
            self.style_colors[g].add(c); self.color_sizes[(g, c)].add(s)
            if v.get('color_name') and (g, c) not in self.style_color:
                self.style_color[(g, c)] = v['color_name']
            if v.get('desc') and g not in self.style_desc and '-' not in (v.get('article') or ''):
                self.style_desc[g] = v['desc']
            if v.get('pc_per_ctn') and s:                       # only real size variants vote, never the generic row
                self.style_pc_ctn_votes[g][str(v['pc_per_ctn']).strip()] += 1
        for g, votes in self.style_pc_ctn_votes.items():      # most common value across the style's variants
            self.style_pc_ctn[g] = max(votes, key=lambda k: votes[k])
        self.ppk_style_desc = {}
        for (g, c, s), v in ppk_gen_exists.items():
            if v.get('desc') and g not in self.ppk_style_desc and '-' not in (v.get('article') or ''):
                self.ppk_style_desc[g] = v['desc']
        # every variant SAP holds for the styles in this file (open-stock styles AND prepack styles):
        # style -> color code -> size -> variant. A prepack's components are exactly such variants.
        self.variant_index, self.style_known, self.fineline, self.mdse_cat = defaultdict(lambda: defaultdict(dict)), set(), defaultdict(set), {}
        for (g, c, sz), v in {**self.ppk_variants, **self.generic_exists}.items():
            self.style_known.add(g)
            if v.get('fineline'): self.fineline[g].add(v['fineline'])
            if v.get('mdse_cat') and g not in self.mdse_cat: self.mdse_cat[g] = v['mdse_cat']                                   # the generic exists (its own head article has no color / size)
            if c and sz:
                self.variant_index[g][str(c).upper()][str(sz).upper()] = v

    def find_variant(self, style, code, size):
        """The variant of this style / color code / size, or None. A size matches as typed, else through the size aliases."""
        sizes = self.variant_index.get(style, {}).get(str(code).strip().upper(), {})
        want = str(size).strip().upper()
        return sizes.get(want) or next((v for k, v in sizes.items() if canonical_size(k) == canonical_size(want)), None)

    def has_color(self, style, code):
        return bool(self.variant_index.get(style, {}).get(str(code).strip().upper()))

    def sizes_on(self, style, code):
        """The sizes SAP has for this style / color code, in size order."""
        have = self.variant_index.get(style, {}).get(str(code).strip().upper(), {})
        return sorted(have, key=lambda z: (SIZE_ORDER.index(z) if z in SIZE_ORDER else 99, z))

    def find_ppk_by_content(self, style, rows):
        """A prepack under this generic whose components are exactly these rows' (color code, size, units per pack).
        Works for one-color prepacks and for complex ones (several rows, one article). '' when none matches."""
        want = set()
        for r in rows:
            sizes, ratio = dash_list(r.get('sizes_str')), dash_list(r.get('ratio'))
            if not sizes or len(sizes) != len(ratio): return ''
            try: want |= {(r['color_code'], sz, int(q)) for sz, q in zip(sizes, ratio) if int(q) > 0}
            except ValueError: return ''
        if not want: return ''
        for art, e in self.ppks_under_generic.items():
            if e['generic'] == style and {(c, sz, q) for c, _, sz, q, _ in e['comps'] if q > 0} == want:
                return art
        return ''

    def detail_for(self, art):
        """ppk_details-shaped record for a prepack found by content, so the same checks apply."""
        e = self.ppks_under_generic.get(art)
        if not e: return {}
        return {'color_name': e['comps'][0][1] if e['comps'] else '', 'hana_desc': e['desc'], 'pc_per_pac': e['pc_per_pac'],
                'variants': [(sort, sz, q) for _, _, sz, q, sort in e['comps']], 'comps': e['comps'],
                'components': e.get('components', []), 'seen': set()}

    def generic_description(self, style):
        """The generic's description: its own article record first, then its variants - never a prepack's."""
        return self.generic_desc.get(style) or self.style_desc.get(style) or self.ppk_style_desc.get(style) or ''

    def currency_of(self, *articles):
        """The currency of SAP's price for the first of these articles that has one ('' when SAP gives none)."""
        for a in articles:
            if a and self.price_ccy.get(a): return self.price_ccy[a]
        return ''

    def os_price(self, style, code, size):
        """SAP price for one OS size plus a plain-English trace. The variant's own price wins, then the generic's.
        Prepack prices are never used here: they are priced separately and can differ from the generic."""
        art = (self.generic_exists.get((style, code, size.upper())) or {}).get('article')
        steps = []
        if art:
            if art in self.prices:
                return self.prices[art], 'variant', f'variant {art}: {self.prices[art]:.2f}'
            steps.append(f'variant {art}: no price row')
        else:
            steps.append(f'variant {style}/{code}/{size}: not found in SAP')
        if style in self.prices:
            steps.append(f'generic {style}: {self.prices[style]:.2f}')
            return self.prices[style], 'generic', ' -> '.join(steps)
        steps.append(f'generic {style}: no price row')
        return None, '', ' -> '.join(steps)

    def os_msrp(self, style, code, size):
        """SAP MSRP (ZVP1) for one OS size: the variant's own, then the generic's."""
        art = (self.generic_exists.get((style, code, size.upper())) or {}).get('article')
        return self.msrp.get(art) if art in self.msrp else self.msrp.get(style)

    def os_color_check(self, style, code, buy_color):
        """Color code + description for one OS row, judged against HANA only -> (check, SAP name compared with, note).
        HANA only shows codes used on ARTICLES: a code created in the color master (WRFCHVAL) but not on any article
        yet looks the same as one that does not exist, so that case is a warning, never 'needs creating'."""
        cmp = lambda c: normalize_color(c).replace(' ', '')
        buy = cmp(buy_color)
        on_style = self.style_color.get((style, code))
        if on_style:
            if not buy: return f'⚠ NO BUY COLOR (SAP={on_style})', on_style, 'The file has no color description for this code.'
            if buy == cmp(on_style): return '✓', on_style, ''
            return f'✗ SAP={on_style}', on_style, f'SAP calls {code} "{on_style}" on this style.'
        names = sorted(self.code_names.get(code, []))
        if names:
            if buy and not any(buy == cmp(n) for n in names):
                return f'✗ SAP={names[0]}', names[0], f'SAP calls {code} "{names[0]}" on other articles.'
            return '⚠ NEW ON STYLE', names[0], f'{code} ({names[0]}) is used on other SAP articles but not on this style yet.'
        if code:
            return ('⚠ NOT ON ANY ARTICLE', '',
                    f'No SAP article uses color code {code}. It may still exist in the color master (WRFCHVAL): '
                    'this check only sees articles, so it cannot say whether the code needs creating.')
        return '⚠ NO COLOR CODE', '', 'The row has no color code.'

    def missing_info(self, style, code, size):
        """Why an OS size row is not in SAP: the whole style, the color on this style, or only this size.
        Always says what SAP does hold, so a wrong 'needs creating' can be spotted at a glance."""
        if style not in self.style_colors:
            return 'style', f'Style {style} is not in SAP - the generic needs creating.'
        if code not in self.style_colors[style]:
            norm = lambda c: re.sub(r'[^A-Z0-9]', '', str(c).upper())
            similar = [c for c in self.style_colors[style] if norm(c) == norm(code)]
            hint = f' SAP has {similar[0]} - the same code spelled differently.' if similar else ''
            known = sorted(self.code_names.get(code, []))
            known_note = (f' The code is already in SAP as "{known[0]}" - add it to this style.' if known
                          else ' The code is not on any article in SAP yet.')
            return 'color', f'Color {code} is not on style {style} in SAP.{hint}{known_note} SAP has: {sap_list(self.style_colors[style])}.'
        return 'size', f'Size {size} is missing for color {code} on {style}. SAP has: {sap_list(self.color_sizes[(style, code)])}.'


def sap_list(items, limit=10):
    items = sorted(items, key=lambda x: (SIZE_ORDER.index(x) if x in SIZE_ORDER else 99, x))
    return ', '.join(items[:limit]) + (f' (+{len(items) - limit} more)' if len(items) > limit else '')


def desc_check_text(buy, sap):
    """Style description against SAP's. A SAP name may carry a ' - COLOR' tail; the name without it also matches."""
    if not sap:
        return '⚠ NO DATA'
    return '✓' if buy.upper() in (sap, sap.rsplit(' - ', 1)[0].strip()) else '✗ MISMATCH'


def check_os_row(r, sap):
    """One open-stock prebuy row (style + color) -> one result per size."""
    style, code, desc = r['style'], r['color_code'], r['desc']
    variants = [sap.generic_exists.get((style, code, sz.upper()), {}).get('article') for sz in r['sizes']]
    currency = sap.currency_of(*[a for a in variants if a], style)
    price, msrp = buy_price_for(r, currency), num(r.get('msrp'))
    r['sap_currency'] = currency                                        # remembered: the buysheet writer puts the price in that column

    sap_desc = sap.generic_description(style)
    desc_check = desc_check_text(desc, sap_desc)
    desc_note = desc_diff_note(desc, sap_desc) if desc_check.startswith('✗') else ''

    buy_pc = num(r.get('pc_ctn'))
    def pc_ctn_for(sz):
        """PC/CTN check for ONE size: that variant's own value; the style's majority only when the variant has none."""
        buy_pc = num((r.get('pc_by_size') or {}).get(sz)) or num(r.get('pc_ctn'))
        v = sap.generic_exists.get((style, code, sz.upper()), {})
        sap_pc, src = (v.get('pc_per_ctn'), f"variant {v.get('article')}") if v.get('pc_per_ctn') else (sap.style_pc_ctn.get(style), 'style majority')
        if buy_pc and sap_pc:
            ok = buy_pc == float(sap_pc)
            return ('✓' if ok else f'✗ SAP={int(float(sap_pc))} BUY={int(buy_pc)}'), (f'SAP {int(float(sap_pc))} per carton ({src})' if not ok else f'{int(buy_pc)} per carton')
        if buy_pc: return '⚠ NO SAP PC/CTN', 'SAP has no pieces-per-carton on this variant or style.'
        if sap_pc: return f'⚠ NO BUY PC/CTN (SAP={int(float(sap_pc))})', f'The prebuy has no PC PER CTN; SAP says {int(float(sap_pc))} ({src}).'
        return '—', ''

    total_qty, pc_os = sum(r['sizes'].values()), r.get('pc_ctn_os')
    pc_by_size = r.get('pc_by_size') or {}
    if pc_by_size:                                            # long layout: every size has its own carton, shown on that size's line
        round_check = ''
    elif pc_os and total_qty:
        try: round_check = '✓' if total_qty % float(pc_os) == 0 else f'✗ {total_qty}/{pc_os}'
        except Exception: round_check = '⚠ NO DATA'
    else:
        round_check = '—'

    def round_for(sz, qty):
        pc = pc_by_size.get(sz)
        if not pc: return '—'
        return '✓' if qty % pc == 0 else f'✗ {qty}/{pc}'

    color_check, sap_color, color_note = sap.os_color_check(style, code, r.get('color', ''))
    buy_season = season_text(r.get('season_label'), style)
    sap_season = next((d_['season'] for (st_, _c, _s), d_ in sap.generic_exists.items() if st_ == style and d_.get('season')), '')
    season_check = '—' if not sap_season else '✓' if sap_season.upper() == buy_season.upper() else f'✗ SAP={sap_season}'     # season / year against the style's SAP assignment
    common = {'season_check': season_check, 'sap_season': sap_season,
        **fineline_fields(sap, r, style),
        'style': style, 'color_code': code, 'buy_price': price, 'buy_msrp': msrp, 'currency': currency,
        'color_check': color_check, 'color_note': color_note, 'buy_color': r.get('color', ''), 'sap_color': sap_color,
        'code_in_sap': bool(sap.code_names.get(code)),
        'season': season_text(r.get('season_label'), style), 'season_src': 'header' if r.get('season_label') else 'style number',
        'customer': r.get('customer', ''), 'segment': r.get('segment', ''), 'delivery': iso_date(r.get('delivery')), 'warehouse': r.get('warehouse', ''),
        'buy_desc': desc, 'sap_desc': sap_desc, 'desc_check': desc_check, 'desc_note': desc_note,
        'round_check': round_check,
    }
    # The same color name under ANOTHER color code of this style (a "second BLACK"): its stock is the same product.
    want = normalize_color(r.get('color') or sap_color or '').replace(' ', '')
    alt = {}
    if want and not color_check.startswith('✗'):        # a code SAP names differently is the file's mistake: its color name cannot be used to go looking for other codes
        for (st_, c_, sz_), d_ in sap.generic_exists.items():
            if st_ == style and c_ != code and normalize_color(d_.get('color_name') or '').replace(' ', '') == want and d_.get('article'):
                alt.setdefault(sz_, []).append({'code': c_, 'color': d_.get('color_name') or '', 'article': d_['article']})
    out = []
    for sz, qty in r['sizes'].items():
        exists = (style, code, sz.upper()) in sap.generic_exists
        kind, reason = ('', '') if exists else sap.missing_info(style, code, sz)
        sap_price, source, trace = sap.os_price(style, code, sz)
        sap_msrp = sap.os_msrp(style, code, sz)
        pc_check, pc_note = pc_ctn_for(sz)
        out.append({**common, 'round_check': round_for(sz, qty) if pc_by_size else round_check, 'size': sz, 'qty': qty, 'sap_article': (sap.generic_exists.get((style, code, sz.upper())) or {}).get('article') or '',
                    'status': '✓ EXISTS' if exists else '✗ MISSING', 'alt_articles': alt.get(sz.upper(), []),
                    'pc_ctn_check': pc_check, 'pc_ctn_note': pc_note,
                    'missing_kind': kind, 'missing_reason': reason, 'sap_price': sap_price, 'price_source': source,
                    'price_check': price_check_text(price, sap_price), 'price_trace': trace,
                    'sap_msrp': sap_msrp, 'msrp_check': price_check_text(msrp, sap_msrp)})
    return out


def ppk_component_list(rows):
    """The components a prepack ORDER needs: unique (color code, size, units per pack), in buysheet order. The color code is the
    COLOR CODE column of each row (an article name can carry an extra letter, a digit, or no color at all). Only when that cell is
    blank is the code taken from the article name, and the component says so. A size with no units in the pack is not a component."""
    out, seen = [], set()
    for x in rows:
        code, from_name = str(x.get('color_code') or '').strip().upper(), False
        if not code:
            name = color_from_ppk(x.get('raw_style') or '').strip().upper()
            code, from_name = (name, True) if len(name) >= 3 else ('', True)
        for sz, q in ppk_components(x):
            try: q = int(q)
            except (TypeError, ValueError): continue
            size = str(sz).strip().upper()
            key = (code, canonical_size(size))
            if q <= 0 or key in seen: continue
            seen.add(key)
            out.append({'color': code, 'color_name': str(x.get('color') or '').strip().upper(), 'size': size, 'qty': q, 'code_from_name': from_name})
    return out

def ppk_readiness(sap, style, comps):
    """For a prepack that is NOT in SAP: can it be created now? Every component (color code + size) must already exist as a
    variant of the generic. Says exactly what is missing: the generic, a color, or sizes of a color; and what SAP has instead."""
    known = style in sap.style_known
    checked = []
    for c in comps:
        v = sap.find_variant(style, c['color'], c['size']) if c['color'] else None
        status = 'ok' if v else 'generic' if not known else 'color' if not (c['color'] and sap.has_color(style, c['color'])) else 'size'
        checked.append(dict(c, article=(v or {}).get('article', ''), status=status))
    bad = [c for c in checked if c['status'] != 'ok']
    colors = list(dict.fromkeys(c['color'] for c in checked))
    needs, parts = [], []
    if not known and checked:
        parts.append(f"generic {style} is not in SAP: create it with color{'s' if len(colors) > 1 else ''} {', '.join(colors) or '?'} "
                     f"and size{'s' if len({c['size'] for c in checked}) > 1 else ''} {', '.join(dict.fromkeys(c['size'] for c in checked))}")
    for code in colors:
        mine = [c for c in bad if c['color'] == code]
        if not mine: continue
        name = next((c['color_name'] for c in checked if c['color'] == code and c['color_name']), '')
        sizes = [c['size'] for c in mine]
        if known and not (code and sap.has_color(style, code)):
            parts.append(f"color {code or '(no color code on the row)'} is not on {style}" + (f" (needs sizes {', '.join(sizes)})" if code else ''))
            needs.append({'kind': 'color', 'style': style, 'color': code, 'color_name': name, 'sizes': sizes})
        elif known:
            have = sap.sizes_on(style, code)
            parts.append(f"{style} {code}: size{'s' if len(sizes) > 1 else ''} {', '.join(sizes)} not in SAP" + (f" (SAP has {', '.join(have)})" if have else ''))
            needs.append({'kind': 'size', 'style': style, 'color': code, 'color_name': name, 'sizes': sizes})
        else:
            needs.append({'kind': 'generic', 'style': style, 'color': code, 'color_name': name, 'sizes': sizes})
    return {'state': 'blocked' if bad else 'ready', 'generic_exists': known, 'components': checked, 'needs': needs, 'mdse_cat': sap.mdse_cat.get(style, ''),
            'message': '; '.join(parts) if parts else 'every component exists: ready to create'}


def fineline_fields(sap, r, style):
    """The FINELINE on the buysheet row against HANA's (the article's external material group). '—' when the row has none."""
    buy = norm_code(r.get('fineline'))
    out = {'buy_fineline': buy, 'sap_fineline': '', 'fineline_check': '—'}
    if not buy: return out
    have = sorted(sap.fineline.get(style, set()))
    out['sap_fineline'] = ', '.join(have)
    out['fineline_check'] = '⚠ NO SAP FINELINE' if not have else '✓' if buy in have else f"✗ SAP={', '.join(have)} BUY={buy}"
    return out

def check_ppk_row(r, sap, group=None):
    """One prepack row -> one result. Descriptions are not checked on prepacks (sales abbreviate them);
    the generic's / prepack's SAP name is carried for display only."""
    art, style, code = r['raw_style'], r['style'], r['color_code']
    buy_color, buy_ratio, buy_sizes, desc = r['color'], r.get('ratio', ''), r.get('sizes_str', ''), r['desc']
    currency = sap.currency_of(art, r['raw_style'], style)
    price, msrp = buy_price_for(r, currency), num(r.get('msrp'))
    r['sap_currency'] = currency
    pc_per_pac, pac_per_ctn, ppk_qty = r.get('pc_per_pac'), r.get('pac_per_ctn'), r.get('ppk_qty')

    exists = art in sap.ppk_exists
    detail = sap.ppk_details.get(art, {})
    match_note = ''
    if not exists:                                           # not under that name: look for it by content under the generic
        found = sap.find_ppk_by_content(style, group or [r])
        if found:
            exists, detail = True, sap.detail_for(found)
            match_note = f'Exists in SAP as {found}' + ('' if found == art else f' (the prebuy would call it {art})')
            art = found
    readiness, ready_check = None, '—'
    if not exists:                                           # not in SAP: can it be created now, from variants that already exist?
        readiness = ppk_readiness(sap, style, ppk_component_list(group or [r]))
        ready_check = '✓ READY' if readiness['state'] == 'ready' else '✗ BLOCKED'
    sap_color = detail.get('color_name', '') if detail else ''
    sap_sizes, sap_ratio = build_ratio_sizes(detail) if detail else ('', '')

    part = color_from_ppk(r['raw_style'])
    cc_match = ('✓' if code in part or part in code else '✗ MISMATCH') if (code and part) else '⚠ NO DATA'

    # prepack descriptions are NOT checked (they are often abbreviated); SAP's is shown for information only
    ppk_full = sap.ppk_desc.get(art, '')
    sap_desc = sap.generic_description(style) or (ppk_full.split(' - ')[0].strip() if ppk_full else '')
    desc_check, desc_note = '—', ''

    sap_price = sap.prices.get(art) if art in sap.prices else sap.prices.get(r['raw_style'])
    sap_msrp = sap.msrp.get(art) if art in sap.msrp else sap.msrp.get(r['raw_style'])
    price_trace = f'PPK article {art}: {sap_price:.2f}' if sap_price is not None else f'PPK article {art}: no price row in MDA_SUP_ZSLS'

    color_note = ''
    if not exists: color_check = '—'
    elif buy_color and sap_color:
        color_check = '✓' if strip_color(buy_color) == strip_color(sap_color) else f'✗ SAP={sap_color}'
        if color_check != '✓': color_note = f'SAP\'s components of {art} are "{sap_color}"; the prebuy says "{buy_color}".'
    else: color_check = '⚠ NO DATA'

    ratio_sum = sum(int(x) for x in buy_ratio.split('-') if x.strip().isdigit()) if buy_ratio else None
    buy_pc = float(pc_per_pac) if pc_per_pac else None
    sap_pc = float(detail.get('pc_per_pac') or 0) if detail and detail.get('pc_per_pac') else None
    if ratio_sum and buy_pc and sap_pc:
        pc_ppk_check = '✓' if ratio_sum == buy_pc == sap_pc else f'✗ RATIO={ratio_sum} BUY={int(buy_pc)} SAP={int(sap_pc)}'
    elif ratio_sum and buy_pc:
        pc_ppk_check = '✓' if ratio_sum == buy_pc else f'✗ RATIO={ratio_sum} BUY={int(buy_pc)}'
    else:
        pc_ppk_check = '⚠ NO DATA'

    sizes_check, ratio_check, ratio_note = compare_ratio(exists, buy_sizes, buy_ratio, detail)

    if ppk_qty and pac_per_ctn and pc_per_pac:
        try:
            full = float(pc_per_pac) * float(pac_per_ctn)
            round_check = '✓' if float(ppk_qty) % full == 0 else f'✗ {ppk_qty}/{full}'
        except Exception: round_check = '⚠ NO DATA'
    else:
        round_check = '—'

    if r.get('complex'):          # one article, several colors: check THIS color's components inside the prepack
        cc_match = '—'                                                       # a complex article carries a suffix, not a color code
        comps = (sap.ppks_under_generic.get(art) or {}).get('comps') or detail.get('comps') or []
        mine = [c for c in comps if c[0] == code]
        if not exists:
            color_check = ratio_check = sizes_check = '—'; ratio_note = ''
        elif comps and not mine:
            color_check = sizes_check = ratio_check = f'✗ NOT IN {art}'
            ratio_note = f'SAP prepack {art} holds colors ' + ', '.join(sorted({f"{c[0]} ({c[1]})" for c in comps})) + f'; {code} is not one of them.'
            color_note = ratio_note
        elif mine:
            sap_color = mine[0][1]
            color_check = '✓' if strip_color(buy_color) == strip_color(sap_color) else f'✗ SAP={sap_color}'
            color_note = '' if color_check == '✓' else f'In {art}, color {code} is "{sap_color}"; the prebuy says "{buy_color}".'
            sizes_check, ratio_check, ratio_note = compare_ratio(True, buy_sizes, buy_ratio, {'variants': [(sort, sz, q) for _, _, sz, q, sort in mine]})
            sap_sizes, sap_ratio = '-'.join(c[2] for c in mine), '-'.join(str(c[3]) for c in mine)
        else:
            color_check = sizes_check = ratio_check = '⚠ NO DATA'; ratio_note = f'SAP returned no components for {art}.'
        whole_sap = sum(c[3] for c in comps) if comps else None
        whole_buy = r.get('pc_per_pac')
        if whole_sap and whole_buy:
            pc_ppk_check = '✓' if int(whole_sap) == int(whole_buy) else f'✗ SAP={int(whole_sap)} PREBUY={int(whole_buy)} (whole prepack)'
        else:
            pc_ppk_check = '⚠ NO DATA' if exists else '—'
        # rounding: this color's pieces must be whole packs, and every color of the prepack must give the SAME number of packs
        row_pc = r.get('row_pc') or 0
        packs = (ppk_qty / row_pc) if (ppk_qty and row_pc) else None
        others = [(x['color_code'], x['ppk_qty'] / x['row_pc']) for x in (group or []) if x.get('row_pc') and x.get('ppk_qty')]
        if packs is None: round_check = '—'
        elif abs(packs - round(packs)) > 1e-9: round_check = f'✗ {ppk_qty}/{row_pc} = {packs:.2f} packs'
        elif any(abs(p - packs) > 1e-9 for _, p in others): round_check = '✗ PACKS DIFFER: ' + ', '.join(f'{c} {p:g}' for c, p in others)
        else:
            full = packs / float(pac_per_ctn) if pac_per_ctn else None
            round_check = '✓' if (full is None or abs(full - round(full)) < 1e-9) else f'✗ {packs:g} packs / {pac_per_ctn} per ctn = {full:.2f} cartons'
        if r.get('qty_conflict'):
            round_check = '⚠ QTY DIFFERS'          # the split used the first PPK QTY typed; the rows do not agree

    return {
        'ppk_article': art, 'prebuy_article': r['raw_style'], 'match_note': match_note, 'color_code': code, 'cc_check': cc_match, 'exists': '✓ EXISTS' if exists else '✗ MISSING',
        'complex': bool(r.get('complex')),
        'season': season_text(r.get('season_label'), style), 'season_src': 'header' if r.get('season_label') else 'style number',
        'pc_ppk_check': pc_ppk_check,
        'buy_price': price, 'sap_price': sap_price, 'price_check': price_check_text(price, sap_price), 'price_trace': price_trace, 'buy_msrp': msrp, 'sap_msrp': sap_msrp, 'msrp_check': price_check_text(msrp, sap_msrp), 'currency': currency,
        'buy_color': buy_color, 'sap_color': sap_color, 'color_check': color_check, 'color_note': color_note,
        'buy_sizes': buy_sizes, 'sap_sizes': sap_sizes, 'sizes_check': sizes_check,
        'buy_ratio': buy_ratio, 'sap_ratio': sap_ratio, 'ratio_check': ratio_check, 'ratio_note': ratio_note,
        'buy_desc': desc, 'sap_desc': sap_desc, 'desc_check': desc_check, 'desc_note': desc_note, 'ppk_desc': ppk_full,
        'round_check': round_check, 'ready_check': ready_check, 'readiness': readiness, **fineline_fields(sap, r, style),
        'components': sorted(((sap.ppks_under_generic.get(art) or {}).get('components') or detail.get('components') or []), key=lambda c: (c.get('code', ''), c.get('sort', ''))),
        'qty': sum(r['sizes'].values()) if r.get('sizes') else (ppk_qty or 0),
        'customer': r.get('customer', ''), 'segment': r.get('segment', ''), 'delivery': iso_date(r.get('delivery')), 'warehouse': r.get('warehouse', ''),
    }


GEN_CHECKS = (('status', 'Exists'), ('color_check', 'Color'), ('price_check', 'ZSLS'), ('msrp_check', 'ZVP1'), ('desc_check', 'Description'), ('season_check', 'Season'), ('pc_ctn_check', 'PC / CTN'), ('round_check', 'Rounding'))
PPK_CHECKS = (('exists', 'Exists'), ('cc_check', 'Color code'), ('color_check', 'Color'), ('sizes_check', 'Sizes'), ('ratio_check', 'Ratio'),
              ('pc_ppk_check', 'PC / PPK'), ('price_check', 'ZSLS'), ('msrp_check', 'ZVP1'), ('round_check', 'Rounding'))

def check_counts(results, checks):
    """For each check: how many rows passed (✓), failed (✗) and were only warned about (⚠). '—' (not applicable) is not counted."""
    out = []
    for key, label in checks:
        vals = [str(r.get(key, '')) for r in results]
        out.append({'key': key, 'label': label, 'ok': sum(v.startswith('✓') for v in vals), 'bad': sum(v.startswith('✗') for v in vals),
                    'warn': sum(v.startswith('⚠') for v in vals), 'info': sum(v.startswith('ℹ') for v in vals),
                    'na': sum(v.startswith('—') or v == '' for v in vals)})
    return out

VARY_PER_ROW = ('qty', 'customer', 'segment', 'delivery', 'warehouse', 'season', 'season_src')   # what legitimately differs between rows of one finding

def collapse_results(results, vary=VARY_PER_ROW):
    """One entry per DISTINCT finding. A buysheet has a DATA row per customer / delivery / color, so the same style or
    prepack is on many rows, and each row is checked against SAP. An identical result must not be listed again and again.
    Rows merge only when everything that was compared (the buy values, SAP's values and every check outcome) is the same;
    what differs between them (customer, delivery, quantity ...) is kept as a count and as lists."""
    groups = {}
    for r in results:
        key = json.dumps({k: v for k, v in r.items() if k not in vary}, sort_keys=True, default=str)
        g = groups.get(key)
        if g is None:
            g = groups[key] = dict(r, count=0, qty=0, customers=[], deliveries=[], segments=[])
        g['count'] += 1
        g['qty'] += r.get('qty') or 0
        for field, bucket in (('customer', 'customers'), ('delivery', 'deliveries'), ('segment', 'segments')):
            v = r.get(field)
            if v and v not in g[bucket]:
                g[bucket].append(v)
    return list(groups.values())

READY_CHECK = (('ready_check', 'Can create now'),)
FINELINE_CHECK = (('fineline_check', 'Fineline'),)

def gen_check_list(results, extra=()):
    return GEN_CHECKS + (FINELINE_CHECK if any(str(r.get('fineline_check', '—')) != '—' for r in results) else ()) + tuple(extra)

def ppk_check_list(results, extra=()):
    """The prepack checks the board shows. 'Can create now' only appears when some prepack is not in SAP."""
    return PPK_CHECKS + (READY_CHECK if any(str(r.get('ready_check', '—')) != '—' for r in results) else ()) + (FINELINE_CHECK if any(str(r.get('fineline_check', '—')) != '—' for r in results) else ()) + tuple(extra)

PPK_ISSUE_KEYS = ('cc_check', 'price_check', 'color_check', 'sizes_check', 'ratio_check', 'desc_check', 'round_check', 'pc_ppk_check')

def check_consistency(rows):
    """A style with more than one price, MSRP or fineline in the same file (REM231TS16CB at CAD 8.79 on one line and 9.50 on another).
    Rows are grouped by generic, so a prepack counts under its generic. Pure: no SAP. Returns one finding per style and kind,
    each with its values and where every value was typed (articles, customers, number of rows)."""
    KINDS = ('price_cad', 'price_usd', 'price', 'msrp', 'fineline')
    groups = defaultdict(lambda: {k: defaultdict(list) for k in KINDS})
    money = lambda v: None if num(v) is None else round(num(v), 2)
    for r in rows:
        g = generic_from_ppk(r['raw_style']).upper()
        where = (str(r['raw_style']).upper(), str(r.get('customer') or '').strip())
        pc, pu = money(r.get('price_cad')), money(r.get('price_usd'))
        vals = [('price_cad', pc), ('price_usd', pu)] if (pc is not None or pu is not None) else [('price', money(r.get('price')))]   # a buysheet with both price columns: each on its own
        for kind, v in vals + [('msrp', money(r.get('msrp'))), ('fineline', str(r.get('fineline') or '').strip() or None)]:
            if v is not None: groups[g][kind][v].append(where)
    out = []
    for g in sorted(groups):
        for kind in KINDS:
            vals = groups[g][kind]
            if len(vals) > 1:
                out.append({'style': g, 'kind': kind, 'values': [
                    {'value': v, 'rows': len(w), 'articles': sorted({a for a, _ in w})[:6], 'customers': sorted({c for _, c in w if c})}
                    for v, w in sorted(vals.items(), key=lambda kv: (isinstance(kv[0], str), kv[0]))]})
    return out


def validate_all(rows):
    sap = SapView(rows)
    generic = [res for r in rows if r['type'] == 'OS' for res in check_os_row(r, sap)]
    def group_key(r):
        """The rows that make up ONE order of a complex prepack: same article, same customer, same delivery. The same article
        can be ordered several times in one file (two customers); each order is checked against itself, not against the others."""
        return (r['raw_style'], str(r.get('customer') or ''), str(r.get('delivery') or ''))
    by_art = defaultdict(list)
    for r in rows:
        if r['type'] == 'PPK': by_art[group_key(r)].append(r)
    ppk = [check_ppk_row(r, sap, by_art[group_key(r)]) for r in rows if r['type'] == 'PPK']

    diag = {
        'os_styles_in_file':           len(sap.generic_styles),
        'ppk_articles_in_file':        len(sap.ppk_articles),
        'sap_variants_returned':       len(sap.generic_exists),
        'sap_styles_with_variants':    len({g for (g, c, s) in sap.generic_exists}),
        'sap_styles_with_pc_per_ctn':  len(sap.style_pc_ctn),
        'sap_ppk_articles_found':      len(sap.ppk_exists),
        'price_keys_queried':          len(sap.price_keys),
        'price_rows_returned':         len(sap.prices),
        'os_size_rows':                len(generic),
        'os_size_rows_with_sap_price': sum(1 for r in generic if r['sap_price'] is not None),
        'ppk_rows':                    len(ppk),
        'ppk_rows_with_sap_price':     sum(1 for r in ppk if r['sap_price'] is not None),
        'os_color_codes_not_on_any_article': sorted({r['color_code'] for r in generic if r['color_check'] == '⚠ NOT ON ANY ARTICLE'})[:20],
        'os_missing_by_kind':          {k: sum(1 for r in generic if r['missing_kind'] == k) for k in ('style', 'color', 'size')},
        'desc_from_generic_record':    sum(1 for st in sap.generic_styles if st in sap.generic_desc),
        'desc_from_variants_fallback': sum(1 for st in sap.generic_styles if st not in sap.generic_desc and st in sap.style_desc),
        'rows_with_buy_price':         sum(1 for r in rows if num(r.get('price'))),
        'rows_with_buy_msrp':          sum(1 for r in rows if num(r.get('msrp'))),
        'os_styles_without_sap_price': sorted({r['style'] for r in generic if r['sap_price'] is None})[:20],
        'ppk_without_sap_price':       [r['ppk_article'] for r in ppk if r['sap_price'] is None][:20],
    }
    generic_u, ppk_u = collapse_results(generic), collapse_results(ppk)
    consistency = check_consistency(rows)
    summary = {
        'gen_exists':  sum(1 for r in generic_u if r['status'] == '✓ EXISTS'),
        'gen_missing': sum(1 for r in generic_u if r['status'] == '✗ MISSING'),
        'rows_os_checked': len(generic), 'rows_ppk_checked': len(ppk), 'lines_os': len(generic_u), 'lines_ppk': len(ppk_u),
        'ppk_exists':  len({r['ppk_article'] for r in ppk if r['exists'] == '✓ EXISTS'}),
        'ppk_missing': len({r['ppk_article'] for r in ppk if r['exists'] == '✗ MISSING'}),
        'ppk_issues':  sum(1 for r in ppk_u if any('✗' in str(r.get(k, '')) for k in PPK_ISSUE_KEYS)),
        'total_os':    sum(1 for r in rows if r['type'] == 'OS'),
        'total_ppk':   len({r['ppk_article'] for r in ppk}),
        'checks_ppk':  check_counts(ppk_u, ppk_check_list(ppk_u)),
        'ppk_ready':   len({r['ppk_article'] for r in ppk if r.get('ready_check') == '✓ READY'}),
        'ppk_blocked': len({r['ppk_article'] for r in ppk if r.get('ready_check') == '✗ BLOCKED'}),
        'checks_gen':  check_counts(generic_u, gen_check_list(generic_u)),
        'multi_price': len({c['style'] for c in consistency if c['kind'].startswith('price')}),
        'multi_msrp': len({c['style'] for c in consistency if c['kind'] == 'msrp'}),
        'multi_fineline': len({c['style'] for c in consistency if c['kind'] == 'fineline'}),
    }
    return {'diag': diag, 'generic': generic, 'ppk': ppk, 'generic_unique': generic_u, 'ppk_unique': ppk_u, 'rows': rows, 'summary': summary, 'consistency': consistency}
