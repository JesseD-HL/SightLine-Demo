"""Small pure functions: numbers, dates, seasons, ratios. No I/O."""

import re
from datetime import datetime
from .config import SEASON_LABEL, SEASON_NAME, SEASON_WORD, SIZE_ALIASES, SIZE_ORDER


# ═══════════════════════════════════════════════════════════════════════════════
# 2. HELPERS - small, pure, no I/O
# ═══════════════════════════════════════════════════════════════════════════════
def num(v):
    """Tolerant number reader: '16.96', '$16.96', '1,200' -> float; blanks and text ('TBD') -> None."""
    if v is None or v == '': return None
    try: return float(str(v).replace('$', '').replace(',', '').strip())
    except ValueError: return None

def up(v):
    """Everything written to the buysheet is upper case."""
    return v.upper() if isinstance(v, str) else v

def iso_date(d):
    return d.strftime('%Y-%m-%d') if hasattr(d, 'strftime') else (str(d)[:10] if d else '')

def mmddyyyy(d):
    """Delivery date as MMDDYYYY for the PR file; empty (never garbage) when it cannot be read."""
    if not d: return ''
    if hasattr(d, 'strftime'): return d.strftime('%m%d%Y')
    t = str(d).strip()
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y', '%Y%m%d', '%m%d%Y'):
        try: return datetime.strptime(t, fmt).strftime('%m%d%Y')
        except ValueError: pass
    return ''

def parse_season(text):
    """'Q2 2027', 'Q2-2027', 'q2/2027', '2027 SUM' or 'SUM 2027' -> 'Q2 2027' ('' when it cannot be read)."""
    t = str(text or '').strip().upper()
    m = re.search(r'\bQ\s*([1-4])\s*[-/ ]?\s*(\d{4})\b', t)
    if m:
        return f"Q{m.group(1)} {m.group(2)}"
    m = re.search(r'\b(20\d{2})\s*[-/ ]?\s*([A-Z]{3,7})\b', t) or re.search(r'\b([A-Z]{3,7})\s*[-/ ]?\s*(20\d{2})\b', t)
    if m:
        year, word = (m.group(1), m.group(2)) if m.group(1).isdigit() else (m.group(2), m.group(1))
        if word in SEASON_WORD:
            return f"Q{SEASON_WORD[word]} {year}"
    return ''

def season_text(label, style):
    """'2027 SUM' for the PR file: from the header season label when there is one, else from the style number."""
    m = re.fullmatch(r'Q([1-4]) (\d{4})', label or '')
    if m:
        return f"{m.group(2)} {SEASON_NAME[m.group(1)]}"
    if no_season_style(style): return f"{NO_SEASON_YEAR} {NO_SEASON_CODE}"
    s_num = style[5] if len(style) > 5 else '1'
    year = '20' + style[3:5] if len(style) > 4 else '2027'
    return f"{year} {SEASON_NAME.get(s_num, 'SPR')}"

# A style with 000 where the year and quarter sit (CPU000FH01) carries no season: it defaults to season 0000, year 9999.
NO_SEASON_CODE, NO_SEASON_YEAR = '0000', '9999'
def no_season_style(style):
    return str(style or '')[3:6] == '000'

def season_from_style(style):
    """'Q2-2027' label for the buysheet's top row, from the style number ('' for a style that carries no season)."""
    if no_season_style(style): return ''
    return f"{SEASON_LABEL.get(style[5], 'Q1')}-20{style[3:5]}" if len(style) > 5 else ''

def canonical_size(raw):
    r = str(raw).strip().upper()
    return SIZE_ALIASES.get(r, r)

def base_style(raw):
    """'SFW271SH01-S470' -> 'SFW271SH01'."""
    return str(raw).strip().split('-')[0]

generic_from_ppk = base_style

def color_from_ppk(raw):
    parts = str(raw).strip().split('-')
    return parts[1] if len(parts) > 1 else ''

def is_ppk(raw):
    return '-' in str(raw).strip() if raw else False

def strip_color(c):
    return re.sub(r'\s*\(.*?\)', '', str(c)).replace(' ', '').upper()

def normalize_color(c):
    c = re.sub(r'\s*\(.*?\)', '', str(c)).strip().upper()
    c = re.sub(r'\s+', ' ', c)
    return re.sub(r'\bHTR\b|\bHTHR\b', 'HEATHER', c).strip()

def build_ratio_sizes(detail):
    """SAP's components for a prepack -> ('XS-S-M', '1-3-4') in SAP's size order."""
    if not detail or not detail.get('variants'):
        return '', ''
    sv = sorted(detail['variants'], key=lambda x: x[0])
    return ('-'.join(str(v[1]) for v in sv),
            '-'.join(str(int(v[2])) if v[2] is not None else '?' for v in sv))

def dash_list(v):
    """'1-3-4-3-2-2', '1, 3, 4' or '1 3 4' -> ['1','3','4'] (upper case, no blanks).
    A real date means Excel turned something like 1-2 into a date: that cannot be read back, so it gives nothing."""
    if v is None or hasattr(v, 'strftime'):
        return []
    return [x for x in re.sub(r'[\s,;/]+', '-', str(v).strip().upper()).split('-') if x]

def split_by_ratio(total, weights):
    """Split a total across components in proportion to their ratio, so the parts always add back to
    exactly the total (largest-remainder rounding: no piece is lost or invented)."""
    whole = sum(weights)
    if total is None or not whole:
        return [None] * len(weights)
    total = int(round(total))
    raw = [total * w / whole for w in weights]
    parts = [int(x) for x in raw]
    for i in sorted(range(len(weights)), key=lambda i: raw[i] - parts[i], reverse=True)[:max(total - sum(parts), 0)]:
        parts[i] += 1
    return parts

def price_check_text(buy, sap):
    """Price check cell. Says WHICH side is missing instead of a generic 'no data'."""
    if buy and sap:
        return '✓' if round(buy, 2) == round(sap, 2) else f'✗ SAP={sap:.2f} BUY={buy:.2f}'
    if buy: return '⚠ NO SAP PRICE'
    if sap: return '⚠ NO BUY PRICE'
    return '—'


def compare_ratio(exists, buy_sizes, buy_ratio, detail):
    """The file's sizes + ratio against the components SAP holds for the prepack, size by size (the order typed does not matter).
    Returns (sizes check, ratio check, a plain-English note saying exactly which sizes differ)."""
    if not exists:
        return '—', '—', ''
    hana = {}
    for _, sz, q in (detail or {}).get('variants', []):
        try:
            hana[str(sz).strip().upper()] = int(round(float(q)))
        except (TypeError, ValueError):
            pass
    if not (buy_sizes and buy_ratio and hana):
        return '⚠ NO DATA', '⚠ NO DATA', ''
    szs = [x for x in buy_sizes.split('-') if x]
    rts = [x for x in buy_ratio.split('-') if x]
    if len(szs) != len(rts):
        return '✗ MISMATCH', '✗ MISMATCH', f"The file has {len(rts)} ratio number{'' if len(rts) == 1 else 's'} but {len(szs)} size{'' if len(szs) == 1 else 's'}."
    if len(set(szs)) != len(szs):
        return '✗ MISMATCH', '✗ MISMATCH', 'A size is listed more than once.'
    try:
        buy = {sz: int(rt) for sz, rt in zip(szs, rts)}
    except ValueError:
        return '✗ MISMATCH', '✗ MISMATCH', 'The ratio has something that is not a whole number.'
    order = sorted(set(buy) | set(hana), key=lambda x: (SIZE_ORDER.index(x) if x in SIZE_ORDER else 99, x))
    diffs = [f"{sz}: file {buy.get(sz, 'none')}, SAP {hana.get(sz, 'none')}" for sz in order if buy.get(sz) != hana.get(sz)]
    return ('✓' if set(buy) == set(hana) else '✗ MISMATCH'), ('✓' if not diffs else '✗ MISMATCH'), '; '.join(diffs)

def desc_diff_note(buy, sap):
    """When two descriptions differ only in punctuation or spacing (7/8" vs 7/8), say so."""
    squash = lambda t: re.sub(r'[^A-Z0-9]', '', str(t).upper())
    return 'differs only in punctuation or spacing' if buy and sap and squash(buy) == squash(sap) else ''


def pad10(v):
    """A PO or requisition number as SAP stores it: digits only, zero-padded to 10."""
    d = re.sub(r'\D', '', str(v or ''))
    return d.zfill(10) if d else ''
