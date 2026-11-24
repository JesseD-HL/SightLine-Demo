"""The BUYSHEETS tab: list the production buysheets on the network share, read one into memory, validate it,
and show a master checklist: reqs created (the HEADER sheet's REQ# column) and POs created (looked up in SAP, per requisition).

READ ONLY.  The share holds production files.  The only things done to it are listing a folder and reading a file's bytes
(open(..., 'rb')).  Nothing here writes, renames, moves or deletes anything on the share, and no copy of a buysheet is saved: the
only file this module writes is its own cache of the reqs it read, buysheet_reqs_cache.json, next to server.py.  SAP is only read."""

import json
import os
import re
import io
import threading
import zipfile
import xml.etree.ElementTree as ET
import traceback
from datetime import datetime
from flask import jsonify, request
from .config import BUYSHEET_ROOT, HERE
from .hana import connect_hana
from .parsers import read_header_table
from .reqcheck import REQ_VIEW
from .routes import validate_bytes, validation_payload
from .webapp import app

EXCEL     = ('.xlsx', '.xlsm')


def _inside_share(path):
    root = os.path.normcase(os.path.realpath(BUYSHEET_ROOT))
    real = os.path.normcase(os.path.realpath(path))
    return real == root or real.startswith(root + os.sep)


def _share_path(rel):
    """A path under the share from a '/'-separated relative name; refuses anything that climbs out of it."""
    path = os.path.normpath(os.path.join(BUYSHEET_ROOT, *[p for p in str(rel).replace('\\', '/').split('/') if p]))
    if not _inside_share(path):
        raise ValueError('That path is outside the buysheet folder.')
    return path


def list_buysheets(year=None, quarter=None):
    """{'years': [...], 'year': '2027', 'quarter_names': [...], 'quarter': 'Q3 2027', 'quarters': [{'quarter': 'Q3 2027', 'files': [...]}]}.
    Only the chosen quarter's folder is read (the newest when none is named): the other quarters are just names.  Listing only."""
    if not BUYSHEET_ROOT:
        raise FileNotFoundError('BUYSHEET_ROOT is not set. Add it to config.json (see config.example.json).')
    if not os.path.isdir(BUYSHEET_ROOT):
        raise FileNotFoundError(f'Cannot reach {BUYSHEET_ROOT}. Is the W: drive connected for the account that runs the server?')
    years = sorted((d for d in os.listdir(BUYSHEET_ROOT) if os.path.isdir(os.path.join(BUYSHEET_ROOT, d)) and re.fullmatch(r'\d{4}', d)), reverse=True)
    year = year if year in years else (years[0] if years else None)
    quarters, names = [], []
    if year:
        ydir = os.path.join(BUYSHEET_ROOT, year)
        names = sorted(d for d in os.listdir(ydir) if os.path.isdir(os.path.join(ydir, d)))
        quarter = quarter if quarter in names else (names[-1] if names else None)
        for q in [quarter] if quarter else []:
            files = []
            for e in os.scandir(os.path.join(ydir, q)):
                if e.is_file() and e.name.lower().endswith(EXCEL) and not e.name.startswith('~$'):     # ~$ = Excel's lock file for an open workbook
                    st = e.stat()
                    files.append({'name': e.name, 'rel': f'{year}/{q}/{e.name}', 'size': st.st_size,
                                  'modified': datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M')})
            files.sort(key=lambda f: f['name'].lower())                 # alphabetical
            quarters.append({'quarter': q, 'files': files})
    return {'years': years, 'year': year, 'quarter_names': names, 'quarter': quarter, 'quarters': quarters, 'root': BUYSHEET_ROOT}


def read_from_share(rel):
    """The share file's bytes, read into memory.  Nothing is saved anywhere: the file is validated straight from memory."""
    src = _share_path(rel)
    if not src.lower().endswith(EXCEL) or not os.path.isfile(src):
        raise FileNotFoundError('Not an Excel buysheet on the share.')
    with open(src, 'rb') as f:                                   # read only
        return f.read()


# ── reading only the HEADER sheet of a workbook, without loading the rest ──────────────────────────────────────────────
_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_RNS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'


def _col_index(ref):
    n = 0
    for ch in ref:
        if not ch.isalpha():
            break
        n = n * 26 + ord(ch.upper()) - 64
    return n - 1


class _Sheet:
    """Just enough of an openpyxl sheet for parsers.read_header_table."""
    def __init__(self, rows):
        self.rows = rows

    def iter_rows(self, min_row=1, max_row=None, values_only=True):
        for r in self.rows[min_row - 1: max_row]:
            yield r


class _Book:
    def __init__(self, sheets):
        self.sheets = sheets
        self.sheetnames = list(sheets)

    def __getitem__(self, name):
        return self.sheets[name]


def header_workbook(source):
    """The HEADER sheets of an .xlsx/.xlsm, read straight from the zip: the sheet XML, plus only the shared strings those sheets use.
    A buysheet is several MB, and almost all of it is the DATA sheet, which is never opened."""
    z = zipfile.ZipFile(source)                                  # a read-only file object: only the parts of the file that are needed get read
    rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    wanted = {}
    for s in ET.fromstring(z.read('xl/workbook.xml')).iter(_NS + 'sheet'):
        if s.get('name', '').strip().upper().startswith('HEADER'):
            target = rels[s.get(_RNS + 'id')].lstrip('/')
            wanted[s.get('name')] = target if target.startswith('xl/') else 'xl/' + target
    parsed, need = {}, set()
    for name, path in wanted.items():
        cells = {}
        for c in ET.fromstring(z.read(path)).iter(_NS + 'c'):
            v, t = c.find(_NS + 'v'), c.get('t')
            ref = c.get('r') or ''
            row = int(''.join(ch for ch in ref if ch.isdigit()) or 0)
            if t == 'inlineStr':
                val = ''.join(x.text or '' for x in c.iter(_NS + 't'))
            elif v is None or v.text is None:
                continue
            elif t == 's':
                val = ('s', int(v.text)); need.add(int(v.text))
            elif t in ('str', 'e', 'b'):
                val = v.text
            else:
                try:
                    f = float(v.text); val = int(f) if f == int(f) else f
                except ValueError:
                    val = v.text
            cells[(row, _col_index(ref))] = val
        parsed[name] = cells
    strings = {}
    if need and 'xl/sharedStrings.xml' in z.namelist():
        top, i = max(need), 0
        with z.open('xl/sharedStrings.xml') as f:
            for _, el in ET.iterparse(f):
                if el.tag == _NS + 'si':
                    if i in need:
                        strings[i] = ''.join(x.text or '' for x in el.iter(_NS + 't'))
                    i += 1
                    el.clear()
                    if i > top:
                        break
    sheets = {}
    for name, cells in parsed.items():
        if not cells:
            sheets[name] = _Sheet([]); continue
        height, width = max(r for r, _ in cells), max(c for _, c in cells) + 1
        grid = [[None] * width for _ in range(height)]
        for (r, c), val in cells.items():
            grid[r - 1][c] = strings.get(val[1]) if isinstance(val, tuple) else val
        sheets[name] = _Sheet([tuple(row) for row in grid])
    return _Book(sheets)



REQ_CACHE = os.path.join(HERE, 'buysheet_reqs_cache.json')       # local only (git-ignored): rel -> {sig, result}
_lock = threading.Lock()
_queue, _worker = [], None          # files waiting to be read, and the thread that reads them


def _load_cache():
    try:
        with open(REQ_CACHE, encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


_req_cache = _load_cache()


def _sig(rel):
    st = os.stat(_share_path(rel))
    return [st.st_mtime, st.st_size]


def cached_req(rel):
    """The remembered result for a file, or None when it is new or has changed since it was read."""
    hit = _req_cache.get(rel)
    try:
        return hit['result'] if hit and hit['sig'] == _sig(rel) else None
    except OSError:
        return None


def req_status(rel):
    """Reqs on the HEADER sheet's table: {'customers': n, 'created': n, 'reqs': {customer: number}}.  Reads the share file's bytes only;
    a file is read again only when it changes."""
    done = cached_req(rel)
    if done is not None:
        return done
    src, sig = _share_path(rel), _sig(rel)
    with open(src, 'rb') as f:                                   # read only
        table = read_header_table(header_workbook(f))
    out = {'customers': len(table), 'created': sum(1 for r in table.values() if r), 'lines': table}      # lines: every customer and its req ('' = none yet)
    with _lock:
        _req_cache[rel] = {'sig': sig, 'result': out}
        try:
            with open(REQ_CACHE, 'w', encoding='utf-8') as f:    # our own cache file, next to server.py
                json.dump(_req_cache, f)
        except OSError:
            pass
    return out


def _work():
    global _worker
    while True:
        with _lock:
            if not _queue:
                _worker = None
                return
            rel = _queue.pop(0)
        try:
            req_status(rel)
        except Exception:
            pass                                                 # unreadable right now (open in Excel, share offline): asked for again next poll


def queue_reqs(rels, first=False):
    """Read these files in the background (skipping the ones already known).  first=True puts them ahead of what is already waiting:
    the quarter on screen is read before the rest."""
    global _worker
    with _lock:
        for rel in (reversed(rels) if first else rels):
            if rel in _queue and first:
                _queue.remove(rel)
            if rel not in _queue:
                _queue.insert(0, rel) if first else _queue.append(rel)
        if _queue and _worker is None:
            _worker = threading.Thread(target=_work, daemon=True)
            _worker.start()


def warm_latest_year():
    """At server start: read every buysheet's reqs for the newest year, so the WARP DRIVE tab is already filled in when it is opened."""
    try:
        names = list_buysheets()['quarter_names']
        for name in reversed(names):                              # newest quarter first
            queue_reqs([f['rel'] for q in list_buysheets(quarter=name)['quarters'] for f in q['files'] if cached_req(f['rel']) is None])
    except Exception:
        pass


def po_counts(reqs):
    """{req: {'lines': n, 'with_po': m, 'pos': [PO numbers]}} from SAP's requisition view: a requisition line that carries a purchase
    order number has been turned into a PO.  Deleted lines do not count.  SAP is only read."""
    reqs = sorted({r for r in reqs if r})
    out = {r: {'lines': 0, 'with_po': 0, 'pos': []} for r in reqs}
    if not reqs:
        return out
    conn = connect_hana()
    try:
        cur = conn.cursor()
        live = "IFNULL(TRIM(\"EBAN_DELETION_INDICATOR\"), '') = ''"
        has_po = "IFNULL(TRIM(\"EBAN_PURCHASE_ORDER_NO\"), '') <> ''"
        for i in range(0, len(reqs), 500):
            part = reqs[i:i + 500]
            marks = ', '.join('?' for _ in part)
            cur.execute(f'SELECT "EBAN_REQUISITION_NO", COUNT(*), SUM(CASE WHEN {has_po} THEN 1 ELSE 0 END) FROM {REQ_VIEW} '
                        f'WHERE "EBAN_REQUISITION_NO" IN ({marks}) AND {live} GROUP BY "EBAN_REQUISITION_NO"', tuple(part))
            for req, n, m in cur.fetchall():
                if req in out:
                    out[req]['lines'], out[req]['with_po'] = int(n or 0), int(m or 0)
            cur.execute(f'SELECT DISTINCT "EBAN_REQUISITION_NO", "EBAN_PURCHASE_ORDER_NO" FROM {REQ_VIEW} '
                        f'WHERE "EBAN_REQUISITION_NO" IN ({marks}) AND {live} AND {has_po} ORDER BY 1, 2', tuple(part))
            for req, po in cur.fetchall():
                if req in out:
                    out[req]['pos'].append(str(po).strip())
    finally:
        conn.close()
    return out


@app.route('/buysheets/list', methods=['GET'])
def buysheets_list():
    try:
        out = list_buysheets(request.args.get('year'), request.args.get('quarter'))
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    return jsonify(out)


@app.route('/buysheets/validate', methods=['POST'])
def buysheets_validate():
    rel = (request.get_json(silent=True) or {}).get('rel', '')
    try:
        validation = validate_bytes(read_from_share(rel), os.path.basename(rel))
    except (ValueError, FileNotFoundError) as e:
        return jsonify({'error': str(e)}), 400
    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500
    return jsonify(validation_payload(validation))


@app.route('/buysheets/reqs', methods=['POST'])
def buysheets_reqs():
    """Req status for the given files, from what has been read so far: a customer line with a req number on the HEADER sheet = created.
    Files not read yet are read in the background; '_pending' says how many are still to come, so the page asks again."""
    out, missing = {}, []
    for rel in (request.get_json(silent=True) or {}).get('rels', []):
        try:
            hit = cached_req(rel)
        except ValueError as e:
            out[rel] = {'error': str(e)[:120]}
            continue
        if hit is None:
            missing.append(rel)
        else:
            out[rel] = hit
    queue_reqs(missing, first=True)
    out['_pending'] = len(missing)
    return jsonify(out)


@app.route('/buysheets/pos', methods=['POST'])
def buysheets_pos():
    """PO status per buysheet: for every customer line that has a req on the HEADER sheet, how many of that requisition's lines SAP has
    turned into purchase orders.  Files whose reqs have not been read yet are left out ('_pending' says how many)."""
    rels = (request.get_json(silent=True) or {}).get('rels', [])
    known, pending = {}, 0
    for rel in rels:
        try:
            hit = cached_req(rel)
        except ValueError:
            hit = None
        if hit is None:
            pending += 1
        else:
            known[rel] = hit
    try:
        sap = po_counts({r for hit in known.values() for r in hit['lines'].values() if r})
    except Exception as e:
        return jsonify({'error': str(e)[:200]}), 500
    out = {'_pending': pending}
    for rel, hit in known.items():
        lines, done = {}, 0
        for cust, req in hit['lines'].items():
            c = sap.get(req) if req else None
            lines[cust] = {'req': req, 'lines': c['lines'] if c else 0, 'with_po': c['with_po'] if c else 0, 'pos': c['pos'] if c else []}
            done += bool(c and c['lines'] and c['with_po'] == c['lines'])
        out[rel] = {'customers': len(lines), 'done': done, 'lines': lines}
    return jsonify(out)


threading.Thread(target=warm_latest_year, daemon=True).start()      # fills the cache while the server starts

