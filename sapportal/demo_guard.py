"""Keeps the demo closed. Imported last, so every route is registered before the guard looks at them.

 * Only an allow-list of read-only, demo-safe endpoints answers. Everything else (anything that could drive a browser, SAP, a file share or a
   subprocess) is 404, whether or not the code behind it exists in this copy.
 * Only requests for localhost are answered (the Host header is checked, which also stops DNS-rebinding), and a browser request that comes from
   another website is refused (the Origin / Referer is checked on every request that changes state or downloads a file).
 * Uploads are capped, responses carry no-store and anti-framing headers, errors never show a stack trace or a file path.
 * No credential exists anywhere in this program: the database is invented data in memory."""

import re
from flask import Response, jsonify, request
from .webapp import app

app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024          # 8 MB per upload

ALLOWED = {'/', '/health', '/validate', '/revalidate', '/validation', '/stock', '/stock/clear', '/creation_templates', '/validation_report', '/generate',
           '/req_csv', '/delivery_flow/preview', '/delivery_flow/download', '/hold_check', '/hold_check/export', '/delivery_split',
           '/dpci_list/state', '/dpci_list/check', '/dpci_list/load', '/dpci_list/refresh', '/dpci_list/export', '/dpci_list/export_list'}
LOCAL_HOSTS = re.compile(r'^(localhost|127\.0\.0\.1|\[::1\])(:\d{1,5})?$', re.I)


def _host_of(url):
    m = re.match(r'^[a-z]+://([^/]+)', url or '', re.I)
    return m.group(1) if m else ''


@app.before_request
def guard():
    if not LOCAL_HOSTS.match(request.host or ''):
        return Response('This demo only answers on localhost.', status=403, mimetype='text/plain')
    if request.path not in ALLOWED:
        return jsonify({'error': 'Not available in the demo.'}), 404
    origin = request.headers.get('Origin') or request.headers.get('Referer') or ''
    if origin and not LOCAL_HOSTS.match(_host_of(origin)):
        return Response('Cross-site requests are refused.', status=403, mimetype='text/plain')
    return None


@app.after_request
def harden(resp):
    resp.headers['Cache-Control'] = 'no-store'
    resp.headers['X-Content-Type-Options'] = 'nosniff'
    resp.headers['X-Frame-Options'] = 'DENY'
    resp.headers['Referrer-Policy'] = 'no-referrer'
    resp.headers.pop('Access-Control-Allow-Origin', None)       # no cross-origin reads, ever
    return resp


@app.errorhandler(413)
def too_big(_):
    return jsonify({'error': 'That file is larger than the demo accepts (8 MB).'}), 413


@app.errorhandler(Exception)
def no_details(exc):
    code = getattr(exc, 'code', 500) if hasattr(exc, 'code') else 500
    return jsonify({'error': 'The demo could not do that.' if code >= 500 else str(exc)}), code if isinstance(code, int) else 500
