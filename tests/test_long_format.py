"""A buysheet whose DATA tab is one row per size (SIZE + QTY columns) parses into the same rows as the wide layout."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from openpyxl import Workbook
import server

wb = Workbook(); ws = wb.active; ws.title = 'DATA'
ws.append(['CUSTOMER', 'STYLE', 'COLOR CODE', 'COLOR', 'DESCRIPTION', 'DELIVERY', 'PRICE', 'MSRP', 'SIZE ', 'QTY ', 'PC PER CTN OS'])
for sz, q, c in [('S', 18, 6), ('M', 60, 15), ('L', 84, 14)]:
    ws.append(['MAPLES.COM - MEN - OS', 'RVM273HD12', 'S143', 'BLACK', 'HOODIE', None, 35, 100, sz, q, c])
ws.append(['MAPLES_EAST - MEN - OS', 'RVM273HD12', 'S143', 'BLACK', 'HOODIE', None, 35, 100, 'S', 6, 6])
rows, customers = server.parse_buysheet(wb)
assert len(rows) == 2, len(rows)                               # one row per customer, sizes folded back together
r = rows[0]
assert r['sizes'] == {'S': 18, 'M': 60, 'L': 84} and r['size_order'] == ['S', 'M', 'L'], r['sizes']
assert rows[1]['sizes'] == {'S': 6}
print('long format parse ok')
