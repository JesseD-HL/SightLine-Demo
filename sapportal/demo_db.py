"""The demo's stand-in for SAP HANA: a small in-memory SQLite database made of invented data, with the same view and column names the
real queries use. Nothing here talks to a network, and there is no credential anywhere in this program.

Invented company: Northwind Apparel, brand SUMMIT. Invented customers: Bigbox, Clubmart, Maple's, Birch, Dalton's, and so on."""

import random
import sqlite3
import threading

ARTMAS = '"_SYS_BIC"."Acme.Data_Universe.Raw_Complete/MDA_RAW_ARTMAS"'
ZSLS = '"_SYS_BIC"."Acme.Data_Universe.Raw_Supporting/MDA_SUP_ZSLS"'
CONDITIONS = '"_SYS_BIC"."Acme.Data_Universe.Raw_Supporting/MDA_SUP_ARTMAS_PRICING_CONDITIONS"'

SIZES = ['S', 'M', 'L', 'XL', 'XXL']
SORT = {'XS': '10', 'S': '20', 'M': '30', 'L': '40', 'XL': '50', 'XXL': '60'}
COLORS = {'S100': 'BLACK', 'S102': 'BLACK', 'S110': 'NAVY', 'S120': 'OXFORD GREY HEATHER', 'S130': 'WHITE', 'S140': 'CRIMSON', 'S150': 'FOREST GREEN',
          'S160': 'GRANITE HEATHER', 'S165': 'GRANITE GREY HEATHER', 'S170': 'DEEP PLUM', 'S180': 'SKY BLUE'}
# generic -> (description, price, msrp, pieces per carton, season, colors on the style)
STYLES = {
    'SPM273HD01': ('POWERBLEND PULLOVER HOODIE "S" LOGO', 16.75, 50, 24, ['S100', 'S110', 'S120', 'S140']),
    'SPM273FZ02': ('POWERBLEND FULL ZIP HOODIE', 18.50, 55, 24, ['S100', 'S110', 'S150', 'S160']),
    'SPM273PA03': ('POWERBLEND OPEN HEM SWEATPANTS', 14.25, 45, 36, ['S100', 'S110', 'S120']),
    'SPM273TS04': ('JERSEY CREWNECK TEE', 7.00, 20, 48, ['S100', 'S130', 'S110', 'S140', 'S180']),
    'SPM273SH05': ('COMMON GROUND CREW SWEATSHIRT', 15.00, 45, 24, ['S120', 'S100', 'S102', 'S150']),
    'SPW273HD11': ('WOMENS POWERBLEND HOODIE', 17.25, 55, 24, ['S100', 'S120', 'S170', 'S165']),
    'SPW273LG12': ('WOMENS HIGH WAIST LEGGING', 12.50, 40, 48, ['S100', 'S110', 'S170']),
    'SPW273QZ13': ('WOMENS QUARTER ZIP TOP', 19.00, 60, 24, ['S100', 'S130', 'S170']),
}
PREPACKS = {            # prepack article -> (generic, color code, ratio by size)
    'SPM273TS04-S110': ('SPM273TS04', 'S110', {'S': 1, 'M': 2, 'L': 2, 'XL': 1}),
    'SPM273TS04-S100': ('SPM273TS04', 'S100', {'S': 1, 'M': 2, 'L': 2, 'XL': 1}),
}
SLOCS = [('H001', 'OTH_H_RFID', 'NJ01'), ('H001', 'ONL', 'ST01'), ('H001', 'OTH', 'GA01'), ('H001', 'CSM', 'NJ01')]

_lock = threading.Lock()
_conn = None


def _build():
    rnd = random.Random(2027)
    db = sqlite3.connect(':memory:', check_same_thread=False)
    for schema in ('_SYS_BIC', 'SAPDEMO'):
        db.execute(f'ATTACH DATABASE \':memory:\' AS "{schema}"')
    db.execute(f'''CREATE TABLE {ARTMAS} ("MARA_ARTICLE" TEXT, "F_GENERIC" TEXT, "MARA_COLOR_CODE" TEXT, "MARA_SIZE_CODE" TEXT, "F_GENERIC_DESCRIPTION" TEXT,
        "WRF_SORT_POSITION" TEXT, "F_PC_PER_CAR_FIX" REAL, "CHARVAL_COLOR_CODE_DESCRIPTION" TEXT, "F_SEASON" TEXT, "F_YEAR" TEXT, "MARA_EXT_MERCH_CAT" TEXT,
        "MARA_MERCHANDISE_CATEGORY" TEXT, "F_ARTICLE_TYPE" TEXT, "MAST_ARTICLE" TEXT, "STPO_PPK_QTY" REAL, "MARM_PC_PER_PAC" REAL,
        "MAKT_ARTICLE_DESCRIPTION" TEXT, "MARA_UPC" TEXT)''')
    db.execute(f'CREATE TABLE {ZSLS} ("ARTICLE" TEXT, "F_AMOUNT_PER_PC" REAL, "VALIDITY_START_DATE" TEXT, "VALIDITY_END_DATE" TEXT)')
    db.execute(f'''CREATE TABLE {CONDITIONS} ("ARTICLE" TEXT, "CUSTOMER" TEXT, "KONP_AMOUNT" REAL, "KONP_PER" REAL, "VALID_FROM" TEXT, "VALID_TO" TEXT,
        "CONDITION_TYPE" TEXT)''')
    for t in ('ZSAC_ZSLS_A504', 'ZSAC_ZSLS_A996'):
        db.execute(f'CREATE TABLE "SAPDEMO"."{t}" ("MATNR" TEXT, "KONWA" TEXT)')
    db.execute('''CREATE TABLE "SAPDEMO"."ZSD_CUTSOLD_STK" ("MATNR" TEXT, "WERKS" TEXT, "SGT_SCAT" TEXT, "LGORT" TEXT, "CLABS" REAL, "CUMLM" REAL, "CINSM" REAL,
        "CEINM" REAL, "CSPEM" REAL, "CRETM" REAL)''')
    db.execute('CREATE TABLE "SAPDEMO"."KNMT" ("MATNR" TEXT, "KDMAT" TEXT, "KUNNR" TEXT)')

    upc = 810000000000
    for g, (desc, price, msrp, pcctn, codes) in STYLES.items():
        db.execute(f'INSERT INTO {ARTMAS} ("MARA_ARTICLE","F_GENERIC","F_GENERIC_DESCRIPTION","F_ARTICLE_TYPE","F_SEASON","F_YEAR","MARA_EXT_MERCH_CAT",'
                   '"MARA_MERCHANDISE_CATEGORY","MAKT_ARTICLE_DESCRIPTION") VALUES (?,?,?,?,?,?,?,?,?)',
                   (g, g, desc, 'G', 'FALL', '2027', '1000', 'SUMMIT-M' if g[2] == 'M' else 'SUMMIT-W', desc))
        db.execute(f'INSERT INTO {ZSLS} VALUES (?,?,?,?)', (g, price, '20260101', '20271231'))
        db.execute(f'INSERT INTO {CONDITIONS} VALUES (?,?,?,?,?,?,?)', (g, '', msrp, 1, '20260101', '20271231', 'ZVP1'))
        n = 0
        for code in codes:
            for size in SIZES:
                n += 1; upc += 7
                art = f'{g}{n:03d}'
                db.execute(f'INSERT INTO {ARTMAS} VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                           (art, g, code, size, desc, SORT[size], pcctn, COLORS[code], 'FALL', '2027', '1000', 'SUMMIT-M' if g[2] == 'M' else 'SUMMIT-W',
                            'V', None, None, None, desc + ' ' + COLORS[code] + ' ' + size, str(upc)))
                db.execute(f'INSERT INTO {ZSLS} VALUES (?,?,?,?)', (art, price, '20260101', '20271231'))
                db.execute(f'INSERT INTO {CONDITIONS} VALUES (?,?,?,?,?,?,?)', (art, '', msrp, 1, '20260101', '20271231', 'ZVP1'))
                db.execute('INSERT INTO "SAPDEMO"."ZSAC_ZSLS_A504" VALUES (?,?)', (art, 'USD'))
                db.execute('INSERT INTO "SAPDEMO"."KNMT" VALUES (?,?,?)', (art, f'{rnd.randint(1, 99):03d}  {rnd.randint(1, 99):02d}  {rnd.randint(1, 9999):04d}', 'BIGBOX'))
                if rnd.random() < 0.7:                                    # about a third of the articles have no stock at all
                    for werks, seg, sloc in SLOCS:
                        if rnd.random() < 0.6:
                            inv = rnd.choice([0, 24, 48, 96, 144, 240, 480, 960, 1440])
                            if inv:
                                db.execute('INSERT INTO "SAPDEMO"."ZSD_CUTSOLD_STK" VALUES (?,?,?,?,?,?,?,?,?,?)',
                                           (art, werks, seg, sloc, inv, rnd.choice([0, 0, 0, 12]), 0, 0, rnd.choice([0, 0, 0, 6]), 0))
    for ppk, (g, code, ratio) in PREPACKS.items():
        for size, qty in ratio.items():
            db.execute(f'INSERT INTO {ARTMAS} VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                       (f'{g}{code}{size}', g, code, size, STYLES[g][0], SORT[size], STYLES[g][3], COLORS[code], 'FALL', '2027', '1000', 'SUMMIT-M',
                        'C', ppk, qty, sum(ratio.values()), STYLES[g][0] + ' PREPACK', None))
    db.commit()
    return db


class _Conn:
    """What the rest of the program sees as a HANA connection: cursors on the shared invented database, and a close() that closes nothing."""
    def cursor(self):
        return _conn.cursor()

    def close(self):
        pass

    def commit(self):
        pass


def connect():
    global _conn
    with _lock:
        if _conn is None:
            _conn = _build()
    return _Conn()
