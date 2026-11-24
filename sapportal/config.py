"""Settings and constants: where HANA is, what a size looks like, season codes, the build stamp, the shared cache."""

import os


# ═══════════════════════════════════════════════════════════════════════════════
# 1. CONFIG + CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))      # the folder that holds server.py (config.json, portal.html, templates/ live there)

_SETTINGS, SETTINGS_SOURCE = {}, 'the built-in demo'
HANA_HOST = HANA_USER = HANA_PASSWORD = ''
HANA_PORT = 0
STOCK_VIEW    = 'SAPDEMO.ZSD_CUTSOLD_STK'              # live stock: the invented stock view
BUYSHEET_ROOT = ''                                     # no file share in the demo
TOOLS_DIR     = HERE                                   # where portal.html lives

# Bump whenever this file changes. The portal shows it next to its own PAGE_BUILD so a stale file is obvious.
BUILD = "SIGHTLINE // DEMO-1"

_cache = {}          # the last validation: rows, customers, validation (one session, one file at a time)

KNOWN_SIZES = {'XS','S','M','L','XL','XXL','2XL','3XL','4XL','5XL','6XL',
               '1X','2X','3X','4X','5X','OS','ONE SIZE','LT','XLT','2XLT','3XLT','4XLT','5XLT'}
SIZE_ORDER  = ['XS','S','M','L','XL','XXL','2XL','3XL','4XL','5XL','6XL',
               '1X','2X','3X','4X','5X','LT','XLT','2XLT','3XLT','4XLT','5XLT','OS','ONE SIZE']

# Only SPELLED-OUT sizes are translated ('SMALL' -> 'S'). Labels such as 2XL / XXL are never rewritten:
# every size is used exactly as the prebuy types it, because SAP's code for a size depends on the brand.
SIZE_ALIASES = {'SMALL':'S','MEDIUM':'M','LARGE':'L',
                'X SMALL':'XS','X-SMALL':'XS','EXTRA SMALL':'XS',
                'X LARGE':'XL','X-LARGE':'XL','EXTRA LARGE':'XL',
                'XX LARGE':'XXL','XX-LARGE':'XXL'}

SEASON_LABEL = {'1':'Q1','2':'Q2','3':'Q3','4':'Q4'}                       # style[5] -> quarter
SEASON_NAME  = {'1':'SPR','2':'SUM','3':'FALL','4':'HOL'}                  # Q1 SPR, Q2 SUM, Q3 FALL, Q4 HOL
SEASON_WORD  = {'SPR':1,'SPRING':1,'SUM':2,'SUMMER':2,'FALL':3,'AUTUMN':3,'HOL':4,'HOLIDAY':4}

GENDER_PREFIX = {
    'RVM':'MEN','RDM':'MEN','REM':'MEN','CPM':'MEN','CLM':'MEN','ANM':'MEN','HEB':'MEN',
    'RVW':'LADIES','RDW':'LADIES','REW':'LADIES','CPW':'LADIES','CLW':'LADIES','HEW':'LADIES',
    'RVB':'BOYS','RDB':'BOYS','REB':'BOYS','CPB':'BOYS','REG':'BOYS','RDG':'BOYS',
    'CAM':'MEN','CFW':'LADIES','CAB':'BOYS',
}
# weights / volume per (gender, 2-digit product code) for the creation templates
WEIGHTS_LOOKUP = {
    ('MEN',    'OW'): {'HOODIE':(1.7,0.16),'HOOD':(1.7,0.16),'DEFAULT':(1.1,0.067)},
    ('MEN',    'HD'): {'DEFAULT':(1.7,0.16)},
    ('MEN',    'FH'): {'DEFAULT':(1.7,0.16)},
    ('MEN',    'TL'): {'DEFAULT':(0.8,0.16)},
    ('MEN',    'TS'): {'POLO':(0.8,0.08),'DEFAULT':(0.8,0.067)},
    ('MEN',    'SS'): {'DEFAULT':(0.8,0.067)},
    ('MEN',    'TE'): {'DEFAULT':(0.8,0.067)},
    ('MEN',    'TX'): {'DEFAULT':(0.8,0.067)},
    ('MEN',    'JK'): {'DEFAULT':(1.1,0.16)},
    ('MEN',    'PA'): {'FLEECE':(1.4,0.16),'DEFAULT':(1.03,0.067)},
    ('MEN',    'SH'): {'DEFAULT':(0.45,0.16)},
    ('LADIES', 'OW'): {'HOODIE':(1.4,0.13),'HOOD':(1.4,0.13),'DEFAULT':(0.9,0.067)},
    ('LADIES', 'HD'): {'DEFAULT':(1.4,0.13)},
    ('LADIES', 'FH'): {'DEFAULT':(1.4,0.13)},
    ('LADIES', 'TL'): {'DEFAULT':(0.8,0.08)},
    ('LADIES', 'TS'): {'DEFAULT':(0.3968,0.067)},
    ('LADIES', 'SS'): {'DEFAULT':(0.3968,0.067)},
    ('LADIES', 'TE'): {'DEFAULT':(0.3968,0.067)},
    ('LADIES', 'TX'): {'DEFAULT':(0.3527,0.067)},
    ('LADIES', 'BR'): {'TANK':(0.3527,0.067),'DEFAULT':(0.23,0.067)},
    ('LADIES', 'JK'): {'DEFAULT':(1.1,0.067)},
    ('LADIES', 'SD'): {'SKIRT':(0.31,0.067),'DEFAULT':(0.6,0.067)},
    ('LADIES', 'LE'): {'DEFAULT':(0.5732,0.067)},
    ('LADIES', 'PA'): {'FLEECE':(0.77,0.1),'DEFAULT':(0.77,0.1)},
    ('LADIES', 'SH'): {'DEFAULT':(0.35,0.067)},
    ('BOYS',   'OW'): {'DEFAULT':(1.36,0.128)},
    ('BOYS',   'HD'): {'DEFAULT':(1.36,0.128)},
    ('BOYS',   'FH'): {'DEFAULT':(1.36,0.128)},
    ('BOYS',   'TL'): {'DEFAULT':(0.64,0.064)},
    ('BOYS',   'TS'): {'DEFAULT':(0.64,0.064)},
    ('BOYS',   'SS'): {'DEFAULT':(0.64,0.064)},
    ('BOYS',   'TE'): {'DEFAULT':(0.64,0.064)},
    ('BOYS',   'TX'): {'DEFAULT':(0.64,0.064)},
    ('BOYS',   'JK'): {'DEFAULT':(0.88,0.128)},
    ('BOYS',   'PA'): {'FLEECE':(1.12,0.112),'DEFAULT':(0.824,0.0536)},
    ('BOYS',   'SH'): {'DEFAULT':(0.36,0.0536)},
}

def get_weights(style, description=''):
    """(weight, volume) for a style from the lookup above, keyed by gender prefix and product code."""
    style = str(style).strip().upper()
    desc  = str(description).strip().upper()
    gender = GENDER_PREFIX.get(style[:3])
    lookup = WEIGHTS_LOOKUP.get((gender, style[6:8])) if gender else None
    if not lookup:
        return None, None
    for kw, vals in lookup.items():
        if kw != 'DEFAULT' and kw in desc:
            return vals
    return lookup.get('DEFAULT', (None, None))
