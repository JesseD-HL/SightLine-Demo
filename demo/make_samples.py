"""Writes the demo's sample files into demo/samples/: an open-stock buysheet, a hold master and a DPCI list. All names and numbers are invented.
Each file has a few deliberate mistakes so the checks have something to find. Run:  python demo/make_samples.py"""
import datetime as dt
import os
import random

from openpyxl import Workbook

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'samples')
os.makedirs(OUT, exist_ok=True)
rnd = random.Random(7)
D = lambda m, d: dt.datetime(2027, m, d)

STYLES = {  # style -> (description, price, msrp, colors)
    'SPM273HD01': ('POWERBLEND PULLOVER HOODIE "S" LOGO', 16.75, 50, {'S100': 'BLACK', 'S110': 'NAVY', 'S120': 'OXFORD GREY HEATHER', 'S140': 'CRIMSON'}),
    'SPM273FZ02': ('POWERBLEND FULL ZIP HOODIE', 18.50, 55, {'S100': 'BLACK', 'S110': 'NAVY', 'S150': 'FOREST GREEN', 'S160': 'GRANITE HEATHER'}),
    'SPM273PA03': ('POWERBLEND OPEN HEM SWEATPANTS', 14.25, 45, {'S100': 'BLACK', 'S110': 'NAVY', 'S120': 'OXFORD GREY HEATHER'}),
    'SPM273TS04': ('JERSEY CREWNECK TEE', 7.00, 20, {'S100': 'BLACK', 'S130': 'WHITE', 'S110': 'NAVY', 'S140': 'CRIMSON'}),
    'SPW273HD11': ('WOMENS POWERBLEND HOODIE', 17.25, 55, {'S100': 'BLACK', 'S120': 'OXFORD GREY HEATHER', 'S170': 'DEEP PLUM', 'S165': 'GRANITE HEATHER'}),
    'SPW273LG12': ('WOMENS HIGH WAIST LEGGING', 12.50, 40, {'S100': 'BLACK', 'S110': 'NAVY', 'S170': 'DEEP PLUM'}),
}
SIZES = ['S', 'M', 'L', 'XL', 'XXL']


def buysheet():
    wb = Workbook(); ws = wb.active; ws.title = 'DATA'
    ws.append(['CUSTOMER', 'STYLE', 'SEGMENT', 'COLOR CODE', 'COLOR', 'DESCRIPTION', 'DELIVERY PROD', 'PRICE', 'MSRP', 'PC PER CTN OS'] + SIZES)
    for cust, seg, when in (('BIGBOX', 'OTH_H_RFID', D(7, 1)), ('CLUBMART ECOM', 'ONL', D(7, 15)), ('MAPLE\'S EAST', 'OTH', D(8, 1))):
        for style, (desc, price, msrp, colors) in STYLES.items():
            for code, name in colors.items():
                if rnd.random() < 0.35:
                    continue
                units = [rnd.choice([24, 48, 72, 96, 144, 192]) for _ in SIZES]
                ws.append([cust, style, seg, code, name, desc, when, price, msrp, 24] + units)
    # deliberate mistakes, so the checks have something to find
    ws.append(['BIGBOX', 'SPM273HD01', 'OTH_H_RFID', 'S110', 'NAVY', 'POWERBLEND PULLOVER HOODIE "S" LOGO', D(7, 1), 14.00, 50, 24, 24, 48, 48, 24, 24])         # price is below SAP's
    ws.append(['BIGBOX', 'SPM273PA03', 'OTH_H_RFID', 'S160', 'GRANITE HEATHER', 'POWERBLEND OPEN HEM SWEATPANTS', D(7, 1), 14.25, 45, 24, 48, 48, 48, 48, 48])         # not on this style yet
    ws.append(['CLUBMART ECOM', 'SPW273HD11', 'ONL', 'S165', 'GRANITE HEATHER', 'WOMENS POWERBLEND HOODIE', D(7, 15), 17.25, 55, 24, 24, 48, 48, 24, 24])            # SAP calls S165 something else
    ws.append(['MAPLE\'S EAST', 'SPM273TS04', 'OTH', 'S190', 'SLATE', 'JERSEY CREWNECK TEE', D(8, 1), 7.00, 20, 48, 48, 48, 48, 48, 48])                           # a color that does not exist yet
    wb.save(os.path.join(OUT, 'Demo Open Stock Buysheet.xlsx'))


def hold_master():
    wb = Workbook(); ws = wb.active; ws.title = 'MASTER'
    ws.append(['Division', 'Customer', 'Hold Name', 'Generic', 'Description', 'Color Code', 'Color Desc', 'Size', 'Qty', 'Article', 'Segment', 'Delivery',
               'Cancel', 'Price', 'MSRP', 'Source file'])
    for division, cust, seg, hold in (('Mens', 'Bigbox', 'OTH_H_RFID', 'HOLD_Q327M_BIGBOX'), ('Mens', 'Birch', 'OTH', 'HOLD_Q327M_BIRCH'), ('Womens', 'Bigbox', 'ONL', 'HOLD_Q327W_BIGBOX')):
        for style, (desc, price, msrp, colors) in STYLES.items():
            if (division == 'Mens') != style.startswith('SPM'):
                continue
            for code, name in colors.items():
                for size in SIZES:
                    ws.append([division, cust, hold, style, desc, code, name, size, rnd.choice([12, 24, 48, 96, 144]), None, seg, D(7, 1), D(7, 1), price, msrp,
                               f'{cust} hold order entry.xlsx'])
    ws.append(['Mens', 'Birch', 'HOLD_Q327M_BIRCH', 'SPM273HD01', 'POWERBLEND PULLOVER HOODIE', 'S120', 'OXFORD HEATHER', 'M', 48, None, 'OTH', D(7, 1), D(7, 1), 16.75, 50, 'Birch hold.xlsx'])   # name differs
    ws.append(['Womens', 'Bigbox', 'HOLD_Q327W_BIGBOX', 'SPW273HD11', 'WOMENS POWERBLEND HOODIE', 'S165', 'GRANITE HEATHER', 'M', 24, None, 'ONL', D(7, 1), D(7, 1), 17.25, 55, 'Bigbox hold.xlsx'])  # code S165 is another color
    wb.save(os.path.join(OUT, 'Demo Hold Master.xlsx'))


def dpci_list():
    wb = Workbook(); ws = wb.active; ws.title = 'DPCI LIST'
    ws.append(['Generic Manufacturer Style #', 'Component DPCI', 'Description', 'COLOR', 'SIZE'])
    for style in ('SPM273HD01', 'SPM273TS04'):
        desc, _, _, colors = STYLES[style]
        for code, name in list(colors.items())[:3]:
            for size in SIZES:
                ws.append([style, f'{rnd.randint(1, 99):03d}-{rnd.randint(1, 99):02d}-{rnd.randint(1, 9999):04d}', desc, name, size])
    wb.save(os.path.join(OUT, 'Demo DPCI List.xlsx'))


buysheet(); hold_master(); dpci_list()
print('wrote', sorted(os.listdir(OUT)))
