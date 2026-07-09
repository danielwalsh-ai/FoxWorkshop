"""
Classification brain — ported from fox_daily_report_builder.py.

Decides, for each purchase-order line:
  Sheet  -> which division tab / cover row it belongs to
  Area   -> the vehicle-area breakdown row (from the merged vehicle lookup)
  Plate  -> plate-year bucket (pre24 / 24plate / 25plate)

Nothing here is Autovolt- or Windows-specific; it just needs a CSV and the
vehicle lookup dicts.
"""
import re
import pandas as pd
from vrm_lookup import load_lookup

# ── COVER STRUCTURE (unchanged from original) ───────────────────────
COVER_TOP_ROW = {
    'Fox Wagons': 3, 'Leyland Wagons': 5, 'JA Jackson': 7, "JJ O'Grady Civils": 9,
    'J FISHER': 11, 'NMS CIVIL': 13, 'Plant': 15, 'Graphics': 17, 'PPE- Workwear': 19,
    'Tyres': 21, 'Misc': 23, 'Assets For sale': 25, 'Asphalt Plant ': 27,
    'J Fisher Plant': 29, 'NMS Plant': 31, 'Damage': 33, 'Windscreen & Glass': 35,
}
SCANIA_ROW = 39; VOLVO_ROW = 41; CAPITAL_ROW = 43
PRE24_ROW = 44; R24_ROW = 45; R25_ROW = 46
INFO_ROWS = [SCANIA_ROW, VOLVO_ROW, CAPITAL_ROW, PRE24_ROW, R24_ROW, R25_ROW]
COVER_BOTTOM_ROW = {
    '8 ALI BODY': 47, '8 EV': 48, '8 WHEELERS': 49, 'ARTICS': 50, 'CONCRETE MIXER': 51,
    'WORKSHOP': 52, 'GRABS': 53, 'HOOKS': 54, 'SWEEPER': 55, 'TRAILERS': 56, 'UNIDENTIFIED': 57,
    'BEAVER TAIL': 58, 'CAR': 59, 'FUEL TANKER': 60, 'PICK UP': 61, 'SHUNTER': 62,
    'TIPPER': 63, 'VAN': 64, 'PLANT': 65, 'JAY/FABSHOP': 66, 'TYRES': 67, 'TARMAC/ASPHALT': 68,
}
BOT_TOTAL_ROW = 69
TOP_SHEETS = set(COVER_TOP_ROW.keys())
sheet_to_row = {**{s: r for s, r in COVER_TOP_ROW.items()}, 'Capital': 43}
PLATE_ROWS = {'pre24': PRE24_ROW, '24plate': R24_ROW, '25plate': R25_ROW}


# ── HELPERS ─────────────────────────────────────────────────────────
def extract_reg(ref):
    if not ref or str(ref) in ['nan', 'None', '']:
        return None
    m = re.search(r'[A-Z]{2}\d{2}[A-Z]{3}|[A-Z]\d{3}[A-Z]{3}', str(ref).upper().replace(' ', ''))
    return m.group(0) if m else None


def reg_year(ref):
    """UK registration YEAR from a standard age-identifier plate — but only
    2021-2026 (fleet trucks). Everything else (older plates, private/cherished
    plates, future years) returns None and is ignored. Add 2027 here in time.
    e.g. PO26xxx / PO76xxx -> 2026, PN21xxx / PN71xxx -> 2021."""
    reg = extract_reg(ref)
    if not reg:
        return None
    m = re.match(r'[A-Z]{2}(\d{2})[A-Z]{3}', reg)
    if not m:
        return None
    num = int(m.group(1))
    year = 2000 + (num - 50) if num >= 51 else 2000 + num
    return year if 2021 <= year <= 2026 else None


def get_area(row, reg_to_area):
    cr = str(row.get('Custom Ref', '')); sr = str(row.get('Supplier Ref', ''))
    for ref in [cr, sr]:
        if any(ref.upper().replace(' ', '').startswith(p) for p in ['FBT', 'HTR', 'MSS']):
            return 'TRAILERS'
    for ref in [cr, sr]:
        reg = extract_reg(ref)
        if reg and reg in reg_to_area:
            return reg_to_area[reg]
    for ref in [cr, sr]:
        ru = ref.upper()
        if 'ASPHALT PLANT' in ru:
            return 'PLANT'
        if ru.strip() == 'PLANT':
            return 'PLANT'
        if 'WASTE TYRE' in ru:
            return 'TYRES'
        if any(k in ru for k in ['STOCK', 'WORKSHOP', 'MISSING', 'CONSUMABLE',
                                 'WATERING', 'WHEEL MARKER', 'LOLER', 'SEAL KIT']):
            return 'WORKSHOP'
    return 'UNIDENTIFIED'


def get_plate(row, reg_plate):
    cr = str(row.get('Custom Ref', '')); sr = str(row.get('Supplier Ref', ''))
    for ref in [cr, sr]:
        reg = extract_reg(ref)
        if reg and reg in reg_plate:
            return reg_plate[reg]
    return None


def classify(row):
    sup_u = str(row.get('Supplier', '')).upper()
    ref_u = str(row.get('Supplier Ref', '')).upper()
    part_u = str(row.get('Part Name', '')).upper()
    asgn = str(row.get('Assigned Depot', '')).strip()
    tgt = str(row.get('Target Depot', '')).strip()
    PLANT_F = ['WIRTGEN', 'FLUID POWER', 'REDGOLD', 'LVS IMAS', 'COUNTY CONVEYORS', 'HARPSCREEN',
               'EXCAVATOR SPARES', 'BANNER EQUIPMENT', 'DANLINE', 'ASTRAK', 'SWAN COMMERCIALS']
    PLANT_R = ['PLANER', 'PLANERS', 'PLANT', 'HYDRAULIC OIL', 'WORKSHOP - HYD OIL', 'RAM RESEAL',
               'KING 22', 'KIN 10', 'CARRINGTON', 'STOCK - HYDRAULIC', 'TEREX', 'RETRO KIT FOR PLANER']
    TYRE_S = {'JB 4 TYRES', 'DMH TYRES', 'GT TYRES', 'TOMLINSON ROAD TYRES', 'PETERBOROUGH TYRE',
              'SOLIDEAL', 'A L TYRES', 'AJ TYRES', 'MICK BUTLER', 'ABCO TYRES', 'CHANTERS', 'ASHTON TYRE'}
    GLASS = ['GLASS REPAIR', 'NEW SCREEN', 'WINDSCREEN', 'REPAIR TO GLASS',
             'FIRST REPAIR TO GLASS', 'REPLACEMENT SCREEN']
    DM = {'Fox Brothers': 'Fox Wagons', 'Hurt Plant': 'Leyland Wagons',
          'J.A.Jackson': 'JA Jackson', "JJ O'Grady Civils": "JJ O'Grady Civils",
          'Plant yard (swarf house)': 'Plant', 'Asphalt plant': 'Asphalt Plant ',
          'PPE': 'PPE- Workwear', 'Tyres': 'Tyres', 'Workshop (misc)': 'Misc',
          'Assets for sales (Sold)': 'Assets For sale'}
    if sup_u.startswith('EWT') or asgn.upper().startswith('CAPITAL'):
        return 'Capital'
    if sup_u == 'WHITTLE GRAPHICS':
        return 'Graphics'
    if sup_u == 'AC PLANT GLAZING' or any(k in ref_u for k in GLASS):
        return 'Windscreen & Glass'
    if 'DAMAGE' in asgn.upper() or any(k in ref_u for k in ['DAMAGE', 'ACCIDENT', 'DAMAGED']):
        return 'Damage'
    if 'REPAIR & PAINT' in ref_u or 'REPAIR AND PAINT' in part_u:
        return 'Damage'
    if 'DSD' in asgn.upper():
        return 'Misc'
    if asgn == 'J FISHER':
        if any(k in sup_u for k in PLANT_F):
            return 'J Fisher Plant'
        if ref_u.startswith('DT') or (ref_u.startswith('CP') and not ref_u.startswith('CTR')):
            return 'J Fisher Plant'
        if any(k in ref_u for k in PLANT_R):
            return 'J Fisher Plant'
        if any(t in sup_u.replace(' ', '') for t in ['JB4TYRES', 'DMHTYRES', 'GTTYRES', 'ALTYRES', 'CHANTERS', 'ASHTONTYRE']):
            return 'Tyres'
        if 'TYRE' in ref_u or 'TYRE' in part_u:
            return 'Tyres'
        return 'J FISHER'
    if asgn == 'NMS CIVIL ENGINEERING LTD':
        if 'WIRTGEN' in sup_u:
            return 'NMS Plant'
        if any(k in ref_u for k in ['PLANT', 'TELEHANDLER', 'PLANER']):
            return 'NMS Plant'
        if sup_u in TYRE_S or 'TYRE' in part_u:
            return 'Tyres'
        return 'NMS CIVIL'
    if sup_u in TYRE_S or 'TYRE' in part_u:
        return 'Tyres'
    if asgn in DM:
        return DM[asgn]
    if not asgn or asgn in ['nan', 'None', '']:
        if tgt in DM:
            return DM[tgt]
    return 'Misc'


def process_csv(csv_path, report_date, reg_to_area=None, reg_plate=None):
    """Load CSV, clean, keep only report_date rows, and classify each line."""
    if reg_to_area is None or reg_plate is None:
        reg_to_area, _, reg_plate = load_lookup()
    df = pd.read_csv(csv_path)
    df['Date'] = pd.to_datetime(df['PO Created Date'])
    df = df[~df['Supplier Ref'].astype(str).str.upper().str.contains('CANCEL', na=False)]
    df = df.drop_duplicates(subset=['PO No', 'Supplier', 'Part Name', 'Cost', 'PO Created Date'])
    df = df[df['Date'].dt.date == report_date]
    df = df.copy()
    df['Sheet'] = df.apply(classify, axis=1)
    df['Area'] = df.apply(lambda r: get_area(r, reg_to_area), axis=1)
    df['Plate'] = df.apply(lambda r: get_plate(r, reg_plate), axis=1)
    return df


def process_all(csv_path, reg_to_area=None, reg_plate=None):
    """Like process_csv but keeps ALL dates (for history backfill); adds a
    per-row ReportDate (date of the PO)."""
    if reg_to_area is None or reg_plate is None:
        reg_to_area, _, reg_plate = load_lookup()
    df = pd.read_csv(csv_path)
    df['Date'] = pd.to_datetime(df['PO Created Date'])
    df = df[~df['Supplier Ref'].astype(str).str.upper().str.contains('CANCEL', na=False)]
    df = df.drop_duplicates(subset=['PO No', 'Supplier', 'Part Name', 'Cost', 'PO Created Date'])
    df = df.copy()
    df['Sheet'] = df.apply(classify, axis=1)
    df['Area'] = df.apply(lambda r: get_area(r, reg_to_area), axis=1)
    df['Plate'] = df.apply(lambda r: get_plate(r, reg_plate), axis=1)
    df['ReportDate'] = df['Date'].dt.date
    return df
