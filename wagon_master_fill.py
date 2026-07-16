"""
Fox Group / Clive Hurt Plant Hire - Daily Wagon Earnings master sheet filler.

Fills one day column in the Daily Wagon Earnings master workbook from:
  1. The daily wagon earnings file (e.g. 02.07.2026.xlsx) - Wagons tab
  2. The fox transaction report (parts/costs)          - Cover tab

Confirmed mapping rules (02/07/2026 trial, signed off by DW):
  - Per-wagon earnings: Wagons tab, reg in col A, value in col C.
    Text statuses (VOR / MN / BD / ABSENCE) written verbatim so master COUNTIFs work.
  - Regs in master with no row in the daily file are left blank.
  - NO WAGONS: Wagons tab count cell (A121 in current template).
  - PARTS    = Cover 'Leyland Wagons' row, day column. Leyland ONLY, never Fox Wagons.
  - WORKSHOP = 0 unless explicitly provided.
  - TYRES    = Cover 'Tyres' row (row 21 area section), day column.
  - Overheads/Fuel/Wages/Plant hire/Tax/Other/EBITDA-costs: standing absolute formulas
    copied from the last populated column (point at the hidden monthly block IX:JC).
  - Missing days between last populated column and target date are left empty.

CRITICAL: the master has 200+ SharePoint external links and chart sheets.
A plain openpyxl load/save DESTROYS the cached link values and the charts.
All writes therefore happen at raw sheet-XML level inside the xlsx zip.
workbook.xml gets fullCalcOnLoad=1 and calcChain is dropped so Excel
recalculates everything cleanly on first open.

Usage:
    python wagon_master_fill.py MASTER.xlsx DAILY.xlsx [TRANSACTIONS.xlsx] -o OUT.xlsx
    (date is read from the daily file's Wagons!P2; override with --date DD/MM/YYYY)

Exit code non-zero and no output file if any verification check fails.
"""

import argparse
import json
import re
import shutil
import sys
import zipfile
from datetime import date, datetime, timedelta

from lxml import etree
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.formula.translate import Translator

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
def q(t): return f'{{{NS}}}{t}'

EPOCH = date(1899, 12, 30)

# ---- master layout (current template) -------------------------------------
VEHICLE_ROWS = (list(range(4, 20)) + [23] + list(range(27, 73)) + list(range(76, 100))
                + list(range(103, 112)) + list(range(115, 121)) + list(range(124, 138))
                + list(range(141, 150)) + list(range(161, 167)))
FORMULA_ROWS = [20, 21, 24, 25, 73, 74, 100, 101, 112, 113, 121, 122, 138, 139, 150, 151,
                155, 156, 157, 158, 167, 168, 170, 171, 172, 173, 174, 175, 179, 180, 181,
                182, 183, 186, 189, 192, 193, 194, 195, 196, 197, 198, 200]
ROW_TOTAL_EARNINGS, ROW_NO_WAGONS = 168, 169
ROW_PARTS, ROW_WORKSHOP, ROW_TYRES = 176, 177, 178
DATE_ROW, DAYNAME_ROW = 2, 1
EARNINGS_SUM_GROUPS = [range(4, 20), [23], range(27, 73), range(76, 100), range(103, 112),
                       range(115, 121), range(124, 138), range(141, 150), range(161, 167)]

# ---- source extraction -----------------------------------------------------

def read_daily_file(path, date_override=None, value_col=3):
    """Return (report_date, {REG: value}, wagon_count) from the daily wagon file.

    value_col/date_override let a weekend sheet (Sat in col C, Sun in col D,
    date in Q2 not P2) be filled one day at a time."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Wagons']
    report_date = date_override or ws['P2'].value
    if isinstance(report_date, datetime):
        report_date = report_date.date()
    mapping, count = {}, None
    for r in range(3, ws.max_row + 1):
        a, c = ws.cell(row=r, column=1).value, ws.cell(row=r, column=value_col).value
        if isinstance(a, str) and a.strip() == 'VEHICLE':
            break  # reached the TARGET table at the bottom - stop
        if isinstance(a, str) and a.strip():
            mapping[a.strip().upper()] = c
        elif isinstance(a, (int, float)) and ws.cell(row=r + 2, column=1).value is None \
                and ws.cell(row=r, column=value_col).value is None:
            count = int(a)  # standalone grand-count row (A121 style)
    if count is None:  # fallback: cell two below last category subtotal
        for r in range(ws.max_row, 3, -1):
            v = ws.cell(row=r, column=1).value
            if isinstance(v, (int, float)) and v > 50:
                count = int(v); break
    return report_date, mapping, count


def read_transaction_report(path, report_date):
    """Return (parts, workshop, tyres) for the given date from the Cover tab.
    PARTS = Leyland Wagons row ONLY. TYRES = 'Tyres' area row. WORKSHOP = 0."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['Cover']
    day_col = 1 + report_date.day  # col B = 1st of the month
    hdr = str(ws.cell(row=2, column=day_col).value or '')
    if not hdr.startswith(str(report_date.day)):
        raise ValueError(f"Cover day header mismatch: expected day {report_date.day}, "
                         f"found {hdr!r} in col {get_column_letter(day_col)}")
    leyland = tyres = None
    for r in range(3, 40):
        label = str(ws.cell(row=r, column=1).value or '').strip().lower()
        if label == 'leyland wagons':
            leyland = ws.cell(row=r, column=day_col).value
        elif label == 'tyres':
            tyres = ws.cell(row=r, column=day_col).value
    if leyland is None or tyres is None:
        raise ValueError("Could not locate 'Leyland Wagons' and/or 'Tyres' rows on Cover tab")
    return float(leyland), 0.0, float(tyres)

# ---- master helpers --------------------------------------------------------

def master_state(path):
    """Return (sheet_xml_path, date_cols, last_date_col, last_populated_col, row->reg map)."""
    wb = openpyxl.load_workbook(path, data_only=False)
    wbv = openpyxl.load_workbook(path, data_only=True)
    ws, wsv = wb['DAILY'], wbv['DAILY']
    date_cols, last_date_col, last_pop_col = {}, None, None
    for c in range(3, ws.max_column + 2):
        d = wsv.cell(row=DATE_ROW, column=c).value
        if isinstance(d, datetime):
            date_cols[d.date()] = c
            last_date_col = c
        f = ws.cell(row=ROW_TOTAL_EARNINGS, column=c).value
        if f not in (None, ''):
            last_pop_col = c
    regs = {}
    for r in VEHICLE_ROWS:
        a = ws.cell(row=r, column=1).value
        if isinstance(a, str) and a.strip():
            regs[r] = a.strip().upper()
    z = zipfile.ZipFile(path)
    wbxml = z.read('xl/workbook.xml').decode()
    rels = z.read('xl/_rels/workbook.xml.rels').decode()
    rid = re.search(r'<sheet[^>]*name="DAILY"[^>]*r:id="(rId\d+)"', wbxml).group(1)
    target = re.search(rf'<Relationship[^>]*Id="{rid}"[^>]*Target="([^"]+)"', rels).group(1)
    z.close()
    return 'xl/' + target, date_cols, last_date_col, last_pop_col, regs, ws, wsv


def col_for_date(date_cols, last_date_col, target):
    """Exact date match in row 2 wins; otherwise extend past the last date column.
    The master's date row has occasional skipped days, so never assume consecutiveness."""
    if target in date_cols:
        return date_cols[target], False
    last_date = max(date_cols)
    if target <= last_date:
        raise ValueError(f"{target} falls inside the existing date range but has no "
                         f"column in row 2 - master dates skip this day; check manually")
    return last_date_col + (target - last_date).days, True

# ---- XML writer ------------------------------------------------------------

class SheetXmlEditor:
    def __init__(self, xml_bytes):
        self.tree = etree.fromstring(xml_bytes)
        self.sheetData = self.tree.find(q('sheetData'))
        self.rowmap = {int(r.get('r')): r for r in self.sheetData.findall(q('row'))}

    def _cell(self, rowel, ref):
        for c in rowel.findall(q('c')):
            if c.get('r') == ref:
                return c
        return None

    def style_of(self, rownum, colletter):
        rel = self.rowmap.get(rownum)
        if rel is None: return None
        c = self._cell(rel, f'{colletter}{rownum}')
        return c.get('s') if c is not None else None

    def write(self, rownum, colletter, *, value=None, formula=None, text=None, style=None):
        rel = self.rowmap.get(rownum)
        if rel is None:
            rel = etree.SubElement(self.sheetData, q('row')); rel.set('r', str(rownum))
            self.rowmap[rownum] = rel
        ref = f'{colletter}{rownum}'
        old = self._cell(rel, ref)
        if old is not None: rel.remove(old)
        c = etree.Element(q('c')); c.set('r', ref)
        if style: c.set('s', style)
        if formula is not None:
            etree.SubElement(c, q('f')).text = formula.lstrip('=')
        elif text is not None:
            c.set('t', 'inlineStr')
            etree.SubElement(etree.SubElement(c, q('is')), q('t')).text = text
        elif value is not None:
            etree.SubElement(c, q('v')).text = repr(value) if isinstance(value, float) else str(value)
        col_idx = column_index_from_string(colletter)
        for existing in rel.findall(q('c')):
            ecol = column_index_from_string(re.match(r'([A-Z]+)', existing.get('r')).group(1))
            if ecol > col_idx:
                existing.addprevious(c); return
        rel.append(c)

    def fix_dimension(self, colletter):
        dim = self.tree.find(q('dimension'))
        if dim is not None:
            dim.set('ref', re.sub(r':[A-Z]+(\d+)$', rf':{colletter}\1', dim.get('ref')))

    def tobytes(self):
        return etree.tostring(self.tree, xml_declaration=True, encoding='UTF-8', standalone=True)

# ---- main fill -------------------------------------------------------------

def fill_master(master, daily, transactions, out, date_override=None, value_col=3):
    report_date, earnings, wagon_count = read_daily_file(daily, date_override, value_col)
    parts = workshop = tyres = None
    if transactions:
        parts, workshop, tyres = read_transaction_report(transactions, report_date)

    (sheet_path, date_cols, last_date_col, last_pop_col, regs, ws, wsv) = master_state(master)
    target_col, needs_extension = col_for_date(date_cols, last_date_col, report_date)
    # Per-column guard: refuse only if the TARGET column itself already holds data.
    # (The old `target_col <= last_pop_col` test blocked backfilling an empty gap that
    #  sits before a later populated column — e.g. filling 23-30 Jun when 2 Jul is done.)
    if ws.cell(row=ROW_TOTAL_EARNINGS, column=target_col).value not in (None, ''):
        raise ValueError(f"{report_date} maps to column {get_column_letter(target_col)} "
                         f"which is already populated - refusing to overwrite")
    # Source column for styles + standing formulas = nearest populated column strictly
    # BEFORE the target (the adjacent real day), not the global last populated column.
    source_col = next((c for c in range(target_col - 1, 2, -1)
                       if ws.cell(row=ROW_TOTAL_EARNINGS, column=c).value not in (None, '')),
                      last_pop_col)
    TL = get_column_letter(target_col)
    src_col_letter = get_column_letter(source_col)

    fills, unmatched_master, matched = {}, [], set()
    for r, reg in regs.items():
        if reg in earnings:
            fills[r] = earnings[reg]; matched.add(reg)
        else:
            unmatched_master.append(reg)
    lost = {k: v for k, v in earnings.items() if k not in matched
            and isinstance(v, (int, float)) and v != 0}
    if lost:
        raise ValueError(f"Daily file regs with earnings not present in master (would be lost): {lost}")

    expected_total = round(sum(v for v in fills.values() if isinstance(v, (int, float))), 2)

    zin = zipfile.ZipFile(master)
    ed = SheetXmlEditor(zin.read(sheet_path))

    # extend dates + day names past the last existing date column if needed
    if needs_extension:
        s2 = ed.style_of(DATE_ROW, get_column_letter(last_date_col))
        s1 = ed.style_of(DAYNAME_ROW, get_column_letter(last_date_col))
        d, c = max(date_cols), last_date_col
        while d < report_date:
            d += timedelta(days=1); c += 1
            cl = get_column_letter(c)
            ed.write(DATE_ROW, cl, value=(d - EPOCH).days, style=s2)
            ed.write(DAYNAME_ROW, cl, formula=f'TEXT({cl}2, "dddd")', style=s1)

    for r, v in fills.items():
        st = ed.style_of(r, src_col_letter)
        if isinstance(v, str):
            ed.write(r, TL, text=v, style=st)
        else:
            ed.write(r, TL, value=v, style=st)

    for r in FORMULA_ROWS:
        f = ws.cell(row=r, column=source_col).value
        if not (isinstance(f, str) and f.startswith('=')):
            raise ValueError(f"Expected formula at {src_col_letter}{r}, found {f!r}")
        ed.write(r, TL, formula=Translator(f, origin=f'{src_col_letter}{r}').translate_formula(f'{TL}{r}'),
                 style=ed.style_of(r, src_col_letter))

    ed.write(ROW_NO_WAGONS, TL, value=wagon_count, style=ed.style_of(ROW_NO_WAGONS, src_col_letter))
    if parts is not None:
        ed.write(ROW_PARTS, TL, value=round(parts, 2), style=ed.style_of(ROW_PARTS, src_col_letter))
        ed.write(ROW_WORKSHOP, TL, value=round(workshop, 2), style=ed.style_of(ROW_WORKSHOP, src_col_letter))
        ed.write(ROW_TYRES, TL, value=round(tyres, 2), style=ed.style_of(ROW_TYRES, src_col_letter))
    ed.fix_dimension(TL)

    # repackage preserving everything else; force recalc on open, drop calcChain
    wb_tree = etree.fromstring(zin.read('xl/workbook.xml'))
    calcPr = wb_tree.find(q('calcPr'))
    if calcPr is None:
        calcPr = etree.SubElement(wb_tree, q('calcPr'))
    calcPr.set('fullCalcOnLoad', '1')
    ct = zin.read('[Content_Types].xml').decode().replace(
        '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.spreadsheetml.calcChain+xml"/>', '')
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == 'xl/calcChain.xml':
                continue
            if item.filename == sheet_path:
                data = ed.tobytes()
            elif item.filename == 'xl/workbook.xml':
                data = etree.tostring(wb_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
            elif item.filename == '[Content_Types].xml':
                data = ct.encode()
            else:
                data = zin.read(item.filename)
            zout.writestr(item, data)
    zin.close()

    # verify what we wrote reads back correctly
    wchk = openpyxl.load_workbook(out, data_only=False)['DAILY']
    tcol = target_col
    back = round(sum(wchk.cell(row=r, column=tcol).value for r, v in fills.items()
                     if isinstance(v, (int, float))), 2)
    assert back == expected_total, f"Read-back total {back} != expected {expected_total}"
    vors = sum(1 for r, v in fills.items() if v == 'VOR')

    return {
        'date': str(report_date), 'column': TL,
        'wagons_filled': len(fills), 'expected_total_earnings': expected_total,
        'vor_count': vors, 'no_wagons': wagon_count,
        'parts': parts, 'workshop': workshop, 'tyres': tyres,
        'master_regs_not_in_daily_file': sorted(unmatched_master),
    }


if __name__ == '__main__':
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    ap.add_argument('master'); ap.add_argument('daily')
    ap.add_argument('transactions', nargs='?', default=None)
    ap.add_argument('-o', '--out', required=True)
    a = ap.parse_args()
    try:
        result = fill_master(a.master, a.daily, a.transactions, a.out)
    except Exception as e:
        print(json.dumps({'status': 'failed', 'error': str(e)}), file=sys.stderr)
        sys.exit(1)
    print(json.dumps({'status': 'ok', **result}, indent=2))
