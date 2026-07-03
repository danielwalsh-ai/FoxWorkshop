"""
Fox Group Daily Report — builder (adapted from fox_daily_report_builder.py).

Differences from the original:
  * date args computed automatically (date_args.py)
  * vehicle areas from the merged master list (vrm_lookup.py via classify.py)
  * per-row budgets READ from the workbook's hidden column 35 (not hardcoded)
  * Windows-safe date formatting
  * blank template auto-generated from a base workbook when needed

Usage:
    python build_report.py <csv> <report_date YYYY-MM-DD> [base_xlsx]
"""
import sys
import shutil
from pathlib import Path
import datetime as dt

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle

import classify as C
from classify import (
    COVER_TOP_ROW, COVER_BOTTOM_ROW, INFO_ROWS, SCANIA_ROW, VOLVO_ROW, CAPITAL_ROW,
    PRE24_ROW, R24_ROW, R25_ROW, BOT_TOTAL_ROW, TOP_SHEETS, sheet_to_row, PLATE_ROWS,
    process_csv,
)
from date_args import compute as compute_dates
from vrm_lookup import load_lookup

HERE = Path(__file__).parent
FIXED_WD = 22  # working days per month (Daniel's convention)
BUDGET_COL = 35
MTD_COL = 33
REMAIN_COL = 37


def fmt_date_long(d: dt.date) -> str:
    return f"{d.day} {d:%B %Y}"          # e.g. "2 July 2026" (no %-d, Windows-safe)


def make_blank_template(base_path: Path, out_path: Path):
    """Create a reusable blank template from a real report: keep structure,
    labels, formulas and budgets; wipe all daily data + tab rows."""
    shutil.copy(base_path, out_path)
    wb = load_workbook(out_path)
    cover = wb['Cover']
    data_rows = list(range(3, 36, 2)) + INFO_ROWS + list(COVER_BOTTOM_ROW.values())
    for r in data_rows:
        for c in range(2, 33):                     # daily columns
            cover.cell(r, c, 0)
        v33 = cover.cell(r, MTD_COL).value
        if not (isinstance(v33, str) and v33.startswith('=')):
            cover.cell(r, MTD_COL, None)           # clear static MTD (keep formulas)
        cover.cell(r, REMAIN_COL, None)            # clear remaining (recomputed)
    for tab in wb.sheetnames:
        if tab == 'Cover':
            continue
        ws = wb[tab]
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row)
    wb.save(out_path)
    return out_path


def build(csv_path, report_date: dt.date, base_path: Path, fixed_wd=FIXED_WD):
    info = compute_dates(report_date, fixed_wd=fixed_wd)
    today_col = info['today_col']
    days_elapsed = info['days_elapsed']
    days_remaining = info['days_remaining']
    wd = info['working_days_in_month']
    is_new_month = info['is_new_month']

    out_xlsx = HERE / f"fox_transaction_report_{report_date:%d-%m-%Y}.xlsx"
    out_pdf = HERE / f"daily_kpi_report_{report_date:%d-%m-%Y}.pdf"
    report_date_long = fmt_date_long(report_date)

    print(f"Building {report_date_long}  col={today_col} elapsed={days_elapsed} "
          f"remaining={days_remaining} wd={wd} new_month={is_new_month}")

    # ── load data ──
    reg_to_area, reg_make, reg_plate = load_lookup()
    df = process_csv(csv_path, report_date, reg_to_area, reg_plate)
    print(f"CSV rows for {report_date}: {len(df)}  £{df['Cost'].sum():,.2f}")

    # ── choose starting workbook ──
    if is_new_month:
        blank = HERE / "blank_template.xlsx"
        if not blank.exists():
            print("Generating blank template from base...")
            make_blank_template(base_path, blank)
        shutil.copy(blank, out_xlsx)
    else:
        shutil.copy(base_path, out_xlsx)

    wb = load_workbook(out_xlsx)
    cover = wb['Cover']

    # mid-month: clear tabs back to header (today's rows get written fresh)
    if not is_new_month:
        for tab in wb.sheetnames:
            if tab == 'Cover':
                continue
            ws = wb[tab]
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row)

    def gcv(r, c):
        v = cover.cell(r, c).value
        if v is None or (isinstance(v, str) and v.startswith('=')):
            return 0.0
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    # ── write cover (today's column) ──
    # reset today's column first so a re-run can never double-count
    reset_rows = list(range(3, 36, 2)) + INFO_ROWS + list(COVER_BOTTOM_ROW.values())
    for r in reset_rows:
        cover.cell(r, today_col, 0)

    for _, row in df.iterrows():
        cost = float(row['Cost'])
        sheet = row['Sheet']
        crow = sheet_to_row.get(sheet)
        if crow:
            cover.cell(crow, today_col, round(gcv(crow, today_col) + cost, 2))
        if sheet in TOP_SHEETS:
            at = row['Area']
            if at in COVER_BOTTOM_ROW:
                brow = COVER_BOTTOM_ROW[at]
                cover.cell(brow, today_col, round(gcv(brow, today_col) + cost, 2))
        sup = str(row.get('Supplier', '')).upper()
        if 'SCANIA' in sup:
            cover.cell(SCANIA_ROW, today_col, round(gcv(SCANIA_ROW, today_col) + cost, 2))
        elif 'VOLVO' in sup or 'THOMAS HARDIE' in sup:
            cover.cell(VOLVO_ROW, today_col, round(gcv(VOLVO_ROW, today_col) + cost, 2))
        if sheet in TOP_SHEETS:
            pc = row['Plate']
            if pc and pc in PLATE_ROWS:
                pr = PLATE_ROWS[pc]
                cover.cell(pr, today_col, round(gcv(pr, today_col) + cost, 2))

    # zero-fill today's column
    all_data_rows = list(range(3, 36, 2)) + INFO_ROWS + list(COVER_BOTTOM_ROW.values())
    for r in all_data_rows:
        if cover.cell(r, today_col).value is None:
            cover.cell(r, today_col, 0)

    # ── MTD (col 33) + Remaining (col 37), budgets read from col 35 ──
    def get_mtd(rn):
        return round(sum(gcv(rn, c) for c in range(2, today_col + 1)), 2)

    for rn in list(COVER_TOP_ROW.values()) + [CAPITAL_ROW]:
        cover.cell(rn, MTD_COL, get_mtd(rn))
    for rn in list(COVER_TOP_ROW.values()) + [CAPITAL_ROW]:
        budget = gcv(rn, BUDGET_COL)
        if budget:
            cover.cell(rn, REMAIN_COL, round(budget - get_mtd(rn), 2))

    # ── write tabs (today's rows only) ──
    for sheet_name in df['Sheet'].unique():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        nr = 2
        for _, row in df[df['Sheet'] == sheet_name].iterrows():
            def s(v):
                import pandas as pd
                return None if pd.isna(v) else v
            ws.cell(nr, 1, str(row.get('Supplier', '')) or '')
            ws.cell(nr, 2, str(row.get('Supplier Source Depot', '')) or '')
            ws.cell(nr, 3, str(row.get('365 No', '')) or '')
            ws.cell(nr, 4, str(row.get('Supplier PN', '')) or '')
            ws.cell(nr, 5, str(row.get('Part Name', '')) or '')
            ws.cell(nr, 6, float(row.get('Cost', 0)))
            ws.cell(nr, 7, str(row.get('PO No', '')) or '')
            ws.cell(nr, 8, s(row.get('Attached Order No')))
            ws.cell(nr, 9, s(row.get('Attached Customer')))
            ws.cell(nr, 10, str(row.get('PO Created Date', '')) or '')
            ws.cell(nr, 11, str(row.get('Supplier/ Collection?', '')) or '')
            ws.cell(nr, 12, s(row.get('Item Count')))
            ws.cell(nr, 13, s(row.get('Goods Received')))
            ws.cell(nr, 14, str(row.get('Target Depot', '')) or '')
            ws.cell(nr, 15, str(row.get('Assigned Depot', '')) or '')
            ws.cell(nr, 16, str(row.get('Supplier Ref', '')) or '')
            ws.cell(nr, 17, str(row.get('Custom Ref', '')) or '')
            ws.cell(nr, 18, row.get('Area', '') or '')
            nr += 1

    wb.save(out_xlsx)

    # ── verify balance ──
    wb2 = load_workbook(out_xlsx)
    cover2 = wb2['Cover']

    def g2(r, c):
        v = cover2.cell(r, c).value
        if v is None or (isinstance(v, str) and v.startswith('=')):
            return 0.0
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    top_rows = list(range(3, 36, 2))
    bot_rows = list(COVER_BOTTOM_ROW.values())
    top = sum(g2(r, today_col) for r in top_rows)
    bot = sum(g2(r, today_col) for r in bot_rows)
    diff = abs(top - bot)
    print(f"Balance: top £{top:,.2f}  bottom £{bot:,.2f}  "
          f"{'BALANCED' if diff < 0.01 else f'GAP £{diff:,.2f}'}")

    _build_pdf(out_pdf, cover2, g2, report_date, report_date_long,
               today_col, days_elapsed, days_remaining, wd)
    print(f"Saved: {out_xlsx.name}")
    print(f"Saved: {out_pdf.name}")
    daily_total = top
    return out_xlsx, out_pdf, diff, daily_total, report_date_long, df


# ── PDF (ported from original, budgets read from sheet) ─────────────
def _build_pdf(out_pdf, cover2, g2, report_date, REPORT_DATE, TODAY_COL,
               DAYS_ELAPSED, DAYS_REMAINING, WD):
    NAVY = colors.HexColor('#24214a'); ORANGE = colors.HexColor('#eb941f')
    BLUE = colors.HexColor('#00579e'); WHITE = colors.white
    LIGHT = colors.HexColor('#F2F3F7'); GREY = colors.HexColor('#666666')
    MID = colors.HexColor('#CCCCCC'); GREEN = colors.HexColor('#27ae60')
    PURPLE = colors.HexColor('#8e44ad'); AMBER = colors.HexColor('#f39c12')
    RED = colors.HexColor('#c0392b'); AMBG = colors.HexColor('#FFF3CD')
    RDBG = colors.HexColor('#FFCCCC'); TEAL = colors.HexColor('#16a085')

    def fmt(v):  return f"£{v:,.2f}"
    def fmts(v): return f"£{v:,.0f}"

    daily_total = sum(g2(r, TODAY_COL) for r in COVER_TOP_ROW.values())
    daily_damage = g2(33, TODAY_COL); daily_glass = g2(35, TODAY_COL)
    daily_capital = g2(43, TODAY_COL); daily_tyres = g2(21, TODAY_COL)
    mtd_total = sum(g2(r, c) for r in COVER_TOP_ROW.values() for c in range(2, TODAY_COL + 1))
    mtd_damage = sum(g2(33, c) for c in range(2, TODAY_COL + 1))
    mtd_tyres = sum(g2(21, c) for c in range(2, TODAY_COL + 1))
    mtd_capital = sum(g2(43, c) for c in range(2, TODAY_COL + 1))
    daily_avg = mtd_total / DAYS_ELAPSED if DAYS_ELAPSED > 0 else 0

    day_nz = {}
    for ci in range(2, TODAY_COL + 1):
        lbl = cover2.cell(2, ci).value or str(ci)
        tot = sum(g2(r, ci) for r in COVER_TOP_ROW.values())
        if tot > 0:
            day_nz[lbl] = tot
    bd_lbl = max(day_nz, key=day_nz.get) if day_nz else '-'
    bd_val = day_nz.get(bd_lbl, 0)

    # budgets from the workbook (hidden col 35)
    BUDGET_ROWS = [(3, 'Fox Wagons'), (5, 'Leyland Wagons'), (11, 'J FISHER (Trucks)'),
                   (13, 'NMS CIVIL'), (21, 'Tyres'), (29, 'J Fisher Plant'), (31, 'NMS Plant')]
    budget_data = [(name, g2(row, BUDGET_COL), sum(g2(row, c) for c in range(2, TODAY_COL + 1)))
                   for row, name in BUDGET_ROWS]
    div_rows = [(name, g2(row, TODAY_COL)) for name, row in COVER_TOP_ROW.items() if g2(row, TODAY_COL) > 0]
    day_rows = []
    for ci in range(2, TODAY_COL + 1):
        lb = cover2.cell(2, ci).value or ''
        dt2 = sum(g2(r, ci) for r in COVER_TOP_ROW.values())
        mo = REPORT_DATE.split()[-2][:3]
        if dt2 > 0:
            day_rows.append((f"{lb} {mo}", dt2, g2(33, ci), g2(21, ci), g2(43, ci)))

    W, H = A4; M = 12 * mm; HH = 22 * mm; FH = 10 * mm; CW = W - 2 * M
    YT = H - HH - 5 * mm; YB = FH + 4 * mm
    CARD_H = 22 * mm; RH = 6.5 * mm; RH_B = 7.5 * mm; LH = 6 * mm; G = 5 * mm
    n_half = (len(div_rows) + 1) // 2
    DIV_H = (n_half + 1) * RH; BGT_H = (len(budget_data) + 1) * RH_B; NOTE_H = 7 * mm
    P1_FIXED = LH + G + CARD_H + G + LH + G + DIV_H + G + LH + G + NOTE_H + G / 2 + BGT_H
    PAD1 = max((YT - YB - P1_FIXED) / 3, 2 * mm)
    MTD_H = CARD_H + 3 * mm + CARD_H; DAY_H = (len(day_rows) + 2) * RH
    PAD2 = max((YT - YB - LH - G - MTD_H - G - LH - G - DAY_H) / 2, 2 * mm)

    cv = canvas.Canvas(str(out_pdf), pagesize=A4)

    def chrome(pg, total):
        cv.setFillColor(NAVY); cv.rect(0, H - HH, W, HH, fill=1, stroke=0)
        cv.setFillColor(ORANGE); cv.rect(0, H - HH, W, 1.5 * mm, fill=1, stroke=0)
        cv.setFillColor(WHITE); cv.setFont('Helvetica-Bold', 16)
        cv.drawString(M, H - 10 * mm, 'Fox Group Ltd — Daily Workshop Spend Report')
        cv.setFont('Helvetica', 9)
        cv.drawString(M, H - 17.5 * mm, f'{REPORT_DATE}   |   Prepared by Knowles Farm Ltd')
        cv.drawRightString(W - M, H - 17.5 * mm, f'Page {pg} of {total}')
        cv.setFillColor(NAVY); cv.rect(0, 0, W, FH, fill=1, stroke=0)
        cv.setFillColor(WHITE); cv.setFont('Helvetica', 7.5)
        cv.drawString(M, 4 * mm, 'Confidential — Fox Group Ltd   |   Prepared by Knowles Farm Ltd')
        cv.drawRightString(W - M, 4 * mm, f'Generated {REPORT_DATE}')

    def sec_lbl(y, text):
        cv.setFillColor(NAVY); cv.setFont('Helvetica-Bold', 11); cv.drawString(M, y, text)
        cv.setFillColor(ORANGE); cv.rect(M, y - 2 * mm, CW, 0.7 * mm, fill=1, stroke=0)

    def draw_cards(y, ch, items, n):
        gap = 3 * mm; kw = (CW - (n - 1) * gap) / n
        for i, (title, val, col, sub) in enumerate(items):
            kx = M + i * (kw + gap); ky = y - ch
            cv.setFillColor(col); cv.roundRect(kx, ky, kw, ch, 2 * mm, fill=1, stroke=0)
            cv.setFillColor(WHITE); cv.setFont('Helvetica-Bold', 7.5)
            cv.drawString(kx + 3 * mm, ky + ch - 5.5 * mm, title)
            cv.setFont('Helvetica-Bold', 16); cv.drawCentredString(kx + kw / 2, ky + ch - 13 * mm, val)
            cv.setFillColor(colors.HexColor('#FFFFFF50', hasAlpha=True))
            cv.rect(kx + 3 * mm, ky + 6.5 * mm, kw - 6 * mm, 0.3 * mm, fill=1, stroke=0)
            cv.setFillColor(WHITE); cv.setFont('Helvetica', 6.5)
            cv.drawCentredString(kx + kw / 2, ky + 3.5 * mm, sub)

    def base_tbl():
        return [('BACKGROUND', (0, 0), (-1, 0), NAVY), ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'), ('ROWBACKGROUNDS', (0, 1), (-1, -1), [LIGHT, WHITE]),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'), ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                ('LINEBELOW', (0, 0), (-1, 0), 0.5, ORANGE), ('GRID', (0, 0), (-1, -1), 0.2, MID),
                ('TOPPADDING', (0, 0), (-1, -1), 1.5), ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5),
                ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3)]

    # Page 1
    chrome(1, 2); y = YT
    sec_lbl(y, "TODAY'S SPEND"); y -= LH + G
    draw_cards(y, CARD_H, [("Today's Total", fmt(daily_total), BLUE, ''),
        ('Damage', fmt(daily_damage), RED, ''), ('Tyres', fmt(daily_tyres), NAVY, ''),
        ('Windscreen/Glass', fmt(daily_glass), TEAL, ''), ('Capital', fmt(daily_capital), ORANGE, '')], 5)
    y -= CARD_H + G + PAD1

    sec_lbl(y, 'Spend by Division'); y -= LH + G
    left_div = div_rows[:n_half]; right_div = div_rows[n_half:]
    while len(right_div) < len(left_div):
        right_div.append(('', ''))
    div_data = [['Division', 'Spend', 'Division', 'Spend']]
    for i in range(len(left_div)):
        lr = left_div[i]; rr = right_div[i]
        div_data.append([lr[0], fmt(lr[1]) if lr[1] else '', rr[0], fmt(rr[1]) if rr[1] else ''])
    dt_tbl = Table(div_data, colWidths=[71 * mm, 22 * mm, 71 * mm, 22 * mm], rowHeights=RH)
    dt_tbl.setStyle(TableStyle(base_tbl() + [('ALIGN', (2, 0), (2, -1), 'LEFT'),
                                             ('LINEAFTER', (1, 0), (1, -1), 0.5, MID)]))
    dt_tbl.wrapOn(cv, W, H); dt_tbl.drawOn(cv, M, y - DIV_H)
    y -= DIV_H + G + PAD1

    mo_name = REPORT_DATE.split()[-2]
    sec_lbl(y, f'Budget Tracker — {mo_name} {REPORT_DATE.split()[-1]}'); y -= LH + G
    cv.setFillColor(GREY); cv.setFont('Helvetica', 8.5)
    cv.drawString(M, y, f'{DAYS_ELAPSED} day{"s" if DAYS_ELAPSED > 1 else ""} elapsed — '
                        f'{DAYS_REMAINING} working days remaining of {WD}')
    y -= NOTE_H + G / 2
    bgt_cols = [52 * mm, 26 * mm, 28 * mm, 28 * mm, 52 * mm]
    bgt_data = [['Division', 'Budget', 'MTD Spend', 'Remaining', 'Daily Avg to stay on budget']]
    bgt_sty = base_tbl()
    for i, (name, budget, spent) in enumerate(budget_data):
        rem = budget - spent; dn = rem / DAYS_REMAINING if DAYS_REMAINING > 0 else 0
        bgt_data.append([name, fmts(budget), fmt(spent), fmt(rem), fmt(dn)])
        bg = LIGHT if i % 2 == 0 else WHITE
        bgt_sty += [('BACKGROUND', (0, i + 1), (2, i + 1), bg), ('BACKGROUND', (4, i + 1), (4, i + 1), bg),
                    ('TOPPADDING', (0, i + 1), (-1, i + 1), 2), ('BOTTOMPADDING', (0, i + 1), (-1, i + 1), 2)]
        if rem < 0:
            bgt_sty += [('BACKGROUND', (3, i + 1), (3, i + 1), RDBG), ('TEXTCOLOR', (3, i + 1), (3, i + 1), RED),
                        ('FONTNAME', (3, i + 1), (3, i + 1), 'Helvetica-Bold')]
        elif budget > 0 and spent / budget > 0.7:
            bgt_sty += [('BACKGROUND', (3, i + 1), (3, i + 1), AMBG), ('TEXTCOLOR', (3, i + 1), (3, i + 1), AMBER),
                        ('FONTNAME', (3, i + 1), (3, i + 1), 'Helvetica-Bold')]
        else:
            bgt_sty += [('TEXTCOLOR', (3, i + 1), (3, i + 1), GREEN), ('FONTNAME', (3, i + 1), (3, i + 1), 'Helvetica-Bold')]
    btbl = Table(bgt_data, colWidths=bgt_cols, rowHeights=RH_B)
    btbl.setStyle(TableStyle(bgt_sty)); btbl.wrapOn(cv, W, H); btbl.drawOn(cv, M, y - BGT_H)

    # Page 2
    cv.showPage(); chrome(2, 2); y = YT
    sec_lbl(y, f'Month-to-Date — {mo_name} {REPORT_DATE.split()[-1]}'); y -= LH + G
    draw_cards(y, CARD_H, [('MTD Total', fmt(mtd_total), BLUE, f'{DAYS_ELAPSED} of {WD} working days'),
        ('MTD Damage', fmt(mtd_damage), RED, ''), ('MTD Tyres', fmt(mtd_tyres), NAVY, '')], 3)
    y -= CARD_H + 3 * mm
    draw_cards(y, CARD_H, [('MTD Capital', fmt(mtd_capital), ORANGE, ''),
        ('Daily Average', fmt(daily_avg), GREEN, f'Target {fmts(mtd_total / WD)} over {WD} days'),
        ('Biggest Day', fmt(bd_val), PURPLE, f'{bd_lbl}')], 3)
    y -= CARD_H + G + PAD2

    sec_lbl(y, 'Day-by-Day Summary'); y -= LH + G
    dcols = [28 * mm, 40 * mm, 32 * mm, 32 * mm, 32 * mm, 22 * mm]
    ddata = [['Date', 'Total Spend', 'Damage', 'Tyres', 'Capital', 'Daily Avg']]
    for dlbl, tot, dmg, tyr, cap in day_rows:
        ddata.append([dlbl, fmt(tot), fmt(dmg), fmt(tyr), fmt(cap), fmt(tot)])
    n_dd = len(ddata)
    ddata.append(['MTD', fmt(mtd_total), fmt(mtd_damage), fmt(mtd_tyres), fmt(mtd_capital), fmt(daily_avg)])
    dsty = base_tbl()
    dsty += [('BACKGROUND', (0, n_dd), (-1, n_dd), colors.HexColor('#E0E3EE')),
             ('FONTNAME', (0, n_dd), (-1, n_dd), 'Helvetica-Bold'),
             ('LINEABOVE', (0, n_dd), (-1, n_dd), 0.5, NAVY)]
    day_tbl = Table(ddata, colWidths=dcols, rowHeights=RH)
    day_tbl.setStyle(TableStyle(dsty))
    day_tbl.wrapOn(cv, W, H); day_tbl.drawOn(cv, M, y - len(ddata) * RH)
    cv.save()


def main():
    if len(sys.argv) < 3:
        print("Usage: python build_report.py <csv> <report_date YYYY-MM-DD> [base_xlsx]")
        sys.exit(1)
    csv_path = sys.argv[1]
    report_date = dt.date.fromisoformat(sys.argv[2])
    base_path = Path(sys.argv[3]) if len(sys.argv) > 3 else (HERE / "base_report.xlsx")
    build(csv_path, report_date, base_path)


if __name__ == "__main__":
    main()
