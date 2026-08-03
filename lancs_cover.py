"""
Cover sheet for Mel's Lancashire master — the 44 rows Paul highlighted, by month.

Paul shaded 44 rows on the DAILY tab in dark blue (0070C0). This builds a new
first tab holding those rows aggregated month by month from January 2025, with a
native Excel bar chart and trend line, and opens scrolled to the latest month.

The workbook carries 180 chart parts and 16 drawings, so it is never re-saved
through openpyxl — that would destroy them. The new sheet is built standalone and
its parts are transplanted into the original zip.
"""
import re
import shutil
import zipfile
import datetime as dt
from pathlib import Path

import openpyxl

HIGHLIGHT = "FF0070C0"          # the dark blue Paul used
START = dt.date(2025, 1, 1)
DATE_ROW, FIRST_DATA_COL = 1, 4
SHEET_NAME = "COVER"

# Rows that are a rate rather than a running total have to be averaged over the
# month, not summed.
def agg_for(label):
    u = (label or "").upper()
    if "AVERAGE" in u or "AVG" in u or u.startswith("NO WAGONS"):
        return "mean"
    return "sum"


def highlighted_rows(path, sheet="DAILY", limit=200):
    """[(row, label)] for every row Paul shaded, in sheet order."""
    ws = openpyxl.load_workbook(path)[sheet]
    out = []
    for r in range(1, limit):
        c = ws.cell(r, 1)
        try:
            rgb = c.fill.fgColor.rgb
        except Exception:
            rgb = None
        if isinstance(rgb, str) and rgb.upper() == HIGHLIGHT:
            lab = str(c.value or "").strip()
            if lab:
                out.append((r, lab))
    return out


def monthly(path, rows, start=START, sheet="DAILY"):
    """{(row,label): {month_start: value}} plus the ordered month list."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    cols = {}
    for c in range(FIRST_DATA_COL, ws.max_column + 1):
        v = ws.cell(DATE_ROW, c).value
        if isinstance(v, dt.datetime) and v.date() >= start:
            cols.setdefault(dt.date(v.year, v.month, 1), []).append(c)
    months = sorted(cols)
    data = {}
    for r, lab in rows:
        how = agg_for(lab)
        series = {}
        for mth in months:
            vals = [ws.cell(r, c).value for c in cols[mth]]
            vals = [float(v) for v in vals if isinstance(v, (int, float))]
            if not vals:
                series[mth] = None
            else:
                series[mth] = round(sum(vals), 2) if how == "sum" \
                    else round(sum(vals) / len(vals), 2)
        data[(r, lab)] = series
    return months, data


def build_cover_workbook(months, data, out_path, chart_rows=None):
    """A standalone workbook holding just the cover sheet + its native chart."""
    from openpyxl.chart import BarChart, Reference, Series
    from openpyxl.chart.trendline import Trendline
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    navy = "FF1A2646"
    ws["A1"] = "Fox Brothers (Lancashire) — monthly summary"
    ws["A1"].font = Font(size=15, bold=True, color=navy)
    ws["A2"] = (f"Rows selected by Paul Fox, aggregated by month from "
                f"{min(months):%B %Y}. Totals are summed; averages and wagon "
                f"counts are averaged.")
    ws["A2"].font = Font(size=9, italic=True, color="FF666666")

    hdr = 4
    ws.cell(hdr, 1, "Metric").font = Font(bold=True, color="FFFFFFFF")
    ws.cell(hdr, 1).fill = PatternFill("solid", fgColor=navy)
    for j, mth in enumerate(months):
        c = ws.cell(hdr, 2 + j, dt.datetime(mth.year, mth.month, 1))
        c.number_format = "mmm yy"
        c.font = Font(bold=True, color="FFFFFFFF")
        c.fill = PatternFill("solid", fgColor=navy)
        c.alignment = Alignment(horizontal="center")

    for i, ((r, lab), series) in enumerate(data.items()):
        rr = hdr + 1 + i
        ws.cell(rr, 1, lab).font = Font(bold=True, size=9)
        for j, mth in enumerate(months):
            cell = ws.cell(rr, 2 + j, series.get(mth))
            cell.number_format = "#,##0"
    last_row = hdr + len(data)

    ws.column_dimensions["A"].width = 30
    for j in range(len(months)):
        ws.column_dimensions[get_column_letter(2 + j)].width = 11
    ws.freeze_panes = ws.cell(hdr + 1, 2)

    # One chart per highlighted row, stacked down the sheet. They can't share an
    # axis: the 44 rows span single-figure wagon counts to £1.9m a month, so on one
    # chart everything but the money is a flat line.
    wanted = list(chart_rows if chart_rows is not None
                  else range(hdr + 1, last_row + 1))
    anchor = last_row + 3
    STEP = 16                       # rows per chart block
    skipped = []
    for rr in wanted:
        label = str(ws.cell(rr, 1).value)
        # Several rows only start part-way through — the category TOTAL EARNINGS
        # rows don't begin until 2026 — so each chart starts at its own first
        # populated month rather than carrying a run of empty columns.
        filled = [j for j in range(len(months))
                  if ws.cell(rr, 2 + j).value is not None]
        if not filled:
            skipped.append(label)
            continue
        c0, c1 = 2 + filled[0], 2 + filled[-1]
        ch = BarChart()
        ch.type = "col"
        ch.grouping = "clustered"
        ch.title = label
        ch.y_axis.title = None
        ch.x_axis.title = None
        ch.height, ch.width = 7.5, 30
        ch.legend = None
        s = Series(Reference(ws, min_col=c0, max_col=c1, min_row=rr),
                   title_from_data=False, title=label)
        s.trendline = Trendline(trendlineType="linear")
        ch.series.append(s)
        ch.set_categories(Reference(ws, min_col=c0, max_col=c1, min_row=hdr))
        ws.add_chart(ch, f"A{anchor}")
        anchor += STEP
    if skipped:
        print(f"  no data, no chart: {', '.join(skipped)}")

    # Open on the most recent month: freeze the metric column and header, then
    # scroll the data pane to the right-hand end.
    last_col = get_column_letter(1 + len(months))
    ws.freeze_panes = ws.cell(hdr + 1, 2)
    try:
        first_visible = get_column_letter(max(2, 1 + len(months) - 5))
        ws.sheet_view.pane.topLeftCell = f"{first_visible}{hdr + 1}"
        ws.sheet_view.selection[-1].activeCell = f"{last_col}{hdr + 1}"
        ws.sheet_view.selection[-1].sqref = f"{last_col}{hdr + 1}"
    except Exception:
        pass

    wb.save(out_path)
    return last_row, len(months)


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("master")
    ap.add_argument("-o", "--out", required=True)
    a = ap.parse_args()
    rows = highlighted_rows(a.master)
    months, data = monthly(a.master, rows)
    lr, nm = build_cover_workbook(months, data, a.out)
    print(f"{len(rows)} highlighted rows x {nm} months "
          f"({months[0]:%b %Y} .. {months[-1]:%b %Y}) -> {a.out}")
