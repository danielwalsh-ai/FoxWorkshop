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
FOX_NAVY = "1A2646"      # Fox Group navy, sampled from the pack
TREND_RED = "C00000"      # trend line

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
    from openpyxl.chart.axis import ChartLines
    from openpyxl.chart.trendline import Trendline
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.line import LineProperties
    from openpyxl.drawing.text import (Paragraph, ParagraphProperties,
                                       CharacterProperties)
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    navy = "FF" + FOX_NAVY
    ws.sheet_view.showGridLines = False          # cleaner behind the charts
    ws["A1"] = "Fox Brothers (Lancashire)"
    ws["A1"].font = Font(name="Calibri", size=20, bold=True, color=navy)
    ws["A2"] = "Monthly summary"
    ws["A2"].font = Font(name="Calibri", size=13, color=navy)
    ws["A3"] = (f"Rows selected by Paul Fox, by month from {min(months):%B %Y}. "
                f"Totals summed; averages and wagon counts averaged.")
    ws["A3"].font = Font(name="Calibri", size=9, italic=True, color="FF7F7F7F")
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 18

    hdr = 5
    thin = Side(style="thin", color="FFD9D9D9")
    edge = Border(bottom=thin)
    hc = ws.cell(hdr, 1, "Metric")
    hc.font = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
    hc.fill = PatternFill("solid", fgColor=navy)
    hc.alignment = Alignment(vertical="center")
    for j, mth in enumerate(months):
        c = ws.cell(hdr, 2 + j, dt.datetime(mth.year, mth.month, 1))
        c.number_format = "mmm yy"
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
        c.fill = PatternFill("solid", fgColor=navy)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[hdr].height = 20

    band = PatternFill("solid", fgColor="FFF4F6F9")
    for i, ((r, lab), series) in enumerate(data.items()):
        rr = hdr + 1 + i
        nc = ws.cell(rr, 1, lab)
        nc.font = Font(name="Calibri", size=9.5, bold=True, color="FF262626")
        nc.border = edge
        for j, mth in enumerate(months):
            cell = ws.cell(rr, 2 + j, series.get(mth))
            cell.number_format = "#,##0;-#,##0;\"–\""
            cell.font = Font(name="Calibri", size=9.5, color="FF262626")
            cell.alignment = Alignment(horizontal="right")
            cell.border = edge
        if i % 2:                                # subtle banding, easier to read across
            for j in range(len(months) + 1):
                ws.cell(rr, 1 + j).fill = band
    last_row = hdr + len(data)

    ws.column_dimensions["A"].width = 32
    for j in range(len(months)):
        ws.column_dimensions[get_column_letter(2 + j)].width = 10.5
    # No frozen panes. A frozen column A clips any chart anchored there — you'd
    # have to widen the column to see it — and a frozen header row leaves month
    # headings floating over empty space once you scroll past the table.

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
        ch.height, ch.width = 7.0, 26
        ch.legend = None
        ch.gapWidth = 55                       # chunkier bars, less white space
        ch.overlap = -10

        s = Series(Reference(ws, min_col=c0, max_col=c1, min_row=rr),
                   title_from_data=False, title=label)
        # every bar the same Fox navy, no outline
        s.graphicalProperties = GraphicalProperties(solidFill=FOX_NAVY)
        s.graphicalProperties.line.noFill = True
        s.trendline = Trendline(trendlineType="linear")
        s.trendline.graphicalProperties = GraphicalProperties()
        s.trendline.graphicalProperties.line = LineProperties(solidFill=TREND_RED, w=22000)
        ch.series.append(s)
        ch.set_categories(Reference(ws, min_col=c0, max_col=c1, min_row=hdr))

        # quiet axes: thin grey gridlines, no clutter
        ch.y_axis.majorGridlines = ChartLines(
            spPr=GraphicalProperties(ln=LineProperties(solidFill="D9D9D9", w=6000)))
        ch.x_axis.majorGridlines = None
        ch.y_axis.numFmt = "#,##0"
        for ax in (ch.x_axis, ch.y_axis):
            ax.majorTickMark = "none"
            ax.minorTickMark = "none"
            ax.spPr = GraphicalProperties(ln=LineProperties(solidFill="BFBFBF", w=6000))
            ax.txPr = RichText(
                p=[Paragraph(pPr=ParagraphProperties(
                    defRPr=CharacterProperties(sz=800, solidFill="595959")), endParaRPr=None)])
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ws.add_chart(ch, f"A{anchor}")
        anchor += STEP
    if skipped:
        print(f"  no data, no chart: {', '.join(skipped)}")

    # Land on the latest month without freezing anything: just park the cursor
    # there so Excel scrolls it into view, and leave the charts unclipped.
    last_col = get_column_letter(1 + len(months))
    try:
        ws.sheet_view.selection[0].activeCell = f"{last_col}{hdr + 1}"
        ws.sheet_view.selection[0].sqref = f"{last_col}{hdr + 1}"
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
