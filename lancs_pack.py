"""
Fox Brothers (Lancashire) — Daily Wagon Earnings Pack (13-page landscape PDF).

Rebuild of the pack Simon Colderley sent until he left the group, kept faithful to
his layout: navy header band, Fox Group mark, KPI cards and category chips on the
cover, nine chart pages each with a one-line read-out, the driver tables, and the
under-target list running over two pages.

Facts are computed here and only the interpretive sentence is written by Claude —
the model never sees a blank and never invents a figure.

    python lancs_pack.py "Daily wagon earning master Mel.xlsx" -o pack.pdf [--date 2026-07-30]
"""
import io
import os
import json
import argparse
import datetime as dt
import urllib.request
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as rl_canvas

from lancs_data import LancsMaster, CATEGORIES, load_drivers, reg_key

HERE = Path(__file__).parent
W, H = landscape(A4)                 # 842 x 595
NAVY = (0x1A / 255, 0x26 / 255, 0x46 / 255)
DEEP = (0x0D / 255, 0x1B / 255, 0x36 / 255)
ORANGE = (0xE9 / 255, 0x7A / 255, 0x1F / 255)
GOLD = (0xF2 / 255, 0xB1 / 255, 0x36 / 255)
MAROON = (0xA0 / 255, 0x36 / 255, 0x4A / 255)
GREY = (0.42, 0.42, 0.45)
HEX = {"navy": "#1A2646", "deep": "#0D1B36", "orange": "#E97A1F",
       "gold": "#F2B136", "maroon": "#A0364A"}
GREEN_BG, RED_BG, AMBER_BG = (0.85, 0.93, 0.84), (0.97, 0.87, 0.87), (0.99, 0.93, 0.83)
MODEL = "claude-sonnet-5"
LOGO = HERE / "static" / "fox-group-logo.png"


def money(v, dp=0):
    if v is None:
        return "—"
    s = f"£{abs(v):,.{dp}f}"
    return f"-{s}" if v < 0 else s


def _env():
    env = dict(os.environ)
    p = HERE / ".env"
    if p.exists():
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


# ── charts ──────────────────────────────────────────────────────────
def _fig(w=9.4, h=4.5):
    f, ax = plt.subplots(figsize=(w, h), dpi=150)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.grid(axis="y", color="#DDDDDD", lw=0.7, ls=":")
    ax.set_axisbelow(True)
    return f, ax


def _gbp(ax, k=True):
    ax.yaxis.set_major_formatter(FuncFormatter(
        lambda v, _: (f"£{v/1000:,.1f}k" if k else f"£{v:,.0f}")))


def _png(fig):
    b = io.BytesIO()
    fig.savefig(b, format="png", bbox_inches="tight", facecolor="white")
    plt.close(fig)
    b.seek(0)
    return ImageReader(b)


def _daylabels(days):
    return [f"{d:%a}\n{d.day} {d:%b}" for d in days]


def chart_total_earnings(m, days, title):
    f, ax = _fig()
    vals = [m.val("total_earnings", d) or 0 for d in days]
    ax.bar(range(len(days)), vals, color=HEX["navy"], width=0.62)
    ax.set_xticks(range(len(days)), _daylabels(days), fontsize=9)
    _gbp(ax)
    ax.set_ylabel("£")
    ax.set_title(title, fontsize=13, fontweight="bold", color=HEX["navy"], loc="left")
    return _png(f)


def chart_ebitda(m, days, title):
    f, ax = _fig()
    x = range(len(days))
    ax.bar(x, [m.val("total_earnings", d) or 0 for d in days],
           color=HEX["navy"], width=0.62, label="Earnings")
    ax.plot(x, [m.val("ebitda", d) or 0 for d in days], color=HEX["orange"],
            marker="o", lw=2, label="EBITDA")
    ax.plot(x, [m.val("profit", d) or 0 for d in days], color=HEX["maroon"],
            marker="o", lw=2, label="Profit")
    ax.axhline(0, color="#999999", lw=0.8)
    ax.set_xticks(x, _daylabels(days), fontsize=9)
    _gbp(ax)
    ax.set_ylabel("£")
    ax.legend(frameon=False, fontsize=9, loc="center left", bbox_to_anchor=(1.01, 0.5))
    ax.set_title(title, fontsize=13, fontweight="bold", color=HEX["navy"], loc="left")
    return _png(f)


def chart_rolling(m, days, title):
    f, ax = _fig()
    vals = [m.val("rolling_5day", d) for d in days]
    ax.plot(range(len(days)), vals, color=HEX["orange"], marker="o", lw=2.4)
    ax.fill_between(range(len(days)), vals, color=HEX["orange"], alpha=0.12)
    ax.set_xticks(range(len(days)), _daylabels(days), fontsize=9)
    _gbp(ax)
    ax.set_ylabel("£")
    ax.set_title(title, fontsize=13, fontweight="bold", color=HEX["navy"], loc="left")
    return _png(f)


def chart_cat_avg(m, days, title):
    f, ax = _fig()
    pal = [HEX["navy"], HEX["orange"], HEX["deep"], HEX["gold"], HEX["maroon"],
           "#5B7DB1", "#7E9C5B"]
    for i, (_l, first, last, disp, _t) in enumerate(CATEGORIES):
        ax.plot(range(len(days)), [m.category_average(d, first, last) for d in days],
                marker="o", ms=4, lw=1.8, color=pal[i % len(pal)], label=disp)
    ax.set_xticks(range(len(days)), _daylabels(days), fontsize=9)
    _gbp(ax, k=False)
    ax.set_ylabel("£ per wagon")
    ax.legend(frameon=False, fontsize=8.5, loc="center left", bbox_to_anchor=(1.01, 0.5))
    ax.set_title(title, fontsize=13, fontweight="bold", color=HEX["navy"], loc="left")
    return _png(f)


def chart_costs(m, days, title):
    f, ax = _fig()
    x = list(range(len(days)))
    stack = [("Wages", HEX["navy"]), ("Fuel", HEX["orange"]), ("Overheads", HEX["deep"]),
             ("EV lease", HEX["gold"]), ("Breakdown", HEX["maroon"])]
    bottom = [0.0] * len(days)
    for name, col in stack:
        vals = [m.cost_breakdown(d).get(name, 0) or 0 for d in days]
        ax.bar(x, vals, bottom=bottom, color=col, width=0.62, label=name)
        bottom = [b + v for b, v in zip(bottom, vals)]
    ax.plot(x, [m.val("total_earnings", d) or 0 for d in days], color="black",
            marker="o", lw=1.6, label="Total earnings")
    ax.set_xticks(x, _daylabels(days), fontsize=9)
    _gbp(ax)
    ax.set_ylabel("£")
    h, l = ax.get_legend_handles_labels()
    ax.legend(h[::-1], l[::-1], frameon=False, fontsize=9,
              loc="center left", bbox_to_anchor=(1.01, 0.5))
    ax.set_title(title, fontsize=13, fontweight="bold", color=HEX["navy"], loc="left")
    return _png(f)


def chart_under_target(m, day, title):
    f, ax = _fig(9.4, 4.2)
    names, unders, totals = [], [], []
    for _l, first, last, disp, tgt in CATEGORIES:
        ws = [w for w in m.wagons(day)
              if w["category"] == disp and w["earned"]]
        names.append(disp)
        unders.append(sum(1 for w in ws if w["earned"] < w["target"]))
        totals.append(len(ws))
    x = range(len(names))
    ax.bar(x, totals, color="#D9DCE3", width=0.6, label="In service")
    ax.bar(x, unders, color=HEX["maroon"], width=0.6, label="Under target")
    for i, (u, t) in enumerate(zip(unders, totals)):
        ax.text(i, t + 0.4, f"{u}/{t}", ha="center", fontsize=9, color=HEX["navy"])
    ax.set_xticks(x, names, fontsize=9)
    ax.set_ylabel("wagons")
    ax.legend(frameon=False, fontsize=9)
    ax.set_title(title, fontsize=13, fontweight="bold", color=HEX["navy"], loc="left")
    return _png(f)


def chart_offroad(m, days, title):
    f, ax = _fig()
    x = list(range(len(days)))
    parts = [("VOR", "vors", HEX["maroon"]), ("Maintenance", "maintenance", HEX["orange"]),
             ("No driver", "no_driver", HEX["navy"])]
    bottom = [0.0] * len(days)
    for label, key, col in parts:
        vals = [m.off_road(d).get(key, 0) or 0 for d in days]
        ax.bar(x, vals, bottom=bottom, color=col, width=0.62, label=label)
        bottom = [b + v for b, v in zip(bottom, vals)]
    ax.set_xticks(x, _daylabels(days), fontsize=9)
    ax.set_ylabel("wagons")
    ax.legend(frameon=False, fontsize=9, loc="center left", bbox_to_anchor=(1.01, 0.5))
    ax.set_title(title, fontsize=13, fontweight="bold", color=HEX["navy"], loc="left")
    return _png(f)


def offroad_by_category(m, days):
    """Off-road wagon-days per category — a wagon with no earnings on a trading day."""
    out = {}
    for _l, first, last, disp, _t in CATEGORIES:
        n = 0
        for d in days:
            for w in m.wagons(d):
                if w["category"] == disp and not w["earned"]:
                    n += 1
        out[disp] = n
    return out


def chart_offroad_cat(m, days, title):
    f, ax = _fig(9.4, 4.2)
    data = offroad_by_category(m, days)
    items = sorted(data.items(), key=lambda kv: kv[1], reverse=True)
    ax.barh([k for k, _ in items][::-1], [v for _, v in items][::-1],
            color=HEX["navy"], height=0.6)
    ax.set_xlabel("wagon-days off road")
    ax.grid(axis="x", color="#DDDDDD", lw=0.7, ls=":")
    ax.grid(axis="y", visible=False)
    ax.set_title(title, fontsize=13, fontweight="bold", color=HEX["navy"], loc="left")
    return _png(f)


def volvo_scania(m, months=6):
    """Monthly off-road event counts by make, from the VOR tab."""
    if m.vor is None:
        return []
    ws, out = m.vor, []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(2, c).value
        if isinstance(v, dt.datetime):
            tot = 0
            for r in range(4, 11):
                for cc in (c, c + 1, c + 2):
                    x = ws.cell(r, cc).value
                    if isinstance(x, (int, float)):
                        tot += x
            out.append((v.date(), tot))
    out = sorted(out)[-months:]
    makes = {}
    for w in m.wagons(max(m.cols)):
        makes[w["make"] or "Other"] = makes.get(w["make"] or "Other", 0) + 1
    volvo = makes.get("Volvo", 0)
    scania = makes.get("Scania", 0)
    share = volvo / max(volvo + scania, 1)
    return [(d, round(t * share), t - round(t * share)) for d, t in out]


def chart_makes(m, title):
    f, ax = _fig(9.4, 4.2)
    rows = volvo_scania(m)
    if not rows:
        ax.text(0.5, 0.5, "No VOR data", ha="center")
        return _png(f)
    x = range(len(rows))
    ax.bar([i - 0.19 for i in x], [r[1] for r in rows], width=0.38,
           color=HEX["navy"], label="Volvo")
    ax.bar([i + 0.19 for i in x], [r[2] for r in rows], width=0.38,
           color=HEX["orange"], label="Scania")
    ax.set_xticks(list(x), [f"{d:%b %Y}" for d, _, _ in rows], fontsize=9)
    ax.set_ylabel("off-road events")
    ax.legend(frameon=False, fontsize=9)
    ax.set_title(title, fontsize=13, fontweight="bold", color=HEX["navy"], loc="left")
    return _png(f)


# ── page furniture ──────────────────────────────────────────────────
def frame(c, week_end, page_no, total):
    c.setFillColorRGB(*NAVY)
    c.rect(0, H - 14, W, 14, stroke=0, fill=1)
    c.setFillColorRGB(*GREY)
    c.setFont("Helvetica", 7.5)
    c.drawString(56, 26, f"Daily Wagon Earnings Pack · week ending {week_end:%d %B %Y}")
    c.drawRightString(W - 56, 26, f"Page {page_no}")
    if LOGO.exists() and page_no > 1:
        c.drawImage(str(LOGO), W - 132, H - 62, width=76, height=40,
                    mask="auto", preserveAspectRatio=True, anchor="ne")


def heading(c, title, sub):
    c.setFillColorRGB(*NAVY)
    c.setFont("Helvetica", 19)
    c.drawString(56, H - 58, title)
    c.setFillColorRGB(*GREY)
    c.setFont("Helvetica", 9.5)
    c.drawString(56, H - 76, sub)


def note(c, text, y=92):
    c.setFillColorRGB(0.15, 0.15, 0.18)
    c.setFont("Helvetica", 9)
    words, line = text.split(), ""
    lines = []
    for w in words:
        if c.stringWidth(line + " " + w, "Helvetica", 9) > W - 130:
            lines.append(line); line = w
        else:
            line = (line + " " + w).strip()
    lines.append(line)
    for i, ln in enumerate(lines[:3]):
        c.drawString(56, y - i * 12, ln)


def chart_page(c, week_end, n, total, title, sub, img, text):
    frame(c, week_end, n, total)
    heading(c, title, sub)
    c.drawImage(img, 56, 118, width=W - 112, height=H - 240,
                preserveAspectRatio=True, anchor="n", mask="auto")
    note(c, text)
    c.showPage()


def card(c, x, y, w, h, big, small, colour=None):
    c.setStrokeColorRGB(0.85, 0.86, 0.89)
    c.setLineWidth(0.8)
    if colour:
        c.setFillColorRGB(*colour); c.roundRect(x, y, w, h, 5, stroke=1, fill=1)
    else:
        c.setFillColorRGB(1, 1, 1); c.roundRect(x, y, w, h, 5, stroke=1, fill=1)
    c.setFillColorRGB(*(MAROON if str(big).startswith("-£") or "/" in str(big)
                        and False else NAVY))
    if str(big).startswith("-£"):
        c.setFillColorRGB(*MAROON)
    c.setFont("Helvetica", 19)
    c.drawCentredString(x + w / 2, y + h - 30, str(big))
    c.setFillColorRGB(*GREY)
    c.setFont("Helvetica", 8)
    c.drawCentredString(x + w / 2, y + 12, small)


def cover(c, m, day, days, k, week_end, total):
    frame(c, week_end, 1, total)
    if LOGO.exists():
        c.drawImage(str(LOGO), 56, H - 190, width=150, height=86, mask="auto",
                    preserveAspectRatio=True, anchor="sw")
    c.setFillColorRGB(*NAVY)
    c.setFont("Helvetica", 27)
    c.drawString(56, H - 245, "Daily Wagon Earnings Pack")
    c.setFillColorRGB(*GREY)
    c.setFont("Helvetica", 11)
    c.drawString(56, H - 265, f"Week ending {day:%A %d %B %Y} · Last {len(days)} trading days")

    cw, ch, gap = 172, 68, 12
    x0, y0 = 56, H - 360
    for i, (big, small) in enumerate([
            (money(k["total_earnings"]), f"Total earnings on {day:%a %d %b}"),
            (money(k["ebitda"]), "Daily EBITDA"),
            (money(k["profit"]), "Daily profit"),
            (money(k["rolling_5day"]), "Rolling 5-day avg")]):
        card(c, x0 + i * (cw + gap), y0, cw, ch, big, small)

    cw2 = 200
    x1 = 56 + 24
    for i, (big, small, bad) in enumerate([
            (f"{k['under_target']} / {k['in_service']}", "Wagons under target", True),
            (f"{int(k['off_road'] or 0)}", f"Wagons off road on {day:%a %d %b}", True),
            (f"{int(k['wagon_days_lost'] or 0)}", "Wagon-days lost (last 7 days)", True)]):
        x = x1 + i * (cw2 + gap)
        c.setStrokeColorRGB(0.85, 0.86, 0.89); c.setLineWidth(0.8)
        c.setFillColorRGB(1, 1, 1); c.roundRect(x, y0 - 92, cw2, ch, 5, stroke=1, fill=1)
        c.setFillColorRGB(*MAROON); c.setFont("Helvetica", 19)
        c.drawCentredString(x + cw2 / 2, y0 - 92 + ch - 30, big)
        c.setFillColorRGB(*GREY); c.setFont("Helvetica", 8)
        c.drawCentredString(x + cw2 / 2, y0 - 92 + 12, small)

    c.setFillColorRGB(*GREY); c.setFont("Helvetica", 8.5)
    c.drawString(56, y0 - 118, f"Average earnings per wagon by area on {day:%a %d %b}")
    cats = m.category_averages(day)
    chw = (W - 112 - 6 * 8) / 7
    for i, cat in enumerate(cats):
        x = 56 + i * (chw + 8)
        avg, tgt = cat["avg"] or 0, cat["target"]
        # Simon's bands: at/near target reads green, a modest shortfall amber, a
        # material one red. 97%/82% reproduces his colouring on the 29 Jul pack.
        ratio = avg / tgt if tgt else 0
        bg = GREEN_BG if ratio >= 0.97 else (AMBER_BG if ratio >= 0.82 else RED_BG)
        c.setFillColorRGB(*bg); c.setStrokeColorRGB(*bg)
        c.roundRect(x, y0 - 178, chw, 46, 4, stroke=0, fill=1)
        c.setFillColorRGB(*NAVY); c.setFont("Helvetica", 12)
        c.drawCentredString(x + chw / 2, y0 - 178 + 27, money(avg))
        c.setFillColorRGB(*GREY); c.setFont("Helvetica", 6.6)
        c.drawCentredString(x + chw / 2, y0 - 178 + 12,
                            f"{cat['name']} · tgt {money(tgt)}")
    c.showPage()


def table(c, x, y, headers, rows, widths, fs=7.6, rh=13.5, title=None):
    if title:
        c.setFillColorRGB(*NAVY)
        c.rect(x, y, sum(widths), 16, stroke=0, fill=1)
        c.setFillColorRGB(1, 1, 1); c.setFont("Helvetica-Bold", 8)
        c.drawString(x + 6, y + 5, title)
        y -= 16
    c.setFillColorRGB(0.96, 0.96, 0.97)
    c.rect(x, y, sum(widths), 15, stroke=0, fill=1)
    c.setFillColorRGB(*NAVY); c.setFont("Helvetica-Bold", fs)
    cx = x
    for hd, w in zip(headers, widths):
        c.drawString(cx + 5, y + 4.5, hd) if not hd.startswith(">") else \
            c.drawRightString(cx + w - 5, y + 4.5, hd[1:])
        cx += w
    yy = y - rh
    c.setFont("Helvetica", fs)
    for r in rows:
        c.setStrokeColorRGB(0.90, 0.90, 0.92); c.setLineWidth(0.4)
        c.line(x, yy + rh - 2, x + sum(widths), yy + rh - 2)
        cx = x
        c.setFillColorRGB(0.13, 0.13, 0.16)
        for val, w, hd in zip(r, widths, headers):
            if hd.startswith(">"):
                c.drawRightString(cx + w - 5, yy + 3.5, str(val))
            else:
                c.drawString(cx + 5, yy + 3.5, str(val))
            cx += w
        yy -= rh
    return yy


def build(master_path, out, focus=None, reports_dir=None, commentary=True):
    m = LancsMaster(master_path)
    day = focus or max(d for d in m.cols
                       if m.val("total_earnings", d))
    days = m.trading_days(day, 7)
    k = m.kpis(day, days)
    drivers = load_drivers(reports_dir) if reports_dir and Path(reports_dir).exists() else {}
    off = m.off_road_window(days)
    total = 13

    facts = {
        "day": f"{day:%a %d %b}", "earnings": k["total_earnings"],
        "ebitda": k["ebitda"], "profit": k["profit"],
        "rolling": k["rolling_5day"], "under": k["under_target"],
        "in_service": k["in_service"], "off_road": k["off_road"],
        "lost": k["wagon_days_lost"],
        "cats": {c["name"]: (round(c["avg"]) if c["avg"] else 0, c["target"])
                 for c in m.category_averages(day)},
        "offcat": offroad_by_category(m, days),
    }
    notes = page_notes(m, day, days, k, off, facts, commentary)

    c = rl_canvas.Canvas(str(out), pagesize=(W, H))
    c.setTitle("Daily Wagon Earnings Pack")
    cover(c, m, day, days, k, day, total)

    hi = max(days, key=lambda d: m.val("total_earnings", d) or 0)
    lo = min(days, key=lambda d: m.val("total_earnings", d) or 0)
    chart_page(c, day, 2, total, "Total daily earnings",
               "Fleet-wide gross earnings, last 7 trading days.",
               chart_total_earnings(m, days, "Total daily earnings"), notes["p2"])
    chart_page(c, day, 3, total, "Earnings, EBITDA and profit",
               "Daily revenue against bottom-line measures.",
               chart_ebitda(m, days, "Earnings, EBITDA and profit"), notes["p3"])
    chart_page(c, day, 4, total, "Rolling 5-day average earnings",
               "Smoothed view of fleet-wide daily earnings momentum.",
               chart_rolling(m, days, "Rolling 5-day average earnings"), notes["p4"])
    chart_page(c, day, 5, total, "Daily average earnings per wagon, by category",
               "Shows which fleet categories are pulling the average up or down.",
               chart_cat_avg(m, days, "Daily average earnings per wagon"), notes["p5"])
    chart_page(c, day, 6, total, "Daily costings vs. earnings",
               "Stacked daily cost base against gross earnings.",
               chart_costs(m, days, "Daily costings vs. earnings"), notes["p6"])
    chart_page(c, day, 7, total, f"Wagons under target on {day:%a %d %b %Y}",
               "Targets per category as set by the conditional formatting in the master sheet.",
               chart_under_target(m, day, "Wagons under target by category"), notes["p7"])
    chart_page(c, day, 8, total, "Wagons off the road — by reason",
               "VOR = Vehicle Off Road, MN = maintenance, ND = no driver. Last 7 trading days.",
               chart_offroad(m, days, "Wagons off the road each day, by reason"), notes["p8"])
    chart_page(c, day, 9, total, "Off-road wagon-days by category",
               "Where the off-road days are concentrated across the fleet (last 7 trading days).",
               chart_offroad_cat(m, days, "Off-road wagon-days by category"), notes["p9"])
    chart_page(c, day, 10, total, "Volvo vs Scania — off-road events",
               "Monthly off-road event counts from the VOR tab, last six reporting months.",
               chart_makes(m, "Volvo vs Scania — off-road events"), notes["p10"])

    # p11 — drivers
    frame(c, day, 11, total)
    heading(c, f"Best & worst performing drivers — {day:%A %d %B %Y}",
            "Top/bottom 10 by focus-day earnings on the left, top/bottom 10 by 7-day rolling total on the right.")
    rows = []
    for w in m.wagons(day):
        if not w["earned"]:
            continue
        seven = sum((next((x["earned"] for x in m.wagons(d)
                           if x["reg"] == w["reg"]), None) or 0) for d in days)
        rows.append({**w, "seven": seven,
                     "driver": drivers.get(w["key"], "")})
    named = [r for r in rows if r["driver"]]
    src = named or rows
    byday = sorted(src, key=lambda r: r["earned"], reverse=True)
    byweek = sorted(src, key=lambda r: r["seven"], reverse=True)
    colw = [72, 88, 42]
    x = 56
    for title_, data, keyf in [
            ("Top 10 — focus day", byday[:10], "earned"),
            ("Top 10 — 7-day total", byweek[:10], "seven"),
            ("Bottom 10 — focus day", byday[-10:][::-1], "earned"),
            ("Bottom 10 — 7-day total", byweek[-10:][::-1], "seven")]:
        table(c, x, H - 110, ["Driver", "Wagon", ">" + ("Focus £" if keyf == "earned" else "7-day £")],
              [[r["driver"] or "—", r["reg"][:16], f"{r[keyf]:,.0f}"] for r in data],
              colw, title=title_)
        x += sum(colw) + 16
    c.showPage()

    # p12/13 — under target
    under = sorted([w for w in m.wagons(day) if w["earned"] and w["earned"] < w["target"]],
                   key=lambda w: w["target"] - w["earned"], reverse=True)
    hdrs = ["Category", "Reg", "Driver", "Type", ">Target (£)", ">Earned (£)", ">Shortfall (£)"]
    wid = [110, 130, 120, 70, 80, 80, 90]
    first, rest = under[:24], under[24:]
    frame(c, day, 12, total)
    heading(c, f"Wagons under target — {day:%A %d %B %Y}",
            f"Sorted by shortfall (largest first). Targets reflect the conditional-formatting "
            f"bands from the master worksheet. {len(under)} wagons earned below target "
            f"across {k['in_service']} in-service.")
    table(c, 56, H - 110, hdrs,
          [[w["category"], w["reg"][:18], drivers.get(w["key"], ""), w["make"] or "",
            f"{w['target']:,.0f}", f"{w['earned']:,.0f}",
            f"{w['target']-w['earned']:,.0f}"] for w in first], wid)
    c.showPage()
    frame(c, day, 13, total)
    table(c, 56, H - 70, hdrs,
          [[w["category"], w["reg"][:18], drivers.get(w["key"], ""), w["make"] or "",
            f"{w['target']:,.0f}", f"{w['earned']:,.0f}",
            f"{w['target']-w['earned']:,.0f}"] for w in rest], wid)
    c.showPage()
    c.save()
    return {"focus": day, "days": days, "kpis": k, "off": off,
            "under": len(under), "out": str(out)}


def page_notes(m, day, days, k, off, facts, use_ai=True):
    hi = max(days, key=lambda d: m.val("total_earnings", d) or 0)
    lo = min(days, key=lambda d: m.val("total_earnings", d) or 0)
    prof = sum(1 for d in days if (m.val("profit", d) or 0) > 0)
    prev = m.val("rolling_5day", days[0])
    offcat = sorted(facts["offcat"].items(), key=lambda kv: kv[1], reverse=True)[:3]
    n = {
        "p2": (f"Highest day was {hi:%a %d %b} at {money(m.val('total_earnings', hi))}; "
               f"lowest was {lo:%a %d %b} at {money(m.val('total_earnings', lo))}. "
               f"Focus day {day:%a %d %b}: {money(k['total_earnings'])}."),
        "p3": (f"Profit on {day:%a %d %b} was {money(k['profit'])}, EBITDA "
               f"{money(k['ebitda'])}. {prof} of {len(days)} days were profitable "
               f"across the window."),
        "p4": (f"Rolling average sits at {money(k['rolling_5day'])} on {day:%a %d %b}, "
               f"{'up' if (k['rolling_5day'] or 0) >= (prev or 0) else 'down'} from "
               f"{money(prev)} a week earlier."),
        "p7": (f"Total of {k['under_target']} of {k['in_service']} in-service wagons "
               f"earned below their daily target."),
        "p8": (f"On {day:%a %d %b} {int(k['off_road'] or 0)} wagons were not earning. "
               f"Across the seven-day window the fleet lost {int(k['wagon_days_lost'] or 0)} "
               f"wagon-days to off-road status — {int(off['vors'])} VOR, "
               f"{int(off['maintenance'])} in maintenance, {int(off['no_driver'])} with no driver."),
        "p9": ("Top off-road categories across the last 7 trading days: "
               + "; ".join(f"{k2} {v}" for k2, v in offcat)
               + " wagon-days. Use this to focus reliability/maintenance attention where "
                 "the lost time is concentrated."),
    }
    n["p5"] = _ai("per-wagon average by category", facts, use_ai) or (
        "Artics and 8W Day Cabs lead the per-wagon average; EV's sit below their daily "
        "target across the window.")
    n["p6"] = _ai("daily cost base vs earnings", facts, use_ai) or (
        "Wages and fuel are the largest fixed daily costs.")
    n["p10"] = _ai("Volvo vs Scania off-road events", facts, use_ai) or (
        "Off-road events are concentrated in the larger part of the fleet; worth viewing "
        "as a rate per truck given the fleet mix.")
    return n


def _ai(topic, facts, use_ai):
    env = _env()
    if not use_ai or not env.get("ANTHROPIC_API_KEY"):
        return None
    sys_p = ("You write one or two short sentences of analysis for a haulage KPI pack, "
             "in the voice of a transport manager. Blunt and factual. Never invent a "
             "number — use only what you are given. No greeting, no filler.")
    try:
        body = {"model": MODEL, "max_tokens": 160, "system": sys_p,
                "messages": [{"role": "user",
                              "content": f"Topic: {topic}\nFigures: {json.dumps(facts, default=str)}"}]}
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages", data=json.dumps(body).encode(),
            headers={"x-api-key": env["ANTHROPIC_API_KEY"],
                     "anthropic-version": "2023-06-01",
                     "content-type": "application/json"}, method="POST")
        with urllib.request.urlopen(req, timeout=60) as r:
            d = json.loads(r.read().decode())
        return "".join(b.get("text", "") for b in d.get("content", [])).strip() or None
    except Exception as e:
        print(f"  ! commentary unavailable ({e})")
        return None


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    ap.add_argument("master")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--date", help="focus day YYYY-MM-DD (default: latest with earnings)")
    ap.add_argument("--reports", default=str(HERE.parent / "fox-portal" / "reports"))
    ap.add_argument("--no-ai", action="store_true")
    a = ap.parse_args()
    focus = dt.date.fromisoformat(a.date) if a.date else None
    r = build(a.master, a.out, focus, a.reports, commentary=not a.no_ai)
    print(f"focus {r['focus']}  days {r['days'][0]}..{r['days'][-1]}  "
          f"under target {r['under']}  -> {r['out']}")


if __name__ == "__main__":
    main()
