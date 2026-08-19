"""Trend charts for the daily workshop spend report (Paul Fox request).

  daily_chart   — spend per day across the current month, with a trend line and
                  the month's average-per-day reference.
  monthly_chart — average spend per day for the last 6 months (rolling: the
                  oldest month drops off as each new month starts).

Both render to PNG for embedding in the PDF and the spreadsheet. Fox brand:
navy #24214A bars, orange #EB941F accent. Single measure, so no legend.
"""
import os
import datetime as dt
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import psycopg2

HERE = Path(__file__).parent
NAVY = "#24214A"
ORANGE = "#EB941F"
INK = "#3A3A50"
MUTED = "#8C8CA6"
GRID = "#E7E7EF"


def _conn():
    url = os.environ.get("WORKSHOP_DATABASE_URL")
    if not url:
        for line in (HERE / ".env").read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("WORKSHOP_DATABASE_URL="):
                url = line.strip().split("=", 1)[1]
    return psycopg2.connect(url)


def _style(ax):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(GRID)
    ax.spines["bottom"].set_color(GRID)
    ax.tick_params(colors=MUTED, labelsize=8, length=0)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"£{v/1000:.0f}k" if v >= 1000 else f"£{v:.0f}"))
    ax.grid(axis="y", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)


def daily_chart(report_date: dt.date, out_png: Path) -> Path:
    first = report_date.replace(day=1)
    with _conn() as c, c.cursor() as cur:
        cur.execute("""SELECT report_date, ROUND(SUM(cost),2) FROM transactions
                       WHERE report_date>=%s AND report_date<=%s GROUP BY 1 ORDER BY 1""",
                    (first, report_date))
        rows = cur.fetchall()
    days = [r[0].day for r in rows]
    vals = [float(r[1] or 0) for r in rows]
    avg = sum(vals) / len(vals) if vals else 0

    fig, ax = plt.subplots(figsize=(9.2, 3.3), dpi=150)
    ax.bar(days, vals, width=0.66, color=NAVY, zorder=3)
    if len(days) >= 2:
        z = np.polyfit(days, vals, 1)
        xs = np.array([min(days), max(days)])
        ax.plot(xs, z[0] * xs + z[1], color=ORANGE, linewidth=2.4, zorder=5)
    ax.axhline(avg, color=MUTED, linewidth=1.0, linestyle=(0, (4, 3)), zorder=2)
    ax.annotate(f"avg £{avg:,.0f}/day", xy=(min(days) if days else 1, avg),
                xytext=(2, 4), textcoords="offset points", ha="left",
                fontsize=8, color=MUTED, zorder=6)
    # figures ON the bars — white inside tall bars (clear of the trend line),
    # small and above the very short ones that can't hold a label.
    mx = max(vals) if vals else 1
    for xi, v in zip(days, vals):
        if v > mx * 0.28:
            ax.annotate(f"£{v:,.0f}", xy=(xi, v), xytext=(0, -6), textcoords="offset points",
                        ha="center", va="top", rotation=90, fontsize=7, color="white",
                        fontweight="bold", zorder=6)
        else:
            ax.annotate(f"£{v:,.0f}", xy=(xi, v), xytext=(0, 3), textcoords="offset points",
                        ha="center", rotation=90, va="bottom", fontsize=6.5, color=INK, zorder=6)
    ax.set_xticks(days)
    ax.set_xlabel("Day of month", fontsize=8, color=MUTED)
    _style(ax)
    ax.set_title(f"Spend per day — {report_date:%B %Y}   (orange = trend)",
                 fontsize=11, color=NAVY, fontweight="bold", loc="left", pad=8)
    fig.tight_layout()
    fig.savefig(out_png, bbox_inches="tight")
    plt.close(fig)
    return out_png


def monthly_chart(report_date: dt.date, out_png: Path, months: int = 6) -> Path:
    with _conn() as c, c.cursor() as cur:
        cur.execute("""SELECT to_char(report_date,'YYYY-MM') ym,
                              COUNT(DISTINCT report_date) d, ROUND(SUM(cost),2) t
                       FROM transactions WHERE report_date<=%s
                       GROUP BY 1 ORDER BY 1""", (report_date,))
        rows = cur.fetchall()
    series = [(ym, float(t or 0) / max(d, 1)) for ym, d, t in rows][-months:]
    labels = [dt.datetime.strptime(ym, "%Y-%m").strftime("%b %y") for ym, _ in series]
    vals = [v for _, v in series]

    fig, ax = plt.subplots(figsize=(9.2, 3.0), dpi=150)
    x = np.arange(len(vals))
    ax.bar(x, vals, width=0.55, color=NAVY, zorder=3)
    for xi, v in zip(x, vals):
        ax.annotate(f"£{v:,.0f}", xy=(xi, v), xytext=(0, 5), textcoords="offset points",
                    ha="center", fontsize=9, color=INK, fontweight="bold", zorder=6)
    ax.set_xticks(x)
    ax.set_xticklabels(labels)
    _style(ax)
    ax.margins(y=0.18)
    ax.set_title(f"Average spend per day — last {len(vals)} months",
                 fontsize=11, color=NAVY, fontweight="bold", loc="left", pad=8)
    fig.tight_layout()
    fig.savefig(out_png, bbox_inches="tight")
    plt.close(fig)
    return out_png


def _avg_bar(labels, vals, title, out_png, width=8.6, height=2.15, rot=0):
    fig, ax = plt.subplots(figsize=(width, height), dpi=150)
    x = np.arange(len(vals))
    ax.bar(x, vals, width=0.62, color=NAVY, zorder=3)
    for xi, v in zip(x, vals):
        ax.annotate(f"£{v:,.0f}", xy=(xi, v), xytext=(0, 4), textcoords="offset points",
                    ha="center", fontsize=8, color=INK, fontweight="bold", zorder=6)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=rot, ha="right" if rot else "center", fontsize=8)
    _style(ax)
    ax.margins(y=0.20)
    ax.set_title(title, fontsize=11, color=NAVY, fontweight="bold", loc="left", pad=8)
    fig.tight_layout()
    fig.savefig(out_png, bbox_inches="tight")
    plt.close(fig)
    return out_png


def age_avg_chart(spd, out_png):
    order = [2021, 2022, 2023, 2024, 2025, 2026, "other"]
    lbl = {2021: "2021", 2022: "2022", 2023: "2023", 2024: "2024",
           2025: "2025", 2026: "2026", "other": "Older"}
    labels = [lbl[k] for k in order]
    vals = [spd["age_ytd"].get(k, 0) for k in order]
    return _avg_bar(labels, vals, "Avg spend per day — by age range (YTD)", out_png)


def area_avg_chart(spd, out_png, top=10):
    items = sorted(spd["area_ytd"].items(), key=lambda x: -x[1])[:top]
    labels = [a.title() for a, _ in items]
    vals = [v for _, v in items]
    return _avg_bar(labels, vals, "Avg spend per day — by area (YTD, top 10)", out_png, rot=30)


if __name__ == "__main__":
    d = dt.date(2026, 8, 19)
    daily_chart(d, HERE / "chart_daily.png")
    monthly_chart(d, HERE / "chart_monthly.png")
    import queries
    spd = queries.spend_per_day(d)
    age_avg_chart(spd, HERE / "chart_age_avg.png")
    area_avg_chart(spd, HERE / "chart_area_avg.png")
    print("wrote 4 charts")
