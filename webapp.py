"""
Workshop & Maintenance dashboard — FastAPI web app (the hub tile).
Reads live from the Workshop Postgres database.

Run locally:
    uvicorn webapp:app --reload --port 8080
"""
import datetime as dt
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import queries

HERE = Path(__file__).parent
(HERE / "static").mkdir(exist_ok=True)

app = FastAPI(title="Fox Group — Workshop & Maintenance")
app.mount("/static", StaticFiles(directory=str(HERE / "static")), name="static")
templates = Jinja2Templates(directory=str(HERE / "templates"))


def _fmt(v):
    try:
        return f"£{float(v):,.2f}"
    except (TypeError, ValueError):
        return "£0.00"


def _fmt0(v):
    try:
        return f"£{float(v):,.0f}"
    except (TypeError, ValueError):
        return "£0"


templates.env.filters["money"] = _fmt
templates.env.filters["money0"] = _fmt0


@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request, month: str | None = None):
    months = queries.available_months()
    if not month or month not in months:
        month = months[0] if months else dt.date.today().strftime("%Y-%m")
    y, m = int(month[:4]), int(month[5:7])

    ctx = queries.overview(y, m)
    trend = queries.monthly_totals(18)
    peak = max((t["total"] for t in trend), default=1) or 1
    month_label = dt.date(y, m, 1).strftime("%B %Y")

    return templates.TemplateResponse(request, "dashboard.html", {
        "month": month,
        "month_label": month_label,
        "months": months,
        "trend": trend,
        "trend_peak": peak,
        "transactions": queries.recent_transactions(y, m, 60),
        **ctx,
    })


@app.get("/api/overview")
def api_overview(month: str | None = None):
    months = queries.available_months()
    month = month if month in months else (months[0] if months else None)
    if not month:
        return {"error": "no data"}
    y, m = int(month[:4]), int(month[5:7])
    return {"month": month, **queries.overview(y, m)}


@app.get("/health")
def health():
    return {"ok": True}
