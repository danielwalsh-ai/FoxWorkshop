"""
One-off history backfill: pull everything Autovolt holds, MONTH BY MONTH
(a single 11-year query times the server out), and load it into the Workshop
database, each line tagged with its own PO date.

Usage:
    python backfill.py            # from 2022-01 to today
    python backfill.py 2024-01    # from a given YYYY-MM to today
"""
import sys
import datetime as dt
from calendar import monthrange
from pathlib import Path

from playwright.sync_api import sync_playwright
from scrape_autovolt import load_env
from classify import process_all
from vrm_lookup import load_lookup
import db

HERE = Path(__file__).parent


def month_ranges(start_year, start_month, end_date):
    y, m = start_year, start_month
    while (y < end_date.year) or (y == end_date.year and m <= end_date.month):
        first = dt.date(y, m, 1)
        last = dt.date(y, m, monthrange(y, m)[1])
        if last > end_date:
            last = end_date
        yield first.isoformat(), last.isoformat()
        m += 1
        if m > 12:
            m, y = 1, y + 1


def main():
    start_s = sys.argv[1] if len(sys.argv) > 1 else "2022-01"
    sy, sm = int(start_s.split("-")[0]), int(start_s.split("-")[1])
    today = dt.date.today()

    env = load_env()
    url_base = env["AUTOVOLT_URL"].rstrip("/")
    reg_to_area, _, reg_plate = load_lookup()
    tmp = HERE / "backfill_month.csv"

    total_rows, grand, first_data = 0, 0.0, None

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context()
        page = ctx.new_page()
        page.goto(f"{url_base}/login", wait_until="domcontentloaded", timeout=60000)
        page.get_by_role("textbox", name="E-Mail Address").fill(env["AUTOVOLT_EMAIL"])
        page.get_by_role("textbox", name="Password").fill(env["AUTOVOLT_PASSWORD"])
        page.get_by_role("button", name="Login").click()
        page.wait_for_load_state("networkidle", timeout=60000)
        if "/login" in page.url:
            raise RuntimeError("Login failed — check .env credentials.")

        for start, end in month_ranges(sy, sm, today):
            tag = start[:7]
            url = (f"{url_base}/reports/tzp/run?report=supplierTransactionReport"
                   f"&start={start}&end={end}")
            try:
                resp = ctx.request.get(url, timeout=180000)
                if not resp.ok:
                    print(f"  {tag}: HTTP {resp.status} — skipped")
                    continue
                tmp.write_bytes(resp.body())
            except Exception as e:
                print(f"  {tag}: error {e!r} — skipped")
                continue

            df = process_all(tmp, reg_to_area, reg_plate)
            if df.empty:
                print(f"  {tag}: (no data)")
                continue
            n = db.ingest_bulk(df)
            total_rows += n
            grand += float(df["Cost"].sum())
            first_data = first_data or tag
            print(f"  {tag}: {n} rows  £{df['Cost'].sum():,.2f}")

        ctx.close()
        browser.close()

    print(f"\nBackfill complete: {total_rows} rows | earliest data {first_data} | total £{grand:,.2f}")


if __name__ == "__main__":
    main()
