"""
Autovolt scraper — logs in and downloads the Supplier Transaction Report CSV
for a given date range, saving it as SupplierTransaction.csv.

Usage:
    python scrape_autovolt.py                      # defaults to yesterday
    python scrape_autovolt.py 2026-07-02           # single day
    python scrape_autovolt.py 2026-06-01 2026-06-30  # a range

The report name + URL pattern were discovered from the Autovolt Reports page:
    /reports/tzp/run?report=supplierTransactionReport&start=YYYY-MM-DD&end=YYYY-MM-DD
"""
import os
import sys
import datetime as dt
from pathlib import Path
from playwright.sync_api import sync_playwright

HERE = Path(__file__).parent
OUT_CSV = HERE / "SupplierTransaction.csv"


def load_env():
    """OS env vars (prod) win; local .env fills gaps for development."""
    env = dict(os.environ)
    p = HERE / ".env"
    if p.exists():
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def scrape(start: str, end: str, out_path: Path = OUT_CSV) -> Path:
    env = load_env()
    url_base = env["AUTOVOLT_URL"].rstrip("/")
    report_url = (
        f"{url_base}/reports/tzp/run"
        f"?report=supplierTransactionReport&start={start}&end={end}"
    )

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        ctx = browser.new_context(accept_downloads=True)
        page = ctx.new_page()

        # --- log in ---
        page.goto(f"{url_base}/login", wait_until="domcontentloaded", timeout=60000)
        page.get_by_role("textbox", name="E-Mail Address").fill(env["AUTOVOLT_EMAIL"])
        page.get_by_role("textbox", name="Password").fill(env["AUTOVOLT_PASSWORD"])
        page.get_by_role("button", name="Login").click()
        page.wait_for_load_state("networkidle", timeout=60000)
        if "/login" in page.url:
            raise RuntimeError("Login failed — still on the login page. Check .env credentials.")

        # --- fetch the CSV directly (reuses the logged-in session cookies) ---
        resp = ctx.request.get(report_url, timeout=180000)
        if not resp.ok:
            raise RuntimeError(f"Report download failed: HTTP {resp.status} for {report_url}")
        body = resp.body()
        out_path.write_bytes(body)

        ctx.close()
        browser.close()

    rows = max(0, len(out_path.read_text(encoding="utf-8", errors="replace").splitlines()) - 1)
    print(f"Downloaded {out_path.name}: {rows} data row(s) for {start} to {end}")
    return out_path


def main():
    args = sys.argv[1:]
    if len(args) == 0:
        yesterday = dt.date.today() - dt.timedelta(days=1)
        start = end = yesterday.isoformat()
    elif len(args) == 1:
        start = end = args[0]
    else:
        start, end = args[0], args[1]
    scrape(start, end)


if __name__ == "__main__":
    main()
