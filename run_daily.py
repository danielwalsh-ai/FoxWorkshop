"""
Daily orchestrator — the single command the 6pm schedule runs.

  1. Scrape the day's CSV from Autovolt.
  2. Classify it and load it into the database (idempotent).
  3. Build the workbook + PDF from the database (month-to-date always correct).
  4. Email the PDF + XLSX (unless the workbook fails its balance check).

Usage:
    python run_daily.py --scheduled     # 6pm cron: only proceeds at 18:00 London
    python run_daily.py                 # report on today
    python run_daily.py 2026-07-02      # report on a specific date
    python run_daily.py --no-email 2026-07-02
"""
import sys
import datetime as dt
from pathlib import Path

from scrape_autovolt import scrape
from classify import process_csv
from build_report import build
from send_email import send_report
import db

HERE = Path(__file__).parent


def pick_report_date(today: dt.date) -> dt.date:
    """The report is for the SAME day the automation runs (data closes at 5:30pm)."""
    return today


def run(report_date: dt.date, send=True, to_me=False):
    print(f"=== Daily run: reporting on {report_date} ({report_date:%A})"
          f"{'' if send else '  [DRY RUN - no email]'} ===")

    # 1. scrape the day (runs every day — weekends included, even if zero)
    scrape(report_date.isoformat(), report_date.isoformat())
    csv_path = HERE / "SupplierTransaction.csv"

    # refresh the Autocheck fleet map (best-effort) so plate->area is current
    try:
        from fetch_autocheck_fleet import fetch as _fetch_fleet
        _fetch_fleet()
    except Exception as _e:
        print(f"autocheck fleet refresh skipped: {_e}")

    # 2. classify + load into the database (the builder reads it back)
    df = process_csv(csv_path, report_date)
    try:
        db.ingest(df, report_date)
    except Exception as e:
        print(f"WARN: database ingest failed: {e}")

    # 3. build workbook + PDF from the database
    out_xlsx, out_pdf, diff, daily_total, date_long = build(report_date)

    # 4. email unless the workbook didn't balance
    if diff >= 0.01:
        print(f"!! Balance gap £{diff:,.2f} — NOT emailing. Investigate first.")
        return 1
    if send:
        send_report(out_pdf, out_xlsx, date_long, f"Today's total spend: £{daily_total:,.2f}", to_me=to_me)
        # nightly line review — never blocks or fails the report itself
        try:
            from line_review import review, send_review_email
            from send_email import load_env
            _env = load_env()
            _n, _findings = review(report_date, _env)
            send_review_email(date_long, _n, _findings, _env)
        except Exception as _e:
            print(f"line review skipped: {_e}")
    else:
        print(f"[DRY RUN] would email — today's total £{daily_total:,.2f}")
    return 0


def main():
    argv = sys.argv[1:]
    scheduled = "--scheduled" in argv          # used by the Coolify cron
    send = "--no-email" not in argv
    to_me = "--to-me" in argv                  # full real send, but only to Daniel
    dates = [a for a in argv if not a.startswith("--")]

    if scheduled:
        # cron fires at 17:00 and 18:00 UTC to cover BST/GMT; only run at 18:00 London
        from zoneinfo import ZoneInfo
        now = dt.datetime.now(ZoneInfo("Europe/London"))
        if now.hour != 18:
            print(f"[guard] London time {now:%Y-%m-%d %H:%M} is not 18:00 — skipping this fire.")
            sys.exit(0)
        report_date = now.date()
    elif dates:
        report_date = dt.date.fromisoformat(dates[0])
    else:
        report_date = pick_report_date(dt.date.today())

    sys.exit(run(report_date, send=send, to_me=to_me))


if __name__ == "__main__":
    main()
