"""
Daily orchestrator — the single command the scheduler runs each morning.

  1. Work out which day to report on (the last working day).
  2. Scrape that day's CSV from Autovolt.
  3. Build the workbook + PDF (carrying month-to-date forward).
  4. Email the PDF + XLSX (unless the workbook fails its balance check).
  5. Save the workbook as the base for tomorrow's run.

Usage:
    python run_daily.py                # auto: report on the last working day
    python run_daily.py 2026-07-02     # force a specific report date
"""
import sys
import shutil
import datetime as dt
from pathlib import Path

from date_args import last_working_day, is_working_day, compute
from scrape_autovolt import scrape
from build_report import build
from send_email import send_report

HERE = Path(__file__).parent
STATE = HERE / "state" / "current_report.xlsx"   # previous day's workbook (tomorrow's base)
FALLBACK_BASE = HERE / "blank_template.xlsx"   # safe seed (no real figures) when no prior state


def pick_report_date(today: dt.date) -> dt.date:
    """The report is for the SAME day the automation runs (Mon-Fri at 6pm)."""
    return today


def run(report_date: dt.date, send=True):
    print(f"=== Daily run: reporting on {report_date} ({report_date:%A})"
          f"{'' if send else '  [DRY RUN - no email]'} ===")
    if not is_working_day(report_date):
        print(f"{report_date} is a weekend / bank holiday — skipping (no report).")
        return 0
    info = compute(report_date, fixed_wd=22)

    # 1+2. scrape that single day
    scrape(report_date.isoformat(), report_date.isoformat())
    csv_path = HERE / "SupplierTransaction.csv"

    # choose base: month-start builds fresh; otherwise use yesterday's workbook
    if info["is_new_month"]:
        base = FALLBACK_BASE
    else:
        base = STATE if STATE.exists() else FALLBACK_BASE
        print(f"Base workbook: {base}")

    # 3. build
    out_xlsx, out_pdf, diff, daily_total, date_long, df = build(csv_path, report_date, base)

    # 3b. write the classified rows into the Workshop database (feeds the tile + AI)
    try:
        import db
        db.ingest(df, report_date)
    except Exception as e:
        print(f"WARN: database ingest failed (report still emailed): {e}")

    # 4. email — but not if the workbook didn't balance
    if diff >= 0.01:
        print(f"!! Balance gap £{diff:,.2f} — NOT emailing. Investigate before sending.")
        return 1
    if send:
        send_report(out_pdf, out_xlsx, date_long, f"Today's total spend: £{daily_total:,.2f}")
    else:
        print(f"[DRY RUN] would email — today's total £{daily_total:,.2f}")

    # 5. persist as tomorrow's base
    STATE.parent.mkdir(exist_ok=True)
    shutil.copy(out_xlsx, STATE)
    print(f"Saved base for next run: {STATE}")
    return 0


def main():
    args = [a for a in sys.argv[1:] if a != "--no-email"]
    send = "--no-email" not in sys.argv[1:]
    if args:
        report_date = dt.date.fromisoformat(args[0])
    else:
        report_date = pick_report_date(dt.date.today())
    sys.exit(run(report_date, send=send))


if __name__ == "__main__":
    main()
