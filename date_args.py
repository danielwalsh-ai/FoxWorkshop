"""
Work out the report's date arguments automatically from a calendar, so nobody
has to type them each day.

Working day = Monday-Friday, excluding England & Wales bank holidays.
Cover sheet layout: the 1st working day of the month is column 2, so
    today_col = (working days from 1st of month up to & incl. report date) + 1
"""
import datetime as dt

# England & Wales bank holidays (extend yearly as needed)
BANK_HOLIDAYS = {
    # 2026
    dt.date(2026, 1, 1), dt.date(2026, 4, 3), dt.date(2026, 4, 6),
    dt.date(2026, 5, 4), dt.date(2026, 5, 25), dt.date(2026, 8, 31),
    dt.date(2026, 12, 25), dt.date(2026, 12, 28),
    # 2027
    dt.date(2027, 1, 1), dt.date(2027, 3, 26), dt.date(2027, 3, 29),
    dt.date(2027, 5, 3), dt.date(2027, 5, 31), dt.date(2027, 8, 30),
    dt.date(2027, 12, 27), dt.date(2027, 12, 28),
}


def is_working_day(d: dt.date) -> bool:
    return d.weekday() < 5 and d not in BANK_HOLIDAYS


def last_working_day(on_or_before: dt.date) -> dt.date:
    """The most recent working day on or before the given date."""
    d = on_or_before
    while not is_working_day(d):
        d -= dt.timedelta(days=1)
    return d


def working_days_in_month(year: int, month: int) -> int:
    d = dt.date(year, month, 1)
    count = 0
    while d.month == month:
        if is_working_day(d):
            count += 1
        d += dt.timedelta(days=1)
    return count


def working_day_index(report_date: dt.date) -> int:
    """How many working days from the 1st of the month up to & including report_date."""
    d = dt.date(report_date.year, report_date.month, 1)
    count = 0
    while d <= report_date:
        if is_working_day(d):
            count += 1
        d += dt.timedelta(days=1)
    return count


def compute(report_date: dt.date, fixed_wd: int | None = None) -> dict:
    idx = working_day_index(report_date)
    wd_total = fixed_wd if fixed_wd else working_days_in_month(report_date.year, report_date.month)
    return {
        "report_date": report_date.isoformat(),
        "today_col": idx + 1,          # col 2 = 1st working day
        "days_elapsed": idx,
        "days_remaining": wd_total - idx,
        "working_days_in_month": wd_total,
        "is_new_month": idx == 1,      # report date is the 1st working day
        "is_working_day": is_working_day(report_date),
    }


if __name__ == "__main__":
    tests = [
        dt.date(2026, 7, 1),   # 1st working day (Wed) -> new month, col 2
        dt.date(2026, 7, 2),   # 2nd working day (Thu) -> col 3, matches known example
        dt.date(2026, 7, 3),   # 3rd working day (Fri) -> col 4
        dt.date(2026, 7, 6),   # Monday
        dt.date(2026, 7, 31),  # month end
    ]
    for d in tests:
        print(d, d.strftime("%a"), "->", compute(d))
    print()
    print("Last working day on/before Mon 6 Jul 2026:", last_working_day(dt.date(2026, 7, 6)))
    print("Working days in July 2026:", working_days_in_month(2026, 7))
    print("Working days in Aug 2026 (has 31 Aug bank hol):", working_days_in_month(2026, 8))
