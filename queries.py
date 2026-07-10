"""Read queries for the Workshop & Maintenance dashboard + API."""
import datetime as dt
from db import get_conn

# Divisions that carry a monthly budget (others show spend only)
BUDGET_DIVISIONS = ['Fox Wagons', 'Leyland Wagons', 'J FISHER', 'NMS CIVIL',
                    'Tyres', 'J Fisher Plant', 'NMS Plant']


def _bounds(y, m):
    first = dt.date(y, m, 1)
    nm, ny = (1, y + 1) if m == 12 else (m + 1, y)
    return first, dt.date(ny, nm, 1)


def available_months():
    with get_conn() as c, c.cursor() as cur:
        cur.execute("SELECT DISTINCT to_char(report_date,'YYYY-MM') FROM transactions ORDER BY 1 DESC")
        return [r[0] for r in cur.fetchall()]


def overview(y, m):
    first, nxt = _bounds(y, m)
    with get_conn() as c, c.cursor() as cur:
        cur.execute("""SELECT division, ROUND(SUM(cost),2) FROM transactions
                       WHERE report_date >= %s AND report_date < %s
                       GROUP BY division ORDER BY 2 DESC""", (first, nxt))
        divs = cur.fetchall()
        cur.execute("""SELECT area, ROUND(SUM(cost),2) FROM transactions
                       WHERE report_date >= %s AND report_date < %s
                       GROUP BY area ORDER BY 2 DESC""", (first, nxt))
        areas = [{"area": a or 'UNIDENTIFIED', "total": float(t)} for a, t in cur.fetchall()]
        cur.execute("""SELECT report_date, ROUND(SUM(cost),2) FROM transactions
                       WHERE report_date >= %s AND report_date < %s
                       GROUP BY report_date ORDER BY report_date""", (first, nxt))
        daily = [{"date": d, "total": float(t)} for d, t in cur.fetchall()]
        cur.execute("SELECT division, budget FROM budgets WHERE year=%s AND month=%s", (y, m))
        budgets = {d: float(b) for d, b in cur.fetchall()}

    div_map = {d: float(t) for d, t in divs}
    total = round(sum(div_map.values()), 2)

    # budget tracker rows
    budget_rows = []
    for name in BUDGET_DIVISIONS:
        spent = div_map.get(name, 0.0)
        bud = budgets.get(name, 0.0)
        rem = bud - spent
        pct = (spent / bud * 100) if bud else 0
        budget_rows.append({"division": name, "budget": bud, "spent": round(spent, 2),
                            "remaining": round(rem, 2), "pct": round(pct, 1)})

    div_rows = [{"division": d, "total": float(t)} for d, t in divs]
    biggest_day = max(daily, key=lambda r: r["total"]) if daily else None

    return {
        "total": total,
        "damage": div_map.get("Damage", 0.0),
        "tyres": div_map.get("Tyres", 0.0),
        "capital": div_map.get("Capital", 0.0),
        "windscreen": div_map.get("Windscreen & Glass", 0.0),
        "div_rows": div_rows,
        "area_rows": areas,
        "budget_rows": budget_rows,
        "daily": daily,
        "days": len(daily),
        "biggest_day": biggest_day,
    }


def monthly_totals(n=18):
    with get_conn() as c, c.cursor() as cur:
        cur.execute("""SELECT to_char(report_date,'YYYY-MM') ym, ROUND(SUM(cost),2)
                       FROM transactions GROUP BY ym ORDER BY ym DESC LIMIT %s""", (n,))
        rows = cur.fetchall()[::-1]
    return [{"ym": ym, "total": float(t)} for ym, t in rows]


def month_rows(y, m):
    """Every transaction line of the month — for rebuilding the workbook cover."""
    first, nxt = _bounds(y, m)
    with get_conn() as c, c.cursor() as cur:
        cur.execute("""SELECT report_date, division, area, plate, supplier, cost, vehicle_reg
                       FROM transactions WHERE report_date >= %s AND report_date < %s""",
                    (first, nxt))
        return cur.fetchall()


def day_tab_rows(d):
    """A single day's transactions, for the division tabs in the workbook."""
    with get_conn() as c, c.cursor() as cur:
        cur.execute("""SELECT division, supplier, supplier_source_depot, system_no, supplier_pn,
                              part_name, cost, po_no, attached_order_no, attached_customer,
                              po_created_date, supply_type, item_count, goods_received,
                              target_depot, assigned_depot, supplier_ref, custom_ref, area
                       FROM transactions WHERE report_date = %s ORDER BY division""", (d,))
        cols = [dd[0] for dd in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]


def reg_year_split(report_date):
    """Vehicle spend by registration year — today + month-to-date.
    Returns (today_by_year, mtd_by_year) dicts keyed by year int."""
    from collections import defaultdict
    from classify import reg_year, TOP_SHEETS
    TOP = {s.strip() for s in TOP_SHEETS}
    first, _ = _bounds(report_date.year, report_date.month)
    with get_conn() as c, c.cursor() as cur:
        cur.execute("""SELECT report_date, division, vehicle_reg, cost FROM transactions
                       WHERE report_date >= %s AND report_date <= %s""", (first, report_date))
        rows = cur.fetchall()
    today, mtd = defaultdict(float), defaultdict(float)
    for rd, division, reg, cost in rows:
        if (division or '').strip() not in TOP:
            continue
        reg = (reg or '').strip()
        if not reg:
            continue
        key = reg_year(reg) or 'other'   # 2021-2026 int, else 'other' (older/private)
        cost = float(cost or 0)
        mtd[key] += cost
        if rd == report_date:
            today[key] += cost
    return dict(today), dict(mtd)


def recent_transactions(y, m, limit=60):
    first, nxt = _bounds(y, m)
    with get_conn() as c, c.cursor() as cur:
        cur.execute("""SELECT po_created_date, supplier, part_name, division, area,
                              vehicle_reg, cost, po_no
                       FROM transactions
                       WHERE report_date >= %s AND report_date < %s
                       ORDER BY cost DESC LIMIT %s""", (first, nxt, limit))
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in cur.fetchall()]
