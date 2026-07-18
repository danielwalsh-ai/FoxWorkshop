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


def parts_category_split(report_date):
    """Parts-category breakdown for the 2025 & 2026-plate trucks.

    Answers Paul's ask: "categorise the spend on the trucks into part
    categories, interested in the 25 plate spend."  Unlike reg_year_split this
    is ALL divisions (incl. Capital) — the biggest chunk of new-truck fit-out
    (safety/camera/radar systems, livery) is booked to Capital, so restricting
    to the workshop top-sheets would hide exactly what he wants to see.

    Returns a dict:
      ltd / mtd  -> {2025: {category: £}, 2026: {category: £}}
      total_ltd / total_mtd -> {2025: £, 2026: £}
      trucks_ltd / trucks_mtd -> {2025: n, 2026: n}
    """
    from collections import defaultdict
    from classify import reg_year
    from parts_category import categorise
    first, _ = _bounds(report_date.year, report_date.month)
    with get_conn() as c, c.cursor() as cur:
        cur.execute("""SELECT report_date, vehicle_reg, part_name, cost
                       FROM transactions WHERE report_date <= %s""", (report_date,))
        rows = cur.fetchall()
    ltd = {2025: defaultdict(float), 2026: defaultdict(float)}
    mtd = {2025: defaultdict(float), 2026: defaultdict(float)}
    tr_ltd = {2025: set(), 2026: set()}
    tr_mtd = {2025: set(), 2026: set()}
    for rd, reg, part, cost in rows:
        reg = (reg or '').strip()
        if not reg:
            continue
        y = reg_year(reg)
        if y not in (2025, 2026):
            continue
        cost = float(cost or 0)
        cat = categorise(part)
        ltd[y][cat] += cost
        tr_ltd[y].add(reg)
        if rd >= first:
            mtd[y][cat] += cost
            tr_mtd[y].add(reg)
    return {
        'ltd': {y: dict(ltd[y]) for y in (2025, 2026)},
        'mtd': {y: dict(mtd[y]) for y in (2025, 2026)},
        'total_ltd': {y: round(sum(ltd[y].values()), 2) for y in (2025, 2026)},
        'total_mtd': {y: round(sum(mtd[y].values()), 2) for y in (2025, 2026)},
        'trucks_ltd': {y: len(tr_ltd[y]) for y in (2025, 2026)},
        'trucks_mtd': {y: len(tr_mtd[y]) for y in (2025, 2026)},
    }


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


def hook_fleet():
    """Hook registrations from vehicle_master.xlsx (AREA == HOOKS), file order."""
    from openpyxl import load_workbook
    from pathlib import Path
    wb = load_workbook(Path(__file__).parent / 'vehicle_master.xlsx', read_only=True)
    ws = wb.active
    return [str(r[0]).strip().upper().replace(' ', '')
            for r in ws.iter_rows(min_row=2, values_only=True)
            if r[0] and str(r[2] or '').strip().upper() == 'HOOKS']


def hook_split(report_date):
    """Per-registration hook spend for the month (area = HOOKS, top-sheet
    divisions only — mirrors exactly what feeds the HOOKS area row).
    Returns (fleet_regs, per_reg, unmatched) where
      per_reg   -> {reg: {report_date: £}}
      unmatched -> {report_date: £}   (area HOOKS but no reg recorded)"""
    from collections import defaultdict
    from classify import TOP_SHEETS
    TOP = {s.strip() for s in TOP_SHEETS}
    fleet = hook_fleet()
    first, _ = _bounds(report_date.year, report_date.month)
    with get_conn() as c, c.cursor() as cur:
        cur.execute("""SELECT report_date, division, vehicle_reg, cost FROM transactions
                       WHERE report_date >= %s AND report_date <= %s AND area = 'HOOKS'""",
                    (first, report_date))
        rows = cur.fetchall()
    per_reg = defaultdict(lambda: defaultdict(float))
    unmatched = defaultdict(float)
    for rd, division, reg, cost in rows:
        if (division or '').strip() not in TOP:
            continue
        cost = float(cost or 0)
        reg = (reg or '').strip().upper().replace(' ', '')
        if reg:
            per_reg[reg][rd] += cost
        else:
            unmatched[rd] += cost
    return fleet, {k: dict(v) for k, v in per_reg.items()}, dict(unmatched)


FN_GROUPS = {'J FISHER': ('J FISHER', 'J Fisher Plant'),
             'NMS': ('NMS CIVIL', 'NMS Plant')}
FN_BANDS = ['2021', '2022', '2023', '2024', '2025', '2026',
            'Older / private plates', 'Unregistered / Plant']


def fisher_nms_split(report_date):
    """Age-band x parts-category matrices for J Fisher (trucks + plant) and
    NMS (civil + plant) — PF request 16/07/2026, 'same way as pages 3-4'.
    Registered vehicles band by plate year (2021-2026, else Older/private);
    lines with no reg (plant kit, stock, consumables) -> 'Unregistered / Plant'
    per DW instruction 18/07 (no plant-age lookup).
    Returns {group: {'mtd': {cat: {band: £}}, 'today': {cat: {band: £}},
                     'mtd_total': £, 'today_total': £}}"""
    from collections import defaultdict
    from classify import reg_year, extract_reg
    from parts_category import categorise
    first, _ = _bounds(report_date.year, report_date.month)
    divs = tuple(d for pair in FN_GROUPS.values() for d in pair)
    with get_conn() as c, c.cursor() as cur:
        cur.execute("""SELECT report_date, division, vehicle_reg, part_name, cost
                       FROM transactions
                       WHERE report_date >= %s AND report_date <= %s
                         AND division IN %s""", (first, report_date, divs))
        rows = cur.fetchall()
    out = {g: {'mtd': defaultdict(lambda: defaultdict(float)),
               'today': defaultdict(lambda: defaultdict(float)),
               'mtd_total': 0.0, 'today_total': 0.0} for g in FN_GROUPS}
    div_to_group = {d.strip(): g for g, pair in FN_GROUPS.items() for d in pair}
    for rd, division, reg, part, cost in rows:
        g = div_to_group.get((division or '').strip())
        if not g:
            continue
        cost = float(cost or 0)
        reg = (reg or '').strip()
        if not reg:
            band = 'Unregistered / Plant'
        else:
            y = reg_year(reg)
            band = str(y) if y else 'Older / private plates'
        cat = categorise(part)
        out[g]['mtd'][cat][band] += cost
        out[g]['mtd_total'] += cost
        if rd == report_date:
            out[g]['today'][cat][band] += cost
            out[g]['today_total'] += cost
    for g in out:
        out[g]['mtd'] = {k: dict(v) for k, v in out[g]['mtd'].items()}
        out[g]['today'] = {k: dict(v) for k, v in out[g]['today'].items()}
        out[g]['mtd_total'] = round(out[g]['mtd_total'], 2)
        out[g]['today_total'] = round(out[g]['today_total'], 2)
    return out
