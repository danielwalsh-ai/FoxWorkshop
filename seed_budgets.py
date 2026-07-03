"""Seed the known monthly budgets for every month that has transaction data.
(Budgets currently come from the report workbook's hidden column; adjust here
or in the budgets table as they change.)"""
import db

BUDGETS = {
    'Fox Wagons': 30000, 'Leyland Wagons': 30000, 'J FISHER': 35000,
    'NMS CIVIL': 5467, 'Tyres': 46125, 'J Fisher Plant': 14000, 'NMS Plant': 1770,
}

with db.get_conn() as conn, conn.cursor() as cur:
    cur.execute("SELECT DISTINCT date_trunc('month', report_date)::date FROM transactions ORDER BY 1")
    months = [r[0] for r in cur.fetchall()]

for mdate in months:
    db.upsert_budgets(BUDGETS, mdate.year, mdate.month)

print(f"Seeded budgets for {len(months)} months.")
