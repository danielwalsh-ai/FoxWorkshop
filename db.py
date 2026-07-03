"""
Workshop database — schema + ingestion.

Tables:
  transactions  one row per purchase-order line, per report day, with our
                division/area/plate classification added.
  budgets       monthly budget per division.

The daily automation calls ingest() after building the report; the web
dashboard and the hub's AI read these tables.
"""
import os
import sys
import datetime as dt
from pathlib import Path

import pandas as pd
import psycopg2
import psycopg2.extras

HERE = Path(__file__).parent


def _load_env():
    """Real environment variables (production/Coolify) win; a local .env file
    fills in the gaps for development."""
    env = dict(os.environ)
    p = HERE / ".env"
    if p.exists():
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def get_conn():
    return psycopg2.connect(_load_env()["WORKSHOP_DATABASE_URL"])


SCHEMA = """
CREATE TABLE IF NOT EXISTS transactions (
  id BIGSERIAL PRIMARY KEY,
  report_date DATE NOT NULL,
  supplier TEXT,
  supplier_source_depot TEXT,
  system_no TEXT,
  supplier_pn TEXT,
  part_name TEXT,
  cost NUMERIC(12,2),
  surcharge NUMERIC(12,2),
  po_no TEXT,
  attached_order_no TEXT,
  attached_customer TEXT,
  po_created_date TIMESTAMP,
  supply_type TEXT,
  item_count INTEGER,
  goods_received TEXT,
  target_depot TEXT,
  assigned_depot TEXT,
  supplier_ref TEXT,
  custom_ref TEXT,
  division TEXT,
  area TEXT,
  plate TEXT,
  vehicle_reg TEXT,
  ingested_at TIMESTAMP DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS idx_tx_report_date ON transactions(report_date);
CREATE INDEX IF NOT EXISTS idx_tx_division    ON transactions(division);
CREATE INDEX IF NOT EXISTS idx_tx_area        ON transactions(area);

CREATE TABLE IF NOT EXISTS budgets (
  division TEXT NOT NULL,
  year INT NOT NULL,
  month INT NOT NULL,
  budget NUMERIC(12,2) NOT NULL,
  PRIMARY KEY (division, year, month)
);
"""

TX_COLS = [
    "report_date", "supplier", "supplier_source_depot", "system_no", "supplier_pn",
    "part_name", "cost", "surcharge", "po_no", "attached_order_no", "attached_customer",
    "po_created_date", "supply_type", "item_count", "goods_received", "target_depot",
    "assigned_depot", "supplier_ref", "custom_ref", "division", "area", "plate", "vehicle_reg",
]


def init_schema():
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute(SCHEMA)
    print("Schema ready (transactions, budgets).")


def _s(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s or None


def _num(v):
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def _int(v):
    n = _num(v)
    return int(n) if n is not None else None


def _row_tuple(r, report_date):
    from classify import extract_reg
    reg = extract_reg(str(r.get("Custom Ref", ""))) or extract_reg(str(r.get("Supplier Ref", "")))
    return (
        report_date, _s(r.get("Supplier")), _s(r.get("Supplier Source Depot")),
        _s(r.get("365 No")), _s(r.get("Supplier PN")), _s(r.get("Part Name")),
        _num(r.get("Cost")), _num(r.get("Surcharge")), _s(r.get("PO No")),
        _s(r.get("Attached Order No")), _s(r.get("Attached Customer")),
        _s(r.get("PO Created Date")), _s(r.get("Supplier/ Collection?")),
        _int(r.get("Item Count")), _s(r.get("Goods Received")), _s(r.get("Target Depot")),
        _s(r.get("Assigned Depot")), _s(r.get("Supplier Ref")), _s(r.get("Custom Ref")),
        _s(r.get("Sheet")), _s(r.get("Area")), _s(r.get("Plate")), reg,
    )


def _insert(cur, rows):
    if rows:
        psycopg2.extras.execute_values(
            cur, f"INSERT INTO transactions ({','.join(TX_COLS)}) VALUES %s", rows, page_size=500)


def ingest(df, report_date: dt.date):
    """Idempotent: replace all rows for report_date with the given dataframe."""
    rows = [_row_tuple(r, report_date) for _, r in df.iterrows()]
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM transactions WHERE report_date = %s", (report_date,))
        _insert(cur, rows)
    print(f"Ingested {len(rows)} rows for {report_date}")
    return len(rows)


def ingest_bulk(df):
    """Backfill: df carries a per-row 'ReportDate'. Replaces the covered date span."""
    rows = [_row_tuple(r, r["ReportDate"]) for _, r in df.iterrows()]
    dmin, dmax = df["ReportDate"].min(), df["ReportDate"].max()
    with get_conn() as conn, conn.cursor() as cur:
        cur.execute("DELETE FROM transactions WHERE report_date BETWEEN %s AND %s", (dmin, dmax))
        _insert(cur, rows)
    print(f"Bulk-ingested {len(rows)} rows across {dmin} .. {dmax}")
    return len(rows)


def upsert_budgets(budget_map, year, month):
    """budget_map: {division: amount}."""
    with get_conn() as conn, conn.cursor() as cur:
        for division, amount in budget_map.items():
            cur.execute(
                """INSERT INTO budgets (division, year, month, budget)
                   VALUES (%s,%s,%s,%s)
                   ON CONFLICT (division, year, month) DO UPDATE SET budget = EXCLUDED.budget""",
                (division, year, month, amount),
            )
    print(f"Upserted {len(budget_map)} budgets for {year}-{month:02d}")


def _cli():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "init"
    if cmd == "init":
        init_schema()
    elif cmd == "load":
        from classify import process_csv
        csv, date_s = sys.argv[2], sys.argv[3]
        d = dt.date.fromisoformat(date_s)
        df = process_csv(csv, d)
        ingest(df, d)
    elif cmd == "totals":
        date_s = sys.argv[2]
        with get_conn() as conn, conn.cursor() as cur:
            cur.execute(
                """SELECT division, ROUND(SUM(cost),2) FROM transactions
                   WHERE report_date=%s GROUP BY division ORDER BY 2 DESC""",
                (dt.date.fromisoformat(date_s),),
            )
            total = 0.0
            for div, amt in cur.fetchall():
                print(f"  {div:<20} £{amt:,.2f}")
                total += float(amt)
            print(f"  {'TOTAL':<20} £{total:,.2f}")
    else:
        print("Usage: python db.py [init | load <csv> <date> | totals <date>]")


if __name__ == "__main__":
    _cli()
