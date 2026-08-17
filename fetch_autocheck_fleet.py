"""Pull the latest Autocheck fleet report CSV from Gmail -> autocheck_fleet.csv.

The daily report uses this file to map vehicle plates to areas (see
autocheck_areas.py), so run_daily refreshes it each morning. Best-effort: if the
mailbox is unreachable it leaves the previous snapshot in place and the report
still runs.
"""
import imaplib
import email
import os
import sys
from pathlib import Path

HERE = Path(__file__).parent
OUT = HERE / "autocheck_fleet.csv"


def _env():
    e = dict(os.environ)
    envf = HERE / ".env"
    if envf.exists():
        for line in envf.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                e.setdefault(k, v)
    return e


def fetch(out_path: Path = OUT) -> Path | None:
    env = _env()
    user = env.get("GMAIL_USER")
    pw = env.get("GMAIL_APP_PASSWORD")
    if not user or not pw:
        print("fetch_autocheck_fleet: no Gmail creds, skipping")
        return None
    m = imaplib.IMAP4_SSL("imap.gmail.com")
    try:
        m.login(user, pw)
        m.select("INBOX")
        typ, data = m.search(None, '(FROM "noreply@auto-check.io" SUBJECT "autocheck vehicle report")')
        ids = data[0].split()
        if not ids:
            print("fetch_autocheck_fleet: no Autocheck email found")
            return None
        typ, msg_data = m.fetch(ids[-1], "(RFC822)")   # newest
        msg = email.message_from_bytes(msg_data[0][1])
        for part in msg.walk():
            fn = part.get_filename() or ""
            if fn.lower().endswith(".csv"):
                out_path.write_bytes(part.get_payload(decode=True))
                rows = max(0, len(out_path.read_text(encoding="utf-8", errors="replace").splitlines()) - 1)
                print(f"fetch_autocheck_fleet: saved {out_path.name} ({rows} vehicles)")
                return out_path
        print("fetch_autocheck_fleet: no CSV attachment on the latest email")
        return None
    finally:
        try:
            m.logout()
        except Exception:
            pass


if __name__ == "__main__":
    sys.exit(0 if fetch() else 1)
