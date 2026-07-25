"""
Wagon master automation — Katie's run sheets in, updated master out to Paul.

Watches Daniel's Gmail for Katie Ward's "Wagon earnings ..." emails, fills every
run sheet attached (she batches days and re-sends corrections), pulls parts/tyres
from the workshop transaction report, tidies the formatting, then emails the
master to Paul Fox with Daniel copied in.

Design notes:
  * A run sheet's days come from C2/D2, not P2/Q2 — a weekend sheet carries
    Saturday in col C and Sunday in col D, and its P2 is empty.
  * Katie re-sends corrections ("Please use this one"). Within a batch the newest
    email wins; a day already in the master is replaced and the change reported.
  * Parts/tyres come from the workshop report's Cover tab, which is month-to-date,
    so one file covers any day of that month. No database needed.
  * Nothing reaches Paul unless every day passes validation. If anything looks
    wrong the run emails Daniel alone and leaves the stored master untouched.

Usage:
    python wagon_auto.py --seed "Daily wagon earnings 23rd July.xlsx"   # first time only
    python wagon_auto.py --once --dry-run    # check + build, send nothing
    python wagon_auto.py --once              # run now
    python wagon_auto.py --scheduled         # cron: quiet when there's nothing new
"""
import os
import re
import ssl
import sys
import json
import email
import shutil
import smtplib
import imaplib
import argparse
import zipfile
import tempfile
import datetime as dt
import urllib.request
from pathlib import Path
from email.message import EmailMessage

import openpyxl

from openpyxl.utils import get_column_letter

from wagon_master_fill import (fill_master, master_state, read_transaction_report,
                               SheetXmlEditor, VEHICLE_ROWS)
import tidy_master

HERE = Path(__file__).parent
STATE_DIR = HERE / "state"
STATE_FILE = STATE_DIR / "wagon_auto.json"
MASTER_FILE = STATE_DIR / "wagon_master.xlsx"

KATIE = "katie.ward@hurtplant.co.uk"
SEND_TO = ["paulfox@foxbrothers.co.uk"]
SEND_CC = ["daniel.walsh@kfltd.uk"]
MODEL = "claude-sonnet-5"

ROW_TOTAL_EARNINGS, ROW_PARTS, ROW_WORKSHOP, ROW_TYRES = 168, 176, 177, 178
MAX_PLAUSIBLE_DAY = 250_000      # a decimal slip would sail past this
MIN_PLAUSIBLE_WEEKDAY = 20_000   # weekdays have run £75k–£97k all month
DATE_WINDOW_DAYS = 90            # a mistyped year must not land in last year's column


def load_env():
    env = dict(os.environ)
    p = HERE / ".env"
    if p.exists():
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


ENV = load_env()


def ordinal(n):
    suffix = "th" if 11 <= n % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def master_filename(last_date):
    return f"Daily wagon earnings {ordinal(last_date.day)} {last_date:%B}.xlsx"


# ── state ───────────────────────────────────────────────────────────
def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    return {"processed_message_ids": [], "last_run": None}


def save_state(state):
    STATE_DIR.mkdir(exist_ok=True)
    state["last_run"] = dt.datetime.now().isoformat(timespec="seconds")
    STATE_FILE.write_text(json.dumps(state, indent=2), encoding="utf-8")


# ── Gmail ───────────────────────────────────────────────────────────
def _attachments(msg, tmpdir, want=".xlsx"):
    out = []
    for part in msg.walk():
        name = part.get_filename()
        if not name or not name.lower().endswith(want):
            continue
        if name.startswith("~$"):          # Excel lock file
            continue
        p = Path(tmpdir) / name
        p.write_bytes(part.get_payload(decode=True))
        out.append(p)
    return out


def fetch_katie_emails(tmpdir, since_days=21):
    """Newest-last list of {id, subject, date, files} from Katie's run-sheet emails."""
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(ENV["GMAIL_USER"], ENV["GMAIL_APP_PASSWORD"])
    M.select("INBOX")
    since = (dt.date.today() - dt.timedelta(days=since_days)).strftime("%d-%b-%Y")
    typ, data = M.search(None, f'(SINCE "{since}" FROM "{KATIE}" SUBJECT "Wagon earnings")')
    found = []
    for uid in data[0].split():
        typ, md = M.fetch(uid, "(RFC822)")
        msg = email.message_from_bytes(md[0][1])
        d = email.utils.parsedate_to_datetime(msg.get("Date"))
        sub = str(email.header.make_header(email.header.decode_header(msg.get("Subject", ""))))
        box = Path(tmpdir) / (msg.get("Message-ID", uid.decode()).strip("<>").replace("/", "_")[:60])
        box.mkdir(parents=True, exist_ok=True)
        files = _attachments(msg, box)
        if files:
            found.append({"id": msg.get("Message-ID", uid.decode()),
                          "subject": sub, "date": d, "files": files})
    M.logout()
    found.sort(key=lambda m: m["date"])       # oldest first — newest correction wins
    return found


def report_date_of(name):
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})", name)
    if not m:
        return None
    try:
        return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def _covers(report_day, target):
    """The Cover tab is month-to-date as of the report's own day, so a report only
    holds `target` if it is from the same month and dated on or after it."""
    return (report_day is not None and report_day.year == target.year
            and report_day.month == target.month and report_day >= target)


def fetch_transaction_report(tmpdir, target):
    """Newest workshop report that actually covers `target`, from Gmail. None if absent."""
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(ENV["GMAIL_USER"], ENV["GMAIL_APP_PASSWORD"])
    M.select('"[Gmail]/All Mail"')
    since = target.strftime("%d-%b-%Y")
    typ, data = M.search(None, f'(SINCE "{since}" SUBJECT "Daily Workshop Spend Report")')
    for uid in reversed(data[0].split()):     # newest first
        typ, md = M.fetch(uid, "(RFC822)")
        msg = email.message_from_bytes(md[0][1])
        box = Path(tmpdir) / f"tx_{uid.decode()}"
        box.mkdir(parents=True, exist_ok=True)
        for f in _attachments(msg, box):
            if _covers(report_date_of(f.name), target):
                M.logout()
                return f
    M.logout()
    return None


def fetch_master_from_gmail(tmpdir, since_days=30):
    """Newest 'Daily wagon earnings *.xlsx' attachment in the mailbox.

    Saves hauling a 1.6MB workbook onto the server by hand — email it to yourself
    (or let a previous run's email to Paul serve) and seed straight from Gmail."""
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(ENV["GMAIL_USER"], ENV["GMAIL_APP_PASSWORD"])
    M.select('"[Gmail]/All Mail"')
    since = (dt.date.today() - dt.timedelta(days=since_days)).strftime("%Y/%m/%d")
    typ, data = M.search(None, "X-GM-RAW",
                         f'"has:attachment filename:xlsx after:{since}"')
    uids = data[0].split()
    for uid in reversed(uids):                       # newest first
        typ, md = M.fetch(uid, "(RFC822)")
        msg = email.message_from_bytes(md[0][1])
        box = Path(tmpdir) / f"seed_{uid.decode()}"
        box.mkdir(parents=True, exist_ok=True)
        for f in _attachments(msg, box):
            if f.name.lower().startswith("daily wagon earnings"):
                M.logout()
                return f, msg.get("Subject", ""), email.utils.parsedate_to_datetime(
                    msg.get("Date"))
    M.logout()
    return None, None, None


def ensure_master(tmpdir, state):
    """Make sure there's a master to work from, rebuilding from Gmail if not.

    The container has no persistent volume, so a redeploy wipes state/. That's fine:
    every run emails the master to Paul, so Gmail already holds the latest copy.
    Recover it, and treat everything Katie sent before that email as already done —
    which reconstructs exactly the state the wipe destroyed."""
    if MASTER_FILE.exists():
        return True
    found, subj, when = fetch_master_from_gmail(tmpdir)
    if not found:
        return False
    STATE_DIR.mkdir(exist_ok=True)
    shutil.copy(found, MASTER_FILE)
    state["watermark"] = when.isoformat()
    # Persist immediately. If this run finds nothing new it returns early, and a later
    # run would otherwise see the master present, skip recovery, and — with no
    # watermark saved — reprocess every email Katie has sent.
    save_state(state)
    print(f"No stored master — recovered {found.name} from Gmail "
          f"(sent {when:%Y-%m-%d %H:%M}).")
    print("  Katie's earlier emails treated as already done.")
    return True


def transaction_report_for(date, tmpdir, _cache={}):
    """Prefer a locally-built report that covers the day; otherwise fetch from Gmail."""
    if date in _cache:
        return _cache[date]
    local = [(report_date_of(f.name), f) for f in HERE.glob("fox_transaction_report_*.xlsx")]
    local = [(d, f) for d, f in local if _covers(d, date)]
    if local:
        _cache[date] = max(local)[1]
        return _cache[date]
    try:
        _cache[date] = fetch_transaction_report(tmpdir, date)
    except Exception as e:
        print(f"  ! could not fetch a transaction report for {date}: {e}")
        _cache[date] = None
    return _cache[date]


# ── run sheets ──────────────────────────────────────────────────────
def days_in_sheet(path):
    """[(date, value_col)] — C2/D2 carry the day(s); a weekend sheet has both."""
    ws = openpyxl.load_workbook(path, data_only=True)["Wagons"]
    days = []
    for col, ref in ((3, "C2"), (4, "D2")):
        v = ws[ref].value
        if isinstance(v, dt.datetime):
            days.append((v.date(), col))
    if not days:                                    # older sheets: fall back to P2/Q2
        for ref in ("P2", "Q2"):
            v = ws[ref].value
            if isinstance(v, dt.datetime):
                days.append((v.date(), 3))
                break
    return days


def filename_date(name):
    """Katie names sheets DD.MM.YYYY — a useful cross-check on the in-sheet date."""
    m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", name)
    if not m:
        return None
    try:
        return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def collect_days(messages):
    """One entry per date across every attachment; a later email supersedes an earlier.

    Dates outside a sane window are set aside rather than filled — a mistyped year
    would otherwise land in a column from a previous year."""
    floor = dt.date.today() - dt.timedelta(days=DATE_WINDOW_DAYS)
    today = dt.date.today()
    by_date, rejected = {}, []
    for msg in messages:                            # already oldest-first
        for f in msg["files"]:
            try:
                found = days_in_sheet(f)
            except Exception as e:
                rejected.append(f"{f.name}: could not be read ({e})")
                continue
            if not found:
                rejected.append(f"{f.name}: no date found in C2/D2 or P2/Q2")
            for d, col in found:
                if d < floor or d > today:
                    fn = filename_date(f.name)
                    hint = (f" — the file is named {fn:%d.%m.%Y}, so this looks like a "
                            f"typo in the sheet's date cell") if fn and fn != d else ""
                    rejected.append(
                        f"{f.name}: sheet date reads {d:%d %b %Y}, which is outside the "
                        f"last {DATE_WINDOW_DAYS} days{hint}. Day skipped.")
                    continue
                by_date[d] = {"date": d, "file": f, "value_col": col,
                              "subject": msg["subject"]}
    return [by_date[d] for d in sorted(by_date)], rejected


# ── costs catch-up ──────────────────────────────────────────────────
def topup_costs(master, tmpdir):
    """Fill parts/tyres for any recent weekday that has earnings but no costs.

    Katie's run sheets often arrive before that evening's workshop report, so a day
    can land with earnings only. This picks them up once the report exists, which
    keeps the master converging without anyone having to notice."""
    sheet_path, date_cols, _, _, _, ws, wsv = master_state(str(master))
    floor = dt.date.today() - dt.timedelta(days=DATE_WINDOW_DAYS)
    todo = []
    for d, c in sorted(date_cols.items()):
        if d < floor or d > dt.date.today() or d.weekday() >= 5:
            continue                                   # weekend cost is averaged into Mon–Fri
        f = ws.cell(ROW_TOTAL_EARNINGS, c).value
        if not (isinstance(f, str) and f.startswith("=")):
            continue                                   # no earnings there yet
        v = ws.cell(ROW_PARTS, c).value
        if isinstance(v, (int, float)) and v:
            continue                                   # already carries a real figure
        todo.append((d, c))
    donor = max((c for c in date_cols.values()
                 if isinstance(ws.cell(ROW_PARTS, c).value, (int, float))
                 and ws.cell(ROW_PARTS, c).value), default=None)
    if not todo or donor is None:
        return None, []

    z = zipfile.ZipFile(str(master))
    ed = SheetXmlEditor(z.read(sheet_path))
    donor_L = get_column_letter(donor)
    filled = []
    for d, c in todo:
        tx = transaction_report_for(d, tmpdir)
        if not tx:
            continue
        try:
            parts, workshop, tyres = read_transaction_report(str(tx), d)
        except Exception as e:
            print(f"  ! costs for {d}: {e}")
            continue
        if not parts and not tyres:
            continue
        L = get_column_letter(c)
        for row, val in ((ROW_PARTS, parts), (ROW_WORKSHOP, workshop), (ROW_TYRES, tyres)):
            ed.write(row, L, value=round(float(val), 2), style=ed.style_of(row, donor_L))
        filled.append(d)
    if not filled:
        z.close()
        return None, []
    out = Path(tmpdir) / "topped_up.xlsx"
    data = ed.tobytes()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist():
            zo.writestr(it, data if it.filename == sheet_path else z.read(it.filename))
    z.close()
    return out, filled


# ── build ───────────────────────────────────────────────────────────
def build(master, days, tmpdir):
    results, current = [], master
    for i, day in enumerate(days):
        tx = transaction_report_for(day["date"], tmpdir)
        out = Path(tmpdir) / f"step_{i:02d}.xlsx"
        r = fill_master(str(current), str(day["file"]), str(tx) if tx else None, str(out),
                        date_override=day["date"], value_col=day["value_col"], replace=True)
        r["subject"] = day["subject"]
        r["had_costs"] = tx is not None
        results.append(r)
        current = out
        flag = "revised" if r["replaced"] else "new"
        print(f"  {day['date']} ({day['date']:%a}) col {r['column']}  "
              f"£{r['expected_total_earnings']:,.2f}  {flag}")
    final = Path(tmpdir) / "final.xlsx"
    tidy_master.tidy(str(current), str(final))
    return final, results


def validate(results):
    problems = []
    today = dt.date.today()
    for r in results:
        d = dt.date.fromisoformat(r["date"])
        t = r["expected_total_earnings"]
        if d > today:
            problems.append(f"{d} is in the future")
        if not r["no_wagons"] or r["no_wagons"] < 50:
            problems.append(f"{d}: wagon count reads {r['no_wagons']!r}, expected ~109")
        if t > MAX_PLAUSIBLE_DAY:
            problems.append(f"{d}: £{t:,.0f} is implausibly high — check for a stray decimal")
        if d.weekday() < 5 and 0 < t < MIN_PLAUSIBLE_WEEKDAY:
            problems.append(f"{d}: weekday total is only £{t:,.0f}, expected £75k–£95k")
    return problems


def month_context(path, upto):
    """Facts for the commentary, computed here so the model never invents a number."""
    ws = openpyxl.load_workbook(path)["DAILY"]
    wv = openpyxl.load_workbook(path, data_only=True)["DAILY"]
    cols = {wv.cell(2, c).value.date(): c for c in range(3, wv.max_column + 2)
            if isinstance(wv.cell(2, c).value, dt.datetime)}
    days = []
    for d, c in sorted(cols.items()):
        if d.month != upto.month or d.year != upto.year or d > upto:
            continue
        f = ws.cell(ROW_TOTAL_EARNINGS, c).value
        if not (isinstance(f, str) and f.startswith("=")):
            continue
        total = sum(float(wv.cell(r, c).value) for r in VEHICLE_ROWS
                    if isinstance(wv.cell(r, c).value, (int, float)))
        days.append((d, round(total, 2)))
    weekdays = [(d, t) for d, t in days if d.weekday() < 5]
    if not weekdays:
        return {}
    first5 = weekdays[:5]
    last5 = weekdays[-5:]
    return {
        "weekday_count": len(weekdays),
        "weekday_avg": round(sum(t for _, t in weekdays) / len(weekdays)),
        "best": max(weekdays, key=lambda x: x[1]),
        "worst": min(weekdays, key=lambda x: x[1]),
        "avg_first5": round(sum(t for _, t in first5) / len(first5)),
        "avg_last5": round(sum(t for _, t in last5) / len(last5)),
        "month_total": round(sum(t for _, t in days)),
    }


# ── commentary ──────────────────────────────────────────────────────
VOICE = """You write as Daniel Walsh, who consults for Fox Group and maintains the
wagon earnings master. You are writing 2-4 short sentences to Paul Fox, the CEO.

How Daniel writes to Paul (real examples):
  "All in the sheet up to Wednesday 8th. Tuesday was a monster. £97k total, best day
   on the sheet, and the 7 day average is now running at £84k against £76k at the
   start of the month. PN72EFE has been VOR all week."
  "Updated master."
  "Tidy this and added older than 2021 spend and hook spend by registration at the bottom."

Rules:
  - Blunt, Northern, factual. No greeting, no sign-off, no "please find attached".
  - Never invent a number. Use only the figures given to you.
  - Round to the nearest £k in prose (£85k, not £85,446.25).
  - Lead with what changed, then anything genuinely worth flagging.
  - If a day was revised, say so plainly.
  - No corporate filler. Never say "I hope this finds you well" or "Let me know if".
"""


def commentary(results, ctx):
    if not ENV.get("ANTHROPIC_API_KEY"):
        return ""
    lines = []
    for r in results:
        d = dt.date.fromisoformat(r["date"])
        bit = f"{d:%A} {d.day} {d:%B}: £{r['expected_total_earnings']:,.0f}"
        if r["replaced"] and r["previous_total"] is not None \
                and abs(r["previous_total"] - r["expected_total_earnings"]) >= 1:
            bit += f" (revised, was £{r['previous_total']:,.0f})"
        elif r["replaced"]:
            bit += " (re-sent, unchanged)"
        if not r["had_costs"]:
            bit += " [no parts/tyres available yet]"
        lines.append(bit)
    facts = "Days just added:\n" + "\n".join(f"  {l}" for l in lines)
    if ctx:
        facts += (f"\n\nMonth to date: {ctx['weekday_count']} weekdays, average "
                  f"£{ctx['weekday_avg']:,}. Best £{ctx['best'][1]:,.0f} on "
                  f"{ctx['best'][0]:%A} {ctx['best'][0].day}. Weakest "
                  f"£{ctx['worst'][1]:,.0f} on {ctx['worst'][0]:%A} {ctx['worst'][0].day}. "
                  f"Average of the first five weekdays £{ctx['avg_first5']:,}, "
                  f"of the last five £{ctx['avg_last5']:,}.")
    body = {"model": MODEL, "max_tokens": 400, "system": VOICE,
            "messages": [{"role": "user", "content": facts}]}
    try:
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=json.dumps(body).encode(),
            headers={"x-api-key": ENV["ANTHROPIC_API_KEY"],
                     "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            method="POST")
        with urllib.request.urlopen(req, timeout=60) as r:
            data = json.loads(r.read().decode())
        return "".join(b.get("text", "") for b in data.get("content", [])).strip()
    except Exception as e:
        print(f"  ! commentary unavailable ({e}) — sending the table alone")
        return ""


# ── email ───────────────────────────────────────────────────────────
def rows_html(results):
    out = []
    for r in results:
        d = dt.date.fromisoformat(r["date"])
        note = ""
        if r["replaced"] and r["previous_total"] is not None \
                and abs(r["previous_total"] - r["expected_total_earnings"]) >= 1:
            note = f' <span style="color:#b45309">revised from £{r["previous_total"]:,.2f}</span>'
        cost = (f'<td align="right">£{r["parts"]:,.2f}</td>'
                f'<td align="right">£{r["tyres"]:,.2f}</td>') if r["had_costs"] \
            else '<td align="right">—</td><td align="right">—</td>'
        out.append(
            f'<tr><td>{d:%a} {d.day} {d:%b}{note}</td>'
            f'<td align="right">£{r["expected_total_earnings"]:,.2f}</td>{cost}</tr>')
    return "\n".join(out)


def send(final_path, results, note, dry_run=False):
    last = max(dt.date.fromisoformat(r["date"]) for r in results)
    fname = master_filename(last)
    subject = f"Wagon earnings — master updated to {ordinal(last.day)} {last:%B}"
    html = f"""
    <div style="font-family:Arial,Helvetica,sans-serif;color:#24214a;font-size:15px">
      {'<p>' + note.replace(chr(10) + chr(10), '</p><p>') + '</p>' if note else ''}
      <table cellpadding="6" style="border-collapse:collapse;font-size:14px">
        <tr style="background:#24214a;color:#fff">
          <th align="left">Day</th><th align="right">Earnings</th>
          <th align="right">Parts</th><th align="right">Tyres</th>
        </tr>
        {rows_html(results)}
      </table>
      <p style="font-size:13px;color:#666">Master attached, updated to
         {ordinal(last.day)} {last:%B}.</p>
      <hr style="border:none;border-top:1px solid #ccc">
      <p style="font-size:12px;color:#888">Updated automatically from Katie's run sheets
         by danielwalsh.ai</p>
    </div>"""
    plain = (note + "\n\n" if note else "") + "\n".join(
        f"{dt.date.fromisoformat(r['date']):%a} {dt.date.fromisoformat(r['date']).day} "
        f"{dt.date.fromisoformat(r['date']):%b}: "
        f"£{r['expected_total_earnings']:,.2f}" for r in results)

    m = EmailMessage()
    m["From"] = f"Daniel Walsh <{ENV['GMAIL_USER']}>"
    m["To"] = ", ".join(SEND_TO)
    m["Cc"] = ", ".join(SEND_CC)
    m["Subject"] = subject
    m.set_content(plain)
    m.add_alternative(html, subtype="html")
    m.add_attachment(Path(final_path).read_bytes(), maintype="application",
                     subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     filename=fname)
    if dry_run:
        print(f"\n[DRY RUN] would send to {', '.join(SEND_TO)} cc {', '.join(SEND_CC)}")
        print(f"[DRY RUN] subject: {subject}")
        print(f"[DRY RUN] attachment: {fname}")
        print(f"\n{plain}\n")
        return fname
    ctx = ssl.create_default_context()
    with smtplib.SMTP("smtp.gmail.com", 587) as s:
        s.starttls(context=ctx)
        s.login(ENV["GMAIL_USER"], ENV["GMAIL_APP_PASSWORD"])
        s.send_message(m)
    print(f"  sent to {', '.join(SEND_TO)} (cc {', '.join(SEND_CC)}) — {fname}")
    return fname


def notify_daniel(subject, body, dry_run=False):
    """Anything Daniel needs to see but Paul doesn't."""
    if dry_run:
        print(f"\n[DRY RUN] would email Daniel — {subject}\n{body}\n")
        return
    m = EmailMessage()
    m["From"] = f"Daniel Walsh <{ENV['GMAIL_USER']}>"
    m["To"] = ENV["GMAIL_USER"]
    m["Subject"] = subject
    m.set_content(body)
    ctx = ssl.create_default_context()
    with smtplib.SMTP("smtp.gmail.com", 587) as s:
        s.starttls(context=ctx)
        s.login(ENV["GMAIL_USER"], ENV["GMAIL_APP_PASSWORD"])
        s.send_message(m)
    print(f"  emailed Daniel — {subject}")


def alert_daniel(problems, results, dry_run=False):
    """Something looks wrong — tell Daniel only, send Paul nothing."""
    lines = "\n".join(f"  - {p}" for p in problems)
    body = ("The wagon master was built but NOT sent to Paul — these checks failed:\n\n"
            f"{lines}\n\nThe stored master is unchanged and these emails will be "
            "picked up again on the next run once the run sheets are corrected.\n\n"
            "Days in this batch:\n"
            + "\n".join(f"  {r['date']}: £{r['expected_total_earnings']:,.2f} "
                        f"({r['no_wagons']} wagons)" for r in results))
    notify_daniel("Wagon master — held back, needs a look", body, dry_run)


# ── main ────────────────────────────────────────────────────────────
def run(dry_run=False, since_days=21, out_dir=None):
    state = load_state()
    with tempfile.TemporaryDirectory() as tmp:
        if not ensure_master(tmp, state):
            print("No stored master and none found in Gmail. Seed it first:\n"
                  '  python wagon_auto.py --seed "<path to current master>"')
            return 2
        # Costs first: days filled before that evening's workshop report catch up here.
        master_in = MASTER_FILE
        topped, topped_days = topup_costs(MASTER_FILE, tmp)
        if topped:
            print("costs caught up for: "
                  + ", ".join(f"{d:%d %b}" for d in topped_days))
            master_in = topped
            if not dry_run:
                shutil.copy(topped, MASTER_FILE)
                master_in = MASTER_FILE

        msgs = fetch_katie_emails(tmp, since_days)
        mark = state.get("watermark")
        mark = dt.datetime.fromisoformat(mark) if mark else None
        new = [m for m in msgs
               if m["id"] not in state["processed_message_ids"]
               and (mark is None or m["date"] > mark)]
        if not new:
            print("Nothing new from Katie.")
            return 0
        print(f"{len(new)} new email(s) from Katie:")
        for m in new:
            print(f"  {m['date']:%Y-%m-%d %H:%M}  {m['subject']}  "
                  f"({len(m['files'])} attachment(s))")

        days, rejected = collect_days(new)
        if rejected:
            print("\nset aside:")
            for r in rejected:
                print(f"  - {r}")
        if not days:
            print("No usable run sheets in those emails.")
            if rejected:
                notify_daniel("Wagon run sheets — nothing usable",
                              "Katie's latest email(s) had no run sheet I could use:\n\n"
                              + "\n".join(f"  - {r}" for r in rejected), dry_run)
            if not dry_run:
                state["processed_message_ids"] += [m["id"] for m in new]
                save_state(state)
            return 0

        print(f"\nfilling {len(days)} day(s):")
        try:
            final, results = build(master_in, days, tmp)
        except Exception as e:
            print(f"BUILD FAILED: {e}")
            alert_daniel([f"build failed: {e}"], [], dry_run)
            return 1

        problems = validate(results)
        if problems:
            print("\nvalidation failed:")
            for p in problems:
                print(f"  - {p}")
            alert_daniel(problems, results, dry_run)
            return 1

        last = max(dt.date.fromisoformat(r["date"]) for r in results)
        note = commentary(results, month_context(final, last))
        fname = send(final, results, note, dry_run)

        if rejected:      # good days went to Paul; the data issue is Daniel's to chase
            notify_daniel(
                "Wagon run sheets — one to check with Katie",
                "The master went to Paul with the days that were fine. These were set "
                "aside:\n\n" + "\n".join(f"  - {r}" for r in rejected)
                + "\n\nOnce Katie re-sends a corrected sheet it will be picked up "
                  "automatically on the next run.", dry_run)

        if not dry_run:
            shutil.copy(final, MASTER_FILE)           # promote only after a clean send
            state["processed_message_ids"] += [m["id"] for m in new]
            state["processed_message_ids"] = state["processed_message_ids"][-500:]
            save_state(state)
        if out_dir:
            Path(out_dir).mkdir(parents=True, exist_ok=True)
            dest = Path(out_dir) / fname
            shutil.copy(final, dest)
            print(f"  copied to {dest}")
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    ap.add_argument("--seed", metavar="XLSX", help="store this file as the starting master")
    ap.add_argument("--seed-from-gmail", action="store_true",
                    help="seed from the newest 'Daily wagon earnings' attachment in Gmail")
    ap.add_argument("--once", action="store_true", help="run now")
    ap.add_argument("--scheduled", action="store_true", help="run from cron")
    ap.add_argument("--dry-run", action="store_true", help="build but send nothing")
    ap.add_argument("--since-days", type=int, default=21)
    ap.add_argument("--out-dir", help="also drop the finished master here")
    a = ap.parse_args()

    if a.seed or a.seed_from_gmail:
        STATE_DIR.mkdir(exist_ok=True)
        if a.seed_from_gmail:
            with tempfile.TemporaryDirectory() as tmp:
                found, subj, when = fetch_master_from_gmail(tmp)
                if not found:
                    print("No 'Daily wagon earnings *.xlsx' attachment found in Gmail.\n"
                          "Email the current master to yourself, then run this again.")
                    return 1
                print(f"Found {found.name}\n  from: {subj}\n  sent: {when:%Y-%m-%d %H:%M}")
                shutil.copy(found, MASTER_FILE)
                a.seed = found.name
        else:
            shutil.copy(a.seed, MASTER_FILE)
        st = load_state()
        # The seeded master is already up to date, so everything Katie has sent so far
        # counts as done — otherwise the first run would refill weeks of history.
        with tempfile.TemporaryDirectory() as tmp:
            seen = fetch_katie_emails(tmp, a.since_days)
        st["processed_message_ids"] = [m["id"] for m in seen]
        st.pop("watermark", None)      # this master IS current; no recovery point needed
        save_state(st)
        print(f"Seeded master from {a.seed} -> {MASTER_FILE}")
        print(f"Marked {len(seen)} existing email(s) from Katie as already done.")
        print("From here it only acts on new ones.")
        return 0
    if not (a.once or a.scheduled):
        ap.print_help()
        return 1
    return run(dry_run=a.dry_run, since_days=a.since_days, out_dir=a.out_dir)


if __name__ == "__main__":
    sys.exit(main())
