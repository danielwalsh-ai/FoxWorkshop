"""
Lancashire pack automation — Mel's email in, Simon's pack out.

Simon Colderley built this pack off Mel Vose's daily master and sent it round.
He has left, so the trigger left with him. This watches for Mel's email, rebuilds
the pack from the attached workbook, and sends it to the same distribution list.

Kept deliberately identical to Simon's: same subject format, same body, same
recipients, same attachment name. Only the sign-off changes.

Usage:
    python lancs_auto.py --once --dry-run     # build, send nothing
    python lancs_auto.py --once --to-me       # real send, to Daniel only
    python lancs_auto.py --once               # live
    python lancs_auto.py --scheduled          # cron; quiet when nothing new
"""
import os
import ssl
import sys
import json
import email
import shutil
import smtplib
import imaplib
import argparse
import tempfile
import datetime as dt
from pathlib import Path
from email.message import EmailMessage

import lancs_pack
from lancs_pack import money

HERE = Path(__file__).parent
STATE_DIR = HERE / "state"
STATE_FILE = STATE_DIR / "lancs_auto.json"

MEL = "mel@foxbrothers.co.uk"
SUBJECT_MATCH = "Daily wagon earning master"
# Simon's distribution list, unchanged.
SEND_TO = ["mel@foxbrothers.co.uk", "paulfox@foxbrothers.co.uk",
           "mark.hierons@foxgroup.co", "darren@foxbrothers.co.uk",
           "samuel@foxbrothers.co.uk", "katie@foxbrothers.co.uk",
           "stuartsweet@foxbrothers.co.uk"]
SEND_CC = ["barry.hope@foxgroup.co", "mike.yates@foxgroup.co", "liam@foxgroup.co",
           "Richard.Kirwin@foxgroup.co", "daniel.walsh@kfltd.uk"]
SIGN_OFF = "Daniel"
IMAP_TIMEOUT = 90
MIN_PLAUSIBLE_DAY = 20_000
MAX_PLAUSIBLE_DAY = 250_000


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


def load_state():
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    return {"processed_message_ids": [], "last_run": None}


def save_state(s):
    STATE_DIR.mkdir(exist_ok=True)
    s["last_run"] = dt.datetime.now().isoformat(timespec="seconds")
    STATE_FILE.write_text(json.dumps(s, indent=2), encoding="utf-8")


def _connect():
    M = imaplib.IMAP4_SSL("imap.gmail.com", timeout=IMAP_TIMEOUT)
    M.login(ENV["GMAIL_USER"], ENV["GMAIL_APP_PASSWORD"])
    return M


def fetch_mel(tmpdir, since_days=10, wanted=None):
    """Headers first; the workbook is ~2MB so only download what's new."""
    M = _connect()
    try:
        M.select("INBOX")
        since = (dt.date.today() - dt.timedelta(days=since_days)).strftime("%d-%b-%Y")
        typ, data = M.search(None, f'(SINCE "{since}" FROM "{MEL}")')
        heads = []
        for uid in data[0].split():
            typ, md = M.fetch(uid, "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID DATE SUBJECT)])")
            if not md or not md[0]:
                continue
            h = email.message_from_bytes(md[0][1])
            subj = str(email.header.make_header(
                email.header.decode_header(h.get("Subject", ""))))
            if SUBJECT_MATCH.lower() not in subj.lower() or subj.lower().startswith("recall"):
                continue
            raw = h.get("Date")
            heads.append({"uid": uid, "subject": subj,
                          "id": (h.get("Message-ID") or uid.decode()).strip(),
                          "date": email.utils.parsedate_to_datetime(raw) if raw else None})
        heads = [h for h in heads if h["date"]]
        heads.sort(key=lambda m: m["date"])
        todo = [h for h in heads if wanted is None or wanted(h)]
        out = []
        for h in todo:
            typ, md = M.fetch(h["uid"], "(RFC822)")
            msg = email.message_from_bytes(md[0][1])
            box = Path(tmpdir) / h["id"].strip("<>").replace("/", "_")[:60]
            box.mkdir(parents=True, exist_ok=True)
            files = []
            for part in msg.walk():
                nm = part.get_filename()
                if nm and nm.lower().endswith(".xlsx") and not nm.startswith("~$"):
                    p = box / nm
                    p.write_bytes(part.get_payload(decode=True))
                    files.append(p)
            if files:
                out.append({**h, "files": files})
        return out, heads
    finally:
        try:
            M.logout()
        except Exception:
            pass


def body_text(r):
    k, off = r["kpis"], r["off"]
    d = r["focus"]
    return (
        "Hi All,\n\n"
        f"Today's Lancashire daily wagon earnings pack for {d:%a} {d.day} {d:%b %Y} "
        "is attached.\n\n"
        f"Total earnings: {money(k['total_earnings'])}\n"
        f"EBITDA: {money(k['ebitda'])}\n"
        f"Profit: {money(k['profit'])}\n"
        f"Rolling 5-day average: {money(k['rolling_5day'])}\n"
        f"Wagons under target: {k['under_target']} of {k['in_service']} in service\n"
        f"Wagons off the road on the day: {int(k['off_road'] or 0)}\n"
        f"Wagon-days lost in the last 7 trading days: {int(k['wagon_days_lost'] or 0)}\n\n"
        "Full breakdown in the attached PDF.\n\n"
        f"{SIGN_OFF}\n")


def body_html(r):
    k = r["kpis"]
    d = r["focus"]
    return f"""<html><body>
<p>Hi All,</p>
<p>Today's Lancashire daily wagon earnings pack for {d:%a} {d.day} {d:%b %Y} is attached.</p>
<p>Total earnings: {money(k['total_earnings'])}<br>
EBITDA: {money(k['ebitda'])}<br>
Profit: {money(k['profit'])}<br>
Rolling 5-day average: {money(k['rolling_5day'])}<br>
Wagons under target: {k['under_target']} of {k['in_service']} in service<br>
Wagons off the road on the day: {int(k['off_road'] or 0)}<br>
Wagon-days lost in the last 7 trading days: {int(k['wagon_days_lost'] or 0)}</p>
<p>Full breakdown in the attached PDF.</p>
<p>{SIGN_OFF}</p>
</body></html>"""


def validate(r):
    k = r["kpis"]
    p = []
    if k["total_earnings"] is None:
        p.append("no total earnings for the focus day")
    else:
        if k["total_earnings"] > MAX_PLAUSIBLE_DAY:
            p.append(f"earnings {money(k['total_earnings'])} implausibly high")
        if 0 < k["total_earnings"] < MIN_PLAUSIBLE_DAY:
            p.append(f"earnings {money(k['total_earnings'])} implausibly low")
    if not k["in_service"] or k["in_service"] < 40:
        p.append(f"only {k['in_service']} wagons in service")
    if r["focus"] > dt.date.today():
        p.append(f"focus day {r['focus']} is in the future")
    return p


def send(r, pdf, dry_run=False, to_me=False):
    d = r["focus"]
    subject = f"Daily Wagon Earnings Pack — {d:%a} {d.day} {d:%b %Y}"
    to = [ENV["GMAIL_USER"]] if to_me else SEND_TO
    cc = [] if to_me else SEND_CC
    if to_me:
        subject = "[TEST — would go to the Lancashire list] " + subject
    m = EmailMessage()
    m["From"] = f"Daniel Walsh <{ENV['GMAIL_USER']}>"
    m["To"] = ", ".join(to)
    if cc:
        m["Cc"] = ", ".join(cc)
    m["Subject"] = subject
    m.set_content(body_text(r))
    m.add_alternative(body_html(r), subtype="html")
    m.add_attachment(Path(pdf).read_bytes(), maintype="application", subtype="pdf",
                     filename=Path(pdf).name)
    if dry_run:
        print(f"\n[DRY RUN] to {', '.join(to)}")
        print(f"[DRY RUN] cc {', '.join(cc) if cc else '-'}")
        print(f"[DRY RUN] subject: {subject}")
        print(f"[DRY RUN] attach: {Path(pdf).name}\n")
        print(body_text(r))
        return
    ctx = ssl.create_default_context()
    with smtplib.SMTP("smtp.gmail.com", 587) as s:
        s.starttls(context=ctx)
        s.login(ENV["GMAIL_USER"], ENV["GMAIL_APP_PASSWORD"])
        s.send_message(m)
    print(f"  sent to {', '.join(to)}" + (f" (cc {len(cc)})" if cc else ""))


def notify(subject, body, dry_run=False):
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


def last_pack_sent():
    """When we last sent a pack, taken from Gmail rather than local state.

    The container has no persistent volume, so state/ is wiped on every redeploy.
    Without this a fresh container would treat Mel's latest as unseen and send the
    pack to the whole Lancashire list a second time."""
    M = _connect()
    try:
        M.select('"[Gmail]/All Mail"')
        since = (dt.date.today() - dt.timedelta(days=30)).strftime("%d-%b-%Y")
        typ, data = M.search(
            None, f'(SINCE "{since}" FROM "{ENV["GMAIL_USER"]}" '
                  f'SUBJECT "Daily Wagon Earnings Pack")')
        newest = None
        for uid in data[0].split():
            typ, md = M.fetch(uid, "(BODY.PEEK[HEADER.FIELDS (DATE SUBJECT)])")
            if not md or not md[0]:
                continue
            h = email.message_from_bytes(md[0][1])
            subj = str(email.header.make_header(
                email.header.decode_header(h.get("Subject", ""))))
            if subj.strip().startswith("[TEST"):        # --to-me runs don't count
                continue
            raw = h.get("Date")
            if raw:
                d = email.utils.parsedate_to_datetime(raw)
                newest = d if newest is None or d > newest else newest
        return newest
    finally:
        try:
            M.logout()
        except Exception:
            pass


def run(dry_run=False, to_me=False, since_days=10, out_dir=None, force=False):
    state = load_state()
    done = set(state["processed_message_ids"])
    if not force and not state["processed_message_ids"]:
        mark = last_pack_sent()
        if mark:
            state["watermark"] = mark.isoformat()
            save_state(state)
            print(f"No local state — last pack we sent was {mark:%Y-%m-%d %H:%M}; "
                  f"only newer mail from Mel will be acted on.")
    wm = state.get("watermark")
    wm = dt.datetime.fromisoformat(wm) if wm else None
    with tempfile.TemporaryDirectory() as tmp:
        new, heads = fetch_mel(tmp, since_days,
                               wanted=(lambda h: True) if force
                               else (lambda h: h["id"] not in done
                                     and (wm is None or h["date"] > wm)))
        if not new:
            print(f"Nothing new from Mel ({len(heads)} email(s) checked).")
            return 0
        msg = new[-1]                       # newest wins; Mel re-sends corrections
        print(f"{len(new)} new email(s) from Mel; using "
              f"{msg['date']:%Y-%m-%d %H:%M}  {msg['subject']}")
        book = msg["files"][-1]
        pack_dir = Path(out_dir) if out_dir else Path(tmp)
        pack_dir.mkdir(parents=True, exist_ok=True)
        tmp_pdf = Path(tmp) / "pack.pdf"
        r = lancs_pack.build(str(book), str(tmp_pdf),
                             reports_dir=str(HERE.parent / "fox-portal" / "reports"))
        name = f"Fox-Group-Daily-Wagon-Earnings-{r['focus']:%Y-%m-%d}.pdf"
        pdf = pack_dir / name
        shutil.copy(tmp_pdf, pdf)
        k = r["kpis"]
        print(f"  focus {r['focus']}  earnings {money(k['total_earnings'])}  "
              f"under target {k['under_target']}/{k['in_service']}  -> {name}")

        problems = validate(r)
        if problems:
            print("  validation failed:")
            for p in problems:
                print(f"   - {p}")
            notify("Lancashire pack — held back",
                   "The pack was built but NOT sent — these checks failed:\n\n"
                   + "\n".join(f"  - {p}" for p in problems), dry_run)
            return 1

        send(r, pdf, dry_run, to_me)
        if not dry_run and not to_me:
            state["processed_message_ids"] = (state["processed_message_ids"]
                                              + [m["id"] for m in new])[-300:]
            save_state(state)
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    ap.add_argument("--once", action="store_true")
    ap.add_argument("--scheduled", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--to-me", action="store_true",
                    help="send for real, but to Daniel only")
    ap.add_argument("--force", action="store_true",
                    help="rebuild from the newest email even if already processed")
    ap.add_argument("--since-days", type=int, default=10)
    ap.add_argument("--out-dir")
    a = ap.parse_args()
    if not (a.once or a.scheduled):
        ap.print_help()
        return 1
    return run(a.dry_run, a.to_me, a.since_days, a.out_dir, a.force)


if __name__ == "__main__":
    sys.exit(main())
