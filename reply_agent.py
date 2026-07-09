"""
Reply agent — reads replies to the daily report, actions them, and reports back.

Runs 4x/day (03:00, 09:00, 12:00, 19:00 Europe/London). For each NEW reply:
  * data question  -> query the database, reply with the answer
  * reclassification -> apply an override, rebuild the report, reply w/ the update
  * opinion / judgement -> add a suggested draft to Daniel's digest (never auto-sent)
Then emails Daniel a short digest of the run.

Reads + sends via Gmail (app password); understands replies via Claude.

Usage:
    python reply_agent.py --once            # run now regardless of schedule
    python reply_agent.py --once --dry-run  # interpret + print, don't send/apply
    python reply_agent.py --scheduled       # cron: only run at 3/9/12/19 London
"""
import os
import sys
import ssl
import json
import email
import imaplib
import smtplib
import datetime as dt
import urllib.request
import urllib.error
from pathlib import Path
from email.message import EmailMessage
from email.utils import parseaddr, getaddresses

HERE = Path(__file__).parent
MODEL = "claude-sonnet-5"
OWN_ADDRESSES = {"daniel.walsh@kfltd.uk", "danielwalsh@kfltd.uk",
                 "reports@crossmanwalsh.com"}

SCHEMA_DESC = """
Table: transactions — one row per purchase-order line of workshop/vehicle spend.
Columns:
  report_date (date)        the business day the line belongs to
  supplier (text), supplier_source_depot (text)
  system_no (text)          supplier 365/Sage number
  supplier_pn (text), part_name (text)
  cost (numeric GBP), surcharge (numeric GBP)
  po_no (text), attached_order_no (text), attached_customer (text)
  po_created_date (timestamp), supply_type (text)
  item_count (int), goods_received (text)
  target_depot (text), assigned_depot (text)
  supplier_ref (text), custom_ref (text)
  division (text)   e.g. 'Fox Wagons','Leyland Wagons','JA Jackson','JJ O''Grady Civils',
                    'J FISHER','NMS CIVIL','Plant','Graphics','PPE- Workwear','Tyres','Misc',
                    'Assets For sale','Asphalt Plant ','J Fisher Plant','NMS Plant','Damage',
                    'Windscreen & Glass','Capital'
  area (text)       vehicle area e.g. '8 WHEELERS','ARTICS','HOOKS','TIPPER', or 'UNIDENTIFIED'
  plate (text)      'pre24' | '24plate' | '25plate'
  vehicle_reg (text) the registration, e.g. 'PO26TOJ'.
Table: budgets(division, year, month, budget).
All amounts are GBP. Today is {today}.

REGISTRATION YEAR — to filter/group by a vehicle's plate year, compute it in SQL from vehicle_reg
(NOT from the `plate` column, which only has pre24/24plate/25plate). Use exactly this expression:
  (CASE WHEN substring(vehicle_reg from '[0-9][0-9]')::int >= 51 THEN 1950 ELSE 2000 END)
    + substring(vehicle_reg from '[0-9][0-9]')::int
Only rows where vehicle_reg ~ '^[A-Z][A-Z][0-9][0-9][A-Z][A-Z][A-Z]$' have a valid year.
CRITICAL: many vehicles have personalised/cherished plates that do NOT encode a real year. Only treat
the computed year as valid when it is BETWEEN 2001 AND EXTRACT(YEAR FROM CURRENT_DATE); exclude every
other row from year analysis (never report years like 2038/2040 — they don't exist). "Number of trucks"
= COUNT(DISTINCT vehicle_reg). For spend-by-year, group only over 2021-2026 and bucket earlier valid
years as 'pre-2021'. If a result looks implausible, say you'll double-check rather than stating it.
Example — 22-plate spend this month:
  SELECT ROUND(SUM(cost),2) FROM transactions
  WHERE vehicle_reg ~ '^[A-Z][A-Z][0-9][0-9][A-Z][A-Z][A-Z]$'
    AND (CASE WHEN substring(vehicle_reg from '[0-9][0-9]')::int >= 51 THEN 1950 ELSE 2000 END)
        + substring(vehicle_reg from '[0-9][0-9]')::int = 2022
    AND date_trunc('month', report_date) = date_trunc('month', CURRENT_DATE);
"""


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
# When OFF, the agent drafts replies into Daniel's digest instead of auto-sending.
AUTO_SEND = ENV.get("AGENT_AUTOSEND", "").strip().lower() in ("1", "true", "yes", "on")


def _clean_sql(s):
    s = s.strip()
    if s.startswith("```"):
        s = s[3:]
        if s[:3].lower() == "sql":
            s = s[3:]
        if "```" in s:
            s = s[:s.rindex("```")]
    return s.strip().rstrip(";").strip()


# ── Claude ──────────────────────────────────────────────────────────
def call_claude(system, user, max_tokens=1500):
    body = {"model": MODEL, "max_tokens": max_tokens, "system": system,
            "messages": [{"role": "user", "content": user}]}
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps(body).encode(),
        headers={"x-api-key": ENV["ANTHROPIC_API_KEY"],
                 "anthropic-version": "2023-06-01",
                 "content-type": "application/json"},
        method="POST")
    with urllib.request.urlopen(req, timeout=90) as r:
        data = json.loads(r.read().decode())
    for block in data.get("content", []):
        if block.get("type") == "text":
            return block["text"]
    return ""


def parse_json(text):
    a, b = text.find("{"), text.rfind("}")
    if a == -1 or b == -1:
        return {}
    try:
        return json.loads(text[a:b + 1])
    except json.JSONDecodeError:
        return {}


# ── Gmail (IMAP read) ───────────────────────────────────────────────
def strip_quoted(text):
    out = []
    for line in text.splitlines():
        s = line.strip()
        if s.startswith(">") or s.startswith("On ") and s.endswith("wrote:"):
            break
        if s.startswith("-----Original Message-----") or s.startswith("From:") and "Sent:" in text[:0]:
            break
        out.append(line)
    return "\n".join(out).strip()[:4000]


def fetch_replies(since_days=3):
    M = imaplib.IMAP4_SSL("imap.gmail.com")
    M.login(ENV["GMAIL_USER"], ENV["GMAIL_APP_PASSWORD"])
    M.select("INBOX")
    since = (dt.date.today() - dt.timedelta(days=since_days)).strftime("%d-%b-%Y")
    typ, data = M.search(None, f'(SINCE "{since}" SUBJECT "transaction report")')
    replies = []
    for uid in data[0].split():
        typ, md = M.fetch(uid, "(RFC822)")
        msg = email.message_from_bytes(md[0][1])
        from_addr = parseaddr(msg.get("From", ""))[1].lower()
        if from_addr in OWN_ADDRESSES or not from_addr:
            continue
        # plaintext body
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="replace")
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors="replace")
        replies.append({
            "message_id": msg.get("Message-ID", "").strip(),
            "from": from_addr,
            "from_name": parseaddr(msg.get("From", ""))[0],
            "subject": msg.get("Subject", ""),
            "references": msg.get("References", ""),
            "body": strip_quoted(body),
        })
    M.logout()
    return replies


# ── Gmail (SMTP send) ───────────────────────────────────────────────
def send_email(to_addrs, subject, body, in_reply_to=None, references=None):
    m = EmailMessage()
    m["From"] = f"Daniel Walsh <{ENV['GMAIL_USER']}>"
    m["To"] = ", ".join(to_addrs)
    m["Subject"] = subject
    if in_reply_to:
        m["In-Reply-To"] = in_reply_to
        m["References"] = (references + " " + in_reply_to).strip() if references else in_reply_to
    m.set_content(body)
    ctx = ssl.create_default_context()
    with smtplib.SMTP("smtp.gmail.com", 587) as s:
        s.starttls(context=ctx)
        s.login(ENV["GMAIL_USER"], ENV["GMAIL_APP_PASSWORD"])
        s.send_message(m)


# ── the brain ───────────────────────────────────────────────────────
def classify(reply):
    system = ("You triage replies to a daily vehicle/workshop spend report for Fox Group. "
              "Classify the sender's message and return ONLY JSON:\n"
              '{"type": "data_question" | "reclassification" | "opinion" | "other",\n'
              ' "summary": "one line of what they want",\n'
              ' "reclassification": {"po_no": "...", "division": "...", "area": "..."} (only if reclassification; '
              'division/area must be one of the known values, area optional),\n'
              ' "question": "the data question in plain English" (only if data_question)}\n'
              "data_question = asking for a specific figure/fact answerable from spend data RIGHT NOW "
              "(e.g. 'what did we spend on reg X', 'total 22 plate costs'). reclassification = asking to "
              "move a cost to a different division/area (usually names a PO or describes the item). "
              "opinion = asking your view/judgement, OR requesting a CHANGE/ADDITION to the report itself "
              "(a feature request like 'add trucks-per-year to the report') — these need Daniel's decision, "
              "so classify them as opinion, do NOT auto-answer with figures. other = acknowledgement/thanks.")
    txt = call_claude(system, f"From: {reply['from_name']} <{reply['from']}>\nSubject: {reply['subject']}\n\n{reply['body']}", 600)
    return parse_json(txt)


def answer_data_question(question):
    today = dt.date.today().isoformat()
    sql = _clean_sql(call_claude(
        "You write ONE read-only PostgreSQL SELECT for the given schema. Output ONLY the SQL, no markdown, "
        "no explanation. Use ROUND(SUM(cost),2) for totals. Match vehicle_reg case-insensitively. "
        "For a specific month use report_date ranges. Follow the REGISTRATION YEAR guidance exactly.\n\n"
        + SCHEMA_DESC.format(today=today),
        f"Question: {question}", 500))
    import db
    try:
        cols, rows = db.run_select(sql)
    except Exception as e:
        return None, f"(could not run query: {e})", sql
    data_txt = f"Columns: {cols}\nRows: {rows[:50]}"
    answer = call_claude(
        "You are Daniel Walsh replying to a colleague about vehicle/workshop spend. Write a short, friendly, "
        "professional reply (2-4 sentences) answering their question using ONLY the data. Use £ and real numbers. "
        "Sign off 'Kind regards, Daniel'. No preamble.",
        f"Question: {question}\n\nData:\n{data_txt}", 500)
    return answer, None, sql


def do_reclassification(rc):
    import db
    from build_report import build
    po = (rc.get("po_no") or "").strip()
    division = (rc.get("division") or "").strip() or None
    area = (rc.get("area") or "").strip() or None
    if not po:
        return None, "no PO number identified"
    db.add_override(po, division=division, area=area, note="via reply agent")
    db.apply_overrides()
    d = db.po_report_date(po)
    if d:
        try:
            build(d)
        except Exception as e:
            print(f"  rebuild warning: {e}")
    dest = division or area or "the requested category"
    msg = call_claude(
        "You are Daniel Walsh. Write a short, professional confirmation (1-2 sentences) that a cost has been "
        "reclassified and the report updated. Sign off 'Kind regards, Daniel'.",
        f"PO {po} has been moved to {dest}. The report and dashboard are updated.", 300)
    return msg, None


def draft_opinion(reply):
    return call_claude(
        "You are Daniel Walsh. Draft a brief, professional reply to this message for Daniel to review before "
        "sending. Keep it natural. Sign off 'Kind regards, Daniel'.",
        f"From {reply['from_name']}: {reply['body']}", 500)


# ── main run ────────────────────────────────────────────────────────
def process(dry_run=False):
    import db
    replies = fetch_replies()
    digest = []
    for r in replies:
        mid = r["message_id"]
        if not mid or db.is_processed(mid):
            continue
        cls = classify(r)
        rtype = cls.get("type", "other")
        who = r["from_name"] or r["from"]
        subj = r["subject"] if r["subject"].lower().startswith("re:") else ("Re: " + r["subject"])
        summary = cls.get("summary", "")
        suggested, note, sendable, action = None, "", False, rtype

        if rtype == "data_question":
            answer, err, sql = answer_data_question(cls.get("question", r["body"]))
            suggested = answer
            sendable = bool(answer)
            note = "" if answer else f"couldn't answer ({err})"
            action = "answered" if answer else "answer_failed"

        elif rtype == "reclassification":
            rc = cls.get("reclassification", {}) or {}
            po = (rc.get("po_no") or "").strip()
            dest = rc.get("division") or rc.get("area") or "?"
            if po:
                if not dry_run:
                    msg, err = do_reclassification(rc)
                    suggested = msg
                    note = f"AUTO-APPLIED: PO {po} → {dest} (report rebuilt)"
                else:
                    suggested = f"(confirmation of moving PO {po} to {dest})"
                    note = f"[dry-run] would apply {rc}"
                sendable = bool(suggested)
                action = "reclassified"
            else:
                suggested = draft_opinion(r)
                note = f"NEEDS YOU — no specific PO to auto-apply ({rc})"
                action = "needs_daniel"

        elif rtype == "opinion":
            suggested = draft_opinion(r)
            note = "NEEDS YOUR JUDGEMENT"
            action = "needs_daniel"

        sent = False
        if sendable and AUTO_SEND and not dry_run:
            send_email([r["from"]], subj, suggested, in_reply_to=mid, references=r["references"])
            sent = True

        block = f"• {rtype.upper()} from {who}: {summary}"
        if note:
            block += f"\n   {note}"
        if suggested:
            label = "SENT ✓" if sent else "SUGGESTED REPLY (review & send)"
            block += f"\n   {label}:\n   " + suggested.replace("\n", "\n   ")
        digest.append(block)

        if not dry_run:
            db.mark_processed(mid, rtype, action)

    if digest:
        head = (f"Reply agent — {len(digest)} item(s) at {dt.datetime.now():%H:%M %d %b}  "
                f"(auto-send {'ON' if AUTO_SEND else 'OFF — drafts for you'})")
        report = head + "\n\n" + "\n\n".join(digest)
        print(report)
        if not dry_run:
            send_email([ENV["GMAIL_USER"]], "Reply agent digest", report)
    else:
        print("No new replies to handle.")


def main():
    argv = sys.argv[1:]
    dry = "--dry-run" in argv
    if "--scheduled" in argv:
        from zoneinfo import ZoneInfo
        h = dt.datetime.now(ZoneInfo("Europe/London")).hour
        if h not in (3, 9, 12, 19):
            print(f"[guard] London hour {h} not a check time — skipping.")
            return
    process(dry_run=dry)


if __name__ == "__main__":
    main()
