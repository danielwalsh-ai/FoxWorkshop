"""
Email the daily report (PDF + XLSX) via Resend.

Reads from .env:
    RESEND_API_KEY   your Resend API key
    EMAIL_FROM       verified sender, e.g. 'Fox Reports <reports@kfltd.uk>'
    EMAIL_TO         comma-separated recipient list

Usage:
    python send_email.py <pdf> <xlsx> <report_date_long> ["Today's total"]
"""
import os
import sys
import json
import base64
import urllib.request
import urllib.error
from pathlib import Path

HERE = Path(__file__).parent


def load_env():
    """OS env vars (prod) win; local .env fills gaps for development."""
    env = dict(os.environ)
    p = HERE / ".env"
    if p.exists():
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip()
    return env


def _attach(path):
    data = Path(path).read_bytes()
    return {"filename": Path(path).name, "content": base64.b64encode(data).decode()}


def _fmt_gbp(v):
    return f"£{v:,.0f}"


def _averages_html(spd, floor=100.0):
    """Two compact tables (age range + area) of average spend per active day,
    omitting any line whose YTD/day average is under £floor. Returns '' if spd
    is missing so the email still sends."""
    if not spd:
        return ""
    agel = {2021: "2021 plate", 2022: "2022 plate", 2023: "2023 plate",
            2024: "2024 plate", 2025: "2025 plate", 2026: "2026 plate",
            "other": "Older / private"}
    age_rows = []
    for k in (2021, 2022, 2023, 2024, 2025, 2026, "other"):
        yv = spd["age_ytd"].get(k, 0)
        if yv >= floor:
            age_rows.append((agel[k], yv, spd["age_mtd"].get(k, 0)))
    area_rows = []
    for a, yv in sorted(spd["area_ytd"].items(), key=lambda x: -x[1]):
        if yv >= floor:
            area_rows.append((a.title(), yv, spd["area_mtd"].get(a, 0)))

    def _tbl(title, rows):
        head = (f"<tr><th align='left' style='background:#24214a;color:#fff;"
                f"padding:5px 10px;font-size:12px'>{title}</th>"
                "<th align='right' style='background:#24214a;color:#fff;padding:5px 10px;font-size:12px'>YTD / day</th>"
                "<th align='right' style='background:#24214a;color:#fff;padding:5px 10px;font-size:12px'>This month / day</th></tr>")
        body = ""
        for i, (lbl, yv, mv) in enumerate(rows):
            bg = "#f2f3f7" if i % 2 == 0 else "#ffffff"
            body += (f"<tr style='background:{bg}'>"
                     f"<td style='padding:4px 10px;font-size:12px'>{lbl}</td>"
                     f"<td align='right' style='padding:4px 10px;font-size:12px'>{_fmt_gbp(yv)}</td>"
                     f"<td align='right' style='padding:4px 10px;font-size:12px'>{_fmt_gbp(mv)}</td></tr>")
        return ("<table cellspacing='0' cellpadding='0' "
                "style='border-collapse:collapse;border:1px solid #ddd;margin:0 24px 16px 0'>"
                + head + body + "</table>")

    return (
        "<h3 style='color:#24214a;margin:18px 0 4px'>Average Spend per Day</h3>"
        "<p style='color:#666;margin:0 0 10px;font-size:12px'>Average spend per active day. "
        "YTD from 1 January 2026; this month from the 1st. Lines averaging under £100/day are omitted.</p>"
        "<table cellspacing='0' cellpadding='0'><tr>"
        f"<td valign='top'>{_tbl('By age range', age_rows)}</td>"
        f"<td valign='top'>{_tbl('By area', area_rows)}</td>"
        "</tr></table>"
    )


def send_report(pdf_path, xlsx_path, report_date_long, headline="", to_me=False, spd=None):
    env = load_env()
    api_key = env.get("RESEND_API_KEY", "")
    sender = env.get("EMAIL_FROM", "")
    # danielwalsh@kfltd.uk (no dot) is not a real mailbox — replies bounce.
    # Correct it wherever it appears, regardless of env config (DW 18/07/2026).
    DEAD = "danielwalsh@kfltd.uk"
    GOOD = "daniel.walsh@kfltd.uk"
    # EMAIL_FROM may be formatted as 'Name <address>' — replace anywhere it appears
    import re as _re
    sender = _re.sub(_re.escape(DEAD), GOOD, sender, flags=_re.IGNORECASE)
    if not sender.strip():
        sender = GOOD
    recipients = [e.strip() for e in env.get("EMAIL_TO", "").split(",") if e.strip()]
    recipients = [GOOD if r.lower() == DEAD else r for r in recipients]
    # recipients added in code (env not directly editable from chat) — DW 01/08/2026
    ALWAYS_INCLUDE = ["reports@foxgroup.co"]
    recipients += [a for a in ALWAYS_INCLUDE if a.lower() not in {r.lower() for r in recipients}]
    if to_me:
        recipients = [GOOD]
    seen = set()
    recipients = [r for r in recipients if not (r.lower() in seen or seen.add(r.lower()))]
    reply_to = env.get("EMAIL_REPLY_TO", GOOD)
    if not api_key:
        raise RuntimeError("RESEND_API_KEY is empty in .env")
    if not recipients:
        raise RuntimeError("EMAIL_TO is empty in .env")

    subject = f"Fox Group — Daily Workshop Spend Report — {report_date_long}"
    body_line = f"<p style='font-size:15px'>{headline}</p>" if headline else ""
    html = f"""
    <div style="font-family:Arial,Helvetica,sans-serif;color:#24214a">
      <h2 style="color:#24214a;margin-bottom:4px">Fox Group Ltd — Daily Workshop Spend Report</h2>
      <p style="color:#666;margin-top:0">{report_date_long}</p>
      {body_line}
      <p>The full KPI report (PDF) and transaction workbook (Excel) are attached.</p>
      {_averages_html(spd)}
      <hr style="border:none;border-top:1px solid #ccc">
      <p style="font-size:12px;color:#888">Prepared automatically by danielwalsh.ai</p>
    </div>
    """

    payload = {
        "from": sender,
        "reply_to": reply_to,
        "to": recipients,
        "subject": subject,
        "html": html,
        "attachments": [_attach(pdf_path), _attach(xlsx_path)],
    }
    reply_to = env.get("EMAIL_REPLY_TO", "").strip()
    if reply_to:
        payload["reply_to"] = reply_to
    req = urllib.request.Request(
        "https://api.resend.com/emails",
        data=json.dumps(payload).encode(),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "User-Agent": "fox-report/1.0 (+automation)",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req) as resp:
            body = resp.read().decode()
        print(f"Sent to {', '.join(recipients)}  ->  {body}")
        return body
    except urllib.error.HTTPError as e:
        detail = e.read().decode()
        raise RuntimeError(f"Resend API error {e.code}: {detail}") from None


def _resend_send(env, subject, html, to, attachments=None):
    """Generic Resend send used by auxiliary emails (line review etc.)."""
    api_key = env.get("RESEND_API_KEY", "")
    sender = env.get("EMAIL_FROM", "") or "daniel.walsh@kfltd.uk"
    import re as _re
    sender = _re.sub(_re.escape("danielwalsh@kfltd.uk"), "daniel.walsh@kfltd.uk",
                     sender, flags=_re.IGNORECASE)
    payload = {"from": sender, "to": to, "subject": subject, "html": html,
               "reply_to": "daniel.walsh@kfltd.uk"}
    if attachments:
        payload["attachments"] = attachments
    req = urllib.request.Request(
        "https://api.resend.com/emails",
        data=json.dumps(payload).encode(),
        headers={"Authorization": f"Bearer {api_key}",
                 "Content-Type": "application/json",
                 "User-Agent": "fox-report/1.0 (+automation)"},
        method="POST")
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode()


def main():
    if len(sys.argv) < 4:
        print("Usage: python send_email.py <pdf> <xlsx> <report_date_long> [headline]")
        sys.exit(1)
    pdf, xlsx, date_long = sys.argv[1], sys.argv[2], sys.argv[3]
    headline = sys.argv[4] if len(sys.argv) > 4 else ""
    send_report(pdf, xlsx, date_long, headline)


if __name__ == "__main__":
    main()
