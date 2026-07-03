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


def send_report(pdf_path, xlsx_path, report_date_long, headline=""):
    env = load_env()
    api_key = env.get("RESEND_API_KEY", "")
    sender = env.get("EMAIL_FROM", "")
    recipients = [e.strip() for e in env.get("EMAIL_TO", "").split(",") if e.strip()]
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
      <hr style="border:none;border-top:1px solid #ccc">
      <p style="font-size:12px;color:#888">Prepared automatically by Knowles Farm Ltd.</p>
    </div>
    """

    payload = {
        "from": sender,
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


def main():
    if len(sys.argv) < 4:
        print("Usage: python send_email.py <pdf> <xlsx> <report_date_long> [headline]")
        sys.exit(1)
    pdf, xlsx, date_long = sys.argv[1], sys.argv[2], sys.argv[3]
    headline = sys.argv[4] if len(sys.argv) > 4 else ""
    send_report(pdf, xlsx, date_long, headline)


if __name__ == "__main__":
    main()
