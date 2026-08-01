"""
line_review.py — nightly data-quality check on every classified line
=====================================================================
Added after the Omega crusher misfile (30/07/2026): a £3,290 plant line
shipped on the Leyland Wagons tab because nothing reviews the classifier's
output. This runs after each build and emails Daniel (only) a review of
anything that looks wrong, so errors are caught the same evening instead
of by the client.

Two tiers:
  1. Deterministic rules — catch the failure shapes we've actually seen.
  2. Claude review — every flagged line plus every line >= £500 gets a
     second opinion from the model (same key/pattern as reply_agent).
     If the API is unavailable the rules still run; review never blocks
     or delays the main report send.
"""
import json
import re
import urllib.request

import db
from classify import TOP_SHEETS
from vrm_lookup import load_lookup

WAGON_DIVISIONS = {'Fox Wagons', 'Leyland Wagons', 'JA Jackson', 'J FISHER'}
PLANT_TERMS = ['CRUSHER', 'SCREENER', 'SHOVEL', 'HYUNDAI', 'HYUNADAI', 'WIRTGEN',
               'TELEHANDLER', 'EXCAVATOR', 'DIGGER', 'DUMPER', 'PLANER', 'LOADING SHOVEL',
               'A25G', 'A30G', 'JCB ', 'KOMATSU', 'DOOSAN']
REG_RX = re.compile(r'[A-Z]{2}\d{2}[A-Z]{3}')
BIG_LINE = 2000.0
CLAUDE_THRESHOLD = 500.0
MODEL = "claude-sonnet-5"


def _refs(line):
    return f"{line.get('supplier_ref') or ''} {line.get('custom_ref') or ''}".upper()


def rule_flags(lines, reg_known):
    flags = []
    for ln in lines:
        division = (ln.get('division') or '').strip()
        area = (ln.get('area') or '').strip()
        cost = float(ln.get('cost') or 0)
        refs = _refs(ln)
        reasons = []
        if division in WAGON_DIVISIONS and any(t in refs for t in PLANT_TERMS):
            reasons.append('plant-machinery reference on a wagon division')
        if area == 'UNIDENTIFIED' and cost >= 100:
            reasons.append('unidentified area')
        if division in WAGON_DIVISIONS and cost >= 250 and not REG_RX.search(refs.replace(' ', '')):
            reasons.append('no vehicle registration on a wagon line')
        m = REG_RX.search(refs.replace(' ', ''))
        if m and m.group(0) not in reg_known:
            reasons.append(f'registration {m.group(0)} not in vehicle master')
        if cost >= BIG_LINE:
            reasons.append(f'big-ticket line (>= £{BIG_LINE:,.0f})')
        if reasons:
            flags.append((ln, reasons))
    return flags


def claude_review(candidates, env):
    """One batched call: verdict per line. Returns {idx: verdict} or {} on any failure."""
    key = env.get('ANTHROPIC_API_KEY', '')
    if not key or not candidates:
        return {}
    listing = []
    for i, (ln, reasons) in enumerate(candidates):
        listing.append(
            f"{i}| division={ln.get('division')} | area={ln.get('area')} | "
            f"supplier={ln.get('supplier')} | part={ln.get('part_name')} | "
            f"£{float(ln.get('cost') or 0):,.2f} | refs={_refs(ln)} | rule_flags={'; '.join(reasons) or 'none'}")
    system = (
        "You review classified workshop-spend lines for a haulage group. Divisions: "
        "Fox Wagons / Leyland Wagons / JA Jackson / J FISHER are WAGON fleets; Plant and "
        "J Fisher Plant are PLANT machinery (crushers, shovels, excavators, screeners); "
        "Tyres, Damage, Windscreen & Glass, Capital, Graphics, Misc are categories. "
        "For each line say whether the division/area looks RIGHT or WRONG for the parts "
        "and references shown, and if wrong, where it should go. Respond ONLY with JSON: "
        '{"lines": [{"i": <idx>, "verdict": "OK"|"WRONG"|"UNSURE", "note": "<max 15 words>"}]}')
    body = json.dumps({
        "model": MODEL, "max_tokens": 1500,
        "system": system,
        "messages": [{"role": "user", "content": "\n".join(listing)}],
    }).encode()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages", data=body,
        headers={"x-api-key": key, "anthropic-version": "2023-06-01",
                 "content-type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            out = json.loads(r.read())
        text = "".join(b.get("text", "") for b in out.get("content", []))
        text = text.replace("```json", "").replace("```", "").strip()
        parsed = json.loads(text)
        return {int(l["i"]): (l.get("verdict", "UNSURE"), l.get("note", ""))
                for l in parsed.get("lines", [])}
    except Exception as e:
        print(f"line_review: Claude pass skipped ({e})")
        return {}


def review(report_date, env):
    """Run the full review; returns (n_lines, findings) where findings is a list of
    dicts ready for the email table. Never raises."""
    try:
        lines = db.day_tab_rows(report_date)
        maps = load_lookup()
        reg_to_area = maps[0] if isinstance(maps, tuple) else maps
        reg_known = set(reg_to_area.keys())
        flagged = rule_flags(lines, reg_known)
        # Claude candidates: every rule-flagged line + every line >= threshold
        seen = {id(ln) for ln, _ in flagged}
        candidates = list(flagged)
        for ln in lines:
            if float(ln.get('cost') or 0) >= CLAUDE_THRESHOLD and id(ln) not in seen:
                candidates.append((ln, []))
        verdicts = claude_review(candidates, env)
        findings = []
        for i, (ln, reasons) in enumerate(candidates):
            v, note = verdicts.get(i, (None, ''))
            if not reasons and v in (None, 'OK'):
                continue          # big line that both tiers are happy with
            findings.append({
                'supplier': ln.get('supplier') or '',
                'part': ln.get('part_name') or '',
                'cost': float(ln.get('cost') or 0),
                'division': ln.get('division') or '',
                'area': ln.get('area') or '',
                'refs': _refs(ln)[:60],
                'rules': '; '.join(reasons) or '—',
                'claude': f"{v}: {note}" if v else 'not reviewed',
            })
        return len(lines), findings
    except Exception as e:
        print(f"line_review: review failed safely ({e})")
        return 0, []


def send_review_email(report_date_long, n_lines, findings, env):
    """Email the findings to Daniel only. No findings -> no email."""
    if not findings:
        print(f"line_review: {n_lines} lines checked, 0 findings — no review email")
        return
    rows = "".join(
        f"<tr><td>{f['supplier']}</td><td>{f['part'][:40]}</td>"
        f"<td style='text-align:right'>£{f['cost']:,.2f}</td>"
        f"<td>{f['division']} / {f['area']}</td><td>{f['refs']}</td>"
        f"<td>{f['rules']}</td><td>{f['claude']}</td></tr>"
        for f in findings)
    html = f"""
    <p>Nightly line review for <b>{report_date_long}</b>: {n_lines} lines checked,
    <b>{len(findings)} flagged</b> for a look before anyone queries them.</p>
    <table border="1" cellpadding="4" cellspacing="0"
           style="border-collapse:collapse;font-size:12px">
      <tr style="background:#24214A;color:#fff"><th>Supplier</th><th>Part</th><th>Cost</th>
      <th>Division / Area</th><th>Refs</th><th>Rule flags</th><th>Claude view</th></tr>
      {rows}
    </table>
    <p style="font-size:12px;color:#888">Fix = classifier keyword or vehicle-master entry;
    tomorrow's rebuild self-corrects the month. Prepared automatically by danielwalsh.ai</p>"""
    from send_email import _resend_send
    _resend_send(env, subject=f"Data Quality Review — {report_date_long} — "
                              f"{len(findings)} line(s) flagged",
                 html=html, to=["daniel.walsh@kfltd.uk"])
    print(f"line_review: {len(findings)} findings emailed to Daniel")
