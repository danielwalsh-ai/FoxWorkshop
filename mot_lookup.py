"""DVSA MOT History API — annual-test expiry lookup (covers HGVs via CVS data).

Credentials come from env (os.environ first, then .env fallback):
  MOT_CLIENT_ID, MOT_CLIENT_SECRET, MOT_API_KEY, MOT_SCOPE, MOT_TOKEN_URL

Usage:
    python mot_lookup.py PN72EFA PN22EJC ...        # ad-hoc
    from mot_lookup import lookup_many               # {reg: {...}}

For each reg returns the CURRENT annual-test expiry (latest PASSED test's
expiryDate) plus make/model and the last recorded odometer.
"""
import os
import json
import time
import urllib.request
import urllib.parse
import urllib.error
from pathlib import Path

HERE = Path(__file__).parent
API_BASE = "https://history.mot.api.gov.uk/v1/trade/vehicles/registration/"


def _load_env():
    vals = {}
    envf = HERE / ".env"
    if envf.exists():
        for line in envf.read_text().splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                vals.setdefault(k.strip(), v.strip())
    return lambda k: os.environ.get(k) or vals.get(k)


_token = {"val": None, "exp": 0.0}


def get_token(g):
    now = time.time()
    if _token["val"] and _token["exp"] - 60 > now:
        return _token["val"]
    data = urllib.parse.urlencode({
        "grant_type": "client_credentials", "client_id": g("MOT_CLIENT_ID"),
        "client_secret": g("MOT_CLIENT_SECRET"), "scope": g("MOT_SCOPE")}).encode()
    req = urllib.request.Request(g("MOT_TOKEN_URL"), data=data,
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    tok = json.loads(urllib.request.urlopen(req, timeout=30).read())
    _token["val"] = tok["access_token"]
    _token["exp"] = now + int(tok.get("expires_in", 3600))
    return _token["val"]


def _norm(reg):
    return (reg or "").replace(" ", "").upper()


def lookup_one(reg, g=None, token=None):
    g = g or _load_env()
    token = token or get_token(g)
    url = API_BASE + urllib.parse.quote(_norm(reg))
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}", "x-api-key": g("MOT_API_KEY"),
        "Accept": "application/json"})
    try:
        body = json.loads(urllib.request.urlopen(req, timeout=30).read())
    except urllib.error.HTTPError as e:
        return {"reg": _norm(reg), "expiry": None,
                "error": "not found" if e.code == 404 else f"HTTP {e.code}"}
    tests = body.get("motTests", []) or []
    passed = [t for t in tests if t.get("testResult") == "PASSED" and t.get("expiryDate")]
    expiry = max((t["expiryDate"] for t in passed), default=None)
    latest = max(tests, key=lambda t: t.get("completedDate", ""), default={})
    return {"reg": _norm(reg), "make": body.get("make"), "model": body.get("model"),
            "expiry": expiry, "odometer": latest.get("odometerValue"),
            "odometer_unit": latest.get("odometerUnit"), "error": None}


def lookup_many(regs, pause=0.12):
    g = _load_env()
    token = get_token(g)
    out = []
    for r in regs:
        out.append(lookup_one(r, g, token))
        time.sleep(pause)          # gentle on the rate limit
    return out


def main():
    import sys
    regs = sys.argv[1:]
    if not regs:
        print("usage: python mot_lookup.py REG [REG ...]")
        return
    for r in lookup_many(regs):
        exp = r.get("expiry") or ("— (" + (r.get("error") or "no test") + ")")
        print(f"{r['reg']:<9} {exp:<22} {r.get('make','') or ''} {r.get('model','') or ''}"
              f"  odo {r.get('odometer','')}{r.get('odometer_unit','') or ''}")


if __name__ == "__main__":
    main()
