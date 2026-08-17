"""Build a reg -> Cover-area map from the Autocheck fleet report.

Autocheck has no explicit area column, so we read the vehicle "type" from the
Internal Vehicle Name (falling back to Make/Model) and map it to one of the
report's existing Cover areas. Only high-confidence types are mapped; anything
ambiguous is left unmapped so it surfaces for review rather than being misfiled.
"""
import re
import csv

# type keyword -> Cover area (checked in order, first match wins; most specific first)
TYPE_RULES = [
    (r"TRAILER|DRAWBAR|NOOTEBOOM|\bFBT\d|\bHTR\d|\bNMS\d\b|\bMSS\d", "TRAILERS"),
    (r"ALLY BODY|ALI BODY", "8 ALI BODY"),
    (r"SWEEPER", "SWEEPER"),
    (r"MIXER", "CONCRETE MIXER"),
    (r"GRAB", "GRABS"),
    (r"HOOK|SKIP LOADER", "HOOKS"),
    (r"PICK ?UP", "PICK UP"),
    (r"TELEHANDLER|FORK CARR|SHOVEL|EXCAVATOR|DIGGER|\bPLANT\b|LOADING SHOVEL", "PLANT"),
    (r"FUELLER|FUEL TANKER|\bTANKER\b", "FUEL TANKER"),
    (r"SHUNTER", "SHUNTER"),
    (r"SPRINTER|MAXUS|DELIVER|CRAFTER|TRANSIT|\bVAN\b", "VAN"),
    (r"\bX5\b|X-DRIVE|XDRIVE|\bBMW\b|\bAUDI\b|\bA6\b|\bA4\b|CAR\b", "CAR"),
    (r"TIPPER", "TIPPER"),
    (r"ARTIC|6X2 UNIT|\bUNIT\b|TRACTOR", "ARTICS"),
    (r"8W|8 W|STEEL BODY|STEEL-SLEEPER|SLEEPER|8X4|MIDLAND|6X4", "8 WHEELERS"),
]


def area_for_type(internal_name, make, model):
    t = f"{internal_name or ''} {model or ''} {make or ''}".upper()
    if not t.strip():
        return None
    if re.search(r"\bEV\b", t):
        return "8 EV"
    for pat, area in TYPE_RULES:
        if re.search(pat, t):
            return area
    return None


def _norm(x):
    return re.sub(r"[^A-Z0-9]", "", str(x).upper()) if x else ""


def build_reg_area(fleet_csv):
    """reg -> area for every Autocheck vehicle we can confidently type."""
    out = {}
    for r in csv.DictReader(open(fleet_csv, encoding="utf-8-sig")):
        reg = _norm(r.get("VRM"))
        if not reg:
            continue
        a = area_for_type(r.get("Internal Vehicle Name"), r.get("Make"), r.get("Model"))
        if a:
            out[reg] = a
    return out


if __name__ == "__main__":
    import sys
    m = build_reg_area(sys.argv[1] if len(sys.argv) > 1
                       else r"C:\Users\gemin\Downloads\vehicles-report.csv")
    from collections import Counter
    print(f"mapped {len(m)} Autocheck vehicles to an area")
    for a, n in Counter(m.values()).most_common():
        print(f"  {a:<16} {n}")
