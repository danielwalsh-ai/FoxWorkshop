"""Classify a free-text part description into a spend category.

The Autovolt data has no structured part category — only a free-text
`part_name` (e.g. "BUDGET DRIVE TYRE", "VIPER LIGHTBAR 42\"", "ABS SENSOR").
Paul asked to "categorise the spend on the trucks into part categories".
This maps each description to one of a small, CEO-friendly set of buckets.

Rules are ordered: the FIRST matching pattern wins, so put the more specific
patterns before the generic ones.
"""
import re

# (category, regex).  Evaluated top-to-bottom; first hit wins.
_RULES = [
    # --- Tyres & wheels (incl. bare tyre sizes like 315/80/22.5, 2958022.5) ---
    ("Tyres & Wheels", r"TYRE|PUNCTURE|WHEEL\s*TRIM|\bWHEEL\b|REMOULD|\bRETREAD\b"),
    ("Tyres & Wheels", r"\b\d{3}[\s/\-]?\d{2}[\s/\-]?\d?\d?[\s/\-]?22[.\s]?5\b"),  # tyre sizes
    ("Tyres & Wheels", r"\b295[\s/\-]?80|385[\s/\-]?65|315[\s/\-]?80|"
                       r"PIRELLI|MICHELIN|CONTINENTAL|SCORPION|\b285[\s/\-]?40"),

    # --- Safety, camera & telematics (blind-spot radar, cameras, GPS alarms) ---
    ("Safety, Camera & Telematics", r"RADAR|CAMERA|\bMOIS\b|\bAHD\b|SPLITTER|\bCCTV\b|"
                                    r"\bDVR\b|TELEMATIC|\bGPS\b|\bALARM\b|BLIND\s*SPOT|"
                                    r"QUAD|MONITOR|\bDASH\s*CAM"),

    # --- Brakes, suspension & air ---
    ("Brakes, Suspension & Air", r"\bABS\b|BRAKE|CALIPER|\bDISC\b|BRAKE\s*PAD|\bPADS?\b|"
                                 r"SPRING|SHOCK|DAMPER|DIAPHRAGM|AIR\s*TANK|LEVELL|"
                                 r"SUSPENSION|BUSH|WISHBONE|\bAXLE\b|HUB\b"),

    # --- Lighting & electrical ---
    ("Lighting & Electrical", r"LIGHTBAR|BEACON|HEADLIGHT|HEADLAMP|\bLAMP\b|\bLED\b|\bLIGHT\b|WIPER|"
                              r"\bECU\b|SENSOR|BATTERY|\bCABLE\b|RELAY|\bFUSE\b|WIRING|"
                              r"ROCKER|ANDERSON|DURITE|STARLIGHT|CAMERA|\bCAM\b|"
                              r"SCENE\s*LIGHT|WARNING\s*LAMP|ALTERNATOR|STARTER|"
                              r"\bPLUG\b|COIL(?!\s*SPRING)|MODULE|MEGAFUSE|THIN\s*WALL|"
                              r"TECHLED|\bISL\d|OBLONG|REFLECTOR|INDICATOR"),

    # --- Glass, windscreen & mirrors glass ---
    ("Glass & Windscreen", r"GLASS|WINDSCREEN|WINDOW|\bLENS\b"),

    # --- Body, cab, panels & graphics ---
    ("Body, Cab & Panels", r"BODY|\bWING\b|PILLAR|PANEL|MIRROR|STEP\s*PLATE|VALANCE|"
                           r"BRACKET|SIDEGUARD|GRAPHIC|\bSEAT\b|\bCOVER|GUARD|\bDOOR\b|"
                           r"GRILLE|BUMPER|MUDGUARD|MUDG|CORNER|CATWALK|TRIM(?!\s|$)|"
                           r"FITTING\s*OUT|FITTING\s*OF|\bTRIM\b|STAY\b|WRAP|LIVERY|"
                           r"\bSTEP\b|\bRAMP\b|HEAT\s*SHIELD|VALENCE"),

    # --- Servicing, MOT, testing & diagnostics ---
    ("Servicing, MOT & Testing", r"\bMOT\b|SERVIC|TACHO|DIAGNOSTIC|\bTEST\b|INSPECT|"
                                 r"CALIBRAT|\bLABOUR\b|\bREPAIR\b(?!\s*KIT)"),

    # --- Filters, fluids & cooling ---
    ("Filters, Fluids & Cooling", r"FILTER|\bOIL\b|ADBLUE|COOLANT|RADIATOR|INTERCOOLER|"
                                  r"WATER\s*PUMP|\bFAN\b|GREASE|LUBRICAN|ANTIFREEZE|"
                                  r"AIR\s*CON|AIRCON|\bA/C\b"),

    # --- Load securing, straps & trailer ancillaries ---
    ("Load Securing & Trailer", r"STRAP|RATCHET|CLAMP|\bCHAIN\b|\bSHEET\b|TRAILER|"
                                r"HYDROCLEAR|TIPP|WALKING\s*FLOOR|\bHOSE\b|COUPLING"),

    # --- Fuel & efficiency ---
    ("Fuel & Efficiency", r"FUEL|OPTIMISATION|EMISSION|EXHAUST|\bDPF\b"),

    # --- Tools & consumables ---
    ("Tools & Consumables", r"WRENCH|SOCKET|SPANNER|\bTOOL|DRILL|\bBIT\b"),
]

_COMPILED = [(cat, re.compile(rx, re.I)) for cat, rx in _RULES]

CATEGORIES = [
    "Tyres & Wheels",
    "Safety, Camera & Telematics",
    "Brakes, Suspension & Air",
    "Lighting & Electrical",
    "Glass & Windscreen",
    "Body, Cab & Panels",
    "Servicing, MOT & Testing",
    "Filters, Fluids & Cooling",
    "Load Securing & Trailer",
    "Fuel & Efficiency",
    "Tools & Consumables",
    "Other / Uncategorised",
]


def categorise(part_name: str) -> str:
    text = (part_name or "").strip()
    if not text:
        return "Other / Uncategorised"
    for cat, rx in _COMPILED:
        if rx.search(text):
            return cat
    return "Other / Uncategorised"
