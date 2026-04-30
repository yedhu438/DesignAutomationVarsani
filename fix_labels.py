
content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()

# Insert the colour/size parsing helpers just before build_label_layer
OLD = '''def build_label_layer(label_text):
    """Small black label overlay for top-left corner of zone."""'''

NEW = '''# ─── SKU COLOUR / SIZE PARSING ───────────────────────────────────────────────
# Maps SKU colour codes → readable colour names (matches owner's label format)
COLOUR_MAP = {
    "Blk": "Black",  "Wht": "White",  "Nvy": "Navy",   "Red": "Red",
    "Pnk": "Pink",   "Gry": "Grey",   "Blu": "Blue",   "Grn": "Green",
    "Ylw": "Yellow", "Fus": "Fuchsia","Pur": "Purple",  "Org": "Orange",
    "Bur": "Burgundy","Nat": "Natural","Lav": "Lavender","RBlu":"Royal Blue",
    "SBlu":"Sky Blue","Camo":"Camo",   "TD": "Tie Dye",  "GryM":"Grey Marl",
    "Ivry":"Ivory",   "BPnk":"Baby Pink",
}

def parse_sku_colour_size(sku):
    """
    Extract colour and size from SKU for use in layer labels.
    e.g. MenTee_WhtXL   -> ("White", "XL")
         KidsTee_Blk911 -> ("Black", "9-11")
         AdultPoloTee_RBluM -> ("Royal Blue", "M")
    Returns (colour_str, size_str) — either may be empty string.
    """
    if not sku:
        return "", ""
    # Split on underscore — last segment has colour+size
    parts = sku.split("_")
    if len(parts) < 2:
        return "", ""
    last = parts[-1]

    # Try to match colour codes (longest match first)
    colour_str = ""
    remainder  = last
    for code in sorted(COLOUR_MAP.keys(), key=len, reverse=True):
        if last.startswith(code):
            colour_str = COLOUR_MAP[code]
            remainder  = last[len(code):]
            break

    # Size: whatever remains after the colour code
    # Normalise age sizes: 911 → 9-11, 78 → 7-8, 1213 → 12-13 etc.
    size_raw = remainder.strip()
    size_str = size_raw
    if size_raw.isdigit() and len(size_raw) >= 2:
        mid = len(size_raw) // 2
        size_str = size_raw[:mid] + "-" + size_raw[mid:]

    return colour_str, size_str


def make_zone_label(zone_key, sku, use_sku_detail=True):
    """
    Build the layer label string.
    - use_sku_detail=True  → "Front - White XL"   (different designs per size)
    - use_sku_detail=False → "front"               (identical designs)
    """
    zone_display = zone_key.title()   # "front" → "Front"
    if not use_sku_detail:
        return zone_display
    colour, size = parse_sku_colour_size(sku)
    parts = [zone_display]
    if colour:
        parts.append(colour)
    if size:
        parts.append(size)
    return " - ".join(parts)   # "Front - White XL"


def build_label_layer(label_text):
    """Small black label overlay for top-left corner of zone."""'''

if OLD in content:
    content = content.replace(OLD, NEW)
    open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)
    print("Step 1 OK: colour/size helpers added")
else:
    print("ERROR: build_label_layer not found")
