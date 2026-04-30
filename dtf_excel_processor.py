"""
dtf_excel_processor.py
======================
Processes orders from the exported Excel file (DTF unshipped orders).
Reads BOTH sheets (DTF Orders + DTF Order Details), joins them, and
generates layered PSD files for each order.

Usage:
    python dtf_excel_processor.py
    python dtf_excel_processor.py --limit 10
    python dtf_excel_processor.py --dry-run

Input:
    - Excel file : W:\\test1\\UnshippedDTFOrders_11042026_031657.xlsx
    - Images     : W:\\test1\\DTFUnshippedImages_20260411_031645\\

Output:
    - PSDs : C:\\Varsany\\Output\\DTF_Excel\\<Category>\\<OrderID>_<SKU>.psd
"""

import os, json, struct, io, traceback, argparse, textwrap
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont, ImageOps

# ─── CONFIG ───────────────────────────────────────────────────────────────────

EXCEL_FILE    = r"W:\test 8\UnshippedDTFOrders_20042026_014056.xlsx"
IMAGE_FOLDER  = r"W:\test 8\DTFUnshippedImages_20260420_014001"
OUTPUT_FOLDER = r"C:\Varsany\Output\DTF_Excel"
FONT_FOLDERS  = [r"C:\Varsany\Fonts", r"W:\fonts"]
LOG_FILE      = r"C:\Varsany\Output\DTF_Excel\dtf_excel_log.txt"

PX_PER_CM = 120          # 304 DPI for test — change to 320 for production
DPI       = int(PX_PER_CM * 2.54)

# ─── SKIP LIST ────────────────────────────────────────────────────────────────
# Orders where designs are already ready — do not process
SKIP_ORDER_IDS: set = set()   # cleared — covered by SKU prefix matching below
# SKU prefixes to skip (design already available for entire product type)
SKIP_SKU_PREFIXES = [
    "FballN",             # football kids hoodies (FballN01, FballN02, FballN05...)
    "NewFballN",          # new football kids hoodies
    "FBall",              # football PE bags (FBall01PEBag...)
    "FootballAdultTee_",  # football adult tees
    "FootballKidsTee_",   # football kids tees
    "Memorial_",          # memorial book orders
    "Name01Bckpck",       # backpacks type 1
    "Name02Bckpck",       # backpacks type 2
    "Name03Bckpck",       # backpacks type 3
]

# ─── PRODUCT CANVAS SIZES ─────────────────────────────────────────────────────
def cm(x): return int(round(x * PX_PER_CM))

PRODUCT_CANVAS = {
    "adulttshirt":  {"front": (cm(30),   cm(30)),  "back": (cm(30),   cm(30)),  "pocket": (cm(9), cm(7))},
    "kidstshirt":   {"front": (cm(23),   cm(30)),  "back": (cm(23),   cm(30)),  "pocket": (cm(9), cm(7))},
    "adulthoodie":  {"front": (cm(25),   cm(25)),  "back": (cm(25),   cm(25)),  "pocket": (cm(9), cm(7))},
    "kidshoodie":   {"front": (cm(23),   cm(20)),  "back": (cm(23),   cm(20)),  "pocket": (cm(9), cm(7))},
    "totebag":      {"front": (cm(28),   cm(28)),  "back": (cm(28),   cm(58))},
    "babyvest":     {"front": (cm(15),   cm(17))},
    "buckethat":    {"front": (cm(18),   cm(5))},
    "beanie":       {"front": (cm(9.5),  cm(4.5))},
    "pegbag":       {"front": (cm(23),   cm(14))},  # PE bag / Make Up Bag / Shoe Bag Sports
    "cushion":      {"front": (cm(30),   cm(30))},
    "slipper":      {"front": (cm(6),    cm(6))},
    "backpack":     {"front": (cm(18),   cm(12))},
    "shoebag":      {"front": (cm(14),   cm(14))},
    "knittingbag":  {"front": (cm(25),   cm(21))},
    "hoddiblanket": {"front": (cm(17),   cm(5))},
    "golfcase":     {"front": (cm(15.7), cm(6))},
    "golftowel":    {"front": (cm(17),   cm(17))},
    "seatbelt":     {"front": (cm(18),   cm(4))},
    "memorial":     {"front": (cm(13.5), cm(8.5))},
    "sleepsuit":    {"front": (cm(13),   cm(18))},
    "socks":        {"front": (cm(6),    cm(12))},
    "stringbag":    {"front": (cm(22),   cm(24))},
    "default":      {"front": (cm(30),   cm(30)),  "back": (cm(30),   cm(30)),  "pocket": (cm(9), cm(7))},
}

SKU_MAP = [
    ("MenTee_",              "adulttshirt"), ("WmnTee_",             "adulttshirt"),
    ("PoloTee_",             "adulttshirt"), ("AdultPoloTee_",       "adulttshirt"),
    ("AnyTxtMenTee_",        "adulttshirt"), ("AnyTxtOverSizeTee_",  "adulttshirt"),
    ("LegendSince",          "adulttshirt"), ("FootballAdultTee_",   "adulttshirt"),
    ("Lvrpool",              "adulttshirt"), ("LBalls",              "adulttshirt"),
    ("PinkCymru",            "adulttshirt"), ("Custom_Tee_",         "adulttshirt"),
    ("Anytxt_DogTee",        "adulttshirt"), ("PEngR01PoloJersy_",   "adulttshirt"),
    ("KidsTee_",             "kidstshirt"),  ("CustomKidsTee_",      "kidstshirt"),
    ("FootballKidsTee_",     "kidstshirt"),  ("67BdayT02KidsTee_",   "kidstshirt"),
    ("BirthdaTruck01Kids",   "kidstshirt"),
    ("AnyTxtAdultHood_",     "adulthoodie"), ("AnytxtFleece_",       "adulthoodie"),
    ("FballN",               "kidshoodie"),  ("NewFball",            "kidshoodie"),
    ("Gymnastichoodie",      "kidshoodie"),  ("AnyTxtKidsHood_",     "kidshoodie"),
    ("AnyTxtTote_",          "totebag"),     ("EasterGirlTote_",     "totebag"),
    ("Knitting",             "knittingbag"),
    ("AnyTxtBabyVest_",      "babyvest"),    ("CustomTxt01BabyVest_","babyvest"),
    ("AnyTextHat_",          "buckethat"),
    ("AnytxtPatchBeanie_",   "beanie"),
    ("AnyTxtPEBag_",         "pegbag"),      ("AnyTxtShoeBagSpt_",   "pegbag"),
    ("AnyTxtMakeUpBag_",     "pegbag"),
    ("CustomCushion",        "cushion"),
    ("Memorial_",            "memorial"),
    ("Name01Bckpck",         "backpack"),    ("Name02Bckpck",        "backpack"),
    ("Name03Bckpck",         "backpack"),
    ("AnyTxtShoeBag_",       "shoebag"),
    ("AnyTxtHoddiBlanket_",  "hoddiblanket"),
    ("AnyTxtGolfCase_",      "golfcase"),
    ("AnyTxtGolfTowel_",     "golftowel"),
    ("AnyTxtSeatBelt_",      "seatbelt"),
    ("AnyTxtSleepSuit_",     "sleepsuit"),
    ("AnyTxtSocksAnkl_",     "socks"),       ("AnyTxtSocksSole_",    "socks"),
    ("AnyTxtStringBag_",     "stringbag"),
]

def detect_product(sku):
    for prefix, key in SKU_MAP:
        if sku.startswith(prefix):
            return key
    return "default"

def detect_category(sku):
    s = sku.lower()
    if "kidstee" in s or "footballkids" in s: return "Kids T-Shirt"
    if "polo" in s:                            return "Polo"
    if "hood" in s or "fleece" in s:           return "Hoodie"
    if "knit" in s:                            return "Knitting Bag"
    if "tote" in s or "eastergirl" in s:       return "Tote Bag"
    if "babyvest" in s or "vest" in s:         return "Baby Vest"
    if "bckpck" in s:                          return "Back Pack"
    if "shoebag" in s:                         return "Shoe Bag"
    if "makeupbag" in s:                       return "Make Up Bag"
    if "pegbag" in s or "pebag" in s:          return "PE Bag"
    if "hoddiblanket" in s:                    return "Hoddi Blanket"
    if "golfcase" in s:                        return "Golf Case"
    if "golftowel" in s:                       return "Golf Towel"
    if "seatbelt" in s:                        return "Seat Belt"
    if "cushion" in s:                         return "Cushion"
    if "beanie" in s:                          return "Hat"
    if "buckethat" in s:                       return "Bucket Hat"
    if "memorial" in s:                        return "Memorial"
    if "sleepsuit" in s:                       return "Sleepsuit"
    if "socks" in s:                           return "Socks"
    if "stringbag" in s:                       return "String Bag"
    if "slipper" in s:                         return "Slipper"
    if "dogtee" in s:                          return "Dog Tee"
    return "T-Shirt"

# ─── FONT INDEX ───────────────────────────────────────────────────────────────

FONT_INDEX = {}
for _folder in FONT_FOLDERS:
    if os.path.exists(_folder):
        for _f in os.listdir(_folder):
            if _f.lower().endswith(('.ttf', '.otf')):
                _norm = os.path.splitext(_f)[0].lower()
                _norm = _norm.replace(' ', '').replace('-', '').replace('_', '')
                FONT_INDEX[_norm] = os.path.join(_folder, _f)

FONT_ALIASES = {
    "arial":           "arial",
    "arialbold":       "arial",
    "abel":            "abel",
    "bebasneue":       "bebasneueregular",
    "bebasneuepro":    "bebasneueregular",
    "bebasneuefree":   "bebasneueregular",
    "chewy":           "chewyregular",
    "permanentmarker": "permanentmarkerregular",
    "russoone":        "russooneregular",
    "ultra":           "ultraregular",
    "fondamento":      "fondamento",
    "lato":            "lato",
    "latobold":        "latobold",
    "roboto":          "roboto",
    "verdana":         "verdana",
    # Premium texture fonts — DB display name (spaces stripped) → FONT_INDEX key
    "texturefont":          "smartkids",
    "texture":              "smartkids",
    "blockfont":            "colorfulblocks",
    "colorfulblock":        "colorfulblocks",
    "colorfulblocks":       "colorfulblocks",
    "paintfont":            "paintsplashesrainbow",
    "paintsplashes":        "paintsplashesrainbow",
    "paintsplashesrainbow": "paintsplashesrainbow",
    "mermaidfont":          "wavemermaid",
    "mermaid":              "wavemermaid",
    "wavemermaid":          "wavemermaid",
    "reflectionfont":       "refractionray",
    "reflection":           "refractionray",
    "refractionray":        "refractionray",
    "camofont":             "camoblock",
    "camo":                 "camoblock",
    "camoblock":            "camoblock",
    "spideyfont":           "spiderweb",
    "spidey":               "spiderweb",
    "spiderweb":            "spiderweb",
    "cozyfont":             "cozywinter",
    "cozy":                 "cozywinter",
    "cozywinter":           "cozywinter",
    "footballfont":         "soccerarmy",
    "football":             "soccerarmy",
    "soccerarmy":           "soccerarmy",
    "vinyl":                "vinylfont",
    "vinylfont":            "vinylfont",
    "bouqet":               "bouqetdisplay",
    "bouqetdisplay":        "bouqetdisplay",
    "smartkids":            "smartkids",
}

# Every font in this set is a premium colour/texture font — rendered via Chrome
# to preserve its built-in SVG glyph colours.  Customer colour is IGNORED for
# these fonts; colour comes from the font's own SVG glyph data.
PREMIUM_FONT_KEYS = {
    "smartkids", "colorfulblocks", "paintsplashesrainbow",
    "wavemermaid", "refractionray", "camoblock", "spiderweb",
    "cozywinter", "soccerarmy", "vinylfont", "bouqetdisplay",
}

def _resolve_font_key(font_name):
    if not font_name:
        return None
    norm = font_name.lower().replace(' ', '').replace('-', '').replace('_', '')
    alias = FONT_ALIASES.get(norm)
    if alias is not None:
        return alias
    return norm if norm in FONT_INDEX else None

def get_font(font_name, size_px):
    """Resolve font name to a PIL ImageFont, falling back to Arial."""
    norm = font_name.lower().replace(' ', '').replace('-', '').replace('_', '')
    norm = FONT_ALIASES.get(norm, norm)
    if norm in FONT_INDEX:
        try:
            return ImageFont.truetype(FONT_INDEX[norm], size_px)
        except Exception:
            pass
    for ext in ('.ttf', '.otf'):
        for folder in FONT_FOLDERS:
            p = os.path.join(folder, font_name + ext)
            if os.path.exists(p):
                try:
                    return ImageFont.truetype(p, size_px)
                except Exception:
                    pass
    try:
        return ImageFont.truetype("arial.ttf", size_px)
    except Exception:
        return ImageFont.load_default()

def _test_font_renders(path):
    """Return True if this font produces visible pixels when PIL renders 'A'."""
    try:
        import numpy as _np
        _f   = ImageFont.truetype(path, 200)
        _img = Image.new("RGBA", (300, 300), (0, 0, 0, 0))
        ImageDraw.Draw(_img).text((10, 10), "A", font=_f, fill=(0, 0, 0, 255))
        return _np.array(_img)[:, :, 3].max() > 0
    except Exception:
        return False

SVG_ONLY_FONT_KEYS = {k for k, p in FONT_INDEX.items() if not _test_font_renders(p)}

def _is_svg_only_font(font_name):
    return _resolve_font_key(font_name) in SVG_ONLY_FONT_KEYS

def _is_premium_font(font_name):
    return _resolve_font_key(font_name) in PREMIUM_FONT_KEYS

# Chrome renderer — shared with batch_processor
try:
    from batch_processor import build_text_layer_chrome as _chrome_render
    _CHROME_AVAILABLE = True
except Exception:
    _CHROME_AVAILABLE = False
    _chrome_render = None

# ─── TEXT PARSING HELPERS ─────────────────────────────────────────────────────

def _safe(val):
    """Return None if val is NaN/None/empty, else str(val)."""
    if val is None:
        return None
    try:
        import math
        if isinstance(val, float) and math.isnan(val):
            return None
    except Exception:
        pass
    s = str(val).strip()
    return s if s and s.lower() not in ('nan', 'none', '') else None

def parse_text_lines(row, zone):
    """
    Extract ordered text lines for a zone.
    Prefers TextJSON (preserves line order); falls back to plain Text field.
    e.g. {"Text1":"RASHFORD","Text2":"11"} -> ["RASHFORD", "11"]
    """
    cap   = zone.capitalize()
    tjson = _safe(row.get(f'{cap}TextJSON'))
    traw  = _safe(row.get(f'{cap}Text'))

    if tjson:
        try:
            data = json.loads(tjson)
            if isinstance(data, dict):
                lines = []
                for k in sorted(data.keys()):
                    if k.lower().startswith('text') and _safe(data.get(k)):
                        val = str(data[k]).replace('\\n', '\n')
                        for part in val.split('\n'):
                            if part.strip():
                                lines.append(part.strip())
                if lines:
                    return lines
        except Exception:
            pass

    if traw:
        return [l.strip() for l in traw.split('\n') if l.strip()]

    return []

def parse_font_name(row, zone):
    """Return (font_name, is_premium) for a zone."""
    cap   = zone.capitalize()
    fjson = _safe(row.get(f'{cap}Fonts'))
    if fjson:
        try:
            data    = json.loads(fjson)
            premium = _safe(data.get('PremiumFont'))
            if premium and premium.lower() not in ('no', 'false', ''):
                return premium, True
            normal  = _safe(data.get('NormalFont'))
            if normal:
                return normal, False
        except Exception:
            pass
    return 'Arial', False

def parse_colour_hex(row, zone):
    """Return hex colour string for a zone, e.g. '#d1c9c9'."""
    cap   = zone.capitalize()
    cjson = _safe(row.get(f'{cap}Colours'))
    if cjson:
        try:
            data = json.loads(cjson)
            col  = _safe(data.get('Colour1'))
            if col:
                return col if col.startswith('#') else f'#{col}'
        except Exception:
            pass
    return '#000000'

def hex_to_rgb(hex_col):
    h = hex_col.lstrip('#')
    if len(h) == 3:
        h = ''.join(c*2 for c in h)
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

# ─── BACKGROUND REMOVAL ───────────────────────────────────────────────────────

# Garment colour keywords in SKU → approximate RGB
SKU_COLOUR_MAP = {
    'wht':  (255, 255, 255), 'blk':  (0,   0,   0  ),
    'nvy':  (31,  40,  80 ), 'red':  (200, 30,  30 ),
    'ylw':  (255, 220, 0  ), 'opnk': (255, 130, 100),
    'pnk':  (255, 150, 180), 'gry':  (150, 150, 150),
    'grn':  (30,  130, 30 ), 'org':  (230, 100, 20 ),
    'rblu': (50,  100, 220), 'sblu': (100, 150, 210),
    'blu':  (30,  80,  200), 'nat':  (240, 230, 200),
    'ivry': (240, 230, 200), 'bur':  (130, 30,  50 ),
    'fus':  (200, 50,  140), 'pur':  (120, 50,  170),
    'ppl':  (120, 50,  170), 'teal': (0,   130, 130),
}

def _garment_colour(sku):
    """Return approximate garment RGB from SKU, or None if unknown."""
    s = sku.lower()
    # Longest keys first so 'rblu'/'sblu'/'opnk' are matched before 'blu'/'pnk'
    for key, rgb in sorted(SKU_COLOUR_MAP.items(), key=lambda x: -len(x[0])):
        if key in s:
            return rgb
    return None

def _colour_close(c1, c2, tol=40):
    return all(abs(a - b) <= tol for a, b in zip(c1, c2))

def should_remove_bg(img_path, sku, row, zone):
    """
    Return True if background should be removed for this zone image.
    Priority:
      1. Database flag Is[Zone]BgRemove
      2. Auto: image has a solid uniform background (any colour) AND
               that colour matches the garment colour OR is a plain solid
               colour that would look bad as a rectangle on the garment.
    Logic:
      - Sample all 4 edges of the image
      - If the mean of each edge is within tolerance of a single colour
        → background is solid/uniform → remove it
    """
    # 1. Manual flag from Excel / DB
    cap  = zone.capitalize()
    flag = row.get(f'Is{cap}BgRemove')
    try:
        if int(flag) == 1:
            return True
    except (TypeError, ValueError):
        pass

    if not img_path or not os.path.isfile(img_path):
        return False

    # 2. Auto-detect: solid uniform background → remove only if bg matches garment
    try:
        import numpy as np
        img = Image.open(img_path).convert('RGB')
        arr = np.array(img)

        # Collect every pixel on all 4 edges
        all_edge_px = np.concatenate([
            arr[0,  :, :],   # top row
            arr[-1, :, :],   # bottom row
            arr[:,  0, :],   # left col
            arr[:, -1, :],   # right col
        ], axis=0)

        # Background colour = median of edge pixels (robust to corners/noise)
        bg_r = int(np.median(all_edge_px[:, 0]))
        bg_g = int(np.median(all_edge_px[:, 1]))
        bg_b = int(np.median(all_edge_px[:, 2]))
        bg_colour = (bg_r, bg_g, bg_b)

        # What fraction of edge pixels are close to that background colour?
        tol_px = 30
        matches = np.sum(
            (np.abs(all_edge_px[:, 0].astype(int) - bg_r) <= tol_px) &
            (np.abs(all_edge_px[:, 1].astype(int) - bg_g) <= tol_px) &
            (np.abs(all_edge_px[:, 2].astype(int) - bg_b) <= tol_px)
        )
        solid_fraction = matches / len(all_edge_px)

        # Photos/natural images have varying edges → low fraction.
        # Logos on solid bg have nearly uniform edges → high fraction.
        # Require 85%+ edge pixels to be the same colour before removing.
        if solid_fraction < 0.85:
            return False

        # Extra check: photos have thousands of unique colours; logos have few.
        # Sample every 4th pixel for performance.
        sample = arr[::4, ::4, :3].reshape(-1, 3)
        unique_colours = len(np.unique(sample, axis=0))
        if unique_colours > 5000:
            return False   # Complex/photo image — keep background as-is

        # Solid logo background confirmed — only remove if it matches garment colour
        garment_rgb = _garment_colour(sku)
        if garment_rgb and _colour_close(bg_colour, garment_rgb, tol=40):
            return True
        return False
    except Exception:
        return False

_bg_remove_cache: dict = {}
BG_CACHE_DIR = r"C:\Varsany\Temp\bg_cache"
os.makedirs(BG_CACHE_DIR, exist_ok=True)

def _bg_cache_path(img_path):
    import hashlib
    h = hashlib.md5(img_path.encode()).hexdigest()
    return os.path.join(BG_CACHE_DIR, h + ".png")

def remove_background(img_path):
    if img_path in _bg_remove_cache:
        return _bg_remove_cache[img_path]
    # Check disk cache first
    cache_file = _bg_cache_path(img_path)
    if os.path.exists(cache_file):
        result = Image.open(cache_file).convert('RGBA')
        _bg_remove_cache[img_path] = result
        return result
    import numpy as np
    from rembg import remove as rembg_remove, new_session

    with open(img_path, 'rb') as f:
        raw = f.read()

    # Detect background colour from image corners before removal
    orig = ImageOps.exif_transpose(Image.open(img_path)).convert('RGB')
    ow, oh = orig.size
    sample_pts = [(2, 2), (ow-3, 2), (2, oh-3), (ow-3, oh-3)]
    bg_samples = [orig.getpixel(p) for p in sample_pts]
    bg_r = int(sum(c[0] for c in bg_samples) / 4)
    bg_g = int(sum(c[1] for c in bg_samples) / 4)
    bg_b = int(sum(c[2] for c in bg_samples) / 4)

    # Step 1: rembg with isnet model (better edge quality than default)
    try:
        session = new_session("birefnet-general")
        result  = rembg_remove(raw, session=session)
    except Exception:
        result = rembg_remove(raw)   # fallback to default

    img = Image.open(io.BytesIO(result)).convert('RGBA')
    arr = np.array(img, dtype=np.uint8)

    # Step 2: flood-fill from all 4 edges to catch remaining bg pixels
    from collections import deque
    h, w = arr.shape[:2]
    tol  = 35   # colour tolerance for bg matching
    visited = np.zeros((h, w), dtype=bool)
    queue   = deque()

    # Seed with all edge pixels that are close to bg colour
    for x in range(w):
        for y in [0, h-1]:
            r, g, b, a = arr[y, x]
            if (abs(int(r)-bg_r) <= tol and abs(int(g)-bg_g) <= tol
                    and abs(int(b)-bg_b) <= tol):
                queue.append((y, x))
                visited[y, x] = True
    for y in range(h):
        for x in [0, w-1]:
            r, g, b, a = arr[y, x]
            if (abs(int(r)-bg_r) <= tol and abs(int(g)-bg_g) <= tol
                    and abs(int(b)-bg_b) <= tol and not visited[y, x]):
                queue.append((y, x))
                visited[y, x] = True

    while queue:
        cy, cx = queue.popleft()
        arr[cy, cx, 3] = 0   # make transparent
        for dy, dx in [(-1,0),(1,0),(0,-1),(0,1)]:
            ny, nx = cy+dy, cx+dx
            if 0 <= ny < h and 0 <= nx < w and not visited[ny, nx]:
                r, g, b, a = arr[ny, nx]
                if (abs(int(r)-bg_r) <= tol and abs(int(g)-bg_g) <= tol
                        and abs(int(b)-bg_b) <= tol):
                    visited[ny, nx] = True
                    queue.append((ny, nx))

    # Step 3: remove any remaining pixels that match background colour
    # (handles enclosed bg "islands" inside the logo not reached by flood fill)
    r_ch = arr[:, :, 0].astype(int)
    g_ch = arr[:, :, 1].astype(int)
    b_ch = arr[:, :, 2].astype(int)
    inner_bg = (
        (np.abs(r_ch - bg_r) <= tol) &
        (np.abs(g_ch - bg_g) <= tol) &
        (np.abs(b_ch - bg_b) <= tol) &
        (arr[:, :, 3] > 0)
    )
    arr[inner_bg, 3] = 0

    # Step 4: aggressive near-background cleanup — rembg sometimes leaves
    # bg-coloured pixels with high alpha (200-250). Catch them with a tighter
    # colour match but wider alpha range.
    r_ch = arr[:, :, 0].astype(int)
    g_ch = arr[:, :, 1].astype(int)
    b_ch = arr[:, :, 2].astype(int)
    tight_bg = (
        (np.abs(r_ch - bg_r) <= 60) &
        (np.abs(g_ch - bg_g) <= 60) &
        (np.abs(b_ch - bg_b) <= 60) &
        (arr[:, :, 3] > 0)
    )
    arr[tight_bg, 3] = 0

    # Step 5: hard alpha threshold — remove soft fringing
    alpha = arr[:, :, 3]
    arr[:, :, 3] = np.where(alpha < 128, 0, 255).astype(np.uint8)

    result = Image.fromarray(arr, 'RGBA')
    _bg_remove_cache[img_path] = result
    try:
        result.save(cache_file)
    except Exception:
        pass
    return result

# ─── IMAGE LOOKUP ─────────────────────────────────────────────────────────────

def find_images_for_item(order_item_id):
    """
    Find all image files for an OrderItemID in the image folder.
    Returns dict: { 'front': path, 'back': path, 'pocket': path,
                    'frontpreview': path, 'backpreview': path, ... }
    """
    prefix = str(order_item_id)
    found  = {}
    if not os.path.isdir(IMAGE_FOLDER):
        return found
    for fname in os.listdir(IMAGE_FOLDER):
        if not fname.startswith(prefix):
            continue
        fpath = os.path.join(IMAGE_FOLDER, fname)
        lower = fname.lower()
        if   'frontpreview'  in lower: found['frontpreview']  = fpath
        elif 'backpreview'   in lower: found['backpreview']   = fpath
        elif 'pocketpreview' in lower: found['pocketpreview'] = fpath
        elif 'confirmorder'  in lower: found['confirmorder']  = fpath
        elif '-1-front'      in lower: found['front']         = fpath
        elif '-2-front'      in lower: found.setdefault('front2', fpath)
        elif '-3-front'      in lower: found.setdefault('front3', fpath)
        elif '-4-front'      in lower: found.setdefault('front4', fpath)
        elif '-5-front'      in lower: found.setdefault('front5', fpath)
        elif '-1-back'       in lower: found['back']          = fpath
        elif '-1-pocket'     in lower: found['pocket']        = fpath
        elif '-front.jpg'    in lower and 'front'  not in found: found['front']  = fpath
        elif '-back.jpg'     in lower and 'back'   not in found: found['back']   = fpath
        elif '-pocket.jpg'   in lower and 'pocket' not in found: found['pocket'] = fpath
    return found

# ─── PSD WRITER ───────────────────────────────────────────────────────────────

def _to_channels(img, mode):
    """Split image into raw channel bytes (no compression — matches PSD compression type 0).

    PSD CMYK stores ink coverage inverted: 0 = full ink, 255 = no ink.
    PIL CMYK is the opposite, so we invert all CMYK channels before writing.
    """
    import numpy as np

    def _inv(band):
        """Invert a PIL band for PSD CMYK encoding."""
        return bytes(np.clip(255 - np.frombuffer(band.tobytes(), dtype=np.uint8), 0, 255).astype(np.uint8))

    if mode == 'CMYKA':
        rgba = img.convert('RGBA')
        a    = rgba.split()[3]
        cmyk = rgba.convert('RGB').convert('CMYK')
        c, m, y, k = cmyk.split()
        return {-1: a.tobytes(), 0: _inv(c), 1: _inv(m), 2: _inv(y), 3: _inv(k)}
    if mode == 'CMYK':
        cmyk = img.convert('RGB').convert('CMYK')
        c, m, y, k = cmyk.split()
        return {0: _inv(c), 1: _inv(m), 2: _inv(y), 3: _inv(k)}
    img = img.convert(mode)
    bands = img.split()
    if mode == 'RGBA':
        r, g, b, a = bands
        return {-1: a.tobytes(), 0: r.tobytes(), 1: g.tobytes(), 2: b.tobytes()}
    r, g, b = bands
    return {0: r.tobytes(), 1: g.tobytes(), 2: b.tobytes()}

def _pack_layer_name(name):
    n = name.encode('utf-8')[:63]
    pad = (4 - ((1 + len(n)) % 4)) % 4
    return struct.pack('>B', len(n)) + n + b'\x00' * pad

def write_psd(out_path, canvas_w, canvas_h, layers):
    """
    Write a layered RGB PSD (or PSB for large files) file.
    Auto-switches to PSB when estimated size exceeds 1.5 GB.
    Layers: list of dicts with image/top/left/name/opacity/visible.
    """
    # Estimate uncompressed size: canvas pixels × 4 bytes × layers
    est_bytes = canvas_w * canvas_h * 4 * (len(layers) + 1)
    psb = est_bytes > 1.5 * 1024**3   # use PSB if >1.5 GB estimated

    if psb and out_path.endswith('.psd'):
        out_path = out_path[:-4] + '.psb'

    # PSB uses 8-byte length fields; PSD uses 4-byte
    ch_len_fmt   = '>hQ' if psb else '>hI'   # channel data length in layer record
    sec_len_fmt  = '>Q'  if psb else '>I'    # layer info / lmi section lengths
    version      = 2     if psb else 1

    buf = io.BytesIO()
    p   = buf.write

    p(b'8BPS')
    p(struct.pack('>H', version))
    p(b'\x00' * 6)
    p(struct.pack('>H', 4))       # 4 channels (CMYK)
    p(struct.pack('>II', canvas_h, canvas_w))
    p(struct.pack('>HH', 8, 4))   # 8 bpc, CMYK

    # Image resources: DPI
    p(struct.pack('>I', 0))
    dpi_fixed = DPI << 16
    res_data  = struct.pack('>IHHIHH', dpi_fixed, 1, 1, dpi_fixed, 1, 1)
    res_block = b'8BIM' + struct.pack('>H', 1005) + b'\x00\x00' + struct.pack('>I', len(res_data)) + res_data
    p(struct.pack('>I', len(res_block)))
    p(res_block)

    lr_buf = io.BytesIO()
    ld_buf = io.BytesIO()

    for lyr in layers:
        img    = lyr['image'].convert('RGBA')
        top    = lyr['top']
        left   = lyr['left']
        bottom = top  + img.height
        right  = left + img.width
        flags  = 0 if lyr.get('visible', True) else 2

        ch = _to_channels(img, 'CMYKA')
        ch_order = [-1, 0, 1, 2, 3]   # alpha, C, M, Y, K

        lr = io.BytesIO()
        lr.write(struct.pack('>iiii', top, left, bottom, right))
        lr.write(struct.pack('>H', 5))   # 5 channels: alpha + CMYK
        for cid in ch_order:
            lr.write(struct.pack(ch_len_fmt, cid, len(ch[cid]) + 2))
        lr.write(b'8BIM')
        lr.write(b'norm')
        lr.write(struct.pack('>BB', lyr.get('opacity', 255), 0))
        lr.write(struct.pack('>BB', flags, 0))
        name_bytes = _pack_layer_name(lyr['name'])
        extra = struct.pack('>II', 0, 0) + name_bytes
        lr.write(struct.pack('>I', len(extra)))
        lr.write(extra)
        lr_buf.write(lr.getvalue())

        for cid in ch_order:
            ld_buf.write(struct.pack('>H', 0))   # compression = raw
            ld_buf.write(ch[cid])

    layer_info = struct.pack('>h', len(layers)) + lr_buf.getvalue() + ld_buf.getvalue()
    if len(layer_info) % 4:
        layer_info += b'\x00' * (4 - len(layer_info) % 4)
    lmi = struct.pack(sec_len_fmt, len(layer_info)) + layer_info + struct.pack('>I', 0)
    p(struct.pack(sec_len_fmt, len(lmi)))
    p(lmi)

    # Composite (merged) image — CMYK
    composite = Image.new('RGB', (canvas_w, canvas_h), (255, 255, 255))
    for lyr in layers:
        if lyr.get('visible', True):
            composite.paste(lyr['image'].convert('RGBA'),
                            (lyr['left'], lyr['top']),
                            lyr['image'].convert('RGBA'))
    comp = _to_channels(composite, 'CMYK')
    p(struct.pack('>H', 0))
    for cid in [0, 1, 2, 3]:
        p(comp[cid])

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())
    return out_path   # may have changed extension to .psb

# ─── AI UPSCALER (Real-ESRGAN) ────────────────────────────────────────────────


_ESRGAN_HELPER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'esrgan_helper.py')
_ESRGAN_MODEL  = r'C:\Varsany\realesrgan\RealESRGAN_x4plus.pth'

def _sharpen(img_pil):
    """Apply unsharp mask + sharpness boost for crisp DTF output."""
    from PIL import ImageFilter, ImageEnhance
    mode = img_pil.mode
    # Work on RGB(A) — split alpha to avoid sharpening transparent edges
    if mode == 'RGBA':
        r, g, b, a = img_pil.split()
        rgb = Image.merge('RGB', (r, g, b))
    else:
        rgb = img_pil.convert('RGB')
        a = None
    # Unsharp mask: radius=1.5, percent=160, threshold=3
    rgb = rgb.filter(ImageFilter.UnsharpMask(radius=1.5, percent=160, threshold=3))
    # Additional sharpness enhancer (1.0 = original, 2.0 = double)
    rgb = ImageEnhance.Sharpness(rgb).enhance(1.4)
    if a is not None:
        r, g, b = rgb.split()
        return Image.merge('RGBA', (r, g, b, a))
    return rgb.convert(mode) if mode != 'RGB' else rgb


def upscale_image(img_pil, target_w, target_h):
    """Upscale to at least target size using LANCZOS + sharpening."""
    src_w, src_h = img_pil.size
    ratio = max(target_w / src_w, target_h / src_h)
    if ratio > 1.0:
        nw = max(1, int(src_w * ratio))
        nh = max(1, int(src_h * ratio))
        img_pil = img_pil.resize((nw, nh), Image.LANCZOS)
    return _sharpen(img_pil)


# ─── LAYER BUILDERS ───────────────────────────────────────────────────────────

def build_image_layer(img_path, w, h, do_bg_remove=False, upscale=True):
    """Scale image to fit within zone (aspect-ratio preserved). Optionally removes background."""
    if not img_path or not os.path.isfile(img_path):
        return Image.new('RGBA', (max(1, w), max(1, h)), (0, 0, 0, 0)), 0, 0
    if do_bg_remove:
        try:
            src = remove_background(img_path)
            log(f'    BG removed: {os.path.basename(img_path)}')
            # Trim transparent border so the subject fills the canvas after scaling
            bbox = src.getbbox()
            if bbox:
                src = src.crop(bbox)
                log(f'    Trimmed to content: {src.width}x{src.height}')
        except Exception as e:
            log(f'    BG removal failed ({e}), using original', 'WARN')
            src = Image.open(img_path).convert('RGBA')
    else:
        src = ImageOps.exif_transpose(Image.open(img_path)).convert('RGBA')
        # Auto-crop transparent borders from customer-uploaded images
        bbox = src.getbbox()
        if bbox and (bbox[0] > 0 or bbox[1] > 0 or bbox[2] < src.width or bbox[3] < src.height):
            src = src.crop(bbox)
    # AI upscale to canvas size, then final exact resize if needed.
    ratio_w = w / src.width
    ratio_h = h / src.height
    ratio   = min(ratio_w, ratio_h)
    target_w = max(1, int(src.width  * ratio))
    target_h = max(1, int(src.height * ratio))
    src = upscale_image(src, target_w, target_h) if upscale else src.resize((target_w, target_h), Image.LANCZOS)
    # ESRGAN may slightly overshoot/undershoot — snap to exact target
    if src.width != target_w or src.height != target_h:
        src = src.resize((target_w, target_h), Image.LANCZOS)
    nw, nh = src.width, src.height
    src = src.convert('RGBA')
    # Centre horizontally within zone if image is narrower than zone width
    left_offset = (w - nw) // 2
    return src, 0, left_offset

def _distribute_words(words, n):
    """Split words into n lines with roughly equal total character length."""
    if n >= len(words):
        return words
    total = sum(len(w) for w in words)
    target = total / n
    lines, current, current_len = [], [], 0
    for i, word in enumerate(words):
        current.append(word)
        current_len += len(word)
        if current_len >= target and len(lines) < n - 1:
            lines.append(' '.join(current))
            current, current_len = [], 0
    if current:
        lines.append(' '.join(current))
    return lines


def auto_wrap_lines(lines, font_name, w, h):
    """If a single long line exists, try word-wrapping it to maximise font size."""
    stripped = [l for l in lines if l.strip()]
    if len(stripped) != 1:
        return stripped
    line = stripped[0]
    if len(line) <= 20 or ' ' not in line:
        return stripped
    words = line.split()
    if len(words) < 3:
        return stripped
    best_lines = stripped
    best_size  = find_best_font_size(stripped, font_name, w, h)
    for n in range(2, min(len(words), 6)):
        candidate = _distribute_words(words, n)
        if len(candidate) > 1:
            # Reject orphan layouts: skip if any line is <40% the length of the longest
            lengths = [len(l) for l in candidate]
            if min(lengths) < max(lengths) * 0.4:
                continue
            sz = find_best_font_size(candidate, font_name, w, h)
            if sz > best_size * 1.15:   # only wrap if 15%+ font-size gain
                best_size, best_lines = sz, candidate
    return best_lines


def build_text_layer(text_lines, font_name, colour_hex, w, h, force_size=None):
    """Auto-size text to fill zone width, centre each line."""
    lines = [l for l in text_lines if l.strip()]
    if not lines:
        return Image.new('RGBA', (1, 1), (0, 0, 0, 0)), 0, 0

    # Premium / SVG colour fonts: render via Chrome to preserve built-in glyph
    # colours (spider-web pattern, colour blocks, paint splashes, etc.).
    # Customer colour is ALWAYS ignored — the font's own SVG data defines colour.
    if _CHROME_AVAILABLE and (_is_svg_only_font(font_name) or _is_premium_font(font_name)):
        img = _chrome_render(lines, font_name, None, w)
        if img is not None:
            left = max(0, (w - img.width) // 2)
            # force_size: scale Chrome output proportionally if caller requires it
            if force_size is not None:
                natural = find_best_font_size(lines, font_name, w, h)
                if natural > 0 and force_size != natural:
                    scale = force_size / natural
                    nw = max(1, int(img.width  * scale))
                    nh = max(1, int(img.height * scale))
                    img = img.resize((nw, nh), Image.LANCZOS)
                    left = max(0, (w - img.width) // 2)
            return img, 0, left

    lines = auto_wrap_lines(lines, font_name, w, h)
    best = find_best_font_size(lines, font_name, w, h, force_size)
    return _render_text(lines, font_name, colour_hex, w, best)


def find_best_font_size(text_lines, font_name, w, h, force_size=None):
    """Return the largest font size that fits text_lines in w×h. Or return force_size if given."""
    lines = [l for l in text_lines if l.strip()]
    if not lines:
        return 20
    if force_size is not None:
        return force_size
    avail_w = int(w * 0.92)
    avail_h = int(h * 0.90)
    scratch = Image.new('RGBA', (1, 1))
    draw    = ImageDraw.Draw(scratch)
    lo, hi  = 20, max(w, h)
    best    = lo
    has_emoji = any(_line_has_emoji(l) for l in lines)
    while lo <= hi:
        mid  = (lo + hi) // 2
        font = get_font(font_name, mid)
        if has_emoji:
            sizes  = [_pilmoji_measure(l, font) for l in lines]
            widths = [s[0] for s in sizes]
            line_h = int(max(s[1] for s in sizes) * 1.4) if sizes else 20
            total_h = line_h * len(lines)
        else:
            # Use multiline_textbbox so height measurement matches _render_text exactly
            text   = "\n".join(lines)
            widths = [draw.textbbox((0, 0), l, font=font)[2] - draw.textbbox((0, 0), l, font=font)[0]
                      for l in lines]
            mb     = draw.multiline_textbbox((0, 0), text, font=font, spacing=0, align='center')
            total_h = mb[3] - mb[1]
        if max(widths) <= avail_w and total_h <= avail_h:
            best, lo, hi = mid, mid + 1, hi
        else:
            lo, hi = lo, mid - 1
    return best


# ─── EMOJI RENDERING (pilmoji / Twemoji) ──────────────────────────────────────

def _is_emoji_char(c):
    """Return True if the character is likely an emoji or special Unicode symbol
    that may not be present in standard Latin fonts."""
    cp = ord(c)
    return (
        0x00A9 <= cp <= 0x00AE or       # © ®
        0x2000 <= cp <= 0x27FF or       # Misc symbols, arrows, dingbats
        0x2900 <= cp <= 0x2BFF or       # More arrows, misc symbols
        0x1F1E0 <= cp <= 0x1F1FF or     # Regional indicator symbols (flag emojis 🇯🇲 🇬🇧 etc)
        0x1F300 <= cp <= 0x1FAFF or     # All modern emoji blocks (faces, objects, flags …)
        0xFE00 <= cp <= 0xFE0F or       # Variation selectors (emoji vs text presentation)
        cp == 0x200D                    # ZWJ
    )


def _line_has_emoji(line):
    return any(_is_emoji_char(c) for c in line)


def _emoji_source():
    try:
        from pilmoji.source import GoogleEmojiSource
        return GoogleEmojiSource
    except Exception:
        return None


def _pilmoji_measure(line, font):
    """Return (width, height) for a line that may contain emoji, using pilmoji."""
    try:
        from pilmoji import Pilmoji
        src = _emoji_source()
        tmp = Image.new('RGBA', (1, 1))
        kwargs = {'source': src} if src else {}
        with Pilmoji(tmp, **kwargs) as p:
            return p.getsize(line, font=font)
    except Exception:
        pass
    bb = ImageDraw.Draw(Image.new('RGBA', (1, 1))).textbbox((0, 0), line, font=font)
    return bb[2] - bb[0], bb[3] - bb[1]


def _pilmoji_draw(img, x, y, line, font, fill):
    """Draw a text line with emoji onto img using pilmoji + Google Noto emoji."""
    try:
        from pilmoji import Pilmoji
        src = _emoji_source()
        kwargs = {'source': src} if src else {}
        with Pilmoji(img, **kwargs) as p:
            p.text((x, y), line, font=font, fill=fill)
    except Exception:
        ImageDraw.Draw(img).text((x, y), line, font=font, fill=fill)


def _render_line_to_strip(line, font, colour, use_emoji, pad=200):
    """Render a single text line to its own RGBA image cropped to actual pixel bounds.

    Uses a generous padding so that font bearing / overhang is never clipped.
    Returns a PIL Image (RGBA) containing only the inked pixels.
    """
    r, g, b = colour
    if use_emoji:
        mw, mh   = _pilmoji_measure(line or 'Ag', font)
        est_w    = max(mw, 10) + pad * 2
        est_h    = max(mh, 10) + pad * 2
    else:
        scratch  = Image.new('RGBA', (1, 1))
        bb       = ImageDraw.Draw(scratch).textbbox((0, 0), line or 'Ag', font=font)
        est_w    = max(bb[2] - bb[0], 10) + pad * 2
        est_h    = max(bb[3] - bb[1], 10) + pad * 2
    strip    = Image.new('RGBA', (max(1, est_w), max(1, est_h)), (0, 0, 0, 0))
    if use_emoji:
        _pilmoji_draw(strip, pad, pad, line, font, (r, g, b, 255))
    else:
        ImageDraw.Draw(strip).text((pad, pad), line, font=font, fill=(r, g, b, 255))
    bbox = strip.getbbox()
    if bbox:
        return strip.crop(bbox)
    # blank line — return a transparent sliver with the right height
    return Image.new('RGBA', (1, max(est_h - pad * 2, 10)), (0, 0, 0, 0))


def _render_text(lines, font_name, colour_hex, w, font_size):
    """Render text lines at a specific font size, center-aligned, returns (img, top=0, left)."""
    r, g, b   = hex_to_rgb(colour_hex)
    font      = get_font(font_name, font_size)
    use_emoji = any(_line_has_emoji(l) for l in lines)

    if use_emoji:
        # Emoji: render each line to its own strip, crop to actual pixels, center-paste.
        # pilmoji doesn't support multiline align so we do it manually.
        line_imgs = [_render_line_to_strip(l, font, (r, g, b), use_emoji) for l in lines]
        line_ws   = [img.width  for img in line_imgs]
        line_hs   = [img.height for img in line_imgs]
        max_lw    = max(line_ws) if line_ws else 1
        line_h    = int(max(line_hs) * 1.4)
        margin    = max(10, line_h // 6)
        img_w     = max_lw + 2 * margin
        img_h     = line_h * len(lines) + 2 * margin
        img       = Image.new('RGBA', (max(1, img_w), max(1, img_h)), (0, 0, 0, 0))
        yl        = margin
        for limg, lw in zip(line_imgs, line_ws):
            img.paste(limg, (margin + (max_lw - lw) // 2, yl), limg)
            yl += line_h
    else:
        # Non-emoji: use PIL's built-in multiline_text with align='center'.
        # This is the most reliable way to center lines relative to each other.
        text      = "\n".join(lines)
        scratch   = Image.new('RGBA', (1, 1))
        d_s       = ImageDraw.Draw(scratch)
        spacing   = 0   # PIL's natural line height; matches find_best_font_size
        ref_h     = d_s.textbbox((0, 0), lines[0] or 'Ag', font=font)
        margin    = max(10, (ref_h[3] - ref_h[1]) // 6)
        bb        = d_s.multiline_textbbox((0, 0), text, font=font,
                                           spacing=spacing, align='center')
        img_w     = int(bb[2] - bb[0]) + margin * 2
        img_h     = int(bb[3] - bb[1]) + margin * 2
        img       = Image.new('RGBA', (max(1, img_w), max(1, img_h)), (0, 0, 0, 0))
        ImageDraw.Draw(img).multiline_text(
            (-bb[0] + margin, -bb[1] + margin),
            text, font=font, fill=(r, g, b, 255),
            spacing=0, align='center'
        )

    # Safety clamp: if rendered text wider than canvas, scale down to fit
    if img.width > w * 0.95:
        scale = (w * 0.92) / img.width
        img   = img.resize((int(img.width * scale), int(img.height * scale)), Image.LANCZOS)
        img_w = img.width

    left = (w - img_w) // 2
    return img, 0, left

def build_label_layer(text):
    """Small black label rendered at top of each zone."""
    size = max(18, cm(0.4))
    try:
        f = ImageFont.truetype('arial.ttf', size)
    except Exception:
        f = ImageFont.load_default()
    tmp  = Image.new('RGBA', (1, 1))
    bb   = ImageDraw.Draw(tmp).textbbox((0, 0), text.upper(), font=f)
    tw   = bb[2] - bb[0] + 10
    th   = bb[3] - bb[1] + 6
    img  = Image.new('RGBA', (max(1, tw), max(1, th)), (0, 0, 0, 0))
    ImageDraw.Draw(img).text((5, 3), text.upper(), font=f, fill=(0, 0, 0, 255))
    return img

# ─── LOGGING ──────────────────────────────────────────────────────────────────

def log(msg, level='INFO'):
    ts   = datetime.now().strftime('%H:%M:%S')
    line = f'[{ts}] [{level}] {msg}'
    print(line)
    try:
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception:
        pass

# ─── MAIN PROCESSOR ───────────────────────────────────────────────────────────

ZONES = ['front', 'back', 'pocket', 'sleeve']

def active_zones_for_row(row, images):
    """
    Determine which zones to render for a row.
    Strategy (in priority order):
      1. PrintLocation text keywords  (most reliable in this export)
      2. Image/preview files present in the image folder
      3. Text content in the zone fields
    IsFrontLocation/IsBackLocation/etc. are ignored — they are all False
    in the unshipped-orders Excel export.
    """
    loc = str(row.get('PrintLocation', '') or '').lower()

    # Map keywords -> zones
    # "Front Pocket" means the chest pocket zone, NOT the full-front zone.
    # We must check for this phrase before checking for bare 'front'.
    from_loc: set[str] = set()
    if 'front pocket' in loc:
        from_loc.add('pocket')   # "Front Pocket + Back" → pocket + back, no front
    else:
        if any(k in loc for k in ('front', 'name', 'upload', 'enter', 'text', 'photo')):
            from_loc.add('front')
        if 'pocket' in loc:
            from_loc.add('pocket')
    if 'back' in loc:
        from_loc.add('back')
    if 'sleeve' in loc:
        from_loc.add('sleeve')

    # Build data-presence set (actual customer image or text only — NOT preview)
    # Preview images are reference-only and must not trigger zone creation.
    from_data: set[str] = set()
    for zone in ZONES:
        if zone in images or bool(parse_text_lines(row, zone)):
            from_data.add(zone)

    # Union: include a zone if signalled by either source
    active = from_loc | from_data

    # If nothing detected at all, default to front (avoids empty PSD)
    return active if active else {'front'}

# ─── FOLDER STRUCTURE HELPERS (match batch_processor.py) ─────────────────────

def sku_colour_folder(sku):
    """Return 'black', 'white', or '' (no sub-folder) from SKU colour code."""
    s = (sku or '').lower()
    if 'blk' in s: return 'black'
    if 'wht' in s: return 'white'
    return ''

def is_multizone_rows(rows, images_by_item):
    """True if any row in the group has content in more than one print zone."""
    for row in rows:
        item_id = _safe(row.get('OrderItemID')) or ''
        images  = images_by_item.get(item_id, {})
        active  = active_zones_for_row(row, images)
        if len(active) > 1:
            return True
    return False

def is_emb_rhine_rows(rows):
    """True if any zone font contains 'emb' or 'rhine' (handled manually)."""
    for row in rows:
        for field in ('FrontFonts', 'BackFonts', 'PocketFonts', 'SleeveFonts'):
            val = (_safe(row.get(field)) or '').lower()
            if 'emb' in val or 'rhine' in val:
                return True
    return False

def is_kids_hood_rows(rows):
    """True if any SKU in the group is a kids hoodie."""
    for row in rows:
        sku = (_safe(row.get('SKU')) or '').lower()
        if any(k in sku for k in ('kidshood', 'kidshoo', 'gymhoodie', 'gymnastichoodie')):
            return True
    return False

def make_folder_type(rows, images_by_item):
    """Return the DTF folder type matching batch_processor.py logic."""
    if is_emb_rhine_rows(rows):
        return 'Emb & Rhine'
    if is_multizone_rows(rows, images_by_item):
        return 'Automated'
    if is_kids_hood_rows(rows):
        return 'DTF Kids Hoodie'
    return 'DTF Front'

def process_excel_orders(limit=None, dry_run=False, order_id=None):
    try:
        import pandas as pd
    except ImportError:
        log('pandas not installed. Run: pip install pandas openpyxl', 'ERROR')
        return

    if not os.path.isfile(EXCEL_FILE):
        log(f'Excel file not found: {EXCEL_FILE}', 'ERROR')
        return

    if not os.path.isdir(IMAGE_FOLDER):
        log(f'Image folder not found: {IMAGE_FOLDER}', 'WARN')

    # ── Read both sheets ────────────────────────────────────────────────────
    log(f'Reading Excel: {EXCEL_FILE}')
    df_orders  = pd.read_excel(EXCEL_FILE, sheet_name='DTF Orders')
    df_details = pd.read_excel(EXCEL_FILE, sheet_name='DTF Order Details')
    log(f'  Orders sheet  : {len(df_orders)} rows')
    log(f'  Details sheet : {len(df_details)} rows')

    # Join Details -> Orders on idCustomOrder
    keep_cols = ['idCustomOrder', 'OrderID', 'OrderItemID', 'SKU', 'Quantity']
    df = df_details.merge(df_orders[keep_cols], on='idCustomOrder', how='left')
    log(f'  Joined rows   : {len(df)}')

    # ── Group by OrderID ────────────────────────────────────────────────────
    from collections import OrderedDict
    groups = OrderedDict()
    for _, row in df.iterrows():
        oid = _safe(row.get('OrderID')) or 'UNKNOWN'
        if oid not in groups:
            groups[oid] = []
        groups[oid].append(row)

    if order_id:
        ids = order_id if isinstance(order_id, list) else [order_id]
        groups = {k: v for k, v in groups.items() if k in ids}
        log(f'Filter        : {len(ids)} order(s): {", ".join(ids)}')

    total_orders = len(groups)
    log(f'Unique orders : {total_orders}')
    if limit:
        log(f'Limit applied : processing first {limit}')
    if dry_run:
        log('MODE          : DRY RUN — no files written')
    log('=' * 60)

    today   = datetime.now().strftime('%Y-%m-%d')
    out_dir = os.path.join(OUTPUT_FOLDER, today)
    os.makedirs(out_dir, exist_ok=True)
    ok_count, fail_count, skip_count = 0, 0, 0

    # Pre-build image lookup for all items (used by make_folder_type)
    all_item_ids = [_safe(row.get('OrderItemID')) or '' for _, rows in groups.items() for row in rows]
    images_by_item = {iid: find_images_for_item(iid) for iid in set(all_item_ids) if iid}

    for i, (order_id, rows) in enumerate(groups.items(), 1):
        if limit and i > limit:
            break

        first    = rows[0]
        sku_raw  = _safe(first.get('SKU')) or 'UNKNOWN'
        safe_id  = order_id.replace('/', '-')

        # Skip orders with pre-made designs
        if order_id in SKIP_ORDER_IDS:
            log(f'[{i}/{total_orders}] SKIP (pre-made design): {order_id}')
            skip_count += 1
            continue
        if any(sku_raw.startswith(p) for p in SKIP_SKU_PREFIXES):
            log(f'[{i}/{total_orders}] SKIP (pre-made SKU: {sku_raw}): {order_id}')
            skip_count += 1
            continue
        safe_sku = sku_raw.replace('/', '-').replace('\\', '-')
        skus_str = ' | '.join(_safe(r.get('SKU')) or '?' for r in rows)

        log(f'[{i}/{total_orders}] {order_id}  |  {skus_str}')

        # ── Build all layers ────────────────────────────────────────────────
        PADDING = cm(1)
        GAP     = cm(0.5)
        all_layers = []

        # Each row corresponds to one item (one SKU / OrderItemID)
        row_data = []
        for row in rows:
            item_id   = _safe(row.get('OrderItemID')) or ''
            row_sku   = _safe(row.get('SKU')) or sku_raw
            row_prod  = detect_product(row_sku)
            row_canvas= PRODUCT_CANVAS.get(row_prod, PRODUCT_CANVAS['default'])
            images    = find_images_for_item(item_id) if item_id else {}

            # Multiple SKUs -> include colour/size in labels
            multi_sku = len(rows) > 1

            zones_for_row = []
            active = active_zones_for_row(row, images)
            for zone_key in ZONES:
                if zone_key not in active:
                    continue

                dims      = row_canvas.get(zone_key, row_canvas.get('front', (cm(30), cm(30))))
                zw, zh    = dims
                img_path  = images.get(zone_key)
                prev_path = images.get(f'{zone_key}preview')
                txt_lines = parse_text_lines(row, zone_key)
                font_name, _ = parse_font_name(row, zone_key)
                colour_hex   = parse_colour_hex(row, zone_key)

                # Label text
                if multi_sku:
                    from batch_processor import parse_sku_colour_size
                    try:
                        colour_str, size_str = parse_sku_colour_size(row_sku)
                        parts = [zone_key.title()]
                        if colour_str: parts.append(colour_str)
                        if size_str:   parts.append(size_str)
                        label_txt = ' - '.join(parts)
                    except Exception:
                        label_txt = f'{zone_key.title()} - {row_sku}'
                else:
                    label_txt = zone_key.title()

                do_bg_remove = bool(img_path) and should_remove_bg(img_path, row_sku, row, zone_key)

                zones_for_row.append({
                    'label':        label_txt,
                    'zone_key':     zone_key,
                    'zw': zw, 'zh': zh,
                    'img_path':     img_path,
                    'prev_path':    prev_path,
                    'txt_lines':    txt_lines,
                    'font_name':    font_name,
                    'colour':       colour_hex,
                    'do_bg_remove': do_bg_remove,
                })

            if not zones_for_row:
                log(f'  WARNING: no active zones for item {item_id}', 'WARN')
            # Stack copies for Quantity > 1
            qty = int(row.get('Quantity', 1) or 1)
            for _ in range(qty):
                row_data.append(zones_for_row)

        flat_zones = [z for zones in row_data for z in zones]
        if not flat_zones:
            log(f'  SKIP — no printable zones found', 'WARN')
            skip_count += 1
            continue

        # ── For multi-item orders: unify font size PER ZONE TYPE ────────────────
        # Compute best font size per zone first, then apply the minimum within
        # each zone type so items look consistent (e.g. "Big Bro" / "Little Bro"
        # same size on front, and same size on back — but front and back are
        # sized independently, since back often has more lines and would
        # otherwise shrink the front text unnecessarily).
        LABEL_H = build_label_layer('Front').height + cm(0.3)
        multi_item = len(rows) > 1
        if multi_item:
            sizes_by_zone_key = {}   # (zone_key, zw, zh) -> [font_sizes]
            for zones in row_data:
                for zone in zones:
                    if zone['txt_lines']:
                        wrapped = auto_wrap_lines(zone['txt_lines'], zone['font_name'],
                                                  zone['zw'], zone['zh'])
                        sz = find_best_font_size(wrapped, zone['font_name'],
                                                 zone['zw'], zone['zh'])
                        key = (zone['zone_key'], zone['zw'], zone['zh'])
                        sizes_by_zone_key.setdefault(key, []).append(sz)
            unified_font_sizes = {k: min(szs) for k, szs in sizes_by_zone_key.items()}
        else:
            unified_font_sizes = {}

        # ── Render layers first so we know actual content heights ────────────
        rendered_rows = []
        for zones in row_data:
            rendered_zones_list = []
            for zone in zones:
                zw, zh = zone['zw'], zone['zh']
                zone_layers = []   # (name, pil, top_offset, left_offset, visible)
                content_h   = 0   # actual rendered height (not fixed zh)

                if zone['img_path']:
                    img_pil, it, il = build_image_layer(zone['img_path'], zw, zh, zone.get('do_bg_remove', False))
                    zone_layers.append((f"{zone['label']} CustomerImage", img_pil, it, il, True))
                    content_h = max(content_h, it + img_pil.height)

                if zone['txt_lines']:
                    txt_pil, tt, tl = build_text_layer(
                        zone['txt_lines'], zone['font_name'], zone['colour'], zw, zh,
                        force_size=unified_font_sizes.get((zone['zone_key'], zone['zw'], zone['zh'])))
                    zone_layers.append((f"{zone['label']} CustomerText", txt_pil, tt, tl, True))
                    content_h = max(content_h, tt + txt_pil.height)

                if zone['prev_path']:
                    prev_pil, pt, pl = build_image_layer(zone['prev_path'], zw, zh, upscale=False)
                    zone_layers.append((f"{zone['label']} Preview Reference", prev_pil, pt, pl, False))
                    # preview is hidden — don't include in content_h

                # If no content rendered, fall back to zone height
                if content_h == 0:
                    content_h = zh

                rendered_zones_list.append({**zone, 'zone_layers': zone_layers, 'content_h': content_h})
            rendered_rows.append(rendered_zones_list)

        # ── Calculate canvas from actual content heights ─────────────────────
        max_zw   = max(z['zw'] for zones in rendered_rows for z in zones)
        canvas_w = PADDING + max_zw + PADDING
        canvas_h = PADDING  # top padding only
        for zones in rendered_rows:
            for z in zones:
                canvas_h += LABEL_H + z['content_h'] + GAP
        canvas_h -= GAP        # remove trailing gap after last zone
        canvas_h += cm(1.5)    # fixed 1.5 cm bottom margin
        canvas_h = max(1, int(canvas_h))

        # ── Place layers ────────────────────────────────────────────────────
        # Labels go into a separate list so we can append them LAST,
        # making them the topmost layers in Photoshop (always visible on top).
        label_layers = []
        y = PADDING
        for zones in rendered_rows:
            for zone in zones:
                zw         = zone['zw']
                content_h  = zone['content_h']
                x_centre   = PADDING + (max_zw - zw) // 2

                # Zone label — collected separately, added on top at the end
                lbl     = build_label_layer(zone['label'])
                lbl_top = y
                label_layers.append({'name': f"{zone['label']} Label",
                                     'image': lbl, 'top': lbl_top, 'left': x_centre,
                                     'opacity': 255, 'visible': True})
                img_top = y + LABEL_H

                for (lname, lpil, lt, ll, lvis) in zone['zone_layers']:
                    all_layers.append({'name': lname, 'image': lpil,
                                       'top': img_top + lt, 'left': x_centre + ll,
                                       'opacity': 255, 'visible': lvis})

                y += LABEL_H + content_h + GAP

        # Labels on top so they're never buried under content layers
        all_layers.extend(label_layers)

        # ── Write PSD ───────────────────────────────────────────────────────
        folder_type = make_folder_type(rows, images_by_item)
        colour_sub  = sku_colour_folder(sku_raw)
        if colour_sub:
            cat_dir = os.path.join(out_dir, folder_type, colour_sub)
        else:
            cat_dir = os.path.join(out_dir, folder_type)
        os.makedirs(cat_dir, exist_ok=True)
        suffix   = f'{safe_sku}_{len(rows)}items' if len(rows) > 1 else safe_sku
        out_path = os.path.join(cat_dir, f'{safe_id}_{suffix}.psd')

        if dry_run:
            log(f'  DRY-RUN -> {out_path}  ({len(all_layers)} layers)', 'OK')
            ok_count += 1
            continue

        try:
            out_path = write_psd(out_path, canvas_w, canvas_h, all_layers)
            size_mb = os.path.getsize(out_path) / 1024 / 1024
            log(f'  OK  {size_mb:.1f} MB  -> {out_path}', 'OK')
            ok_count += 1
        except Exception as e:
            log(f'  FAIL  {e}', 'FAIL')
            log(traceback.format_exc()[-500:], 'ERROR')
            fail_count += 1

    log('=' * 60)
    log(f'DONE  {ok_count} OK  |  {fail_count} FAIL  |  {skip_count} SKIP')
    log(f'Output : {OUTPUT_FOLDER}')
    log('=' * 60)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='DTF Excel Order Processor')
    parser.add_argument('--limit',   type=int, default=None, help='Max orders to process')
    parser.add_argument('--dry-run', action='store_true',    help='Preview only, no files written')
    parser.add_argument('--order',   type=str, default=None, action='append', help='Process specific OrderID(s) — can be repeated')
    parser.add_argument('--excel',   type=str, default=None, help='Path to Excel file (overrides EXCEL_FILE)')
    parser.add_argument('--images',  type=str, default=None, help='Path to images folder (overrides IMAGE_FOLDER)')
    parser.add_argument('--dpi',     type=int, default=None, help='Resolution px/cm (120=304dpi, 320=812dpi)')
    args = parser.parse_args()
    if args.excel:
        EXCEL_FILE = args.excel
    if args.images:
        IMAGE_FOLDER = args.images
    if args.dpi:
        PX_PER_CM = args.dpi
    process_excel_orders(limit=args.limit, dry_run=args.dry_run, order_id=args.order or None)
