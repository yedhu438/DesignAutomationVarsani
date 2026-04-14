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

import os, json, struct, io, traceback, argparse
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont

# ─── CONFIG ───────────────────────────────────────────────────────────────────

EXCEL_FILE    = r"W:\test2\UnshippedDTFOrders_14042026_013121.xlsx"
IMAGE_FOLDER  = r"W:\test2\DTFUnshippedImages_20260414_013102"
OUTPUT_FOLDER = r"C:\Varsany\Output\DTF_Excel"
FONT_FOLDERS  = [r"C:\Varsany\Fonts", r"W:\fonts"]
LOG_FILE      = r"C:\Varsany\Output\DTF_Excel\dtf_excel_log.txt"

PX_PER_CM = 120          # 304 DPI for test — change to 320 for production
DPI       = int(PX_PER_CM * 2.54)

# ─── PRODUCT CANVAS SIZES ─────────────────────────────────────────────────────
def cm(x): return int(round(x * PX_PER_CM))

PRODUCT_CANVAS = {
    "adulttshirt":  {"front": (cm(30), cm(30)), "back": (cm(30), cm(30)),  "pocket": (cm(9), cm(7))},
    "kidstshirt":   {"front": (cm(23), cm(30)), "back": (cm(23), cm(30)),  "pocket": (cm(9), cm(7))},
    "adulthoodie":  {"front": (cm(25), cm(25)), "back": (cm(25), cm(25)),  "pocket": (cm(9), cm(7))},
    "kidshoodie":   {"front": (cm(23), cm(20)), "back": (cm(23), cm(20)),  "pocket": (cm(9), cm(7))},
    "totebag":      {"front": (cm(28), cm(28)), "back": (cm(28), cm(58))},
    "babyvest":     {"front": (cm(15), cm(17))},
    "buckethat":    {"front": (cm(18), cm(5))},
    "beanie":       {"front": (cm(9.5), cm(4.5))},
    "pegbag":       {"front": (cm(23), cm(14))},
    "cushion":      {"front": (cm(30), cm(30))},
    "slipper":      {"front": (cm(6),  cm(6))},
    "default":      {"front": (cm(30), cm(30)), "back": (cm(30), cm(30)), "pocket": (cm(9), cm(7))},
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
    ("AnyTxtTote_",          "totebag"),     ("Knitting",            "totebag"),
    ("AnyTxtBabyVest_",      "babyvest"),
    ("AnyTextHat_",          "buckethat"),
    ("AnytxtPatchBeanie_",   "beanie"),
    ("AnyTxtPEBag_",         "pegbag"),
    ("CustomCushion",        "cushion"),
    ("Memorial_",            "default"),
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
    if "tote" in s or "knit" in s:             return "Tote Bag"
    if "babyvest" in s or "vest" in s:         return "Baby Vest"
    if "pegbag" in s or "pebag" in s:          return "PE Bag"
    if "cushion" in s:                         return "Cushion"
    if "beanie" in s:                          return "Hat"
    if "dogtee" in s:                          return "Dog Tee"
    if "memorial" in s:                        return "Memorial"
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
}

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
                lines = [str(data[k]).strip()
                         for k in sorted(data.keys())
                         if k.lower().startswith('text') and _safe(data.get(k))]
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
    'wht': (255, 255, 255), 'blk': (0,   0,   0  ),
    'nvy': (31,  40,  80 ), 'red': (200, 30,  30 ),
    'ylw': (255, 220, 0  ), 'pnk': (255, 150, 180),
    'gry': (150, 150, 150), 'grn': (30,  130, 30 ),
    'org': (230, 100, 20 ), 'blu': (30,  80,  200),
    'rblu':(50,  100, 220), 'sblu':(100, 150, 210),
    'nat': (240, 230, 200), 'ivry':(240, 230, 200),
    'bur': (130, 30,  50 ), 'fus': (200, 50,  140),
    'ppl': (120, 50,  170),
}

def _garment_colour(sku):
    """Return approximate garment RGB from SKU, or None if unknown."""
    s = sku.lower()
    for key, rgb in SKU_COLOUR_MAP.items():
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

    # 2. Auto-detect: solid uniform background on all edges → remove
    try:
        import numpy as np
        img = Image.open(img_path).convert('RGB')
        arr = np.array(img)
        edges = [
            arr[0,  :,  :].mean(axis=0),   # top row
            arr[-1, :,  :].mean(axis=0),   # bottom row
            arr[:,  0,  :].mean(axis=0),   # left col
            arr[:, -1,  :].mean(axis=0),   # right col
        ]
        # Check all 4 edges agree on the same colour
        for i in range(1, 4):
            if not _colour_close(tuple(edges[0].astype(int)),
                                  tuple(edges[i].astype(int)), tol=25):
                return False
        # Edges are uniform — check it's not a gradient or content image
        bg_colour = tuple(edges[0].astype(int))
        garment_rgb = _garment_colour(sku)
        # Remove if: (a) bg matches garment colour, OR (b) bg is any solid
        # colour (black, white, grey etc.) that would appear as a rectangle
        is_solid_neutral = all(abs(bg_colour[0] - bg_colour[c]) < 20
                               for c in range(1, 3))   # r≈g≈b → grey/neutral
        if garment_rgb and _colour_close(bg_colour, garment_rgb, tol=40):
            return True
        if is_solid_neutral:
            return True
        return False
    except Exception:
        return False

def remove_background(img_path):
    """Run rembg on img_path, return PIL RGBA image with background removed."""
    from rembg import remove as rembg_remove
    with open(img_path, 'rb') as f:
        data = f.read()
    result = rembg_remove(data)
    return Image.open(io.BytesIO(result)).convert('RGBA')

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
    """Split image into raw channel bytes (no compression — matches PSD compression type 0)."""
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
    """Write a layered RGB PSD file. Layers: list of dicts with image/top/left/name/opacity/visible."""
    buf = io.BytesIO()
    p   = buf.write

    p(b'8BPS')
    p(struct.pack('>H', 1))       # version 1 = PSD
    p(b'\x00' * 6)
    p(struct.pack('>H', 3))       # 3 channels (RGB)
    p(struct.pack('>II', canvas_h, canvas_w))
    p(struct.pack('>HH', 8, 3))   # 8 bpc, RGB

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

        ch = _to_channels(img, 'RGBA')
        ch_order = [-1, 0, 1, 2]

        lr = io.BytesIO()
        lr.write(struct.pack('>iiii', top, left, bottom, right))
        lr.write(struct.pack('>H', 4))
        for cid in ch_order:
            lr.write(struct.pack('>hI', cid, len(ch[cid]) + 2))
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
            ld_buf.write(struct.pack('>H', 0))   # compression = deflate raw
            ld_buf.write(ch[cid])

    layer_info = struct.pack('>h', len(layers)) + lr_buf.getvalue() + ld_buf.getvalue()
    if len(layer_info) % 4:
        layer_info += b'\x00' * (4 - len(layer_info) % 4)
    lmi = struct.pack('>I', len(layer_info)) + layer_info + struct.pack('>I', 0)
    p(struct.pack('>I', len(lmi)))
    p(lmi)

    # Composite (merged) image
    composite = Image.new('RGB', (canvas_w, canvas_h), (255, 255, 255))
    for lyr in layers:
        if lyr.get('visible', True):
            composite.paste(lyr['image'].convert('RGBA'),
                            (lyr['left'], lyr['top']),
                            lyr['image'].convert('RGBA'))
    comp = _to_channels(composite, 'RGB')
    p(struct.pack('>H', 0))
    for cid in [0, 1, 2]:
        p(comp[cid])

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())

# ─── LAYER BUILDERS ───────────────────────────────────────────────────────────

def build_image_layer(img_path, w, h, do_bg_remove=False):
    """Scale image to fill zone width, top-aligned. Optionally removes background."""
    if not img_path or not os.path.isfile(img_path):
        return Image.new('RGBA', (max(1, w), max(1, h)), (0, 0, 0, 0)), 0, 0
    if do_bg_remove:
        try:
            src = remove_background(img_path)
            log(f'    BG removed: {os.path.basename(img_path)}')
        except Exception as e:
            log(f'    BG removal failed ({e}), using original', 'WARN')
            src = Image.open(img_path).convert('RGBA')
    else:
        src = Image.open(img_path).convert('RGBA')
    ratio = w / src.width
    nw    = w
    nh    = max(1, int(src.height * ratio))
    src   = src.resize((nw, nh), Image.LANCZOS)
    return src, 0, 0

def build_text_layer(text_lines, font_name, colour_hex, w, h, force_size=None):
    """Auto-size text to fill zone width, centre each line."""
    lines = [l for l in text_lines if l.strip()]
    if not lines:
        return Image.new('RGBA', (1, 1), (0, 0, 0, 0)), 0, 0

    best = find_best_font_size(text_lines, font_name, w, h, force_size)
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
    while lo <= hi:
        mid  = (lo + hi) // 2
        font = get_font(font_name, mid)
        widths = [draw.textbbox((0, 0), l, font=font)[2] - draw.textbbox((0, 0), l, font=font)[0]
                  for l in lines]
        bb0    = draw.textbbox((0, 0), lines[0], font=font)
        line_h = int((bb0[3] - bb0[1]) * 1.4)
        total_h = line_h * len(lines)
        if max(widths) <= avail_w and total_h <= avail_h:
            best, lo, hi = mid, mid + 1, hi
        else:
            lo, hi = lo, mid - 1
    return best


def _render_text(lines, font_name, colour_hex, w, font_size):
    """Render text lines at a specific font size, center-aligned, returns (img, top=0, left).

    Each line is centered relative to the widest line. The image is sized
    to the widest line (+ small margin) and then centered in the zone width w.
    This guarantees every line stays centered regardless of text length.
    """
    r, g, b = hex_to_rgb(colour_hex)
    font    = get_font(font_name, font_size)
    scratch = Image.new('RGBA', (1, 1))
    draw    = ImageDraw.Draw(scratch)

    # Measure all lines
    line_bbs = [draw.textbbox((0, 0), l, font=font) for l in lines]
    line_ws  = [bb[2] - bb[0] for bb in line_bbs]
    line_hs  = [bb[3] - bb[1] for bb in line_bbs]
    line_h   = int(max(line_hs) * 1.4)        # uniform line height
    max_lw   = max(line_ws) if line_ws else 1
    margin   = max(10, line_h // 6)

    # Image exactly fits the text block (no excess padding)
    img_w = max_lw + 2 * margin
    img_h = line_h * len(lines) + 2 * margin
    img   = Image.new('RGBA', (max(1, img_w), max(1, img_h)), (0, 0, 0, 0))
    d     = ImageDraw.Draw(img)
    yl    = margin
    for lw, line in zip(line_ws, lines):
        x = margin + (max_lw - lw) // 2   # center line within max_lw
        d.text((x, yl), line, font=font, fill=(r, g, b, 255))
        yl += line_h

    # Center the image in the zone width
    left = (w - img_w) // 2   # may be negative only if font overflows (prevented by binary search)
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

    # Build data-presence set (image file or non-empty text)
    from_data: set[str] = set()
    for zone in ZONES:
        if (zone in images
                or f'{zone}preview' in images
                or bool(parse_text_lines(row, zone))):
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
        groups = {k: v for k, v in groups.items() if k == order_id}
        log(f'Filter        : order {order_id} only')

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
            row_data.append(zones_for_row)

        flat_zones = [z for zones in row_data for z in zones]
        if not flat_zones:
            log(f'  SKIP — no printable zones found', 'WARN')
            skip_count += 1
            continue

        # ── For multi-item orders: unify font size across all text zones ────────
        # Compute best font size per zone first, then apply the minimum so all
        # items look consistent (e.g. "Big Bro" / "Little Bro" same size).
        LABEL_H = build_label_layer('Front').height + cm(0.3)
        multi_item = len(rows) > 1
        if multi_item:
            txt_font_sizes = []
            for zones in row_data:
                for zone in zones:
                    if zone['txt_lines']:
                        sz = find_best_font_size(zone['txt_lines'], zone['font_name'],
                                                 zone['zw'], zone['zh'])
                        txt_font_sizes.append(sz)
            unified_font_size = min(txt_font_sizes) if txt_font_sizes else None
        else:
            unified_font_size = None

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
                        force_size=unified_font_size)
                    zone_layers.append((f"{zone['label']} CustomerText", txt_pil, tt, tl, True))
                    content_h = max(content_h, tt + txt_pil.height)

                if zone['prev_path']:
                    prev_pil, pt, pl = build_image_layer(zone['prev_path'], zw, zh)
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
            write_psd(out_path, canvas_w, canvas_h, all_layers)
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
    parser.add_argument('--order',   type=str, default=None, help='Process a single order ID only')
    args = parser.parse_args()
    process_excel_orders(limit=args.limit, dry_run=args.dry_run, order_id=args.order)
