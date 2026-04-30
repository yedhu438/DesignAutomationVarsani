"""
Varsany Batch Processor
========================
Processes all unprocessed orders from the database.
Generates layered PSD files for each order.

Images : W:\\images\\Jan-Image\\ and W:\\images\\Feb-Image\\
Fonts  : C:\\Varsany\\Fonts\\ + W:\\fonts\\ + system fonts
Output : C:\\Varsany\\Output\\YYYY-MM-DD\\OrderID.psd

Usage:
    python batch_processor.py                  # all unprocessed orders
    python batch_processor.py --limit 10       # first 10 only (test)
    python batch_processor.py --order 203-xxx  # one specific order
    python batch_processor.py --dry-run        # preview, no files written
    python batch_processor.py --dpi 320        # full print resolution
"""

import os, json, struct, io, argparse, traceback, urllib.request, tempfile
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import pyodbc

try:
    from rembg import remove as rembg_remove
    REMBG_AVAILABLE = True
except ImportError:
    REMBG_AVAILABLE = False

# Editable text layer writer (TySh + Txt2 PSD blocks)
try:
    from psd_text_layer import build_editable_text_tagged_blocks, resolve_ps_font_name
    EDITABLE_TEXT_AVAILABLE = True
except ImportError:
    EDITABLE_TEXT_AVAILABLE = False

# ─── CONFIG ───────────────────────────────────────────────────────────────────

DB_CONNECTION = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost\\SQLEXPRESS;"
    "DATABASE=dbAmazonCustomOrders;"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)

IMAGE_FOLDERS = [
    r"W:\images\Feb-Image",
    r"W:\images\Jan-Image",
    r"C:\Varsany\Uploads",
]

# Auto-discover W:\test*\DTFUnshippedImages_* bulk download folders
_W = r"W:\\"
if os.path.exists(_W):
    for _entry in os.listdir(_W):
        if _entry.lower().startswith("test"):
            _test_dir = os.path.join(_W, _entry)
            if os.path.isdir(_test_dir):
                for _sub in os.listdir(_test_dir):
                    if _sub.startswith("DTFUnshippedImages"):
                        IMAGE_FOLDERS.append(os.path.join(_test_dir, _sub))

FONT_FOLDERS = [
    r"C:\Varsany\Fonts",
    r"W:\fonts",
]

OUTPUT_FOLDER = r"C:\Varsany\Output"
LOG_FILE      = r"C:\Varsany\batch_log.txt"

PX_PER_CM = 120
DPI        = int(PX_PER_CM * 2.54)

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ─── BUILD IMAGE INDEX ────────────────────────────────────────────────────────

print("Building image index...")
IMAGE_INDEX = {}
for _folder in IMAGE_FOLDERS:
    if os.path.exists(_folder):
        for _f in os.listdir(_folder):
            IMAGE_INDEX[_f.lower()] = os.path.join(_folder, _f)
print(f"  Indexed {len(IMAGE_INDEX):,} images")

# ─── BUILD FONT INDEX ─────────────────────────────────────────────────────────
# Maps normalised font name -> full file path
# e.g. "bebasneuepro" -> "C:\\Varsany\\Fonts\\BebasNeue-Regular.ttf"

FONT_INDEX = {}
for _folder in FONT_FOLDERS:
    if os.path.exists(_folder):
        for _f in os.listdir(_folder):
            if _f.lower().endswith(('.ttf', '.otf')):
                _norm = os.path.splitext(_f)[0].lower()
                _norm = _norm.replace(' ','').replace('-','').replace('_','')
                FONT_INDEX[_norm] = os.path.join(_folder, _f)

# Explicit aliases so database font names map to actual filenames
FONT_ALIASES = {
    # Standard fonts
    "arial":            "arial",
    "arialbold":        "arial",
    "bebasneuepro":     "bebasneueregular",
    "bebasneuefree":    "bebasneueregular",
    "chewy":            "chewyregular",
    "fondamento":       "fondamentoregular",
    "permanentmarker":  "permanentmarkerregular",
    "russoone":         "russooneregular",
    "ultra":            "ultraregular",
    "lato":             "latoregular",
    "latobold":         "latobold",
    "roboto":           "roboto",
    "verdana":          "verdana",
    # Premium texture fonts (database name → FONT_INDEX key)
    "texturefont":      "smartkids",
    "texture font":     "smartkids",
    "texture":          "smartkids",
    "blockfont":        "colorfulblocks",
    "block font":       "colorfulblocks",
    "colorfulblock":    "colorfulblocks",
    "paintfont":        "paintsplashesrainbow",
    "paint font":       "paintsplashesrainbow",
    "paintsplashes":    "paintsplashesrainbow",
    "mermaidfont":      "wavemermaid",
    "mermaid font":     "wavemermaid",
    "mermaid":          "wavemermaid",
    "reflectionfont":   "refractionray",
    "reflection font":  "refractionray",
    "reflection":       "refractionray",
    "refractionray":    "refractionray",
    "camofont":         "camoblock",
    "camo font":        "camoblock",
    "camo":             "camoblock",
    "spideyfont":       "spiderweb",
    "spidey font":      "spiderweb",
    "spidey":           "spiderweb",
    "cozyfont":         "cozywinter",
    "cozy font":        "cozywinter",
    "cozy":             "cozywinter",
    "footballfont":     "soccerarmy",
    "football font":    "soccerarmy",
    "football":         "soccerarmy",
    "flowerfont":       "tropicalflower",        # install Tropical Flower Regular.ttf if missing
    "flower font":      "tropicalflower",
    "flower":           "tropicalflower",
    "tropicalflower":   "tropicalflower",
    # Vinyl — TTF is installed
    "vinyl":            "vinylfont",
    "vinylFont":        "vinylfont",
    "vinyl font":       "vinylfont",
    "vinylfont":        "vinylfont",
    # Non-print methods (flag to designer — no TTF)
    "rhinestone":       None,
    "rhinestoneFont":   None,
    "rhinestone font":  None,
    "embroidery":       None,
    "embroideryFont":   None,
    "embroidery font":  None,
    "crystalfont":      None,
    "varsanycrystal":   None,
    "varsany crystal":  None,
}

print(f"  Fonts indexed: {list(FONT_INDEX.keys())}")

# Keys in FONT_INDEX that belong to premium texture/specialty fonts
PREMIUM_FONT_KEYS = {
    "smartkids", "colorfulblocks", "paintsplashesrainbow",
    "wavemermaid", "refractionray", "camoblock", "spiderweb",
    "cozywinter", "soccerarmy", "tropicalflower", "vinylfont",
}

def _resolve_font_key(font_name):
    """Return the FONT_INDEX key for font_name, or None if unmappable."""
    if not font_name:
        return None
    norm = font_name.lower().replace(" ", "").replace("-", "").replace("_", "")
    alias = FONT_ALIASES.get(norm)
    if alias is not None:
        return alias
    return norm if norm in FONT_INDEX else None

def has_premium_font(row):
    """Return True if any zone in this order row uses a premium font.
    Checks: (1) JSON PremiumFont field in FrontFonts/BackFonts (post-Feb-2026 format),
            (2) FrontPremiumFont='Yes' DB column (live DB),
            (3) font name resolves to a PREMIUM_FONT_KEY (pre-Feb-2026 fallback).
    """
    # Post-Feb-2026: font stored as JSON {"NormalFont":"...","PremiumFont":"..."}
    for col in ("FrontFonts", "BackFonts", "PocketFonts", "SleeveFonts"):
        if parse_is_premium_font(row.get(col) or ""):
            return True
    # Live DB has dedicated FrontPremiumFont column
    for col in ("FrontPremiumFont", "BackPremiumFont", "PocketPremiumFont", "SleevePremiumFont"):
        val = row.get(col)
        if val and str(val).strip().lower() in ("yes", "1", "true"):
            return True
    # Pre-Feb-2026 fallback: font name matches a known premium key
    for col in ("FrontFonts", "BackFonts", "PocketFonts", "SleeveFonts"):
        key = _resolve_font_key(row.get(col) or "")
        if key and key in PREMIUM_FONT_KEYS:
            return True
    return False


def cm_to_px(cm): return int(round(cm * PX_PER_CM))

# ─── CANVAS SIZES — from owner Canvases.xlsx ──────────────────────────────────
PRODUCT_CANVAS = {
    # T-shirts
    "adulttshirt":    {"front": (cm_to_px(30), cm_to_px(30)), "back": (cm_to_px(30), cm_to_px(30)), "pocket": (cm_to_px(9),  cm_to_px(7))},
    "kidstshirt":     {"front": (cm_to_px(23), cm_to_px(30)), "back": (cm_to_px(23), cm_to_px(30)), "pocket": (cm_to_px(9),  cm_to_px(7))},
    # Hoodies
    "adulthoodie":    {"front": (cm_to_px(25), cm_to_px(25)), "back": (cm_to_px(25), cm_to_px(25)), "pocket": (cm_to_px(9),  cm_to_px(7)), "sleeve": (cm_to_px(9), cm_to_px(7))},
    "kidshoodie":     {"front": (cm_to_px(23), cm_to_px(20)), "back": (cm_to_px(23), cm_to_px(20)), "pocket": (cm_to_px(9),  cm_to_px(7))},
    # Bags
    "totebag":        {"front": (cm_to_px(28), cm_to_px(28)), "back": (cm_to_px(28), cm_to_px(28))},
    "backpack":       {"front": (cm_to_px(18), cm_to_px(12))},
    "makeupbag":      {"front": (cm_to_px(23), cm_to_px(14))},
    "shoebag":        {"front": (cm_to_px(23), cm_to_px(14))},
    "shoebag2":       {"front": (cm_to_px(14), cm_to_px(14))},
    "stringbag":      {"front": (cm_to_px(22), cm_to_px(24))},
    "knittingbag":    {"front": (cm_to_px(25), cm_to_px(21))},
    # Accessories
    "buckethat":      {"front": (cm_to_px(18), cm_to_px(5))},
    "beanie":         {"front": (cm_to_px(9.5),cm_to_px(4.5))},
    "socks":          {"front": (cm_to_px(6),  cm_to_px(12))},
    "seatbelt":       {"front": (cm_to_px(18), cm_to_px(4))},
    # Baby / Kids
    "babyvest":       {"front": (cm_to_px(15), cm_to_px(17))},
    "sleepsuit":      {"front": (cm_to_px(13), cm_to_px(18))},
    "hodieblanket":   {"front": (cm_to_px(17), cm_to_px(5))},
    # Home / Other
    "cushion":        {"front": (cm_to_px(30), cm_to_px(30))},
    "memorialplaque": {"front": (cm_to_px(13), cm_to_px(8))},
    "golftowel":      {"front": (cm_to_px(17), cm_to_px(17))},
    "golfcase":       {"front": (cm_to_px(15), cm_to_px(6))},
    "slipper":        {"front": (cm_to_px(6),  cm_to_px(6))},
    # Default fallback
    "default":        {"front": (cm_to_px(30), cm_to_px(30)), "back": (cm_to_px(30), cm_to_px(30)), "pocket": (cm_to_px(9), cm_to_px(7))},
}

# ─── SKU PREFIX → PRODUCT KEY — from owner Canvases.xlsx ─────────────────────
SKU_MAP = [
    # Adult T-shirt
    ("MenTee_",                       "adulttshirt"),
    ("AnyTxtOverSizeTee_",            "adulttshirt"),
    ("WmnTee_",                       "adulttshirt"),
    ("PoloTee_",                      "adulttshirt"),
    ("AdultPoloTee_",                 "adulttshirt"),
    ("SignLan01_Tee_",                "adulttshirt"),
    ("Custom04_Tee_",                 "adulttshirt"),
    ("LegendSince",                   "adulttshirt"),
    ("AnyTxt",                        "adulttshirt"),
    # Kids T-shirt
    ("KidsTee_",                      "kidstshirt"),
    ("SLan01KidsTee_",                "kidstshirt"),
    ("PerSingleLetter01KidsTee_",     "kidstshirt"),
    ("FootballKids",                  "kidstshirt"),
    ("67BdayT02Kid",                  "kidstshirt"),
    # Adult Hoodie
    ("AnyTxtAdultHood_",              "adulthoodie"),
    ("MenHood_",                      "adulthoodie"),
    ("HandStand",                     "adulthoodie"),
    ("SplitGirl",                     "adulthoodie"),
    ("FballN",                        "adulthoodie"),
    ("NewFball",                      "adulthoodie"),
    # Kids Hoodie
    ("AnyTxtKidsHood_",               "kidshoodie"),
    ("KidsHood_",                     "kidshoodie"),
    # Tote Bag
    ("AnyTxtTote_",                   "totebag"),
    ("Tote",                          "totebag"),
    # Backpack
    ("AnyTxtBckpck_",                 "backpack"),
    ("BckPack",                       "backpack"),
    ("Name01",                        "backpack"),
    # Baby Vest
    ("AnyTxtBabyVest_",               "babyvest"),
    ("BabyVest",                      "babyvest"),
    # Bucket Hat
    ("AnyTextHat_",                   "buckethat"),
    # Beanie
    ("AnytxtBeanie_",                 "beanie"),
    # Make Up Bag
    ("AnyTxtMakUp_",                  "makeupbag"),
    # Hoodie Blanket
    ("AnyTxtBlanketHood_",            "hodieblanket"),
    # Shoe Bag Sports
    ("AnyTxtShoeB_",                  "shoebag"),
    # Slipper
    ("AnyTxtSlip",                    "slipper"),
    # Socks
    ("AnyTxtSocks",                   "socks"),
    # Cushion
    ("PCushion",                      "cushion"),
    # Custom Tee variants (same canvas as standard tees)
    ("CustomKidsTee_",                "kidstshirt"),   # e.g. CustomKidsTee_Blk78
    ("Custom_Tee_",                   "adulttshirt"),  # e.g. Custom_Tee_BlkM
    # Gym / Swim
    ("GymLeo",                        "default"),
    ("SwimSuit",                      "default"),
]

# ─── HELPERS ──────────────────────────────────────────────────────────────────

def log(msg, level="INFO"):
    ts   = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except:
        pass

def find_image(filename):
    if not filename or not filename.strip():
        return None
    fname = filename.strip().lower()
    if fname in IMAGE_INDEX:
        return IMAGE_INDEX[fname]
    base = os.path.splitext(fname)[0]
    for ext in ['.jpg', '.jpeg', '.png', '.webp']:
        if (base + ext) in IMAGE_INDEX:
            return IMAGE_INDEX[base + ext]
    return None

def download_preview(url):
    """Load a preview image from a URL or a local filename, return PIL Image or None."""
    if not url or not url.strip():
        return None
    src = url.strip()
    # Full URL → download
    if src.startswith("http://") or src.startswith("https://"):
        try:
            tmp = tempfile.mktemp(suffix=".jpg")
            urllib.request.urlretrieve(src, tmp)
            img = Image.open(tmp).convert("RGBA")
            try: os.remove(tmp)
            except: pass
            return img
        except Exception as e:
            log(f"    Preview download failed: {e}", "WARN")
            return None
    # Filename → look up in image index (same as customer images on local DB)
    path = find_image(src)
    if path:
        try:
            return Image.open(path).convert("RGBA")
        except Exception as e:
            log(f"    Preview load failed: {e}", "WARN")
    return None

def parse_image_json(json_str):
    if not json_str or not json_str.strip():
        return []
    try:
        d = json.loads(json_str.strip())
        return [d[f"Image{i}"].strip() for i in range(1, 6) if d.get(f"Image{i}", "").strip()]
    except:
        return []

def _parse_font_json(fonts_raw):
    """Parse FrontFonts/BackFonts value. Returns (font_name, is_premium).
    Handles both valid JSON (double quotes) and Python dict repr (single quotes).
    """
    if not fonts_raw:
        return "Arial", False
    s = fonts_raw.strip()
    if s.startswith("{"):
        d = None
        # Try JSON first (double quotes)
        try:
            d = json.loads(s)
        except Exception:
            pass
        # Fallback: Python dict repr (single quotes from DB)
        if d is None:
            try:
                import ast
                d = ast.literal_eval(s)
            except Exception:
                pass
        if d is not None:
            premium = (d.get("PremiumFont") or "").strip()
            normal  = (d.get("NormalFont")  or "").strip()
            if premium and premium.lower() not in ("no", "none", ""):
                return premium, True
            return normal or "Arial", False
    return s or "Arial", False

def parse_font(fonts_raw):
    """Return font name to use (premium takes priority over normal)."""
    return _parse_font_json(fonts_raw)[0]

def parse_is_premium_font(fonts_raw):
    """Return True if this zone uses a premium font."""
    return _parse_font_json(fonts_raw)[1]

def parse_colour(colours_raw):
    if not colours_raw:
        return "#ffffff"
    s = colours_raw.strip()
    if s.startswith("{"):
        try:
            d = json.loads(s)
            return d.get("Colour1") or d.get("colour1") or "#ffffff"
        except:
            pass
    if s.startswith("#"):
        return s
    return "#ffffff"

def parse_texts(raw):
    """Parse customer text, preserving blank lines as spacers (capped at 1 consecutive blank).
    Empty string "" in the returned list = one blank spacer line."""
    if not raw or not raw.strip():
        return []
    raw = raw.strip()
    if "|" in raw and "\n" not in raw:
        return [t.strip() for t in raw.split("|") if t.strip()]
    lines = raw.split("\n")
    result = []
    prev_blank = False
    for line in lines:
        s = line.strip()
        if s:
            result.append(s)
            prev_blank = False
        else:
            # Only add one blank spacer between real text lines (not at start)
            if result and not prev_blank:
                result.append("")
            prev_blank = True
    return result

def hex_to_rgb(hex_col):
    h = hex_col.lstrip("#")
    if len(h) == 3:
        h = "".join(c*2 for c in h)
    try:
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
    except:
        return (255, 255, 255)

def get_font(font_name, size_px):
    norm = font_name.lower().replace(' ','').replace('-','').replace('_','')
    # Check aliases first
    resolved = FONT_ALIASES.get(norm)
    if resolved is None:
        pass  # known non-renderable font (rhinestone etc) - fall through to arial
    elif resolved and resolved in FONT_INDEX:
        try:
            return ImageFont.truetype(FONT_INDEX[resolved], size_px)
        except:
            pass
    # Direct match in index
    if norm in FONT_INDEX:
        try:
            return ImageFont.truetype(FONT_INDEX[norm], size_px)
        except:
            pass
    # Partial match - find any font whose key contains the norm
    for key, path in FONT_INDEX.items():
        if norm in key or key in norm:
            try:
                return ImageFont.truetype(path, size_px)
            except:
                pass
    # System fonts
    system_map = {
        "arial": "arial.ttf", "arialbold": "arialbd.ttf",
        "timesnewroman": "times.ttf", "couriernew": "cour.ttf",
        "verdana": "verdana.ttf", "impact": "impact.ttf",
        "helvetica": "arial.ttf", "georgia": "georgia.ttf",
        "tahoma": "tahoma.ttf",
    }
    if norm in system_map:
        try:
            return ImageFont.truetype(system_map[norm], size_px)
        except:
            pass
    # Final fallback
    try:
        return ImageFont.truetype("arial.ttf", size_px)
    except:
        return ImageFont.load_default()

def detect_category(sku):
    if not sku:
        return "Other"
    s = sku.lower()
    if "polo" in s:                                    return "Polo"
    if "kidstee" in s or "kidstshirt" in s:            return "Kids T-Shirt"
    if "hood" in s:                                    return "Hoodie"
    if "tote" in s:                                    return "Tote Bag"
    if "slipper" in s:                                 return "Slipper"
    if "baby" in s or "vest" in s:                     return "Baby Vest"
    if "backpack" in s or "bckpck" in s:               return "Backpack"
    if "mentee" in s or "_tee_" in s or "wmntee" in s: return "T-Shirt"
    if "hat" in s or "cap" in s or "beanie" in s:      return "Hat"
    if "gym" in s or "leo" in s or "legsui" in s:      return "Gym & Leotard"
    if "dart" in s:                                    return "Dart Case"
    if "towel" in s or "twl" in s:                     return "Towel"
    if "rainsuit" in s or "wellis" in s or "socks" in s or "keychain" in s: return "Accessories"
    if "lan" in s:                                     return "Sign Language"
    return "Other"

def sku_colour_folder(sku):
    """Return 'black', 'white', or '' (save directly, no sub-folder) from SKU colour code."""
    if not sku:
        return ""
    s = sku.lower()
    if "blk" in s:  return "black"
    if "wht" in s:  return "white"
    return ""

def is_multizone_row(row):
    """True if the row has content in more than one print zone (front+back, +sleeve, etc.)."""
    active = 0
    if row.get("IsFrontLocation") or (row.get("FrontImage") or "").strip() or (row.get("FrontText") or "").strip():
        active += 1
    if row.get("IsBackLocation") or (row.get("BackImage") or "").strip() or (row.get("BackText") or "").strip():
        active += 1
    if row.get("IsPocketLocation") or (row.get("PocketImage") or "").strip() or (row.get("PocketText") or "").strip():
        active += 1
    if row.get("IsSleeveLocation") or (row.get("SleeveImage") or "").strip() or (row.get("SleeveText") or "").strip():
        active += 1
    return active > 1

def is_emb_rhine_row(row):
    """True if any zone uses an embroidery or rhinestone font (handled manually, not DTF)."""
    for field in ("FrontFonts", "BackFonts", "PocketFonts", "SleeveFonts"):
        val = (row.get(field) or "").lower()
        if "emb" in val or "rhine" in val:
            return True
    return False

def detect_product(sku):
    """Map SKU to product key using SKU_MAP from owner canvas file.
    Tries each prefix in order — first match wins.
    Falls back to keyword matching, then default.
    """
    if not sku:
        return "default"
    # Direct prefix match from owner SKU_MAP
    for prefix, product_key in SKU_MAP:
        if sku.startswith(prefix):
            return product_key
    # Keyword fallback for edge cases
    s = sku.lower()
    if "kidstee" in s:          return "kidstshirt"
    if "kidshoo" in s:          return "kidshoodie"
    if "hood" in s:             return "adulthoodie"
    if "tote" in s:             return "totebag"
    if "slipper" in s:          return "slipper"
    if "baby" in s:             return "babyvest"
    if "vest" in s:             return "babyvest"
    if "backpack" in s or "bckpck" in s: return "backpack"
    if "beanie" in s:           return "beanie"
    if "hat" in s:              return "buckethat"
    if "tee" in s or "polo" in s: return "adulttshirt"
    return "default"

def get_dims(product, zone):
    spec = PRODUCT_CANVAS.get(product, PRODUCT_CANVAS["default"])
    return spec.get(zone, spec.get("front", (cm_to_px(30), cm_to_px(30))))

# ─── PSD WRITER ───────────────────────────────────────────────────────────────

def _pack_layer_name(s):
    b = s.encode("latin-1", errors="replace")[:255]
    data = bytes([len(b)]) + b
    pad = (4 - len(data) % 4) % 4
    return data + b'\x00' * pad

def _to_channels(img, mode):
    img = img.convert(mode)
    bands = img.split()
    if mode == 'RGBA':
        r, g, b, a = bands
        return {0: r.tobytes(), 1: g.tobytes(), 2: b.tobytes(), -1: a.tobytes()}
    r, g, b = bands
    return {0: r.tobytes(), 1: g.tobytes(), 2: b.tobytes()}

def write_psd(out_path, canvas_w, canvas_h, layers):
    """Write a layered PSD file using psd-tools (Photoshop-compatible).
    Falls back to a raw binary writer for PSB (canvas > 30 000px) since
    psd-tools does not support the PSB (version 2) format.
    """
    # Always write PSD — printing software does not support PSB
    use_psb = False

    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    # ── Raw binary writer ─────────────────────────────────────────────────────
    version = 2 if use_psb else 1

    buf = io.BytesIO()
    p   = buf.write

    p(b'8BPS')
    p(struct.pack('>H', version))
    p(b'\x00' * 6)
    p(struct.pack('>H', 3))
    p(struct.pack('>I', canvas_h))
    p(struct.pack('>I', canvas_w))
    p(struct.pack('>H', 8))
    p(struct.pack('>H', 3))

    p(struct.pack('>I', 0))  # Color mode data (empty)

    dpi_fixed = DPI << 16
    res_data  = struct.pack('>IHHIHH', dpi_fixed, 1, 1, dpi_fixed, 1, 1)
    res_block = (b'8BIM' + struct.pack('>H', 1005) +
                 b'\x00\x00' + struct.pack('>I', len(res_data)) + res_data)
    p(struct.pack('>I', len(res_block)))
    p(res_block)

    lr_buf = io.BytesIO()
    ld_buf = io.BytesIO()

    for lyr in layers:
        img    = lyr['image']
        top    = lyr['top']
        left   = lyr['left']
        bottom = top  + img.height
        right  = left + img.width
        flags  = 0 if lyr.get('visible', True) else 2

        ch       = _to_channels(img, 'RGBA')
        ch_order = [-1, 0, 1, 2]

        lr = io.BytesIO()
        lr.write(struct.pack('>iiii', top, left, bottom, right))
        lr.write(struct.pack('>H', 4))
        for cid in ch_order:
            ch_len = len(ch[cid]) + 2
            if use_psb:
                lr.write(struct.pack('>hQ', cid, ch_len))
            else:
                lr.write(struct.pack('>hI', cid, ch_len))
        lr.write(b'8BIM')
        lr.write(b'norm')
        lr.write(struct.pack('>B', lyr.get('opacity', 255)))
        lr.write(struct.pack('>B', 0))
        lr.write(struct.pack('>B', flags))
        lr.write(b'\x00')
        name_bytes = _pack_layer_name(lyr['name'])
        extra = struct.pack('>I', 0) + struct.pack('>I', 0) + name_bytes
        lr.write(struct.pack('>I', len(extra)))
        lr.write(extra)
        lr_buf.write(lr.getvalue())

        for cid in ch_order:
            ld_buf.write(struct.pack('>H', 0))
            ld_buf.write(ch[cid])

    layer_info = struct.pack('>h', len(layers)) + lr_buf.getvalue() + ld_buf.getvalue()
    if len(layer_info) % 4:
        layer_info += b'\x00' * (4 - len(layer_info) % 4)

    if use_psb:
        lmi = struct.pack('>Q', len(layer_info)) + layer_info + struct.pack('>I', 0)
        p(struct.pack('>Q', len(lmi)))
    else:
        lmi = struct.pack('>I', len(layer_info)) + layer_info + struct.pack('>I', 0)
        p(struct.pack('>I', len(lmi)))
    p(lmi)

    composite = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
    for lyr in layers:
        if lyr.get('visible', True):
            composite.paste(lyr['image'].convert("RGBA"),
                            (lyr['left'], lyr['top']),
                            lyr['image'].convert("RGBA"))
    comp = _to_channels(composite, 'RGB')
    p(struct.pack('>H', 0))
    for cid in [0, 1, 2]:
        p(comp[cid])

    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())
    return out_path

# ─── LAYER BUILDERS ───────────────────────────────────────────────────────────

def build_image_layer(img_path, w, h, sku=None):
    if not img_path or not os.path.isfile(img_path):
        return Image.new("RGBA", (1, 1), (0, 0, 0, 0)), 0, 0
    src = Image.open(img_path).convert("RGBA")

    # Auto background removal: if image background matches garment colour, remove it
    garment_rgb = get_garment_rgb(sku) if sku else None
    if garment_rgb and image_bg_matches_garment(src, garment_rgb):
        log(f"  Auto bg-remove: background matches garment colour {garment_rgb}", "INFO")
        src = remove_background(src, garment_rgb=garment_rgb)

    # Alpha threshold — remove near-transparent noise pixels
    r, g, b, a = src.split()
    a = a.point(lambda x: 0 if x < 128 else x)
    src = Image.merge("RGBA", (r, g, b, a))

    # Crop to content bounding box so transparent borders (left by bg-removal)
    # don't shrink the design when it is scaled to fill the canvas
    bbox = src.getbbox()
    if bbox:
        src = src.crop(bbox)

    # Scale to full zone width — height follows proportionally
    ratio = w / src.width
    nw    = w
    nh    = max(1, int(src.height * ratio))
    src   = src.resize((nw, nh), Image.LANCZOS)
    return src, 0, 0

try:
    from pilmoji import Pilmoji
    PILMOJI_AVAILABLE = True
except ImportError:
    PILMOJI_AVAILABLE = False

# ─── CHROME SVG COLOUR FONT RENDERER ─────────────────────────────────────────
# Fonts like RefractionRay, Smart Kids, Cozy Winter store glyphs as SVG only.
# PIL/FreeType cannot render them. We fall back to headless Chrome which
# supports all OpenType SVG colour fonts natively.

CHROME_EXE = None
for _p in [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]:
    if os.path.exists(_p):
        CHROME_EXE = _p
        break

# Fonts that PIL cannot render visually (SVG/colour-only fonts with empty outlines).
# Detected by: loading OK but rendering 'A' produces zero visible pixels.
SVG_ONLY_FONT_KEYS = set()
def _test_font_renders(path):
    """Return True if this font file produces visible pixels when rendering 'A'."""
    try:
        from PIL import ImageFont as _IF, Image as _Im, ImageDraw as _ID
        import numpy as _np
        _f   = _IF.truetype(path, 200)
        _img = _Im.new("RGBA", (300, 300), (0, 0, 0, 0))
        _ID.Draw(_img).text((10, 10), "A", font=_f, fill=(0, 0, 0, 255))
        return _np.array(_img)[:, :, 3].max() > 0
    except Exception:
        return False  # can't load at all → also needs Chrome

for _key, _path in FONT_INDEX.items():
    if not _test_font_renders(_path):
        SVG_ONLY_FONT_KEYS.add(_key)

if SVG_ONLY_FONT_KEYS:
    print(f"  SVG colour fonts (Chrome renderer): {sorted(SVG_ONLY_FONT_KEYS)}")

def _is_premium_font_name(font_name):
    """Return True if this font is a premium colour font (all go through Chrome)."""
    key = _resolve_font_key(font_name)
    return key in PREMIUM_FONT_KEYS

def _is_svg_only_font(font_name):
    """Return True if this font requires Chrome to render (PIL fails)."""
    key = _resolve_font_key(font_name)
    return key in SVG_ONLY_FONT_KEYS

def build_text_layer_chrome(text_lines, font_name, colour_hex, canvas_w):
    """Render premium SVG colour fonts using Chrome.

    Strategy A (primary): one HTML page per text block with @font-face pointing
    directly at the font file via --allow-file-access-from-files.  Single Chrome
    call per text block — no profile-lock issues, handles kerning automatically.

    Strategy B (fallback): glyph-by-glyph SVG render — one Chrome call per
    character, composited in PIL.  Avoids CSS font-rendering limitations.

    Returns RGBA PIL image, or None on failure.
    """
    if not CHROME_EXE:
        return None
    real_lines = [l for l in text_lines if l.strip()]
    if not real_lines:
        return None

    key       = _resolve_font_key(font_name)
    font_path = FONT_INDEX.get(key, "")
    if not font_path:
        return None

    # Prefer OTF — it carries the full SVG colour glyph data
    otf_path = font_path.rsplit(".", 1)[0] + ".otf"
    if os.path.exists(otf_path):
        font_path = otf_path

    import subprocess, numpy as np, re, html as _html

    os.makedirs(r"C:\Varsany\Temp", exist_ok=True)

    # ── STRATEGY A: glyph-collage SVG (single Chrome call, correct font colours) ─
    # Each unique character is rendered as an inline SVG <img> data-URL element.
    # SVG rendering preserves the font's built-in colour data (unlike CSS @font-face).
    # All glyphs are placed in one HTML page → one Chrome call, no profile-lock issues.
    def _try_collage_render():
        import base64 as _b64

        try:
            import sys as _sys
            _sys.path.insert(
                0, r"C:\Users\yedhu\AppData\Local\Programs\Python\Python313\Lib\site-packages")
            from fontTools.ttLib import TTFont
        except ImportError:
            return None

        try:
            ft_font     = TTFont(font_path)
            svg_table   = ft_font.get('SVG ')
            cmap        = ft_font.getBestCmap()
            glyph_order = ft_font.getGlyphOrder()
            upem        = ft_font['head'].unitsPerEm
            ascender    = ft_font['hhea'].ascent
            descender   = ft_font['hhea'].descent   # negative
            vb_h        = ascender - descender
            hmtx        = ft_font['hmtx'].metrics
        except Exception:
            return None

        if not svg_table:
            return None

        svg_map = {}
        for doc in svg_table.docList:
            for gid in range(doc.startGlyphID, doc.endGlyphID + 1):
                svg_map[gid] = (doc.data,
                                doc.startGlyphID != doc.endGlyphID)  # shared doc?

        # Render glyphs at the actual output size so Chrome produces
        # sharp detail (football patterns, textures etc.) without any
        # upscaling step.  Target: each glyph fills its share of canvas width.
        # We cap at 2400px so the collage HTML doesn't exceed Chrome's limits.
        max_chars = max(len(l) for l in real_lines if l)
        target_w  = int(canvas_w * 0.85 / max(1, max_chars))   # px per char
        # Use ×4 multiplier so glyphs render at high resolution before any
        # scale-to-canvas step — prevents blurring on upscale.
        glyph_h   = max(800, min(2400, target_w * 4))

        # vb_h_render: viewBox height used when rendering each glyph.
        # Must span from top of ascenders (-ascender) down to below baseline.
        # Using ascender*2 captures decorative elements (e.g. Mermaid tails at y=+570)
        # which can extend well below the standard descender line.
        vb_h_render = ascender * 2
        scale       = glyph_h / vb_h_render   # font units → screen pixels

        # ── Collect unique characters and prepare their SVG ──────────────────
        all_chars = sorted(set(ch for l in text_lines for ch in l if ch.strip()))
        char_svg  = {}   # ch → (adv_px, fixed_svg_string)
        char_gid  = {}   # ch → gid

        for ch in all_chars:
            cp         = ord(ch)
            glyph_name = cmap.get(cp)
            if not glyph_name:
                continue
            try:
                gid = glyph_order.index(glyph_name)
            except ValueError:
                continue
            entry = svg_map.get(gid)
            if not entry:
                continue
            svg_raw, shared = entry
            adv_units = hmtx.get(glyph_name, (upem // 3, 0))[0]
            adv_px    = max(1, int(adv_units * scale))
            h_px      = glyph_h   # viewport height = glyph_h (= vb_h_render * scale)

            svg = svg_raw
            # Force a UNIFORM viewBox for every glyph so all letters are
            # rendered at the same visual scale regardless of how the font
            # designer laid out the SVG coordinate space per glyph.
            # The viewBox spans the full advance width horizontally and the
            # standard cap-height range vertically.
            vb_attr = f'viewBox="0 {-ascender} {upem} {vb_h_render}"'
            if 'viewBox' in svg:
                svg = re.sub(r'viewBox="[^"]*"', vb_attr, svg)
            else:
                svg = svg.replace('<svg', f'<svg {vb_attr}', 1)
            # Pixel dimensions: width scales with advance, height = glyph_h.
            # Both are derived from the SAME scale factor so all glyphs have
            # identical height and proportional widths.
            svg = re.sub(r'\s+width="[^"]*"',  '', svg)
            svg = re.sub(r'\s+height="[^"]*"', '', svg)
            svg = svg.replace('<svg', f'<svg width="{adv_px}px" height="{h_px}px"', 1)
            # Shared document: inject CSS to show only this glyph's <g> element
            if shared:
                hide = (f'<style>g{{display:none}}'
                        f'#glyph{gid}{{display:inline}}'
                        f'#glyph\\.{gid}{{display:inline}}</style>')
                svg = re.sub(r'(<svg[^>]*>)', r'\1' + hide, svg, count=1)

            char_svg[ch] = (adv_px, h_px, svg)
            char_gid[ch] = gid

        if not char_svg:
            return None

        # ── Build one-line collage HTML (all unique glyphs side by side) ─────
        # Each glyph is a base64-encoded SVG data-URL <img>.
        # Using <img> means each SVG renders in isolation — no ID conflicts.
        collage_w = sum(v[0] for v in char_svg.values()) + 10
        collage_h = max(v[1] for v in char_svg.values()) + 10

        items_html = ""
        x_pos = {}
        cx = 0
        for ch, (adv_px, h_px, svg_str) in char_svg.items():
            svg_b64 = _b64.b64encode(svg_str.encode("utf-8")).decode("ascii")
            items_html += (
                f'<img style="position:absolute;left:{cx}px;top:0;'
                f'width:{adv_px}px;height:{h_px}px;display:block" '
                f'src="data:image/svg+xml;base64,{svg_b64}">\n'
            )
            x_pos[ch] = cx
            cx += adv_px

        html_src = f"""<!DOCTYPE html>
<html><head><style>
*{{margin:0;padding:0}}html,body{{background:#ffffff;overflow:hidden}}
.c{{position:relative;width:{collage_w}px;height:{collage_h}px}}
</style></head><body>
<div class="c">{items_html}</div>
</body></html>"""

        _pid      = os.getpid()
        html_path = rf"C:\Varsany\Temp\glyph_collage_{_pid}.html"
        png_path  = rf"C:\Varsany\Temp\glyph_collage_{_pid}.png"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_src)
        try:
            if os.path.exists(png_path):
                os.remove(png_path)
        except Exception:
            pass

        cmd = [
            CHROME_EXE, "--headless", "--no-sandbox",
            "--disable-gpu", "--disable-extensions",
            "--no-first-run", "--disable-sync",
            f"--screenshot={png_path}",
            f"--window-size={collage_w},{collage_h}",
            "file:///" + html_path.replace("\\", "/"),
        ]
        try:
            subprocess.run(cmd, capture_output=True, timeout=40)
        except Exception:
            return None

        if not os.path.exists(png_path):
            return None
        try:
            collage_img = Image.open(png_path)
            collage_img.load()
            collage_img = collage_img.convert("RGBA")
        except Exception:
            return None

        # Handle HiDPI: Chrome may return 2× pixels
        dpr = collage_img.width / collage_w if collage_w > 0 else 1.0

        # ── Crop individual glyphs from collage ───────────────────────────────
        glyph_imgs_raw = {}
        for ch, (adv_px, h_px, _) in char_svg.items():
            x0 = int(x_pos[ch] * dpr)
            x1 = int((x_pos[ch] + adv_px) * dpr)
            y1 = int(h_px * dpr)
            crop = collage_img.crop((x0, 0, min(x1, collage_img.width),
                                     min(y1, collage_img.height)))
            if dpr != 1.0:
                crop = crop.resize((adv_px, h_px), Image.LANCZOS)
            arr = np.array(crop)
            white = (arr[:, :, 0] > 240) & (arr[:, :, 1] > 240) & (arr[:, :, 2] > 240)
            arr[white, 3] = 0
            img_rgba = Image.fromarray(arr)
            bbox = img_rgba.getbbox()
            # Crop to visible content bbox so we know the real pixel size
            glyph_imgs_raw[ch] = img_rgba.crop(bbox) if (bbox and arr[:, :, 3].max() > 0) else None

        if not any(v is not None for v in glyph_imgs_raw.values()):
            log(f"  Collage render: all glyphs transparent for {font_name}", "WARN")
            return None

        # Normalise glyphs within each typographic class to a uniform height.
        # Uppercase + digits  → cap_h  (normalise all to the tallest uppercase glyph)
        # Lowercase           → x_h    (normalise all to the tallest lowercase glyph)
        # Punctuation/symbols → keep natural size (never scale up a tiny apostrophe)
        # This preserves the visual uppercase/lowercase distinction.
        import string as _string
        _UPPER = set(_string.ascii_uppercase + _string.digits)
        _LOWER = set(_string.ascii_lowercase)
        _upper_hs = [g.height for ch, g in glyph_imgs_raw.items() if g is not None and ch in _UPPER]
        _lower_hs = [g.height for ch, g in glyph_imgs_raw.items() if g is not None and ch in _LOWER]
        cap_h = max(_upper_hs, default=None)
        x_h   = max(_lower_hs, default=None)
        if cap_h is None: cap_h = x_h or 1
        if x_h   is None: x_h   = cap_h
        max_content_h = cap_h   # line_h is always driven by cap height

        glyph_imgs    = {}
        glyph_base_y  = {}      # per-char vertical offset when pasting into line_img
        for ch, gimg in glyph_imgs_raw.items():
            if gimg is None:
                glyph_imgs[ch] = None
                continue
            if ch in _UPPER:
                target_h = cap_h
                base_y   = 0                        # sit at top of line
            elif ch in _LOWER:
                target_h = x_h
                base_y   = cap_h - x_h             # baseline-align (sit at bottom)
            else:
                # Punctuation/symbol: if the font has an explicit SVG glyph for it,
                # the designer made it a full display character → scale to cap_h.
                # Characters not in the SVG table have no designed size so keep natural.
                if ch in char_svg:
                    target_h = cap_h
                    base_y   = 0
                else:
                    target_h = gimg.height
                    base_y   = max(0, (cap_h - gimg.height) // 2)  # vertically centred
            if gimg.height != target_h:
                sf   = target_h / gimg.height
                nw   = max(1, int(gimg.width * sf))
                gimg = gimg.resize((nw, target_h), Image.LANCZOS)
            glyph_imgs[ch]   = gimg
            glyph_base_y[ch] = base_y

        # ── Compose text lines ────────────────────────────────────────────────
        # Advance width = max(hmtx advance, actual content width + small gap).
        # Using hmtx alone causes overlaps when short glyphs are normalised up
        # to max_content_h (width scales up too, exceeding the advance cell).
        norm_scale = max_content_h / max(1, glyph_h)
        space_em   = hmtx.get('space', hmtx.get('uni0020', (upem // 3, 0)))[0]
        space_w    = max(4, int(space_em * norm_scale * glyph_h / upem))
        min_gap    = max(4, max_content_h // 20)   # ~5% of cap height between letters
        line_h     = max_content_h
        line_imgs  = []
        for line_text in text_lines:
            if not line_text.strip():
                line_imgs.append(None)
                continue
            x = 0
            parts = []
            for ch in line_text:
                gimg = glyph_imgs.get(ch)
                if ch in char_svg:
                    adv_px_at_glyph_h = char_svg[ch][0]
                    adv_hmtx = max(1, int(adv_px_at_glyph_h * norm_scale))
                    # Never let advance be smaller than the glyph content —
                    # that would cause letters to overlap each other.
                    content_w = gimg.width if gimg is not None else 0
                    adv = max(adv_hmtx, content_w + min_gap)
                    gx_offset = max(0, (adv - content_w) // 2)
                else:
                    adv = space_w
                    gx_offset = 0
                parts.append((ch, gimg, x + gx_offset))
                x += adv
            if x <= 0:
                line_imgs.append(None)
                continue
            line_img = Image.new("RGBA", (x, line_h), (0, 0, 0, 0))
            for ch, gimg, gx in parts:
                if gimg is not None:
                    gy = glyph_base_y.get(ch, 0)
                    line_img.paste(gimg, (gx, gy), gimg)
            line_imgs.append(line_img)

        valid = [li for li in line_imgs if li is not None]
        if not valid:
            return None

        # Stack lines — use actual max line width to prevent right-edge clipping
        line_gap   = max(line_h // 3, 40)  # ~33% of cap height between lines
        n_lines    = len(line_imgs)
        total_h    = line_h * n_lines + line_gap * max(0, n_lines - 1)
        max_line_w = max((limg.width for limg in line_imgs if limg is not None), default=canvas_w)
        result     = Image.new("RGBA", (max_line_w, max(1, total_h)), (0, 0, 0, 0))
        y = 0
        for limg in line_imgs:
            if limg is not None:
                cx2 = max(0, (max_line_w - limg.width) // 2)
                result.paste(limg, (cx2, y), limg)
            y += line_h + line_gap

        bbox = result.getbbox()
        if not bbox:
            return None
        result = result.crop(bbox)

        # Scale to fill ~85% of canvas width — same as the PIL renderer does.
        avail_w = int(canvas_w * 0.85)
        if result.width > 0 and result.width != avail_w:
            ratio   = avail_w / result.width
            new_h   = max(1, int(result.height * ratio))
            result  = result.resize((avail_w, new_h), Image.LANCZOS)

        log(f"  Collage render OK for {font_name}: {result.size}", "INFO")
        return result

    result = _try_collage_render()
    if result is not None:
        return result

    log(f"  Collage render failed for {font_name} — falling back to CSS font render", "WARN")

    # ── STRATEGY B: CSS @font-face (single Chrome call — shape only, no SVG colour) ─
    def _try_html_render():
        max_chars  = max(len(l) for l in real_lines if l)
        font_px    = max(80, min(500, int(canvas_w * 0.85 / max(1, max_chars))))
        line_height = int(font_px * 1.25)
        total_h     = line_height * len(text_lines) + font_px

        bg = "#00FE00"   # lime-green background for clean keying

        lines_html = "\n".join(
            f'<div class="tl">{_html.escape(l) if l.strip() else "&nbsp;"}</div>'
            for l in text_lines
        )

        font_url = "file:///" + font_path.replace("\\", "/")
        fmt      = "opentype" if font_path.lower().endswith(".otf") else "truetype"

        html_src = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
*{{margin:0;padding:0;box-sizing:border-box}}
html,body{{background:{bg};width:{canvas_w}px}}
@font-face{{font-family:'PF';src:url('{font_url}') format('{fmt}')}}
.tl{{font-family:'PF',sans-serif;font-size:{font_px}px;
     line-height:{line_height}px;text-align:center;
     width:{canvas_w}px;white-space:pre;overflow:hidden}}
</style></head><body>
{lines_html}
</body></html>"""

        _pid      = os.getpid()
        html_path = rf"C:\Varsany\Temp\premium_font_{_pid}.html"
        png_path  = rf"C:\Varsany\Temp\premium_font_{_pid}.png"

        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_src)
        try:
            if os.path.exists(png_path):
                os.remove(png_path)
        except Exception:
            pass

        cmd = [
            CHROME_EXE, "--headless", "--no-sandbox",
            "--disable-gpu", "--disable-extensions",
            "--allow-file-access-from-files",
            "--no-first-run", "--disable-sync",
            f"--screenshot={png_path}",
            f"--window-size={canvas_w},{max(200, total_h)}",
            "file:///" + html_path.replace("\\", "/"),
        ]
        try:
            subprocess.run(cmd, capture_output=True, timeout=30)
        except Exception:
            return None

        if not os.path.exists(png_path):
            return None
        try:
            img = Image.open(png_path)
            img.load()
            img = img.convert("RGBA")
        except Exception:
            return None

        if img.width > canvas_w:
            scale_f = canvas_w / img.width
            img = img.resize(
                (canvas_w, max(1, int(img.height * scale_f))), Image.LANCZOS)

        arr = np.array(img)
        lime = (arr[:, :, 0] < 30) & (arr[:, :, 1] > 240) & (arr[:, :, 2] < 30)
        arr[lime, 3] = 0

        if arr[:, :, 3].max() == 0:
            return None

        result = Image.fromarray(arr)
        bbox = result.getbbox()
        if not bbox:
            return None
        log(f"  CSS render OK for {font_name}: {result.size} -> bbox {bbox}", "INFO")
        return result.crop(bbox)

    result = _try_html_render()
    if result is not None:
        return result

    log(f"  CSS render also failed for {font_name} — falling back to glyph-by-glyph SVG", "WARN")

    # ── STRATEGY B: glyph-by-glyph SVG render ──────────────────────────────────
    try:
        import sys as _sys
        _sys.path.insert(
            0, r"C:\Users\yedhu\AppData\Local\Programs\Python\Python313\Lib\site-packages")
        from fontTools.ttLib import TTFont
    except ImportError:
        log("fontTools not available — cannot render premium font", "WARN")
        return None

    try:
        ft_font     = TTFont(font_path)
        svg_table   = ft_font.get('SVG ')
        cmap        = ft_font.getBestCmap()
        glyph_order = ft_font.getGlyphOrder()
        upem        = ft_font['head'].unitsPerEm
        ascender    = ft_font['hhea'].ascent
        descender   = ft_font['hhea'].descent   # negative number
        vb_h        = ascender - descender
        hmtx        = ft_font['hmtx'].metrics
    except Exception as e:
        log(f"fontTools load error for {font_name}: {e}", "WARN")
        return None

    if not svg_table:
        log(f"No SVG table in {font_name}", "WARN")
        return None

    # glyph_id → SVG document
    svg_map = {}
    for doc in svg_table.docList:
        for gid in range(doc.startGlyphID, doc.endGlyphID + 1):
            svg_map[gid] = doc.data

    # Target glyph size
    max_chars  = max(len(l) for l in real_lines if l)
    glyph_px   = max(150, min(400, canvas_w // max(1, max_chars)))
    scale      = glyph_px / vb_h
    line_h     = max(10, int(vb_h * scale * 1.1))
    profile_dir = r"C:\Varsany\Temp\chrome_profile"   # shared — calls are sequential

    def render_glyph_svg(gid):
        """Render one glyph → RGBA image, or (None, adv_px) on failure."""
        glyph_name = glyph_order[gid] if gid < len(glyph_order) else None
        adv_units  = hmtx.get(glyph_name, (upem // 3, 0))[0] if glyph_name else upem // 3
        adv_px     = max(1, int(adv_units * scale))
        h_px       = max(1, int(vb_h * scale))

        svg_raw = svg_map.get(gid)
        if not svg_raw:
            return None, adv_px

        svg = svg_raw
        # Fix / add viewBox so Chrome maps the glyph coordinate space correctly
        vb_attr = f'viewBox="0 {descender} {upem} {vb_h}"'
        if 'viewBox' in svg:
            svg = re.sub(r'viewBox="[^"]*"', vb_attr, svg)
        else:
            svg = svg.replace('<svg', f'<svg {vb_attr}', 1)

        # Force explicit pixel size — remove any existing width/height first
        svg = re.sub(r'\s+width="[^"]*"',  '', svg)
        svg = re.sub(r'\s+height="[^"]*"', '', svg)
        svg = svg.replace('<svg', f'<svg width="{adv_px}" height="{h_px}"', 1)

        # For shared SVG documents (multiple glyphs), hide all except this one
        # by injecting a CSS rule that shows only #glyph{gid}
        if svg_map.get(gid) and svg_table.docList:
            for doc in svg_table.docList:
                if doc.startGlyphID <= gid <= doc.endGlyphID and \
                        doc.startGlyphID != doc.endGlyphID:
                    # Shared document — CSS-hide all sibling glyphs
                    hide_css = (
                        f'<style>g{{display:none}} '
                        f'#glyph{gid},#glyph.{gid}{{display:inline}}</style>'
                    )
                    svg = svg.replace('<svg', '<svg', 1)  # no-op anchor
                    svg = re.sub(r'(<svg[^>]*>)', r'\1' + hide_css, svg, count=1)
                    break

        svg_path = rf"C:\Varsany\Temp\glyph_{gid}.svg"
        png_path = rf"C:\Varsany\Temp\glyph_{gid}.png"

        try:
            if os.path.exists(png_path):
                os.remove(png_path)
        except Exception:
            pass

        with open(svg_path, "w", encoding="utf-8") as f:
            f.write(svg)

        cmd = [
            CHROME_EXE, "--headless", "--no-sandbox",
            "--disable-gpu", "--disable-extensions",
            "--no-first-run", "--disable-sync",
            f"--user-data-dir={profile_dir}",
            f"--screenshot={png_path}",
            f"--window-size={adv_px},{h_px}",
            "file:///" + svg_path.replace("\\", "/"),
        ]
        try:
            subprocess.run(cmd, capture_output=True, timeout=20)
        except Exception as e:
            log(f"  Chrome error glyph {gid}: {e}", "WARN")
            return None, adv_px

        if not os.path.exists(png_path):
            return None, adv_px

        try:
            img = Image.open(png_path)
            img.load()
            img = img.convert("RGBA")
        except Exception:
            return None, adv_px

        # Handle HiDPI — resize to expected dimensions
        if img.width != adv_px or img.height != h_px:
            img = img.resize((adv_px, h_px), Image.LANCZOS)

        # Remove white background
        arr = np.array(img)
        white = (arr[:, :, 0] > 240) & (arr[:, :, 1] > 240) & (arr[:, :, 2] > 240)
        arr[white, 3] = 0

        if arr[:, :, 3].max() == 0:
            return None, adv_px  # Glyph rendered no visible pixels

        return Image.fromarray(arr), adv_px

    # Build one image per line
    line_images = []
    for line_text in text_lines:
        if not line_text.strip():
            line_images.append(None)
            continue

        x     = 0
        parts = []
        for ch in line_text:
            cp         = ord(ch)
            glyph_name = cmap.get(cp)
            if not glyph_name:
                adv_px = int((upem // 3) * scale)
                parts.append((None, x, adv_px))
                x += adv_px
                continue
            try:
                gid = glyph_order.index(glyph_name)
            except ValueError:
                adv_px = max(1, int(hmtx.get(glyph_name, (upem // 3, 0))[0] * scale))
                parts.append((None, x, adv_px))
                x += adv_px
                continue
            glyph_img, adv_px = render_glyph_svg(gid)
            parts.append((glyph_img, x, adv_px))
            x += adv_px

        if x <= 0:
            line_images.append(None)
            continue

        line_img = Image.new("RGBA", (x, line_h), (0, 0, 0, 0))
        for glyph_img, gx, adv_px in parts:
            if glyph_img is None:
                continue
            gh      = min(glyph_img.height, line_h)
            paste_y = max(0, (line_h - gh) // 2)
            clip_w  = min(glyph_img.width, line_img.width - gx)
            if clip_w > 0:
                clip = glyph_img.crop((0, 0, clip_w, gh))
                line_img.paste(clip, (gx, paste_y), clip)

        # Save first line as debug reference
        if not line_images:
            try:
                line_img.save(r"C:\Varsany\Temp\debug_line0.png")
            except Exception:
                pass

        line_images.append(line_img)

    valid = [li for li in line_images if li is not None]
    if not valid:
        log(f"  Glyph-SVG renderer: no visible glyphs for {font_name}", "WARN")
        return None

    # Assemble lines on canvas
    line_gap = max(line_h // 3, 40)  # ~33% of cap height between lines
    n_lines  = len(line_images)
    total_h  = line_h * n_lines + line_gap * max(0, n_lines - 1)
    result   = Image.new("RGBA", (canvas_w, max(1, total_h)), (0, 0, 0, 0))
    y        = 0
    for limg in line_images:
        if limg is not None:
            cx = max(0, (canvas_w - limg.width) // 2)
            result.paste(limg, (cx, y), limg)
        y += line_h + line_gap

    bbox = result.getbbox()
    if not bbox:
        log(f"  Glyph-SVG renderer: final result transparent for {font_name}", "WARN")
        return None

    return result.crop(bbox)

def build_text_layer(text_lines, font_name, colour_hex, w, h):
    if not text_lines:
        return Image.new("RGBA", (1, 1), (0, 0, 0, 0)), 0, 0

    # Premium fonts: try Chrome first (Strategy A renders SVG glyph patterns —
    # mermaid scales, camo, football texture etc.).  If Chrome fails (font has no
    # SVG table) we fall through to PIL with black fill.
    # SVG-only fonts (PIL renders blank) also route through Chrome, passing None
    # for colour so the font's own embedded colours are used.
    if _is_svg_only_font(font_name) or _is_premium_font_name(font_name):
        chrome_col = None if _is_svg_only_font(font_name) else colour_hex
        chrome_img = build_text_layer_chrome(text_lines, font_name, chrome_col, w)
        if chrome_img is not None:
            left = max(0, (w - chrome_img.width) // 2)
            return chrome_img, 0, left
        # Chrome failed (no SVG table) → fall through to PIL

    is_premium = _is_premium_font_name(font_name)
    r, g, b    = (0, 0, 0) if is_premium else hex_to_rgb(colour_hex)
    avail_w  = int(w * 0.90)
    real_lines = [l for l in text_lines if l.strip()]
    if not real_lines:
        return Image.new("RGBA", (1, 1), (0, 0, 0, 0)), 0, 0
    scratch  = Image.new("RGBA", (1, 1))
    draw     = ImageDraw.Draw(scratch)
    lo, hi, best = 20, min(int(h * 0.25), h // max(1, len(real_lines))), 60
    while lo <= hi:
        mid   = (lo + hi) // 2
        font  = get_font(font_name, mid)
        # Measure every line — character count != pixel width (e.g. "STRONG" > "I MAKE")
        widest = max(draw.textbbox((0, 0), l, font=font)[2] - draw.textbbox((0, 0), l, font=font)[0]
                     for l in real_lines)
        if widest <= avail_w:
            best = mid
            lo   = mid + 1
        else:
            hi   = mid - 1
    font   = get_font(font_name, best)
    bb0    = draw.textbbox((0, 0), real_lines[0], font=font)
    line_h = int((bb0[3] - bb0[1]) * 1.4)
    pad    = line_h

    tmp_w = w + pad * 2
    tmp_h = line_h * len(text_lines) + pad * 2
    img   = Image.new("RGBA", (tmp_w, tmp_h), (0, 0, 0, 0))
    yl    = pad

    has_emoji = any(ord(c) > 127 for line in text_lines for c in line)

    if has_emoji and PILMOJI_AVAILABLE:
        # pilmoji renders colour emoji (❤️ 😊 etc.) correctly
        with Pilmoji(img) as pil:
            for line in text_lines:
                if line.strip():
                    bb = draw.textbbox((0, 0), line, font=font)
                    lw = bb[2] - bb[0]
                    x  = max(pad, pad + (w - lw) // 2)
                    pil.text((x, yl), line, fill=(r, g, b, 255), font=font)
                yl += line_h
    else:
        d2 = ImageDraw.Draw(img)
        for line in text_lines:
            if line.strip():
                bb = d2.textbbox((0, 0), line, font=font)
                lw = bb[2] - bb[0]
                x  = max(pad, pad + (w - lw) // 2)
                d2.text((x, yl), line, font=font, fill=(r, g, b, 255))
            yl += line_h

    bbox = img.getbbox()
    if bbox:
        margin = max(10, line_h // 6)
        img = img.crop((max(0, bbox[0] - margin), max(0, bbox[1] - margin),
                        min(tmp_w, bbox[2] + margin), min(tmp_h, bbox[3] + margin)))

    left = max(0, (w - img.width) // 2)
    return img, 0, left

# ─── SKU COLOUR / SIZE PARSING ───────────────────────────────────────────────
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



# Maps SKU colour codes -> approximate RGB of the garment
# Used to detect if the image background matches the garment colour
GARMENT_RGB = {
    "Blk":  (20,  20,  20),
    "Wht":  (255, 255, 255),
    "Nvy":  (31,  40,  80),
    "Red":  (200, 30,  30),
    "Pnk":  (255, 150, 180),
    "BPnk": (255, 182, 193),
    "Gry":  (150, 150, 150),
    "GryM": (160, 160, 160),
    "Blu":  (30,  100, 200),
    "RBlu": (65,  105, 225),
    "SBlu": (135, 206, 235),
    "Grn":  (34,  139, 34),
    "Ylw":  (255, 220, 0),
    "Fus":  (255, 0,   144),
    "Pur":  (128, 0,   128),
    "Org":  (255, 140, 0),
    "Bur":  (128, 0,   32),
    "Nat":  (245, 222, 179),
    "Lav":  (230, 190, 255),
    "Ivry": (255, 255, 240),
}

def get_garment_rgb(sku):
    if not sku:
        return None
    parts = sku.split("_")
    if len(parts) < 2:
        return None
    last = parts[-1]
    for code in sorted(GARMENT_RGB.keys(), key=len, reverse=True):
        if last.startswith(code):
            return GARMENT_RGB[code]
    return None

def _is_light_colour(rgb):
    """True if the garment is a light colour (white, ivory, yellow, light pink etc.)"""
    r, g, b = rgb
    brightness = (r * 299 + g * 587 + b * 114) / 1000
    return brightness > 160

def image_bg_matches_garment(img_rgba, garment_rgb, tolerance=40):
    """
    Returns True when edges are solid-background AND interior check passes.

    Interior check only applies for LIGHT garments (white, ivory, yellow etc.):
      - Light garment: if 25%+ of interior also matches garment colour, it's
        complex artwork (e.g. poster with white smoke/mist) — skip removal.
      - Dark garment (black, navy etc.): skip interior check entirely.
        Dark backgrounds naturally fill the whole canvas; rembg handles them cleanly.
    """
    w, h = img_rgba.size
    if w < 20 or h < 20:
        return False
    strip = 5
    arr = img_rgba.load()
    gr, gg, gb = garment_rgb

    def matches(px):
        return abs(px[0]-gr) <= tolerance and abs(px[1]-gg) <= tolerance and abs(px[2]-gb) <= tolerance

    # --- Edge sample (95% must match garment colour) ---
    edge_px = []
    for x in range(0, w, max(1, w // 100)):
        for y in range(strip):
            edge_px.append(arr[x, y])
            edge_px.append(arr[x, h - 1 - y])
    for y in range(0, h, max(1, h // 100)):
        for x in range(strip):
            edge_px.append(arr[x, y])
            edge_px.append(arr[w - 1 - x, y])
    if not edge_px:
        return False
    edge_match = sum(1 for px in edge_px if matches(px)) / len(edge_px)
    if edge_match < 0.95:
        return False

    # --- Interior check — only for light garments ---
    if _is_light_colour(garment_rgb):
        bx0, by0 = int(w * 0.1), int(h * 0.1)
        bx1, by1 = int(w * 0.9), int(h * 0.9)
        interior_px = []
        step_x = max(1, (bx1 - bx0) // 50)
        step_y = max(1, (by1 - by0) // 50)
        for x in range(bx0, bx1, step_x):
            for y in range(by0, by1, step_y):
                interior_px.append(arr[x, y])
        if interior_px:
            interior_match = sum(1 for px in interior_px if matches(px)) / len(interior_px)
            # 80%+ interior matches → almost entirely background colour, no real design
            # visible (e.g. Kingdom poster with white smoke/mist throughout).
            # Badges/logos with white sections inside score 50-65% — still remove those.
            if interior_match >= 0.80:
                return False

    return True

def remove_background_colourkey(img_rgba, garment_rgb, tolerance=40):
    """
    Fast colour-key removal: replace every pixel close to garment_rgb with transparent.
    Works perfectly for flat graphic designs / logos on solid-colour backgrounds.
    Much more accurate than rembg for non-photographic images.
    """
    import numpy as np
    arr = np.array(img_rgba, dtype=np.int32)
    gr, gg, gb = garment_rgb
    # Distance from each pixel to garment colour
    diff = (np.abs(arr[:,:,0] - gr) <= tolerance) & \
           (np.abs(arr[:,:,1] - gg) <= tolerance) & \
           (np.abs(arr[:,:,2] - gb) <= tolerance)
    result = arr.copy().astype(np.uint8)
    result[diff, 3] = 0   # make matching pixels transparent
    return Image.fromarray(result, 'RGBA')

def remove_background(img_rgba, garment_rgb=None):
    """
    For graphic designs on solid backgrounds → fast colour-key removal.
    For photos (light garments) → rembg AI removal.
    """
    if garment_rgb and not _is_light_colour(garment_rgb):
        # Dark garment (black, navy etc.) — always use colour-key, not rembg
        # rembg destroys dark graphic designs; colour-key is precise
        return remove_background_colourkey(img_rgba, garment_rgb)

    # Light garment — use rembg for photo subjects
    if not REMBG_AVAILABLE:
        log("rembg not installed - falling back to colour-key removal", "WARN")
        if garment_rgb:
            return remove_background_colourkey(img_rgba, garment_rgb)
        return img_rgba
    try:
        result = rembg_remove(img_rgba)
        # Validate rembg output: if fewer than 15% of pixels are visible it
        # destroyed a flat graphic/logo — fall back to colour-key removal instead
        import numpy as _np
        alpha = _np.array(result)[:, :, 3]
        visible_pct = (alpha > 10).sum() / max(1, alpha.size)
        if visible_pct < 0.15:
            log(f"  rembg left only {visible_pct*100:.1f}% visible — flat graphic detected, switching to colour-key", "WARN")
            if garment_rgb:
                return remove_background_colourkey(img_rgba, garment_rgb)
        return result
    except Exception as e:
        log(f"rembg failed: {e} — falling back to colour-key", "WARN")
        if garment_rgb:
            return remove_background_colourkey(img_rgba, garment_rgb)
        return img_rgba

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


def build_label_layer(label_text):
    """Small black label overlay for top-left corner of zone."""
    font_size = max(20, cm_to_px(0.5))
    try:
        f = ImageFont.truetype("arial.ttf", font_size)
    except:
        f = ImageFont.load_default()
    # Measure actual bounding box — bb[0/1] may be non-zero (font descenders)
    tmp = Image.new("RGBA", (1, 1))
    bb  = ImageDraw.Draw(tmp).textbbox((0, 0), label_text.upper(), font=f)
    pad = 6
    tw  = bb[2] - bb[0] + pad * 2
    th  = bb[3] - bb[1] + pad * 2
    img = Image.new("RGBA", (max(1, tw), max(1, th)), (0, 0, 0, 0))
    # Offset by bb[0/1] so text is never clipped
    ImageDraw.Draw(img).text((pad - bb[0], pad - bb[1]), label_text.upper(), font=f, fill=(0, 0, 0, 255))
    return img

# ─── ZONE BUILDER ─────────────────────────────────────────────────────────────

def build_zones(row, product):
    preview_map = {
        "front":  row.get("FrontPreviewImage")  or "",
        "back":   row.get("BackPreviewImage")   or "",
        "sleeve": row.get("SleevePreviewImage") or "",
        "pocket": row.get("PocketPreviewImage") or "",
    }

    sku = row.get("SKU") or ""

    def make_zone(label, zone_key, img_filename=None, text_lines=None, font=None, colour=None):
        w, h = get_dims(product, zone_key)
        return {
            "label":        label,
            "zone_key":     zone_key,
            "w":            w,
            "h":            h,
            "img_path":     find_image(img_filename) if img_filename else None,
            "img_filename": img_filename or "",
            "text_lines":   text_lines or [],
            "font":         font   or "",
            "colour":       colour or (0, 0, 0),
            "preview_url":  preview_map.get(zone_key, ""),
            "sku":          sku,
        }

    zones = []

    # Per-zone fonts and colours
    # Premium fonts have built-in colour/texture — customer colour does not apply
    front_font   = parse_font(row.get("FrontFonts")   or "")
    front_colour = parse_colour(row.get("FrontColours") or "")
    back_font    = parse_font(row.get("BackFonts")    or "") or front_font
    back_colour  = parse_colour(row.get("BackColours") or "") or front_colour
    pocket_font  = parse_font(row.get("PocketFonts")  or "") or front_font
    pocket_colour= parse_colour(row.get("PocketColours") or "") or front_colour
    sleeve_font  = parse_font(row.get("SleeveFonts")  or "") or front_font
    sleeve_colour= parse_colour(row.get("SleeveColours") or "") or front_colour

    # FRONT — up to 5 images (front first so it appears at top of canvas)
    front_imgs = parse_image_json(row.get("FrontImageJSON") or "")
    front_text = parse_texts(row.get("FrontText") or "")
    front_img  = row.get("FrontImage") or ""
    if front_imgs:
        for i, fname in enumerate(front_imgs):
            label = "front" if len(front_imgs) == 1 else f"front {i+1}"
            zones.append(make_zone(label, "front", fname, front_text if i == 0 else [], front_font, front_colour))
    elif front_img:
        zones.append(make_zone("front", "front", front_img, front_text, front_font, front_colour))
    elif front_text:
        zones.append(make_zone("front", "front", text_lines=front_text, font=front_font, colour=front_colour))

    # BACK
    back_imgs = parse_image_json(row.get("BackImageJSON") or "")
    back_text = parse_texts(row.get("BackText") or "")
    back_img  = row.get("BackImage") or ""
    if back_imgs:
        zones.append(make_zone("back", "back", back_imgs[0], back_text, back_font, back_colour))
    elif back_img:
        zones.append(make_zone("back", "back", back_img, back_text, back_font, back_colour))
    elif back_text:
        zones.append(make_zone("back", "back", text_lines=back_text, font=back_font, colour=back_colour))

    # POCKET — pocket left + right if 2 images
    pocket_imgs = parse_image_json(row.get("PocketImageJSON") or "")
    pocket_text = parse_texts(row.get("PocketText") or "")
    pocket_img  = row.get("PocketImage") or ""
    if len(pocket_imgs) >= 2:
        zones.append(make_zone("pocket left",  "pocket", pocket_imgs[0], font=pocket_font, colour=pocket_colour))
        zones.append(make_zone("pocket right", "pocket", pocket_imgs[1], font=pocket_font, colour=pocket_colour))
    elif len(pocket_imgs) == 1:
        zones.append(make_zone("pocket", "pocket", pocket_imgs[0], pocket_text, pocket_font, pocket_colour))
    elif pocket_img:
        zones.append(make_zone("pocket", "pocket", pocket_img, pocket_text, pocket_font, pocket_colour))
    elif pocket_text:
        zones.append(make_zone("pocket", "pocket", text_lines=pocket_text, font=pocket_font, colour=pocket_colour))

    # SLEEVE
    sleeve_imgs = parse_image_json(row.get("SleeveImageJSON") or "")
    sleeve_text = parse_texts(row.get("SleeveText") or "")
    sleeve_img  = row.get("SleeveImage") or ""
    if sleeve_imgs:
        zones.append(make_zone("sleeve", "sleeve", sleeve_imgs[0], sleeve_text, sleeve_font, sleeve_colour))
    elif sleeve_img:
        zones.append(make_zone("sleeve", "sleeve", sleeve_img, sleeve_text, sleeve_font, sleeve_colour))
    elif sleeve_text:
        zones.append(make_zone("sleeve", "sleeve", text_lines=sleeve_text, font=sleeve_font, colour=sleeve_colour))

    return zones

# ─── PSD BUILDER ──────────────────────────────────────────────────────────────

def build_psd_for_order(order_id, row, out_path):
    sku      = row.get("SKU") or ""
    product  = detect_product(sku)
    zones    = build_zones(row, product)
    quantity = max(1, int(row.get("Quantity") or 1))

    if not zones:
        return False, "No zones found — no image or text data"

    PADDING  = cm_to_px(1)    # 1 cm border around whole canvas
    GAP      = cm_to_px(0.5)  # gap between different zones (front/back/sleeve)
    QTY_GAP  = cm_to_px(1.0)  # 1 cm gap between quantity copies (for cutting)

    # Label sits in its own small strip above each zone (not on the image)
    lbl_sample = build_label_layer("front")
    LABEL_H    = lbl_sample.height + cm_to_px(2.0)   # label text height + 2cm gap below it

    TEXT_GAP = cm_to_px(0.3)
    max_zw   = max(z["w"] for z in zones)
    canvas_w = PADDING + max_zw + PADDING

    # First pass: pre-build all layers so we know actual content heights
    for zone in zones:
        zw, zh = zone["w"], zone["h"]
        zone["_img"] = zone["_it"] = zone["_il"] = None
        if zone["img_path"]:
            zone["_img"], zone["_it"], zone["_il"] = build_image_layer(zone["img_path"], zw, zh, sku=sku)
        elif zone["img_filename"]:
            log(f"    WARNING image not found: {zone['img_filename']}", "WARN")

        zone["_txt"] = zone["_tt"] = zone["_tl"] = None
        if zone["text_lines"]:
            zone["_txt"], zone["_tt"], zone["_tl"] = build_text_layer(
                zone["text_lines"], zone["font"], zone["colour"], zw, zh)

        zone["_prev"] = zone["_pnw"] = zone["_pnh"] = None
        if zone.get("preview_url"):
            pi = download_preview(zone["preview_url"])
            if pi:
                ratio = min(zw / pi.width, zh / pi.height)
                pnw = max(1, int(pi.width  * ratio))
                pnh = max(1, int(pi.height * ratio))
                zone["_prev"] = pi.resize((pnw, pnh), Image.LANCZOS)
                zone["_pnw"]  = pnw
                zone["_pnh"]  = pnh

        img_h    = zone["_img"].height if zone["_img"] else 0
        txt_h    = zone["_txt"].height if zone["_txt"] else 0
        raw_h    = txt_h + (TEXT_GAP + img_h if txt_h and img_h else img_h)
        # Always use actual content size — no spec_h padding.
        # Canvas adapts to content; printing team cuts by zone label.
        content_h = raw_h if raw_h > 0 else cm_to_px(1)
        # Per-copy spacing: use raw_h so qty repeats don't have huge gaps
        repeat_h  = raw_h if raw_h > 0 else content_h
        zone["_txt_v_offset"] = 0   # no extra centring — 1cm gap from label is enough
        zone["_txt_h"]    = txt_h
        zone["_img_h"]    = img_h
        zone["_content_h"] = content_h
        zone["_repeat_h"]  = repeat_h

    # Canvas height based on actual rendered content (not spec zone height)
    # For quantity > 1, copies are spaced by repeat_h (actual text/image size),
    # but the last copy still occupies at least content_h (spec cutting area).
    def zone_total_h(z, qty):
        if qty <= 1:
            return z["_content_h"]
        return z["_repeat_h"] * (qty - 1) + QTY_GAP * (qty - 1) + z["_content_h"]

    canvas_h = (PADDING
                + sum(LABEL_H + zone_total_h(z, quantity) for z in zones)
                + GAP * (len(zones) - 1)
                + PADDING)

    all_layers = []
    y_cursor   = PADDING

    for zone in zones:
        zw      = zone["w"]
        x_left  = PADDING + (max_zw - zw) // 2

        display_label = zone.get("display_label") or zone["label"]
        lbl = build_label_layer(display_label)
        all_layers.append({
            "name":    f"{display_label} label",
            "image":   lbl,
            "top":     y_cursor,
            "left":    x_left,
            "opacity": 255,
            "visible": True,
        })

        content_start = y_cursor + LABEL_H
        img_pil   = zone["_img"];  it = zone["_it"];  il = zone["_il"]
        txt_pil   = zone["_txt"];  tt = zone["_tt"];  tl = zone["_tl"]
        prev_img  = zone["_prev"]; pnw = zone["_pnw"]; pnh = zone["_pnh"]
        txt_h     = zone["_txt_h"]
        content_h = zone["_content_h"]
        repeat_h  = zone["_repeat_h"]

        for copy_idx in range(quantity):
            # Use repeat_h as spacing between copies so text-only zones aren't
            # stretched to full spec zone height for every copy
            copy_top = content_start + copy_idx * (repeat_h + QTY_GAP)
            suffix   = f" #{copy_idx + 1}" if quantity > 1 else ""

            # Text goes ABOVE the image (matches Amazon preview layout)
            v_off = zone.get("_txt_v_offset", 0)
            if txt_pil:
                _tl_dict = {
                    "name":    f"{display_label} CustomerText{suffix}",
                    "image":   txt_pil,
                    "top":     copy_top + v_off + tt,
                    "left":    x_left + tl,
                    "opacity": 255,
                    "visible": True,
                }
                if EDITABLE_TEXT_AVAILABLE and zone.get("text_lines"):
                    _ps_font = resolve_ps_font_name(zone.get("font", "arial"))
                    _r, _g, _b = hex_to_rgb(zone.get("colour", "#ffffff"))
                    _tl_dict["_text_blocks"] = build_editable_text_tagged_blocks(
                        text="\n".join(zone["text_lines"]),
                        font_name=_ps_font,
                        font_size_px=txt_pil.height,
                        r=_r, g=_g, b=_b,
                        px_per_cm=PX_PER_CM,
                        layer_left=x_left + tl,
                        layer_top=copy_top + v_off + tt,
                        layer_w=txt_pil.width,
                        layer_h=txt_pil.height,
                    )
                all_layers.append(_tl_dict)

            if img_pil:
                img_top = copy_top + txt_h + (TEXT_GAP if txt_pil else 0) + it
                all_layers.append({
                    "name":    f"{display_label} CustomerImage{suffix}",
                    "image":   img_pil,
                    "top":     img_top,
                    "left":    x_left + il,
                    "opacity": 255,
                    "visible": True,
                })

            # Preview reference: only one per zone (not duplicated for every copy)
            if prev_img and copy_idx == 0:
                all_layers.append({
                    "name":    f"{display_label} Preview Reference",
                    "image":   prev_img,
                    "top":     content_start + (content_h - pnh) // 2,
                    "left":    x_left + (zw - pnw) // 2,
                    "opacity": 255,
                    "visible": False,
                })

        y_cursor += LABEL_H + zone_total_h(zone, quantity) + GAP

    out_path = write_psd(out_path, canvas_w, canvas_h, all_layers)

    if not os.path.isfile(out_path):
        return False, "PSD file not written"

    size_mb    = os.path.getsize(out_path) / (1024 * 1024)
    zone_names = [z["label"] for z in zones]
    return True, f"{size_mb:.1f} MB | zones: {zone_names}"


def rows_have_same_design(rows):
    if len(rows) <= 1:
        return True
    def sig(row):
        return (
            (row.get("FrontText") or "").strip(),
            (row.get("FrontImageJSON") or "").strip(),
            (row.get("FrontImage") or "").strip(),
            (row.get("FrontFonts") or "").strip(),
            (row.get("FrontColours") or "").strip(),
        )
    first = sig(rows[0])
    return all(sig(r) == first for r in rows)


def build_merged_psd_for_order_group(order_id, rows, out_path):
    """
    Builds one merged PSD for an order that has multiple items (rows).

    Owner's rules:
      - All items identical design → stack vertically, label = "front" (no SKU detail)
      - Items have different designs → stack vertically, label = "Front - White XL" etc.
      - 1cm gap between copies (for cutting)
    """
    if not rows:
        return False, "No rows"

    same_design = rows_have_same_design(rows)
    log(f"  Order group: {len(rows)} items, same_design={same_design}", "INFO")

    PADDING  = cm_to_px(1)
    QTY_GAP  = cm_to_px(1.0)
    TEXT_GAP = cm_to_px(0.3)
    lbl_h    = build_label_layer("front").height + cm_to_px(2.0)

    # Build zones for every row, attaching display_label and pre-built layers
    # First pass: collect zones and compute canvas width
    all_row_zones = []
    for row in rows:
        sku     = row.get("SKU") or ""
        product = detect_product(sku)
        zones   = build_zones(row, product)
        for z in zones:
            z["display_label"] = (z["label"].title() if same_design
                                  else make_zone_label(z["label"], sku, use_sku_detail=True))
        all_row_zones.append(zones)

    all_zones_flat = [z for zones in all_row_zones for z in zones]
    if not all_zones_flat:
        return False, "No zones in any row"

    # Canvas width — determined by the widest zone (e.g. adult XXL tshirt)
    max_zw   = max(z["w"] for z in all_zones_flat)

    # Second pass: pre-build layers using each zone's own spec width so smaller
    # garments (e.g. kids tshirt 23cm) render proportionally smaller than adult XXL (30cm)
    for zones in all_row_zones:
        for z in zones:
            draw_w = z["w"]   # use this zone's own spec width, not the widest
            zh     = z["h"]

            z["_img"] = z["_it"] = z["_il"] = None
            if z["img_path"]:
                z["_img"], z["_it"], z["_il"] = build_image_layer(z["img_path"], draw_w, zh, sku=z.get("sku"))
            elif z["img_filename"]:
                log(f"    WARNING image not found: {z['img_filename']}", "WARN")

            z["_txt"] = z["_tt"] = z["_tl"] = None
            if z["text_lines"]:
                z["_txt"], z["_tt"], z["_tl"] = build_text_layer(
                    z["text_lines"], z["font"], z["colour"], draw_w, zh)

            z["_prev"] = z["_pnw"] = z["_pnh"] = None
            if z.get("preview_url"):
                pi = download_preview(z["preview_url"])
                if pi:
                    ratio = min(draw_w / pi.width, zh / pi.height)
                    pnw = max(1, int(pi.width  * ratio))
                    pnh = max(1, int(pi.height * ratio))
                    z["_prev"] = pi.resize((pnw, pnh), Image.LANCZOS)
                    z["_pnw"]  = pnw
                    z["_pnh"]  = pnh

            img_h    = z["_img"].height if z["_img"] else 0
            txt_h    = z["_txt"].height if z["_txt"] else 0
            raw_h    = txt_h + (TEXT_GAP + img_h if txt_h and img_h else img_h)
            content_h = raw_h if raw_h > 0 else cm_to_px(1)
            z["_txt_v_offset"] = 0
            z["_txt_h"]    = txt_h
            z["_img_h"]    = img_h
            z["_content_h"] = content_h

    canvas_w = PADDING + max_zw + PADDING

    GAP = cm_to_px(0.5)  # gap between zones within the same item (front/back/sleeve)

    # Canvas height: each zone stacked vertically, gaps between zones and between items
    def item_height(zones):
        if not zones:
            return 0
        return (sum(lbl_h + z["_content_h"] for z in zones)
                + GAP * (len(zones) - 1))

    canvas_h = (PADDING
                + sum(item_height(zones) for zones in all_row_zones)
                + QTY_GAP * (len(all_row_zones) - 1)
                + PADDING)

    all_layers = []
    y_cursor   = PADDING

    for row_idx, (row, zones) in enumerate(zip(rows, all_row_zones)):
        if not zones:
            continue

        for zone_idx, zone in enumerate(zones):
            display_label = zone["display_label"]
            x_left    = PADDING + (max_zw - zone["w"]) // 2
            img_start = y_cursor + lbl_h

            # Label
            lbl = build_label_layer(display_label)
            all_layers.append({
                "name": f"{display_label} label",
                "image": lbl,
                "top": y_cursor,
                "left": x_left,
                "opacity": 255, "visible": True,
            })

            # Text goes ABOVE the image; for text-only zones it is vertically centred
            v_off = zone.get("_txt_v_offset", 0)
            if zone["_txt"]:
                _tl_dict2 = {
                    "name": f"{display_label} CustomerText",
                    "image": zone["_txt"],
                    "top": img_start + v_off + zone["_tt"],
                    "left": x_left + zone["_tl"],
                    "opacity": 255, "visible": True,
                }
                if EDITABLE_TEXT_AVAILABLE and zone.get("text_lines"):
                    _ps_font2 = resolve_ps_font_name(zone.get("font", "arial"))
                    _r2, _g2, _b2 = hex_to_rgb(zone.get("colour", "#ffffff"))
                    _tl_dict2["_text_blocks"] = build_editable_text_tagged_blocks(
                        text="\n".join(zone["text_lines"]),
                        font_name=_ps_font2,
                        font_size_px=zone["_txt"].height,
                        r=_r2, g=_g2, b=_b2,
                        px_per_cm=PX_PER_CM,
                        layer_left=x_left + zone["_tl"],
                        layer_top=img_start + v_off + zone["_tt"],
                        layer_w=zone["_txt"].width,
                        layer_h=zone["_txt"].height,
                    )
                all_layers.append(_tl_dict2)

            # Image below text
            if zone["_img"]:
                img_top = img_start + zone["_txt_h"] + (TEXT_GAP if zone["_txt"] else 0) + zone["_it"]
                all_layers.append({
                    "name": f"{display_label} CustomerImage",
                    "image": zone["_img"],
                    "top": img_top,
                    "left": x_left + zone["_il"],
                    "opacity": 255, "visible": True,
                })

            # Preview reference (invisible)
            if zone["_prev"]:
                pnh       = zone["_pnh"]
                pnw       = zone["_pnw"]
                content_h = zone["_content_h"]
                all_layers.append({
                    "name":    f"{display_label} Preview Reference",
                    "image":   zone["_prev"],
                    "top":     img_start + (content_h - pnh) // 2,
                    "left":    x_left + (max_zw - pnw) // 2,
                    "opacity": 255,
                    "visible": False,
                })

            # Advance y after each zone
            y_cursor += lbl_h + zone["_content_h"]
            if zone_idx < len(zones) - 1:
                y_cursor += GAP  # gap between zones in same item

        if row_idx < len(all_row_zones) - 1:
            y_cursor += QTY_GAP  # gap between items

    out_path = write_psd(out_path, canvas_w, canvas_h, all_layers)

    if not os.path.isfile(out_path):
        return False, "PSD file not written"

    size_mb = os.path.getsize(out_path) / (1024 * 1024)
    labels  = [z.get("display_label") for zones in all_row_zones for z in zones]
    return True, f"{size_mb:.1f} MB | labels: {labels}"


# ─── DATABASE ─────────────────────────────────────────────────────────────────

def get_db():
    return pyodbc.connect(DB_CONNECTION)

def fetch_orders(limit=None, order_id_filter=None, sku_filter=None, multizone=False, reprocess=False, date_filter=None, date_after=None):
    conn  = get_db()
    cur   = conn.cursor()
    if reprocess:
        where = "1=1"   # skip IsDesignComplete filter when reprocessing
    else:
        where = "(d.IsDesignComplete = 0 OR d.IsDesignComplete IS NULL)"
    if order_id_filter:
        if isinstance(order_id_filter, list):
            # Flatten any comma-separated values in the list
            flat = []
            for item in order_id_filter:
                flat.extend(i.strip() for i in item.split(",") if i.strip())
            ids = "','".join(flat)
            where += f" AND o.OrderID IN ('{ids}')"
        else:
            # Also handle comma-separated string
            parts = [i.strip() for i in order_id_filter.split(",") if i.strip()]
            if len(parts) > 1:
                ids = "','".join(parts)
                where += f" AND o.OrderID IN ('{ids}')"
            else:
                where += f" AND o.OrderID = '{order_id_filter}'"
    if sku_filter:
        like_clauses = " OR ".join(f"o.SKU LIKE '%{s}%'" for s in sku_filter.split(","))
        where += f" AND ({like_clauses})"
    if multizone:
        where += " AND d.PrintLocation LIKE '%+%'"
    if date_filter:
        where += f" AND CAST(o.DateAdd AS DATE) = '{date_filter}'"
    if date_after:
        where += f" AND CAST(o.DateAdd AS DATE) > '{date_after}'"
    top = f"TOP {limit}" if limit else ""
    # Check whether the premium font columns exist (live DB has them, local may not)
    try:
        cur.execute("SELECT TOP 0 FrontPremiumFont FROM tblCustomOrderDetails")
        premium_cols = ", d.FrontPremiumFont, d.BackPremiumFont, d.PocketPremiumFont, d.SleevePremiumFont"
    except Exception:
        premium_cols = ""
        conn.close(); conn = get_db(); cur = conn.cursor()

    cur.execute(f"""
        SELECT {top}
            o.OrderID, o.SKU, o.ItemType, o.Quantity,
            d.idCustomOrderDetails, d.PrintLocation,
            d.FrontText, d.FrontFonts, d.FrontColours,
            d.FrontImage, d.FrontImageJSON, d.FrontPreviewImage,
            d.BackText, d.BackFonts, d.BackColours,
            d.BackImage, d.BackImageJSON, d.BackPreviewImage,
            d.PocketText, d.PocketFonts, d.PocketColours,
            d.PocketImage, d.PocketImageJSON, d.PocketPreviewImage,
            d.SleeveText, d.SleeveImage, d.SleeveImageJSON, d.SleevePreviewImage
            {premium_cols}
        FROM tblCustomOrder o
        JOIN tblCustomOrderDetails d ON o.idCustomOrder = d.idCustomOrder
        WHERE {where}
        ORDER BY o.DateAdd ASC
    """)
    cols = [c[0] for c in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    conn.close()
    return rows

def mark_complete(detail_id, out_path):
    conn = get_db()
    conn.cursor().execute("""
        UPDATE tblCustomOrderDetails
        SET IsDesignComplete = 1,
            IsOrderProcess   = 1,
            ProcessBy        = 'BatchProcessor',
            ProcessTime      = GETDATE(),
            AdditionalPSD    = ?
        WHERE idCustomOrderDetails = ?
    """, out_path, detail_id)
    conn.commit()
    conn.close()

# ─── GOOGLE DRIVE UPLOAD ──────────────────────────────────────────────────────

GDRIVE_CREDENTIALS  = r"C:\Varsany\credentials.json"
GDRIVE_TOKEN        = r"C:\Varsany\gdrive_token.json"
GDRIVE_ROOT_FOLDER  = "1ZObOngMUAQo519ThI0vEckR4waKp7bsj"   # shared Drive folder
GDRIVE_SCOPES       = ["https://www.googleapis.com/auth/drive.file"]

_gdrive_service = None   # module-level cache

def get_gdrive_service():
    global _gdrive_service
    if _gdrive_service:
        return _gdrive_service
    try:
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build

        creds = None
        if os.path.exists(GDRIVE_TOKEN):
            creds = Credentials.from_authorized_user_file(GDRIVE_TOKEN, GDRIVE_SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(GDRIVE_CREDENTIALS, GDRIVE_SCOPES)
                creds = flow.run_local_server(port=0)
            with open(GDRIVE_TOKEN, "w") as f:
                f.write(creds.to_json())

        _gdrive_service = build("drive", "v3", credentials=creds)
        return _gdrive_service
    except Exception as e:
        log(f"Google Drive auth failed: {e}", "WARN")
        return None


def gdrive_get_or_create_folder(service, name, parent_id):
    """Return the Drive folder ID for `name` inside `parent_id`, creating it if needed."""
    q = (f"name='{name}' and mimeType='application/vnd.google-apps.folder' "
         f"and '{parent_id}' in parents and trashed=false")
    res = service.files().list(q=q, fields="files(id,name)", pageSize=1).execute()
    items = res.get("files", [])
    if items:
        return items[0]["id"]
    meta = {
        "name":     name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents":  [parent_id],
    }
    folder = service.files().create(body=meta, fields="id").execute()
    return folder["id"]


def gdrive_upload_file(service, local_path, parent_id):
    """Upload a file to Drive, replacing any existing file with the same name."""
    from googleapiclient.http import MediaFileUpload
    fname = os.path.basename(local_path)

    # Delete existing file with same name to avoid duplicates
    q = f"name='{fname}' and '{parent_id}' in parents and trashed=false"
    existing = service.files().list(q=q, fields="files(id)", pageSize=1).execute().get("files", [])
    for f in existing:
        service.files().delete(fileId=f["id"]).execute()

    mime = "application/octet-stream"
    media = MediaFileUpload(local_path, mimetype=mime, resumable=True, chunksize=10 * 1024 * 1024)
    meta  = {"name": fname, "parents": [parent_id]}
    uploaded = service.files().create(body=meta, media_body=media, fields="id,name,size").execute()
    size_mb = int(uploaded.get("size", 0)) / 1_048_576
    return uploaded["id"], size_mb


def gdrive_upload_psd(local_path, date_str, folder_type, colour_sub):
    """Mirror the local folder structure on Drive and upload the PSD."""
    try:
        service = get_gdrive_service()
        if not service:
            return False

        # Build path: ROOT / date / folder_type / [colour_sub]
        date_id   = gdrive_get_or_create_folder(service, date_str,    GDRIVE_ROOT_FOLDER)
        type_id   = gdrive_get_or_create_folder(service, folder_type, date_id)
        parent_id = gdrive_get_or_create_folder(service, colour_sub, type_id) if colour_sub else type_id

        file_id, size_mb = gdrive_upload_file(service, local_path, parent_id)
        log(f"  Drive: uploaded {os.path.basename(local_path)}  ({size_mb:.1f} MB)", "OK")
        return True
    except Exception as e:
        log(f"  Drive upload failed: {e}", "WARN")
        return False


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def run_batch(limit=None, order_id_filter=None, dry_run=False, sku_filter=None, multizone=False, reprocess=False, date_filter=None, date_after=None, upload_gdrive=False):
    log("=" * 60)
    log(f"Varsany Batch Processor  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"Resolution : {PX_PER_CM} px/cm  ({DPI} DPI)")
    log(f"Images     : {len(IMAGE_INDEX):,} files indexed")
    log(f"Fonts      : {list(FONT_INDEX.keys())}")
    if dry_run:
        log("MODE       : DRY RUN — no files written")
    log("=" * 60)

    orders = fetch_orders(limit=limit, order_id_filter=order_id_filter,
                          sku_filter=sku_filter, multizone=multizone, reprocess=reprocess,
                          date_filter=date_filter, date_after=date_after)
    total  = len(orders)
    log(f"Orders to process: {total}")

    if not orders:
        log("Nothing to process.")
        return

    ok_count   = 0
    fail_count = 0
    today      = datetime.now().strftime("%Y-%m-%d")
    out_dir    = os.path.join(OUTPUT_FOLDER, today)
    os.makedirs(out_dir, exist_ok=True)

    # Group rows by OrderID — one PSD per order (may contain multiple SKUs)
    from collections import OrderedDict
    order_groups = OrderedDict()
    for row in orders:
        oid = row["OrderID"]
        if oid not in order_groups:
            order_groups[oid] = []
        order_groups[oid].append(row)

    total_orders = len(order_groups)
    log(f"Unique orders: {total_orders}  (from {total} rows)")

    for i, (order_id, group_rows) in enumerate(order_groups.items(), 1):
        first_row   = group_rows[0]
        safe_id     = order_id.replace("/", "-")
        sku_raw     = first_row.get("SKU") or ""
        sku         = sku_raw.replace("/", "-").replace("\\", "-")
        is_emb_rhine = any(is_emb_rhine_row(r) for r in group_rows)
        is_multi     = any(is_multizone_row(r) for r in group_rows)
        is_kids_hood = any("kidshoo" in (r.get("SKU") or "").lower() or
                           "gymhoodie" in (r.get("SKU") or "").lower() or
                           "kidshood" in (r.get("SKU") or "").lower()
                           for r in group_rows)
        if is_emb_rhine:
            folder_type = "Emb & Rhine"
        elif is_multi:
            folder_type = "Automated"
        elif is_kids_hood:
            folder_type = "DTF Kids Hoodie"
        else:
            folder_type = "DTF Front"
        colour_sub  = sku_colour_folder(sku_raw)
        if colour_sub:
            cat_dir = os.path.join(out_dir, folder_type, colour_sub)
        else:
            cat_dir = os.path.join(out_dir, folder_type)
        os.makedirs(cat_dir, exist_ok=True)

        # Filename: single item → OrderID_SKU.psd, multi-item → OrderID_SKU_Nitems.psd
        if len(group_rows) == 1:
            base_name = f"{safe_id}_{sku}.psd"
        else:
            base_name = f"{safe_id}_{sku}_{len(group_rows)}items.psd"
        base_path = os.path.join(cat_dir, base_name)
        out_path  = base_path
        counter   = 2
        while os.path.exists(out_path):
            out_path = base_path.replace(".psd", f"_{counter}.psd")
            counter += 1

        skus_str = " | ".join(r.get("SKU", "") for r in group_rows)
        log(f"[{i}/{total_orders}] {order_id}  ({len(group_rows)} items)  |  {skus_str}")

        if dry_run:
            for row in group_rows:
                product = detect_product(row.get("SKU") or "")
                zones   = build_zones(row, product)
                for z in zones:
                    status = "FOUND" if z["img_path"] else ("MISSING" if z["img_filename"] else "text-only")
                    log(f"  [{z['label']}]  img={z['img_filename'] or 'none'} ({status})  text={z['text_lines']}", "DRY")
                if not zones:
                    log("  SKIP — no zones", "DRY")
            continue

        try:
            if len(group_rows) == 1:
                ok, msg = build_psd_for_order(order_id, first_row, out_path)
            else:
                ok, msg = build_merged_psd_for_order_group(order_id, group_rows, out_path)

            if ok:
                for row in group_rows:
                    mark_complete(row["idCustomOrderDetails"], out_path)
                log(f"  OK  {msg}", "OK")
                ok_count += 1
                if upload_gdrive and not dry_run:
                    gdrive_upload_psd(out_path, today, folder_type, colour_sub)
            else:
                log(f"  FAIL  {msg}", "FAIL")
                fail_count += 1
        except Exception as e:
            log(f"  ERROR  {e}", "ERROR")
            log(traceback.format_exc()[-400:], "ERROR")
            fail_count += 1

        if i % 50 == 0:
            log(f"--- Progress {i}/{total_orders}  ok={ok_count}  fail={fail_count} ---")

    log("=" * 60)
    log(f"DONE  {ok_count} OK  |  {fail_count} FAILED  |  Output: {out_dir}")
    log("=" * 60)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Varsany Batch PSD Processor")
    parser.add_argument("--limit",      type=int, default=None, help="Max orders to process")
    parser.add_argument("--order",      type=str, default=None, action="append", help="Process specific OrderID(s) — can be repeated")
    parser.add_argument("--dry-run",    action="store_true",    help="Preview only, no files written")
    parser.add_argument("--dpi",        type=int, default=120,  help="Resolution px/cm (120=304dpi, 320=812dpi)")
    parser.add_argument("--sku-filter", type=str, default=None, help="Comma-separated SKU substrings e.g. MenTee,WmnTee")
    parser.add_argument("--multizone",   action="store_true",    help="Only orders with multiple print zones")
    parser.add_argument("--reprocess",   action="store_true",    help="Re-export already-completed orders")
    parser.add_argument("--date",        type=str, default=None, help="Export orders from a specific date e.g. 2026-02-28")
    parser.add_argument("--date-after",  type=str, default=None, help="Export orders placed after a date e.g. 2026-04-10")
    parser.add_argument("--gdrive",      action="store_true",    help="Upload finished PSDs to Google Drive after export")
    args = parser.parse_args()

    PX_PER_CM = args.dpi
    DPI       = int(PX_PER_CM * 2.54)

    run_batch(
        limit          = args.limit,
        order_id_filter= args.order,
        dry_run        = args.dry_run,
        sku_filter     = args.sku_filter,
        multizone      = args.multizone,
        reprocess      = args.reprocess,
        date_filter    = args.date,
        date_after     = args.date_after,
        upload_gdrive  = args.gdrive,
    )
