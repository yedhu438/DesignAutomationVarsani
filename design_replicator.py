"""
design_replicator.py
====================
Replicates the crssoft.co.uk preview design layout using OpenCV + Pillow.

Layout measured precisely from real crssoft preview images using OpenCV:
  - Customer photo fills almost the entire canvas (top 2% to 95%)
  - Text is overlaid ON TOP of the photo at the BOTTOM (70% to 92%)
  - Text is centred horizontally with a dark shadow for legibility
  - For text-only orders: text is centred vertically in the canvas

Usage:
    python design_replicator.py --order 205-8000463-7311543
"""

import os, sys, json, cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime

# ─── CONFIG ───────────────────────────────────────────────────────────────────
IMAGE_FOLDERS = [
    r"W:\images\Feb-Image",
    r"W:\images\Jan-Image",
    r"C:\Varsany\Uploads",
]
FONT_FOLDERS = [
    r"C:\Varsany\Fonts",
    r"W:\fonts",
]
OUTPUT_FOLDER = r"C:\Varsany\Output"

# Build image index
IMAGE_INDEX = {}
for _f in IMAGE_FOLDERS:
    if os.path.exists(_f):
        for _img in os.listdir(_f):
            IMAGE_INDEX[_img.lower()] = os.path.join(_f, _img)

# Build font index
FONT_INDEX = {}
for _f in FONT_FOLDERS:
    if os.path.exists(_f):
        for _fn in os.listdir(_f):
            if _fn.lower().endswith(('.ttf', '.otf')):
                norm = os.path.splitext(_fn)[0].lower().replace(' ','').replace('-','').replace('_','')
                FONT_INDEX[norm] = os.path.join(_f, _fn)

# ─── LAYOUT CONSTANTS (measured from real crssoft previews with OpenCV) ───────
CANVAS_W = 900
CANVAS_H = 900

# Photo fills almost the entire canvas
PHOTO_TOP_PCT    = 0.02
PHOTO_BOTTOM_PCT = 0.95
PHOTO_LEFT_PCT   = 0.05
PHOTO_RIGHT_PCT  = 0.95

# Text overlaid at BOTTOM of photo (measured: 70% to 92%)
TEXT_TOP_PCT    = 0.70
TEXT_BOTTOM_PCT = 0.92

# Photo border
PHOTO_BORDER_PX    = 3
PHOTO_BORDER_COLOR = (220, 220, 220)

# ─── HELPERS ──────────────────────────────────────────────────────────────────

def find_image(filename):
    if not filename: return None
    fname = filename.strip().lower()
    if fname in IMAGE_INDEX: return IMAGE_INDEX[fname]
    base = os.path.splitext(fname)[0]
    for ext in ['.jpg', '.jpeg', '.png', '.webp']:
        if (base + ext) in IMAGE_INDEX:
            return IMAGE_INDEX[base + ext]
    return None

def parse_image_json(json_str):
    if not json_str or not json_str.strip(): return []
    try:
        d = json.loads(json_str.strip())
        return [d[f"Image{i}"].strip() for i in range(1,6) if d.get(f"Image{i}","").strip()]
    except: return []

def parse_font(fonts_raw):
    if not fonts_raw: return "Arial"
    s = fonts_raw.strip()
    if s.startswith("{"):
        try:
            d = json.loads(s)
            return d.get("NormalFont") or d.get("PremiumFont") or "Arial"
        except: pass
    return s or "Arial"

def parse_colour(colours_raw):
    if not colours_raw: return "#ffffff"
    s = colours_raw.strip()
    if s.startswith("{"):
        try:
            d = json.loads(s)
            return d.get("Colour1") or d.get("colour1") or "#ffffff"
        except: pass
    return s if s.startswith("#") else "#ffffff"

def parse_texts(raw):
    if not raw or not raw.strip(): return []
    return [l.strip() for l in raw.replace('\r\n','\n').split('\n') if l.strip()]

def hex_to_rgb(hex_col):
    h = hex_col.lstrip('#')
    if len(h)==3: h = ''.join(c*2 for c in h)
    try: return tuple(int(h[i:i+2],16) for i in (0,2,4))
    except: return (255,255,255)

def get_pil_font(font_name, size_px):
    norm = font_name.lower().replace(' ','').replace('-','').replace('_','')
    if norm in FONT_INDEX:
        try: return ImageFont.truetype(FONT_INDEX[norm], size_px)
        except: pass
    for key, path in FONT_INDEX.items():
        if norm in key or key in norm:
            try: return ImageFont.truetype(path, size_px)
            except: pass
    for sysf in ['arial.ttf', 'arialbd.ttf']:
        try: return ImageFont.truetype(sysf, size_px)
        except: pass
    return ImageFont.load_default()

# ─── CORE DESIGN FUNCTION ─────────────────────────────────────────────────────

def replicate_design(img_path, text_lines, font_name, colour_hex,
                     canvas_w=CANVAS_W, canvas_h=CANVAS_H):
    """
    Replicates the EXACT crssoft.co.uk preview layout:

    Measured with OpenCV from real preview images:
      - Photo fills canvas: top=2%, bottom=95%, left=5%, right=95%
      - Text overlaid ON TOP of photo at BOTTOM: y = 70% to 92%
      - Text centred horizontally, dark shadow for legibility

    For text-only (no image): text centred in canvas.
    """
    has_text  = bool(text_lines)
    has_image = bool(img_path and os.path.isfile(img_path))

    # Transparent canvas
    canvas = Image.new("RGBA", (canvas_w, canvas_h), (255, 255, 255, 0))
    draw   = ImageDraw.Draw(canvas)

    # ── Place photo ───────────────────────────────────────────────────────────
    if has_image:
        photo_top   = int(canvas_h * PHOTO_TOP_PCT)
        photo_bot   = int(canvas_h * PHOTO_BOTTOM_PCT)
        photo_left  = int(canvas_w * PHOTO_LEFT_PCT)
        photo_right = int(canvas_w * PHOTO_RIGHT_PCT)
        pw = photo_right - photo_left
        ph = photo_bot   - photo_top

        # OpenCV LANCZOS4 = highest quality resize
        cv_img = cv2.imread(img_path, cv2.IMREAD_UNCHANGED)
        if cv_img is not None:
            if len(cv_img.shape) > 2 and cv_img.shape[2] == 4:
                cv_img = cv2.cvtColor(cv_img, cv2.COLOR_BGRA2RGBA)
            else:
                cv_img = cv2.cvtColor(cv_img, cv2.COLOR_BGR2RGBA)
            ih, iw = cv_img.shape[:2]
            scale  = min(pw / iw, ph / ih)
            nw, nh = max(1, int(iw*scale)), max(1, int(ih*scale))
            resized   = cv2.resize(cv_img, (nw, nh), interpolation=cv2.INTER_LANCZOS4)
            pil_photo = Image.fromarray(resized, 'RGBA')
        else:
            pil_photo = Image.open(img_path).convert("RGBA")
            ih, iw    = pil_photo.height, pil_photo.width
            scale     = min(pw / iw, ph / ih)
            nw, nh    = max(1, int(iw*scale)), max(1, int(ih*scale))
            pil_photo = pil_photo.resize((nw, nh), Image.LANCZOS)

        # Centre photo, thin border
        x_off = photo_left + (pw - nw) // 2
        y_off = photo_top  + (ph - nh) // 2
        if PHOTO_BORDER_PX > 0:
            bc = PHOTO_BORDER_COLOR
            draw.rectangle([x_off-PHOTO_BORDER_PX, y_off-PHOTO_BORDER_PX,
                            x_off+nw+PHOTO_BORDER_PX, y_off+nh+PHOTO_BORDER_PX],
                           fill=(bc[0], bc[1], bc[2], 255))
        canvas.paste(pil_photo, (x_off, y_off), pil_photo)

    # ── Render text ───────────────────────────────────────────────────────────
    if has_text and text_lines:
        r, g, b    = hex_to_rgb(colour_hex)
        text_color = (r, g, b, 255)

        if has_image:
            # Text zone: bottom of photo (measured 70%-92%)
            txt_zone_top = int(canvas_h * TEXT_TOP_PCT)
            txt_zone_h   = int(canvas_h * (TEXT_BOTTOM_PCT - TEXT_TOP_PCT))
        else:
            # Text only: centred vertically
            txt_zone_top = int(canvas_h * 0.25)
            txt_zone_h   = int(canvas_h * 0.50)

        avail_w = int(canvas_w * 0.90)
        longest = max(text_lines, key=len)

        # Binary search for best font size
        scratch = Image.new("RGBA", (1, 1))
        sd      = ImageDraw.Draw(scratch)
        lo, hi, best = 18, txt_zone_h // max(1, len(text_lines)), 120
        while lo <= hi:
            mid  = (lo + hi) // 2
            font = get_pil_font(font_name, mid)
            bb   = sd.textbbox((0, 0), longest, font=font)
            if (bb[2]-bb[0]) <= avail_w:
                best = mid; lo = mid + 1
            else:
                hi = mid - 1

        font   = get_pil_font(font_name, best)
        bb0    = sd.textbbox((0, 0), text_lines[0], font=font)
        line_h = int((bb0[3]-bb0[1]) * 1.25)
        total_h = line_h * len(text_lines)

        # Vertically centre text within its zone
        yt = txt_zone_top + max(0, (txt_zone_h - total_h) // 2)

        for line in text_lines:
            bb = draw.textbbox((0, 0), line, font=font)
            lw = bb[2] - bb[0]
            xt = max(0, (canvas_w - lw) // 2)
            # Shadow for legibility on any background
            draw.text((xt+2, yt+2), line, font=font, fill=(0, 0, 0, 130))
            draw.text((xt,   yt),   line, font=font, fill=text_color)
            yt += line_h

    return canvas


def replicate_from_row(row, out_path):
    """Build design PNG from a database row dict."""
    font_name  = parse_font(row.get("FrontFonts") or "")
    colour_hex = parse_colour(row.get("FrontColours") or "")
    text_lines = parse_texts(row.get("FrontText") or "")

    front_imgs = parse_image_json(row.get("FrontImageJSON") or "")
    front_img  = row.get("FrontImage") or ""
    img_file   = front_imgs[0] if front_imgs else front_img
    img_path   = find_image(img_file)

    canvas = replicate_design(img_path, text_lines, font_name, colour_hex)
    canvas.save(out_path, "PNG")
    return out_path


# ─── CLI ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse, pyodbc

    parser = argparse.ArgumentParser()
    parser.add_argument("--order", required=True)
    parser.add_argument("--out",   default=None)
    args = parser.parse_args()

    DB = ("DRIVER={ODBC Driver 17 for SQL Server};"
          "SERVER=localhost\\SQLEXPRESS;DATABASE=dbAmazonCustomOrders;"
          "Trusted_Connection=yes;TrustServerCertificate=yes;")

    conn = pyodbc.connect(DB)
    cur  = conn.cursor()
    cur.execute("""
        SELECT o.OrderID, o.SKU, d.FrontText, d.FrontFonts, d.FrontColours,
               d.FrontImage, d.FrontImageJSON
        FROM tblCustomOrder o
        JOIN tblCustomOrderDetails d ON o.idCustomOrder=d.idCustomOrder
        WHERE o.OrderID=?
    """, args.order)
    cols = [c[0] for c in cur.description]
    rows = cur.fetchall()
    conn.close()

    if not rows:
        print(f"Order {args.order} not found")
        sys.exit(1)

    row = dict(zip(cols, rows[0]))
    print(f"Order : {row['OrderID']}")
    print(f"SKU   : {row['SKU']}")
    print(f"Text  : {row['FrontText']}")
    print(f"Font  : {row['FrontFonts']}")
    print(f"Colour: {row['FrontColours']}")

    out = args.out or os.path.join(OUTPUT_FOLDER, f"design_{args.order}.png")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    replicate_from_row(row, out)
    print(f"Saved : {out}")
