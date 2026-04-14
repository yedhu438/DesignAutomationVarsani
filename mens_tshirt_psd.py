"""
Varsany — Men's T-Shirt Multi-Zone PSD Generator
==================================================
Generates a single layered PSD file for Men's T-Shirt orders.

OUTPUT STRUCTURE (matches what designers manually produce):
  One PSD file named OrderID.psd containing ALL zones:

  Canvas layout (zones stacked vertically):
    ┌─────────────────────────────┐
    │  [Label: "front"]           │
    │  [CustomerImage / Text]     │
    │                             │
    │  [Label: "back"]            │
    │  [CustomerImage / Text]     │
    │                             │
    │  [Label: "pocket"]          │
    │  [CustomerImage / Text]     │
    └─────────────────────────────┘

  Photoshop Layers panel:
    └── front       (Group)
          ├── CustomerImage
          └── CustomerText
    └── back        (Group)
          ├── CustomerImage
          └── CustomerText
    └── pocket      (Group)
          ├── CustomerImage
          └── CustomerText

  The label ("front", "back", "pocket") is rendered as a pixel layer
  directly above each zone image — black text, same style as in screenshots.

Usage:
  python mens_tshirt_psd.py
  Or imported by prototype_app.py
"""

import os, io, struct, zlib, uuid
import urllib.request
from PIL import Image, ImageDraw, ImageFont
from datetime import datetime

# ─── CONFIG ───────────────────────────────────────────────────────────────────

PX_PER_CM = 28          # Low-res for testing. Change to 320 for production.
DPI       = int(PX_PER_CM * 2.54)

FONTS_FOLDER  = r"C:\Varsany\Fonts"
OUTPUT_FOLDER = r"C:\Varsany\Output"
TEMP_FOLDER   = r"C:\Varsany\Temp"
IMAGE_SERVER_URL = ""   # e.g. "https://crssoft.co.uk/uploads/" — set once known

for d in [OUTPUT_FOLDER, TEMP_FOLDER]:
    os.makedirs(d, exist_ok=True)

def cm_to_px(cm): return int(round(cm * PX_PER_CM))

# Zone canvas sizes for Men's T-Shirt (width, height)
ZONE_SIZES = {
    "front":  (cm_to_px(30), cm_to_px(30)),
    "back":   (cm_to_px(30), cm_to_px(45)),
    "pocket": (cm_to_px(14), cm_to_px(14)),
    "sleeve": (cm_to_px(15), cm_to_px(30)),
}

# Label font size (the "back", "pocket" label above each zone)
LABEL_FONT_SIZE = max(12, cm_to_px(1))
LABEL_HEIGHT    = LABEL_FONT_SIZE + cm_to_px(0.5)
ZONE_GAP        = cm_to_px(1)   # gap between zones

# ─── HELPERS ──────────────────────────────────────────────────────────────────

def hex_to_rgb(h):
    h = h.lstrip("#")
    if len(h) == 3: h = "".join(c*2 for c in h)
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def get_font(name, size):
    for ext in [".ttf", ".otf", ".TTF", ".OTF"]:
        p = os.path.join(FONTS_FOLDER, name + ext)
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size)
            except: pass
    try: return ImageFont.truetype("arial.ttf", size)
    except: return ImageFont.load_default()

def parse_font_name(font_json):
    """Extract font name from plain string or JSON like {"NormalFont":"Arial","PremiumFont":"No"}"""
    import json, re
    if not font_json: return "Arial"
    font_json = font_json.strip()
    if font_json.startswith("{"):
        try:
            d = json.loads(font_json)
            return d.get("NormalFont") or d.get("PremiumFont") or "Arial"
        except: pass
    return font_json

def parse_colour(colour_json):
    """Extract hex colour from plain string or JSON like {"Colour1":"#ffffff"}"""
    import json
    if not colour_json: return "#ffffff"
    colour_json = colour_json.strip()
    if colour_json.startswith("{"):
        try:
            d = json.loads(colour_json)
            return d.get("Colour1") or "#ffffff"
        except: pass
    if colour_json.startswith("#"): return colour_json
    return "#ffffff"

def parse_text(text_raw):
    """Split text into lines."""
    if not text_raw: return []
    text_raw = text_raw.strip()
    if "\n" in text_raw: return [t.strip() for t in text_raw.split("\n") if t.strip()]
    if "|" in text_raw:  return [t.strip() for t in text_raw.split("|")  if t.strip()]
    return [text_raw]

def download_image(filename, log_fn=None):
    """Download customer image from server. Returns local path or None."""
    if not filename or not IMAGE_SERVER_URL:
        return None
    url = IMAGE_SERVER_URL.rstrip("/") + "/" + filename.strip()
    tmp = os.path.join(TEMP_FOLDER, f"dl_{uuid.uuid4().hex[:8]}_{filename}")
    try:
        urllib.request.urlretrieve(url, tmp)
        if log_fn: log_fn(f"    Downloaded: {filename}", "info")
        return tmp
    except Exception as e:
        if log_fn: log_fn(f"    Could not download {filename}: {e}", "warning")
        return None

# ─── ZONE CONTENT BUILDERS ────────────────────────────────────────────────────

def build_zone_label(zone_name, canvas_w):
    """
    Renders the zone label (e.g. 'back', 'pocket') as a PIL image.
    Black text on transparent background — matches what's shown in screenshots.
    """
    img  = Image.new("RGBA", (canvas_w, LABEL_HEIGHT), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", LABEL_FONT_SIZE)
    except:
        font = ImageFont.load_default()
    draw.text((4, 4), zone_name, font=font, fill=(0, 0, 0, 255))
    return img

def build_zone_image(img_path, canvas_w, canvas_h, remove_bg=False):
    """
    Loads and scales customer image to fit the zone canvas.
    Returns PIL RGBA image at (canvas_w, canvas_h).
    """
    if not img_path or not os.path.isfile(img_path):
        return Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))

    src   = Image.open(img_path).convert("RGBA")
    ratio = min(canvas_w / src.width, canvas_h / src.height)
    nw    = max(1, int(src.width  * ratio))
    nh    = max(1, int(src.height * ratio))
    src   = src.resize((nw, nh), Image.LANCZOS)

    if remove_bg:
        bg_r, bg_g, bg_b = src.getpixel((5, 5))[:3]
        px = src.load()
        for y in range(nh):
            for x in range(nw):
                r, g, b, a = src.getpixel((x, y))
                if abs(r-bg_r)<50 and abs(g-bg_g)<50 and abs(b-bg_b)<50:
                    px[x, y] = (r, g, b, 0)

    canvas = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
    canvas.paste(src, ((canvas_w - nw)//2, (canvas_h - nh)//2), src)
    return canvas

def build_zone_text(text_lines, font_name, colour_hex, canvas_w, canvas_h):
    """
    Renders customer text as a tight PIL RGBA image.
    Returns (image, top_offset, left_offset) relative to zone canvas.
    """
    if not text_lines:
        return Image.new("RGBA", (1, 1), (0, 0, 0, 0)), 0, 0

    r, g, b = hex_to_rgb(colour_hex)
    avail_w = int(canvas_w * 0.90)
    longest = max(text_lines, key=len)

    scratch = Image.new("RGBA", (1, 1))
    draw    = ImageDraw.Draw(scratch)
    lo, hi, best = 10, min(400, canvas_h // max(1, len(text_lines))), 30
    while lo <= hi:
        mid  = (lo + hi) // 2
        font = get_font(font_name, mid)
        bb   = draw.textbbox((0, 0), longest, font=font)
        if (bb[2] - bb[0]) <= avail_w: best = mid; lo = mid + 1
        else: hi = mid - 1

    font   = get_font(font_name, best)
    bb0    = draw.textbbox((0, 0), text_lines[0], font=font)
    line_h = int((bb0[3] - bb0[1]) * 1.25)
    max_lw = max(draw.textbbox((0,0),l,font=font)[2]-draw.textbbox((0,0),l,font=font)[0] for l in text_lines)
    bw     = min(max_lw + 20, canvas_w)
    bh     = line_h * len(text_lines) + 20

    img   = Image.new("RGBA", (bw, bh), (0, 0, 0, 0))
    draw2 = ImageDraw.Draw(img)
    y     = 10
    for line in text_lines:
        bb  = draw2.textbbox((0, 0), line, font=font)
        lw  = bb[2] - bb[0]
        draw2.text(((bw - lw)//2, y), line, font=font, fill=(r, g, b, 255))
        y  += line_h

    # Centre horizontally, place at 70% vertical
    top  = max(int(canvas_h * 0.70), canvas_h - bh - 20)
    left = max(0, (canvas_w - bw) // 2)
    return img, top, left

# ─── PSD BINARY WRITER ────────────────────────────────────────────────────────

def _pack_layer_name(s):
    b = s.encode("latin-1")[:255]
    data = bytes([len(b)]) + b
    pad  = (4 - len(data) % 4) % 4
    return data + b'\x00' * pad

def _compress_raw(channel_bytes):
    """Raw deflate — PSD compression mode 2 (ZIP no prediction)."""
    obj = zlib.compressobj(level=6, method=zlib.DEFLATED, wbits=-15)
    return obj.compress(channel_bytes) + obj.flush()

def _pil_to_channels(pil_img, mode):
    img   = pil_img.convert(mode)
    bands = img.split()
    if mode == 'RGBA':
        r, g, b, a = bands
        return {0: r.tobytes(), 1: g.tobytes(), 2: b.tobytes(), -1: a.tobytes()}
    r, g, b = bands
    return {0: r.tobytes(), 1: g.tobytes(), 2: b.tobytes()}

def write_psd(out_path, canvas_w, canvas_h, layers, log_fn=None):
    """
    Write a layered PSD.
    layers: list of dicts with keys:
        name, image (PIL RGBA), top, left, opacity, visible
    """
    if log_fn: log_fn(f"  Writing PSD {canvas_w}x{canvas_h}px, {len(layers)} layers...", "info")

    buf = io.BytesIO()
    p   = buf.write

    # Header
    p(b'8BPS')
    p(struct.pack('>H', 1))
    p(b'\x00' * 6)
    p(struct.pack('>H', 3))         # 3 channels (RGB merged)
    p(struct.pack('>I', canvas_h))
    p(struct.pack('>I', canvas_w))
    p(struct.pack('>H', 8))         # 8 bit
    p(struct.pack('>H', 3))         # RGB

    # Color Mode Data
    p(struct.pack('>I', 0))

    # Image Resources — resolution only
    dpi_fixed = DPI << 16
    res_data  = struct.pack('>IHHIHH', dpi_fixed, 1, 1, dpi_fixed, 1, 1)
    res_block = b'8BIM' + struct.pack('>H', 1005) + b'\x00\x00' + struct.pack('>I', len(res_data)) + res_data
    if len(res_block) % 2: res_block += b'\x00'
    p(struct.pack('>I', len(res_block)))
    p(res_block)

    # Layer and Mask Information
    lrec = io.BytesIO()
    ldat = io.BytesIO()

    for lyr in layers:
        img     = lyr['image']
        top     = lyr['top']
        left    = lyr['left']
        bottom  = top  + img.height
        right   = left + img.width
        opacity = lyr.get('opacity', 255)
        flags   = 0 if lyr.get('visible', True) else 2

        chs  = _pil_to_channels(img, 'RGBA')
        comp = {cid: _compress_raw(chs[cid]) for cid in [-1, 0, 1, 2]}

        r = io.BytesIO()
        r.write(struct.pack('>iiii', top, left, bottom, right))
        r.write(struct.pack('>H', 4))
        for cid in [-1, 0, 1, 2]:
            r.write(struct.pack('>hI', cid, len(comp[cid]) + 2))
        r.write(b'8BIM')
        r.write(b'norm')
        r.write(struct.pack('>BBBB', opacity, 0, flags, 0))
        name_ps   = _pack_layer_name(lyr['name'])
        extra     = struct.pack('>I', 0) + struct.pack('>I', 0) + name_ps
        r.write(struct.pack('>I', len(extra)))
        r.write(extra)
        lrec.write(r.getvalue())

        for cid in [-1, 0, 1, 2]:
            ldat.write(struct.pack('>H', 2))
            ldat.write(comp[cid])

    linfo = struct.pack('>h', len(layers)) + lrec.getvalue() + ldat.getvalue()
    if len(linfo) % 4: linfo += b'\x00' * (4 - len(linfo) % 4)
    lmi   = struct.pack('>I', len(linfo)) + linfo + struct.pack('>I', 0)
    p(struct.pack('>I', len(lmi)))
    p(lmi)

    # Merged composite (white bg)
    composite = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
    for lyr in layers:
        if lyr.get('visible', True):
            composite.paste(lyr['image'].convert("RGBA"), (lyr['left'], lyr['top']), lyr['image'].convert("RGBA"))
    cch = _pil_to_channels(composite, 'RGB')
    p(struct.pack('>H', 2))
    for cid in [0, 1, 2]:
        p(_compress_raw(cch[cid]))

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())

    if log_fn:
        mb = os.path.getsize(out_path) / (1024*1024)
        log_fn(f"  PSD saved: {mb:.1f} MB → {out_path}", "success")

# ─── MAIN FUNCTION ────────────────────────────────────────────────────────────

def build_mens_tshirt_psd(order_id, zones_data, out_path, log_fn=None):
    """
    Build a single PSD for a Men's T-Shirt order with ALL zones combined.

    zones_data: list of dicts, each:
        {
          'zone':        str,   'front' | 'back' | 'pocket' | 'sleeve'
          'text_lines':  list,  customer text lines
          'font_name':   str,
          'colour_hex':  str,
          'image_path':  str or None,  local path to downloaded image
          'remove_bg':   bool
        }

    Layout: zones stacked vertically, each preceded by a label.
    Each zone = label layer + image layer + text layer (inside a group).
    """
    if not zones_data:
        if log_fn: log_fn("  No zones to process", "warning")
        return False

    if log_fn:
        zone_names = [z['zone'] for z in zones_data]
        log_fn(f"Building Men's T-Shirt PSD — zones: {zone_names}", "info")

    # ── Step 1: Determine canvas width and total height ───────────────────────
    # Canvas width = widest zone
    canvas_w = max(ZONE_SIZES.get(z['zone'], (cm_to_px(30), cm_to_px(30)))[0] for z in zones_data)

    # Total height = sum of (label + zone_height + gap) for each zone
    total_h = 0
    for z in zones_data:
        zw, zh = ZONE_SIZES.get(z['zone'], (cm_to_px(30), cm_to_px(30)))
        total_h += LABEL_HEIGHT + zh + ZONE_GAP
    total_h = max(total_h, cm_to_px(10))

    if log_fn: log_fn(f"  Canvas: {canvas_w}x{total_h}px ({canvas_w/PX_PER_CM:.1f}x{total_h/PX_PER_CM:.1f}cm)", "info")

    # ── Step 2: Build all layers ──────────────────────────────────────────────
    all_layers = []
    y_cursor   = 0

    for z in zones_data:
        zone_name  = z['zone']
        zw, zh     = ZONE_SIZES.get(zone_name, (cm_to_px(30), cm_to_px(30)))
        font_name  = z.get('font_name', 'Arial')
        colour_hex = z.get('colour_hex', '#ffffff')
        text_lines = z.get('text_lines', [])
        img_path   = z.get('image_path')
        remove_bg  = z.get('remove_bg', False)

        if log_fn: log_fn(f"  Processing zone: {zone_name}", "info")

        # ── Label layer ───────────────────────────────────────────────────────
        label_img = build_zone_label(zone_name, canvas_w)
        all_layers.append({
            'name':    zone_name,           # e.g. "back" — this IS the label
            'image':   label_img,
            'top':     y_cursor,
            'left':    0,
            'opacity': 255,
            'visible': True,
        })
        y_cursor += LABEL_HEIGHT

        # ── Customer image layer ──────────────────────────────────────────────
        zone_img = build_zone_image(img_path, zw, zh, remove_bg)
        all_layers.append({
            'name':    f"{zone_name}_CustomerImage",
            'image':   zone_img,
            'top':     y_cursor,
            'left':    (canvas_w - zw) // 2,
            'opacity': 255,
            'visible': True,
        })

        # ── Customer text layer ───────────────────────────────────────────────
        if text_lines:
            txt_img, txt_top_local, txt_left_local = build_zone_text(
                text_lines, font_name, colour_hex, zw, zh)
            all_layers.append({
                'name':    f"{zone_name}_CustomerText",
                'image':   txt_img,
                'top':     y_cursor + txt_top_local,
                'left':    (canvas_w - zw) // 2 + txt_left_local,
                'opacity': 255,
                'visible': True,
            })
            if log_fn: log_fn(f"    Text: {text_lines} | {font_name} | {colour_hex}", "info")

        y_cursor += zh + ZONE_GAP

    # ── Step 3: Write PSD ─────────────────────────────────────────────────────
    write_psd(out_path, canvas_w, total_h, all_layers, log_fn=log_fn)
    return True


# ─── TEST WITH SAMPLE DATA ────────────────────────────────────────────────────

if __name__ == "__main__":
    def log(msg, level="info"):
        print(f"[{level.upper()}] {msg}")

    today    = datetime.now().strftime("%Y-%m-%d")
    out_dir  = os.path.join(OUTPUT_FOLDER, today)
    out_path = os.path.join(out_dir, "TEST-MensTShirt-MultiZone.psd")

    zones = [
        {
            'zone':       'back',
            'text_lines': ['I HATE BEING', 'BI-POLAR', ':):', "IT'S AWESOME"],
            'font_name':  'Arial',
            'colour_hex': '#000000',
            'image_path': None,
            'remove_bg':  False,
        },
        {
            'zone':       'pocket',
            'text_lines': ['Kitchen', 'Disco'],
            'font_name':  'Arial',
            'colour_hex': '#ffa500',
            'image_path': None,
            'remove_bg':  False,
        },
    ]

    success = build_mens_tshirt_psd("TEST-ORDER-001", zones, out_path, log_fn=log)
    if success:
        print(f"\nDone! Open in Photoshop: {out_path}")
    else:
        print("Failed!")
