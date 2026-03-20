"""
Varsany Print Automation — Prototype
=====================================
PSD Engine: Pure Python struct writer (zero NumPy, zero psd-tools for writing)
  - Writes real layered PSD using Python's built-in struct + zlib
  - Zero memory issues — only processes actual content pixels
  - Output: 20-80 MB layered PSD depending on image content
  - Layer structure:
      PSDImage (RGB, 320px/cm)
        ├── Background       (white fill)
        ├── CustomerImage    (PixelLayer — customer graphic, RGBA)
        └── CustomerText     (PixelLayer — text rendered via Pillow, RGBA)

Install:
  pip install flask pyodbc pillow

Run:
  python prototype_app.py   →   http://localhost:5000
"""

import os, uuid, threading, time, struct, zlib, io
from datetime import datetime
from flask import (Flask, render_template_string, request,
                   jsonify, send_from_directory)
from PIL import Image, ImageDraw, ImageFont
import pyodbc

app = Flask(__name__)
app.secret_key = "varsany-prototype-2026"

# ─── CONFIG ───────────────────────────────────────────────────────────────────

DB_CONNECTION = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost\\SQLEXPRESS;"
    "DATABASE=dbAmazonCustomOrders;"
    "Trusted_Connection=yes;"
    "TrustServerCertificate=yes;"
)

UPLOAD_FOLDER = r"C:\Varsany\Uploads"
OUTPUT_FOLDER = r"C:\Varsany\Output"
FONTS_FOLDER  = r"C:\Varsany\Fonts"
TEMP_FOLDER   = r"C:\Varsany\Temp"

for _d in [UPLOAD_FOLDER, OUTPUT_FOLDER, FONTS_FOLDER, TEMP_FOLDER]:
    os.makedirs(_d, exist_ok=True)

# LOW-RES MODE  — change back to 320 (812 DPI) when ready for production
PX_PER_CM = 28                        # ≈72 DPI: canvas ~840px for 30cm zone
DPI        = int(PX_PER_CM * 2.54)   # ~71

def cm_to_px(cm): return int(round(cm * PX_PER_CM))

PRODUCT_CANVAS = {
    "hoodie":     {"front":  (cm_to_px(30), cm_to_px(30)),
                   "back":   (cm_to_px(30), cm_to_px(45)),
                   "sleeve": (cm_to_px(15), cm_to_px(30)),
                   "pocket": (cm_to_px(30), cm_to_px(30))},
    "tshirt":     {"front":  (cm_to_px(30), cm_to_px(30)),
                   "back":   (cm_to_px(30), cm_to_px(45)),
                   "sleeve": (cm_to_px(15), cm_to_px(30)),
                   "pocket": (cm_to_px(30), cm_to_px(30))},
    "kidstshirt": {"front":  (cm_to_px(23), cm_to_px(30)),
                   "back":   (cm_to_px(23), cm_to_px(30))},
    "totebag":    {"front":  (cm_to_px(27), cm_to_px(27)),
                   "back":   (cm_to_px(27), cm_to_px(59))},
    "slipper":    {"front":  (cm_to_px(11), cm_to_px(7))},
    "babyvest":   {"front":  (cm_to_px(15), cm_to_px(17))},
}

progress_logs = {}

# ─── HELPERS ──────────────────────────────────────────────────────────────────

def log_progress(order_id, message, level="info"):
    if order_id not in progress_logs:
        progress_logs[order_id] = []
    entry = {"time": datetime.now().strftime("%H:%M:%S"),
             "message": message, "level": level}
    progress_logs[order_id].append(entry)
    print(f"[{entry['time']}] {message}")


def hex_to_rgb(hex_col):
    h = hex_col.lstrip("#")
    if len(h) == 3: h = "".join(c*2 for c in h)
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def get_font(font_name, size_px):
    for ext in [".ttf", ".otf", ".TTF", ".OTF"]:
        p = os.path.join(FONTS_FOLDER, font_name + ext)
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size_px)
            except: pass
    try: return ImageFont.truetype("arial.ttf", size_px)
    except: return ImageFont.load_default()


def parse_texts(raw):
    if not raw: return []
    if "\n" in raw: return [t.strip() for t in raw.split("\n") if t.strip()]
    if "|"  in raw: return [t.strip() for t in raw.split("|")  if t.strip()]
    return [raw.strip()]


# ─── DATABASE ─────────────────────────────────────────────────────────────────

def get_db():
    return pyodbc.connect(DB_CONNECTION)


def save_order_to_db(order_data):
    conn      = get_db()
    cur       = conn.cursor()
    oid       = str(uuid.uuid4())
    did       = str(uuid.uuid4())
    amazon_id = f"PROTO-{datetime.now().strftime('%Y%m%d%H%M%S')}"

    cur.execute("""
        INSERT INTO tblCustomOrder
        (idCustomOrder, OrderID, OrderItemID, ASIN, SKU, Quantity,
         ItemType, Gender, BuyerName, IsCustomOrderDetailsGet, IsShipped, DateAdd)
        VALUES (?,?,?,?,?,1,?,?,?,1,0,GETDATE())
    """, oid, amazon_id, str(uuid.uuid4())[:8], "PROTO-ASIN",
        order_data["sku"], order_data["product"], "Unisex", "Prototype Customer")

    zone = order_data["zone"]
    z    = zone.capitalize()
    sql = f"""
        INSERT INTO tblCustomOrderDetails
        (idCustomOrderDetails, idCustomOrder, PrintLocation,
         IsFrontLocation, IsBackLocation, IsSleeveLocation, IsPocketLocation,
         {z}Image, {z}Text, {z}Fonts, {z}Colours,
         IsOrderProcess, IsDesignComplete,
         IsFrontPSDDownload, IsBackPSDDownload, IsSleevePSDDownload, IsPocketPSDDownload,
         DateAdd)
        VALUES (?,?,?, ?,?,?,?, ?,?,?,?, 0,0, 0,0,0,0, GETDATE())
    """
    cur.execute(sql,
        did, oid, zone,
        1 if zone == "front"  else 0,
        1 if zone == "back"   else 0,
        1 if zone == "sleeve" else 0,
        1 if zone == "pocket" else 0,
        os.path.basename(order_data.get("image_path", "")),
        (order_data["text"]   or "")[:500],
        (order_data["font"]   or "")[:100],
        (order_data["colour"] or "")[:50])

    conn.commit()
    conn.close()
    return oid, did, amazon_id


def mark_order_complete(detail_id, output_path):
    conn = get_db()
    conn.cursor().execute("""
        UPDATE tblCustomOrderDetails
        SET IsDesignComplete=1, IsOrderProcess=1,
            ProcessBy='AutomationPrototype', ProcessTime=GETDATE(),
            AdditionalPSD=?
        WHERE idCustomOrderDetails=?
    """, output_path, detail_id)
    conn.commit()
    conn.close()


def get_recent_orders():
    try:
        conn = get_db()
        cur  = conn.cursor()
        cur.execute("""
            SELECT TOP 10
                o.OrderID, o.ItemType, d.PrintLocation,
                d.FrontText, d.BackText, d.SleeveText,
                d.IsDesignComplete, d.ProcessTime, d.AdditionalPSD
            FROM tblCustomOrderDetails d
            JOIN tblCustomOrder o ON d.idCustomOrder = o.idCustomOrder
            WHERE o.OrderID LIKE 'PROTO-%'
            ORDER BY d.DateAdd DESC
        """)
        cols = [c[0] for c in cur.description]
        rows = [dict(zip(cols, r)) for r in cur.fetchall()]
        conn.close()
        return rows
    except:
        return []


# ─── PURE PYTHON PSD WRITER ───────────────────────────────────────────────────
#
# Writes a valid Adobe PSD file using only Python struct + zlib.
# Zero NumPy, zero psd-tools, zero external dependencies.
#
# Why this works where psd-tools fails:
#   psd-tools.PSDImage.new() immediately allocates a full-canvas NumPy float32
#   array for its internal compositing engine, even if you never composite.
#   For a 9600×9600 RGBA canvas that's 9600×9600×4×4 = 1.4 GB.
#   This writer only allocates memory for the actual content pixels.
#
# PSD format reference: adobe.com/devnet-apps/photoshop/fileformatashtml/
# Sections written: Header → Color Mode Data → Image Resources →
#                   Layer and Mask Info → Image Data (merged composite)
#
# Layer compression: ZIP without prediction (compression=2).
# Each channel is deflated independently using zlib — typically 10:1
# ratio on print-res images, keeping files well under 100 MB.

def _pack_pascal_string(s: str) -> bytes:
    """Pascal string padded to 2-byte boundary — used in Image Resources."""
    b = s.encode("latin-1")
    data = bytes([len(b)]) + b
    if len(data) % 2 != 0:
        data += b'\x00'
    return data


def _pack_layer_name(s: str) -> bytes:
    """
    Pascal string padded to 4-byte boundary — required for Layer Records.
    PSD spec: "Layer name: Pascal string, padded to a multiple of 4 bytes."
    Image resources use 2-byte padding; layer names use 4-byte padding.
    """
    b = s.encode("latin-1")
    data = bytes([len(b)]) + b
    pad = (4 - len(data) % 4) % 4
    return data + b'\x00' * pad


def _compress_channel_zip(channel_bytes: bytes) -> bytes:
    """
    Raw deflate for PSD compression mode 2 (ZIP without prediction).
    PSD spec requires raw deflate with NO zlib wrapper (no 0x789C header,
    no Adler-32 checksum). zlib.compress() adds both; wbits=-15 strips them.
    """
    obj = zlib.compressobj(level=6, method=zlib.DEFLATED, wbits=-15)
    return obj.compress(channel_bytes) + obj.flush()


def _pil_to_channels(pil_img: Image.Image, mode: str) -> dict:
    """
    Splits a PIL image into per-channel raw bytes dicts.
    mode='RGBA' → keys: 'R','G','B','A'
    mode='RGB'  → keys: 'R','G','B'
    Returns dict of channel_id -> raw bytes.
    Channel IDs: 0=R, 1=G, 2=B, -1=Alpha (transparency mask)
    """
    img = pil_img.convert(mode)
    bands = img.split()
    if mode == 'RGBA':
        r, g, b, a = bands
        return {0: r.tobytes(), 1: g.tobytes(), 2: b.tobytes(), -1: a.tobytes()}
    else:
        r, g, b = bands
        return {0: r.tobytes(), 1: g.tobytes(), 2: b.tobytes()}


def write_psd(out_path: str, canvas_w: int, canvas_h: int,
              layers: list, log_fn=None) -> None:
    """
    Writes a layered PSD file to out_path.

    layers: list of dicts, each:
        {
          'name':   str,           layer name shown in Photoshop
          'image':  PIL.Image,     RGBA PIL image (actual content size)
          'top':    int,           y offset on canvas
          'left':   int,           x offset on canvas
          'opacity': int,          0-255 (255 = fully opaque)
          'visible': bool,
        }

    Canvas is RGB (3 channels). Layers are RGBA (4 channels including alpha).
    The merged composite (Image Data section) is a white background with all
    layers flattened — used by apps that can't read layers.
    """
    if log_fn:
        log_fn(f"Writing PSD: {canvas_w}x{canvas_h}px, {len(layers)} layers", "info")

    buf = io.BytesIO()
    p = buf.write

    # ── Section 1: File Header ────────────────────────────────────────────────
    # Signature, version, reserved, channels, height, width, depth, color_mode
    p(b'8BPS')                              # signature
    p(struct.pack('>H', 1))                 # version: 1 = PSD (not PSB)
    p(b'\x00' * 6)                          # reserved
    p(struct.pack('>H', 3))                 # channels in merged image (RGB=3)
    p(struct.pack('>I', canvas_h))          # height
    p(struct.pack('>I', canvas_w))          # width
    p(struct.pack('>H', 8))                 # bit depth per channel
    p(struct.pack('>H', 3))                 # color mode: 3=RGB

    # ── Section 2: Color Mode Data ────────────────────────────────────────────
    p(struct.pack('>I', 0))                 # length=0 (not indexed/duotone)

    # ── Section 3: Image Resources ────────────────────────────────────────────
    # We write one resource: Resolution Info (1005)
    # ResolutionInfo: hRes(fixed16.16), hResUnit(2=PPI), widthUnit,
    #                 vRes(fixed16.16), vResUnit(2=PPI), heightUnit
    dpi_fixed = (DPI << 16)  # convert int DPI to Fixed 16.16
    res_data = struct.pack('>IHHIHH',
        dpi_fixed, 1, 1,   # hRes, hResUnit=pixels/inch, widthUnit
        dpi_fixed, 1, 1    # vRes, vResUnit=pixels/inch, heightUnit
    )
    res_block = (b'8BIM' +
                 struct.pack('>H', 1005) +     # resource ID
                 b'\x00\x00' +                 # pascal string (empty, 2 bytes)
                 struct.pack('>I', len(res_data)) +
                 res_data)
    if len(res_block) % 2 != 0:
        res_block += b'\x00'

    p(struct.pack('>I', len(res_block)))
    p(res_block)

    # ── Section 4: Layer and Mask Information ─────────────────────────────────
    # Build all layer records first so we know the total size

    layer_records_buf = io.BytesIO()
    layer_data_buf    = io.BytesIO()

    num_layers = len(layers)

    for lyr in layers:
        img    = lyr['image']           # PIL RGBA image
        lname  = lyr['name']
        top    = lyr['top']
        left   = lyr['left']
        bottom = top  + img.height
        right  = left + img.width
        opacity = lyr.get('opacity', 255)
        visible = lyr.get('visible', True)
        flags   = 0 if visible else 2   # bit 1 = invisible

        # Channel list: alpha(-1), R(0), G(1), B(2) — 4 channels for RGBA layer
        channels = _pil_to_channels(img, 'RGBA')
        channel_order = [-1, 0, 1, 2]  # alpha first (PSD convention)

        # Layer record
        lr = io.BytesIO()
        lr.write(struct.pack('>iiii', top, left, bottom, right))
        lr.write(struct.pack('>H', 4))   # num_channels = 4

        # Channel info: (channel_id int16, data_length uint32)
        # data_length = raw bytes + 2 bytes for the compression type uint16 header
        for cid in channel_order:
            data_len = len(channels[cid]) + 2
            lr.write(struct.pack('>hI', cid, data_len))

        lr.write(b'8BIM')                           # blend mode signature
        lr.write(b'norm')                           # blend mode: normal
        lr.write(struct.pack('>B', opacity))        # opacity
        lr.write(struct.pack('>B', 0))              # clipping: 0=base
        lr.write(struct.pack('>B', flags))          # flags
        lr.write(b'\x00')                           # filler

        # Extra data MUST contain three sub-sections in this exact order
        # (per Adobe PSD spec) or Photoshop/GIMP report invalid mask info size:
        #   1. Layer Mask / Adjustment data (4-byte length prefix, can be 0)
        #   2. Layer Blending Ranges (4-byte length prefix, can be 0)
        #   3. Layer Name (pascal string padded to 4-byte boundary — NOT 2-byte)
        name_pascal  = _pack_layer_name(lname)   # 4-byte padded, per spec
        layer_mask   = struct.pack('>I', 0)       # no mask data
        blend_ranges = struct.pack('>I', 0)       # no blending ranges
        extra_data   = layer_mask + blend_ranges + name_pascal

        lr.write(struct.pack('>I', len(extra_data)))
        lr.write(extra_data)

        layer_records_buf.write(lr.getvalue())

        # Channel image data — compression=0 (RAW), same as the merged composite.
        # ZIP mode (2) needs per-row byte-count tables; omitting them corrupts the file.
        for cid in channel_order:
            layer_data_buf.write(struct.pack('>H', 0))   # compression: 0=raw
            layer_data_buf.write(channels[cid])          # raw (uncompressed) bytes

    layer_records_bytes = layer_records_buf.getvalue()
    layer_data_bytes    = layer_data_buf.getvalue()

    # Layer info block = count(int16) + records + channel data
    layer_info = (struct.pack('>h', num_layers) +   # positive = layers have merged alpha
                  layer_records_bytes +
                  layer_data_bytes)

    # Pad layer info to 4-byte boundary
    if len(layer_info) % 4 != 0:
        layer_info += b'\x00' * (4 - len(layer_info) % 4)

    layer_info_block = struct.pack('>I', len(layer_info)) + layer_info

    # Global mask info (empty)
    global_mask = struct.pack('>I', 0)

    lmi_content = layer_info_block + global_mask
    p(struct.pack('>I', len(lmi_content)))
    p(lmi_content)

    # ── Section 5: Image Data (merged/flattened composite) ───────────────────
    # Build a flattened RGB composite: white background + all layers composited
    if log_fn:
        log_fn("Compositing merged preview...", "info")

    composite = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
    for lyr in layers:
        img  = lyr['image'].convert("RGBA")
        top  = lyr['top']
        left = lyr['left']
        composite.paste(img, (left, top), img)

    comp_channels = _pil_to_channels(composite, 'RGB')

    # Write merged image: compression=0 (RAW/uncompressed) — safest for
    # compatibility. ZIP mode (2) requires per-row byte count tables which
    # were absent and caused Photoshop to report "unexpected end-of-file".
    p(struct.pack('>H', 0))  # compression type: 0 = raw
    for cid in [0, 1, 2]:
        p(comp_channels[cid])

    # ── Write to file ──────────────────────────────────────────────────────────
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())

    if log_fn:
        size_mb = os.path.getsize(out_path) / (1024 * 1024)
        log_fn(f"PSD written: {size_mb:.1f} MB", "success")


# ─── LAYER PREPARATION ────────────────────────────────────────────────────────

def _prepare_customer_image(img_path, canvas_w, canvas_h, remove_bg):
    """Loads, scales, optionally removes bg. Returns (PIL RGBA, top, left)."""
    if not img_path or not os.path.isfile(img_path):
        return Image.new("RGBA", (1, 1), (0, 0, 0, 0)), 0, 0

    src   = Image.open(img_path).convert("RGBA")
    ratio = min(canvas_w / src.width, canvas_h / src.height)
    nw    = max(1, int(src.width  * ratio))
    nh    = max(1, int(src.height * ratio))
    src   = src.resize((nw, nh), Image.LANCZOS)

    if remove_bg:
        bg_r, bg_g, bg_b = src.getpixel((4, 4))[:3]
        threshold = 35
        pixels    = src.load()
        for py in range(nh):
            for px in range(nw):
                r, g, b, a = src.getpixel((px, py))
                if (abs(r - bg_r) < threshold and
                    abs(g - bg_g) < threshold and
                    abs(b - bg_b) < threshold):
                    pixels[px, py] = (r, g, b, 0)

    top  = (canvas_h - nh) // 2
    left = (canvas_w - nw) // 2
    return src, top, left


def _prepare_customer_text(text_lines, font_name, colour_hex, canvas_w, canvas_h):
    """Renders text on a tight canvas. Returns (PIL RGBA, top, left)."""
    if not text_lines:
        return Image.new("RGBA", (1, 1), (0, 0, 0, 0)), 0, 0

    r, g, b = hex_to_rgb(colour_hex)
    avail_w = int(canvas_w * 0.90)
    longest = max(text_lines, key=len)

    scratch = Image.new("RGBA", (1, 1))
    draw    = ImageDraw.Draw(scratch)
    lo, hi, best = 20, min(900, canvas_h // max(1, len(text_lines))), 60
    while lo <= hi:
        mid  = (lo + hi) // 2
        font = get_font(font_name, mid)
        bb   = draw.textbbox((0, 0), longest, font=font)
        if (bb[2] - bb[0]) <= avail_w:
            best = mid
            lo   = mid + 1
        else:
            hi = mid - 1

    font    = get_font(font_name, best)
    bb0     = draw.textbbox((0, 0), text_lines[0], font=font)
    line_h  = int((bb0[3] - bb0[1]) * 1.25)

    max_lw = max(
        draw.textbbox((0, 0), line, font=font)[2] -
        draw.textbbox((0, 0), line, font=font)[0]
        for line in text_lines
    )
    block_w = min(max_lw + 40, canvas_w)
    block_h = line_h * len(text_lines) + 40

    text_img = Image.new("RGBA", (block_w, block_h), (0, 0, 0, 0))
    draw2    = ImageDraw.Draw(text_img)
    y_local  = 20
    for line in text_lines:
        bb  = draw2.textbbox((0, 0), line, font=font)
        lw  = bb[2] - bb[0]
        x_local = max(0, (block_w - lw) // 2)
        draw2.text((x_local, y_local), line, font=font, fill=(r, g, b, 255))
        y_local += line_h

    top  = max(int(canvas_h * 0.70), canvas_h - block_h - 60)
    left = max(0, (canvas_w - block_w) // 2)
    return text_img, top, left


def build_layered_psd(order_id, zone, w, h,
                       img_path, text_lines, font_name,
                       colour_hex, remove_bg,
                       out_path, log_fn):
    """
    Builds a layered PSD using the pure Python struct writer.
    No psd-tools, no NumPy, no memory issues.
    """
    try:
        log_fn("Preparing CustomerImage layer...", "info")
        img_pil, img_top, img_left = _prepare_customer_image(
            img_path, w, h, remove_bg)
        log_fn(f"  Image: {img_pil.width}x{img_pil.height}px at ({img_left},{img_top})", "info")

        log_fn("Preparing CustomerText layer...", "info")
        txt_pil, txt_top, txt_left = _prepare_customer_text(
            text_lines, font_name, colour_hex, w, h)
        if text_lines:
            log_fn(f"  Text: {txt_pil.width}x{txt_pil.height}px at ({txt_left},{txt_top})", "info")

        layers = [
            {"name": "CustomerImage", "image": img_pil,
             "top": img_top, "left": img_left, "opacity": 255, "visible": True},
            {"name": "CustomerText",  "image": txt_pil,
             "top": txt_top, "left": txt_left, "opacity": 255, "visible": True},
        ]

        log_fn(f"Writing layered PSD ({w}x{h}px, pure Python)...", "info")
        write_psd(out_path, w, h, layers, log_fn=log_fn)

        if not os.path.isfile(out_path):
            return False, f"PSD not found after write: {out_path}"

        return True, "OK"

    except Exception as e:
        import traceback
        return False, f"PSD write error: {e}\n{traceback.format_exc()[-600:]}"


# ─── AUTOMATION PIPELINE ──────────────────────────────────────────────────────

def run_automation(order_id, detail_id, amazon_id, order_data):
    def log(msg, level="info"):
        log_progress(order_id, msg, level)

    try:
        log("Starting automation pipeline...", "info")

        zone      = order_data["zone"]
        product   = order_data["product"]
        text_raw  = order_data["text"]
        font_name = order_data["font"]
        colour    = order_data["colour"]
        img_path  = order_data.get("image_path", "")
        remove_bg = order_data.get("remove_bg", False)

        spec = PRODUCT_CANVAS.get(product, PRODUCT_CANVAS["tshirt"])
        dims = spec.get(zone, spec.get("front"))
        w, h = dims
        log(f"Canvas: {w}x{h}px ({w/PX_PER_CM:.1f}x{h/PX_PER_CM:.1f}cm) "
            f"| Product: {product} | Zone: {zone}", "info")

        # ── Step 1: Optional rembg background removal ───────────────────────
        prepared_img = img_path
        if img_path and os.path.exists(img_path) and remove_bg:
            log("Trying rembg for background removal...", "info")
            try:
                from rembg import remove as rembg_remove, new_session
                session      = new_session("u2netp")
                src          = Image.open(img_path).convert("RGBA")
                result_img   = rembg_remove(src, session=session)
                prepared_img = os.path.splitext(img_path)[0] + "_nobg.png"
                result_img.save(prepared_img)
                log("Background removed via rembg", "success")
                remove_bg = False
            except ImportError:
                log("rembg not installed — using colour-select removal", "warning")
            except Exception as e:
                log(f"rembg failed ({e}) — using colour-select removal", "warning")

        # ── Step 2: Parse text ───────────────────────────────────────────────
        text_lines = parse_texts(text_raw) if text_raw.strip() else []
        if text_lines:
            log(f"Text: {text_lines} | Font: {font_name} | Colour: {colour}", "info")
        else:
            log("No text — image-only order", "info")

        # ── Step 3: Output path ──────────────────────────────────────────────
        today    = datetime.now().strftime("%Y-%m-%d")
        out_dir  = os.path.join(OUTPUT_FOLDER, today)
        os.makedirs(out_dir, exist_ok=True)
        safe_id  = amazon_id.replace("/", "-")
        out_path = os.path.join(out_dir, f"{safe_id}_{zone}.psd")

        # ── Step 4: Build layered PSD ────────────────────────────────────────
        log("Building layered PSD (pure Python, no NumPy)...", "info")

        success, message = build_layered_psd(
            order_id=order_id, zone=zone, w=w, h=h,
            img_path=prepared_img or img_path,
            text_lines=text_lines, font_name=font_name,
            colour_hex=colour, remove_bg=remove_bg,
            out_path=out_path, log_fn=log,
        )

        if not success:
            log(f"PSD write failed: {message}", "error")
            log("Saving flat PNG fallback...", "warning")
            out_path = out_path.replace(".psd", "_FLAT.png")
            _save_flat_png(w, h, prepared_img or img_path,
                           text_lines, font_name, colour, out_path)
            log(f"Flat PNG saved → {out_path}", "warning")
        else:
            size_mb = os.path.getsize(out_path) / (1024 * 1024)
            log(f"Layered PSD → {out_path} ({size_mb:.1f} MB)", "success")
            log(f"Open in Photoshop — layers: CustomerImage / CustomerText", "success")

        # ── Step 5: Update database ──────────────────────────────────────────
        log("Updating database...", "info")
        mark_order_complete(detail_id, out_path)
        log(f"Order {amazon_id} complete!", "success")
        progress_logs[order_id].append({"done": True, "file": out_path})

    except Exception as e:
        log_progress(order_id, f"Pipeline error: {str(e)}", "error")
        progress_logs[order_id].append({"done": True, "error": str(e)})


def _save_flat_png(w, h, img_path, text_lines, font_name, colour_hex, out_path):
    canvas = Image.new("RGBA", (w, h), (255, 255, 255, 255))
    if img_path and os.path.exists(img_path):
        src   = Image.open(img_path).convert("RGBA")
        ratio = min(w / src.width, h / src.height)
        nw, nh = int(src.width * ratio), int(src.height * ratio)
        src   = src.resize((nw, nh), Image.LANCZOS)
        canvas.paste(src, ((w - nw) // 2, (h - nh) // 2), src)
    if text_lines:
        draw = ImageDraw.Draw(canvas)
        rgb  = hex_to_rgb(colour_hex)
        fo   = get_font(font_name, max(40, h // 12))
        y    = int(h * 0.70)
        for line in text_lines:
            bb = draw.textbbox((0, 0), line, font=fo)
            lw = bb[2] - bb[0]
            draw.text(((w - lw) // 2, y), line, font=fo, fill=rgb + (255,))
            y += int((bb[3] - bb[1]) * 1.25)
    canvas.save(out_path, dpi=(DPI, DPI))


# ─── ROUTES ───────────────────────────────────────────────────────────────────

HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Varsany Print Automation</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:system-ui,sans-serif;background:#0f0f0f;color:#e0e0e0;min-height:100vh}
header{background:#1a1a2e;padding:16px 32px;display:flex;align-items:center;gap:16px;border-bottom:1px solid #333}
header h1{font-size:20px;font-weight:600;color:#fff}
header span{background:#7c3aed;color:#fff;padding:4px 12px;border-radius:20px;font-size:12px}
.container{max-width:1200px;margin:0 auto;padding:32px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:24px}
.card{background:#1a1a1a;border:1px solid #333;border-radius:12px;padding:24px}
.card h2{font-size:16px;font-weight:600;margin-bottom:20px;color:#fff}
.form-group{margin-bottom:16px}
label{display:block;font-size:13px;color:#999;margin-bottom:6px}
input,select,textarea{width:100%;background:#111;border:1px solid #333;color:#e0e0e0;
  padding:10px 14px;border-radius:8px;font-size:14px;outline:none}
input:focus,select:focus{border-color:#7c3aed}
input[type=color]{height:40px;padding:4px;cursor:pointer}
.btn{background:#7c3aed;color:#fff;border:none;padding:12px 24px;border-radius:8px;
  font-size:14px;font-weight:600;cursor:pointer;width:100%}
.btn:hover{background:#6d28d9}
.log-box{background:#0a0a0a;border:1px solid #222;border-radius:8px;
  height:340px;overflow-y:auto;padding:12px;font-family:monospace;font-size:12px}
.log-info{color:#60a5fa}.log-success{color:#34d399}
.log-warning{color:#fbbf24}.log-error{color:#f87171}
.log-time{color:#555;margin-right:8px}
.progress-bar{height:6px;background:#222;border-radius:3px;margin:12px 0}
.progress-fill{height:100%;background:#7c3aed;border-radius:3px;transition:width 0.3s;width:0%}
.status-badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px}
.status-done{background:#064e3b;color:#34d399}
.status-pending{background:#1e1b4b;color:#818cf8}
table{width:100%;border-collapse:collapse;font-size:13px}
th{text-align:left;padding:8px 12px;color:#666;font-weight:500;border-bottom:1px solid #222}
td{padding:8px 12px;border-bottom:1px solid #1a1a1a;color:#ccc}
tr:hover td{background:#1f1f1f}
.preview-box{background:#111;border:1px dashed #333;border-radius:8px;
  min-height:100px;display:flex;align-items:center;justify-content:center;
  color:#555;font-size:13px;margin-top:8px}
.preview-box img{max-width:100%;max-height:180px;border-radius:6px}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.result-file{background:#052e16;border:1px solid #166534;border-radius:8px;
  padding:12px;margin-top:12px;font-size:13px;color:#4ade80;word-break:break-all}
.engine-note{background:#1c1917;border:1px solid #44403c;border-radius:8px;
  padding:10px 14px;margin-bottom:20px;font-size:12px;color:#a8a29e;line-height:1.6}
.engine-note b{color:#a78bfa}
</style>
</head>
<body>
<header>
  <h1>Varsany Print Automation</h1>
  <span>Pure Python PSD Writer</span>
</header>
<div class="container">
  <div class="grid">

    <div class="card">
      <div class="engine-note">
        <b>Engine:</b> Pure Python struct + zlib — zero NumPy, zero external tools.
        ⚠️ <b>LOW-RES MODE</b> (~72 DPI). Raise PX_PER_CM to 320 for production.
        Designers open in Photoshop: CustomerImage + CustomerText layers.
      </div>
      <h2>New Customisation Order</h2>
      <form id="orderForm" enctype="multipart/form-data">
        <div class="two-col">
          <div class="form-group">
            <label>Product Type</label>
            <select name="product" id="productSelect" onchange="updateZones()">
              <option value="tshirt">T-Shirt</option>
              <option value="hoodie">Hoodie</option>
              <option value="kidstshirt">Kids T-Shirt</option>
              <option value="totebag">Tote Bag</option>
              <option value="slipper">Slippers</option>
              <option value="babyvest">Baby Vest</option>
            </select>
          </div>
          <div class="form-group">
            <label>Print Zone</label>
            <select name="zone" id="zoneSelect">
              <option value="front">Front</option>
              <option value="back">Back</option>
              <option value="sleeve">Sleeve</option>
              <option value="pocket">Pocket</option>
            </select>
          </div>
        </div>
        <div class="form-group">
          <label>Upload Image (optional)</label>
          <input type="file" name="image" accept="image/*" onchange="previewImage(this)">
          <div class="preview-box" id="imgPreview">No image selected</div>
        </div>
        <div class="form-group">
          <label>Remove Background?</label>
          <select name="remove_bg">
            <option value="0">No — keep background</option>
            <option value="1">Yes — remove background</option>
          </select>
        </div>
        <div class="form-group">
          <label>Customer Text</label>
          <textarea name="text" rows="3"
            placeholder="Enter text&#10;Use Enter for new line"></textarea>
        </div>
        <div class="two-col">
          <div class="form-group">
            <label>Font (place .ttf in C:\\Varsany\\Fonts\\)</label>
            <select name="font">
              <option>Arial</option>
              <option>Arial Bold</option>
              <option>Impact</option>
              <option>Times New Roman</option>
              <option>Courier New</option>
              <option>Russo One</option>
              <option>Bebas Neue</option>
              <option>Chewy</option>
            </select>
          </div>
          <div class="form-group">
            <label>Text Colour</label>
            <input type="color" name="colour" value="#ffffff">
          </div>
        </div>
        <div class="form-group">
          <label>SKU</label>
          <input type="text" name="sku" value="MenTee_BlkM">
        </div>
        <button type="submit" class="btn" id="submitBtn">
          Submit Order &amp; Generate Layered PSD
        </button>
      </form>
    </div>

    <div class="card">
      <h2>Automation Progress</h2>
      <div id="statusArea" style="color:#555;font-size:13px;text-align:center;padding:40px 0">
        Submit an order to see live progress
      </div>
      <div id="progressArea" style="display:none">
        <div style="font-size:13px;color:#999;margin-bottom:8px">
          Order: <span id="currentOrderId" style="color:#a78bfa"></span>
        </div>
        <div class="progress-bar">
          <div class="progress-fill" id="progressFill"></div>
        </div>
        <div class="log-box" id="logBox"></div>
        <div id="resultBox"></div>
      </div>
    </div>

  </div>

  <div class="card" style="margin-top:24px">
    <h2>Order History</h2>
    <table>
      <thead>
        <tr><th>Order ID</th><th>Product</th><th>Zone</th>
            <th>Text</th><th>Status</th><th>Output File</th></tr>
      </thead>
      <tbody>
        {% for o in orders %}
        <tr>
          <td style="font-family:monospace;font-size:11px">{{o.OrderID}}</td>
          <td>{{o.ItemType}}</td>
          <td>{{o.PrintLocation}}</td>
          <td>{{(o.FrontText or o.BackText or o.SleeveText or '')[:30]}}</td>
          <td>
            {% if o.IsDesignComplete %}
              <span class="status-badge status-done">Done</span>
            {% else %}
              <span class="status-badge status-pending">Pending</span>
            {% endif %}
          </td>
          <td style="font-size:11px;color:#555">{{(o.AdditionalPSD or '')[-50:]}}</td>
        </tr>
        {% endfor %}
        {% if not orders %}
        <tr><td colspan="6" style="color:#555;text-align:center;padding:24px">
          No prototype orders yet</td></tr>
        {% endif %}
      </tbody>
    </table>
  </div>
</div>

<script>
function previewImage(input) {
  const box = document.getElementById('imgPreview');
  if (input.files && input.files[0]) {
    const reader = new FileReader();
    reader.onload = e => { box.innerHTML = '<img src="'+e.target.result+'">'; };
    reader.readAsDataURL(input.files[0]);
  }
}
function updateZones() {
  const zones = {
    hoodie:['front','back','sleeve','pocket'], tshirt:['front','back','sleeve','pocket'],
    kidstshirt:['front','back'], totebag:['front','back'],
    slipper:['front'], babyvest:['front']
  };
  const z = zones[document.getElementById('productSelect').value] || ['front'];
  document.getElementById('zoneSelect').innerHTML = z.map(v =>
    `<option value="${v}">${v.charAt(0).toUpperCase()+v.slice(1)}</option>`).join('');
}
let pollInterval = null;
document.getElementById('orderForm').onsubmit = async function(e) {
  e.preventDefault();
  const btn = document.getElementById('submitBtn');
  btn.disabled = true;
  btn.textContent = 'Generating PSD...';
  const fd  = new FormData(this);
  const res = await fetch('/submit', {method:'POST', body:fd});
  const data = await res.json();
  if (data.error) {
    alert('Error: ' + data.error);
    btn.disabled = false;
    btn.textContent = 'Submit Order & Generate Layered PSD';
    return;
  }
  const orderId = data.order_id;
  document.getElementById('currentOrderId').textContent = orderId;
  document.getElementById('statusArea').style.display   = 'none';
  document.getElementById('progressArea').style.display = 'block';
  document.getElementById('logBox').innerHTML   = '';
  document.getElementById('resultBox').innerHTML = '';
  document.getElementById('progressFill').style.width = '5%';
  let progress = 5, seen = 0;
  pollInterval = setInterval(async () => {
    const r    = await fetch('/progress/' + orderId);
    const logs = await r.json();
    const box  = document.getElementById('logBox');
    for (let i = seen; i < logs.length; i++) {
      const l = logs[i];
      if (l.done !== undefined) {
        clearInterval(pollInterval);
        document.getElementById('progressFill').style.width = '100%';
        btn.disabled = false;
        btn.textContent = 'Submit Order & Generate Layered PSD';
        if (l.file) {
          document.getElementById('resultBox').innerHTML =
            '<div class="result-file">Layered PSD saved:<br>' + l.file + '</div>';
        }
        if (l.error) {
          document.getElementById('resultBox').innerHTML =
            '<div style="background:#450a0a;border:1px solid #7f1d1d;border-radius:8px;'
            +'padding:12px;margin-top:12px;font-size:13px;color:#f87171">Error: '+l.error+'</div>';
        }
        setTimeout(() => location.reload(), 4000);
        break;
      }
      const cls = 'log-'+(l.level||'info');
      box.innerHTML += `<div><span class="log-time">${l.time}</span>`
                     + `<span class="${cls}">${l.message}</span></div>`;
      seen = i + 1;
    }
    box.scrollTop = box.scrollHeight;
    progress = Math.min(progress + 7, 90);
    document.getElementById('progressFill').style.width = progress + '%';
  }, 800);
};
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML, orders=get_recent_orders())


@app.route("/submit", methods=["POST"])
def submit():
    try:
        image_path = ""
        if "image" in request.files:
            f = request.files["image"]
            if f and f.filename:
                ext        = os.path.splitext(f.filename)[1]
                image_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}{ext}")
                f.save(image_path)

        order_data = {
            "product":    request.form.get("product", "tshirt"),
            "zone":       request.form.get("zone", "front"),
            "text":       request.form.get("text", ""),
            "font":       request.form.get("font", "Arial"),
            "colour":     request.form.get("colour", "#ffffff"),
            "sku":        request.form.get("sku", "MenTee_BlkM"),
            "remove_bg":  request.form.get("remove_bg", "0") == "1",
            "image_path": image_path,
        }

        order_id, detail_id, amazon_id = save_order_to_db(order_data)

        threading.Thread(
            target=run_automation,
            args=(order_id, detail_id, amazon_id, order_data),
            daemon=True
        ).start()

        return jsonify({"order_id": order_id, "amazon_id": amazon_id})
    except Exception as e:
        return jsonify({"error": str(e)})


@app.route("/progress/<order_id>")
def progress(order_id):
    return jsonify(progress_logs.get(order_id, []))


@app.route("/output/<path:filename>")
def output_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename)


if __name__ == "__main__":
    print(f"\n{'='*55}")
    print(f"  Varsany Print Automation — Pure Python PSD Writer")
    print(f"  Engine: struct + zlib (zero NumPy)")
    print(f"  Resolution: LOW-RES MODE ({PX_PER_CM} px/cm / {DPI} DPI)")
    print(f"  Expected file size: 1-5 MB per order (raise PX_PER_CM=320 for prod)")
    print(f"  Open: http://localhost:5000")
    print(f"{'='*55}\n")
    app.run(debug=True, port=5000, threaded=True)