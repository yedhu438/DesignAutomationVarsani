"""
Men's T-Shirt Multi-Zone PSD Engine
=====================================
Produces a single PSD file with ALL print zones stacked vertically,
each with a text label above it — exactly like the manual Photoshop workflow.

Canvas layout (top to bottom):
  ┌─────────────────────┐
  │  front              │  ← label text
  │  [front design]     │
  ├─────────────────────┤
  │  back               │  ← label text
  │  [back design]      │
  ├─────────────────────┤
  │  pocket             │  ← label text
  │  [pocket design]    │
  └─────────────────────┘

Each zone is a separate named layer group in the PSD.
Layer names match the image filenames from the database.

Usage:
    from men_tshirt_engine import build_men_tshirt_psd
    build_men_tshirt_psd(order_data, out_path, log_fn)
"""

import os, struct, zlib, io, json, urllib.request, uuid
from PIL import Image, ImageDraw, ImageFont

# ─── CONSTANTS ────────────────────────────────────────────────────────────────

# LOW-RES for testing. Change to 320 for production (812 DPI)
PX_PER_CM = 28
DPI       = int(PX_PER_CM * 2.54)

FONTS_FOLDER = r"C:\Varsany\Fonts"
TEMP_FOLDER  = r"C:\Varsany\Temp"

# Zone canvas sizes in cm (width, height)
ZONE_SIZES_CM = {
    "front":  (30, 30),
    "back":   (30, 45),
    "pocket": (12, 12),   # left/right pocket
    "sleeve": (15, 30),
}

# Gap between zones in pixels + label area height
LABEL_HEIGHT_PX = int(PX_PER_CM * 1.2)   # ~1.2cm for label text
ZONE_GAP_PX     = int(PX_PER_CM * 0.5)   # ~0.5cm gap between zones

# Image server base URL — fill in when confirmed
IMAGE_SERVER_URL = ""   # e.g. "https://crssoft.co.uk/uploads/"


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def cm_to_px(cm):
    return int(round(cm * PX_PER_CM))

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

def parse_font_json(raw):
    """Extract font name from plain string or JSON like {"NormalFont":"Arial"}"""
    if not raw: return "Arial"
    raw = raw.strip()
    if raw.startswith("{"):
        try:
            d = json.loads(raw)
            return d.get("NormalFont") or d.get("PremiumFont") or "Arial"
        except: pass
    return raw

def parse_colour_json(raw):
    """Extract hex colour from plain string or JSON like {"Colour1":"#ffffff"}"""
    if not raw: return "#ffffff"
    raw = raw.strip()
    if raw.startswith("{"):
        try:
            d = json.loads(raw)
            return d.get("Colour1") or "#ffffff"
        except: pass
    return raw

def parse_text(raw):
    """Split multiline text into list of lines."""
    if not raw: return []
    lines = []
    for line in raw.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = line.strip()
        if line:
            lines.append(line)
    return lines

def download_image(filename, log_fn=None):
    """Download image from server. Returns local temp path or None."""
    if not IMAGE_SERVER_URL or not filename:
        return None
    url = IMAGE_SERVER_URL.rstrip("/") + "/" + filename.strip()
    try:
        tmp = os.path.join(TEMP_FOLDER, f"dl_{uuid.uuid4().hex[:8]}_{filename}")
        urllib.request.urlretrieve(url, tmp)
        return tmp
    except Exception as e:
        if log_fn: log_fn(f"  Could not download {filename}: {e}", "warning")
        return None


# ─── ZONE RENDERING ───────────────────────────────────────────────────────────

def render_zone(zone_name, image_filename, text_raw, font_raw, colour_raw,
                canvas_w, canvas_h, log_fn=None):
    """
    Renders a single print zone as a PIL RGBA image at canvas_w x canvas_h.
    Returns the PIL image.
    """
    canvas = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))

    # ── Customer image ──────────────────────────────────────────────────────
    if image_filename:
        img_path = download_image(image_filename, log_fn)
        # Fallback: check if file already exists locally
        if not img_path:
            local = os.path.join(TEMP_FOLDER, image_filename)
            if os.path.exists(local): img_path = local

        if img_path and os.path.exists(img_path):
            try:
                src   = Image.open(img_path).convert("RGBA")
                ratio = min(canvas_w / src.width, canvas_h / src.height)
                nw    = max(1, int(src.width  * ratio))
                nh    = max(1, int(src.height * ratio))
                src   = src.resize((nw, nh), Image.LANCZOS)
                ox    = (canvas_w - nw) // 2
                oy    = (canvas_h - nh) // 2
                canvas.paste(src, (ox, oy), src)
                if log_fn: log_fn(f"  [{zone_name}] Image placed: {nw}x{nh}px", "info")
            except Exception as e:
                if log_fn: log_fn(f"  [{zone_name}] Image error: {e}", "warning")

    # ── Customer text ───────────────────────────────────────────────────────
    text_lines = parse_text(text_raw)
    if text_lines:
        font_name = parse_font_json(font_raw)
        colour    = parse_colour_json(colour_raw)
        r, g, b   = hex_to_rgb(colour)

        avail_w = int(canvas_w * 0.90)
        longest = max(text_lines, key=len)

        scratch = Image.new("RGBA", (1, 1))
        draw    = ImageDraw.Draw(scratch)
        lo, hi, best = 8, min(400, canvas_h // max(1, len(text_lines))), 20
        while lo <= hi:
            mid  = (lo + hi) // 2
            font = get_font(font_name, mid)
            bb   = draw.textbbox((0, 0), longest, font=font)
            if (bb[2] - bb[0]) <= avail_w:
                best = mid; lo = mid + 1
            else:
                hi = mid - 1

        font   = get_font(font_name, best)
        bb0    = draw.textbbox((0, 0), text_lines[0], font=font)
        line_h = int((bb0[3] - bb0[1]) * 1.25)
        total_h = line_h * len(text_lines)
        y_pos   = (canvas_h - total_h) // 2

        draw2 = ImageDraw.Draw(canvas)
        for line in text_lines:
            bb  = draw2.textbbox((0, 0), line, font=font)
            lw  = bb[2] - bb[0]
            x   = max(0, (canvas_w - lw) // 2)
            draw2.text((x, y_pos), line, font=font, fill=(r, g, b, 255))
            y_pos += line_h

        if log_fn: log_fn(f"  [{zone_name}] Text rendered: {text_lines[:2]}", "info")

    return canvas


def render_zone_with_label(zone_name, image_filename, text_raw,
                            font_raw, colour_raw, canvas_w, canvas_h,
                            log_fn=None):
    """
    Renders a zone with a label above it.
    Returns a PIL RGBA image of size (canvas_w, LABEL_HEIGHT_PX + canvas_h).
    """
    total_h = LABEL_HEIGHT_PX + canvas_h
    result  = Image.new("RGBA", (canvas_w, total_h), (0, 0, 0, 0))

    # ── Draw label text (e.g. "back", "pocket") ─────────────────────────────
    draw      = ImageDraw.Draw(result)
    label_font = get_font("Arial", max(8, LABEL_HEIGHT_PX - 4))
    draw.text((4, 2), zone_name, font=label_font, fill=(0, 0, 0, 255))

    # ── Render zone content ──────────────────────────────────────────────────
    zone_img = render_zone(zone_name, image_filename, text_raw,
                           font_raw, colour_raw, canvas_w, canvas_h, log_fn)
    result.paste(zone_img, (0, LABEL_HEIGHT_PX), zone_img)

    return result


# ─── PSD BINARY WRITER ────────────────────────────────────────────────────────
# (Same proven writer used in prototype_app.py)

def _compress_channel(channel_bytes):
    obj = zlib.compressobj(level=6, method=zlib.DEFLATED, wbits=-15)
    return obj.compress(channel_bytes) + obj.flush()

def _pack_layer_name(s):
    b = s.encode("latin-1")[:254]
    data = bytes([len(b)]) + b
    pad  = (4 - len(data) % 4) % 4
    return data + b'\x00' * pad

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
    Writes a layered PSD.
    layers: list of {name, image (PIL RGBA), top, left, visible, opacity}
    """
    if log_fn:
        log_fn(f"Writing PSD {canvas_w}x{canvas_h}px, {len(layers)} layers...", "info")

    buf = io.BytesIO()
    p   = buf.write

    # Section 1: Header
    p(b'8BPS')
    p(struct.pack('>H', 1))
    p(b'\x00' * 6)
    p(struct.pack('>H', 3))
    p(struct.pack('>I', canvas_h))
    p(struct.pack('>I', canvas_w))
    p(struct.pack('>H', 8))
    p(struct.pack('>H', 3))

    # Section 2: Color Mode Data
    p(struct.pack('>I', 0))

    # Section 3: Image Resources (resolution)
    dpi_fixed = DPI << 16
    res_data  = struct.pack('>IHHIHH', dpi_fixed, 1, 1, dpi_fixed, 1, 1)
    res_block = b'8BIM' + struct.pack('>H', 1005) + b'\x00\x00' + struct.pack('>I', len(res_data)) + res_data
    if len(res_block) % 2: res_block += b'\x00'
    p(struct.pack('>I', len(res_block)))
    p(res_block)

    # Section 4: Layer and Mask Info
    lrec_buf = io.BytesIO()
    ldat_buf = io.BytesIO()

    for lyr in layers:
        img     = lyr['image']
        top     = lyr.get('top', 0)
        left    = lyr.get('left', 0)
        bottom  = top  + img.height
        right   = left + img.width
        opacity = lyr.get('opacity', 255)
        visible = lyr.get('visible', True)
        flags   = 0 if visible else 2

        channels      = _pil_to_channels(img, 'RGBA')
        channel_order = [-1, 0, 1, 2]
        compressed    = {cid: _compress_channel(channels[cid]) for cid in channel_order}

        lr = io.BytesIO()
        lr.write(struct.pack('>iiii', top, left, bottom, right))
        lr.write(struct.pack('>H', 4))
        for cid in channel_order:
            lr.write(struct.pack('>hI', cid, len(compressed[cid]) + 2))
        lr.write(b'8BIM' + b'norm')
        lr.write(struct.pack('>BBBB', opacity, 0, flags, 0))

        name_bytes = _pack_layer_name(lyr['name'])
        extra      = struct.pack('>I', 0) + struct.pack('>I', 0) + name_bytes
        lr.write(struct.pack('>I', len(extra)))
        lr.write(extra)
        lrec_buf.write(lr.getvalue())

        for cid in channel_order:
            ldat_buf.write(struct.pack('>H', 2))
            ldat_buf.write(compressed[cid])

    layer_info = (struct.pack('>h', len(layers))
                  + lrec_buf.getvalue()
                  + ldat_buf.getvalue())
    if len(layer_info) % 4:
        layer_info += b'\x00' * (4 - len(layer_info) % 4)
    lmi = struct.pack('>I', len(layer_info)) + layer_info + struct.pack('>I', 0)
    p(struct.pack('>I', len(lmi)))
    p(lmi)

    # Section 5: Merged composite (white bg)
    composite = Image.new("RGB", (canvas_w, canvas_h), (255, 255, 255))
    for lyr in layers:
        if lyr.get('visible', True):
            composite.paste(lyr['image'].convert("RGBA"),
                            (lyr.get('left', 0), lyr.get('top', 0)),
                            lyr['image'].convert("RGBA"))
    comp_ch = _pil_to_channels(composite, 'RGB')
    p(struct.pack('>H', 2))
    for cid in [0, 1, 2]:
        p(_compress_channel(comp_ch[cid]))

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())

    size_mb = os.path.getsize(out_path) / (1024 * 1024)
    if log_fn: log_fn(f"PSD saved: {size_mb:.1f} MB → {out_path}", "success")


# ─── MAIN FUNCTION ────────────────────────────────────────────────────────────

def build_men_tshirt_psd(order_data, out_path, log_fn=None):
    """
    Builds a single multi-zone PSD for a Men's T-shirt order.

    order_data keys:
        front_image, front_text, front_fonts, front_colours
        back_image,  back_text,  back_fonts,  back_colours
        pocket_image, pocket_text, pocket_fonts, pocket_colours
        sleeve_image, sleeve_text, sleeve_fonts, sleeve_colours
        print_location  e.g. "Front + Back", "Front Pocket + Back"

    The PSD canvas stacks all active zones vertically with labels.
    Each zone gets its own named layer.
    """
    print_location = order_data.get("print_location", "").lower()

    # ── Determine which zones are active ────────────────────────────────────
    has_front  = bool(order_data.get("front_image")  or order_data.get("front_text"))
    has_back   = bool(order_data.get("back_image")   or order_data.get("back_text"))
    has_pocket = bool(order_data.get("pocket_image") or order_data.get("pocket_text"))
    has_sleeve = bool(order_data.get("sleeve_image") or order_data.get("sleeve_text"))

    # Also detect from print_location string if flags not set
    if not any([has_front, has_back, has_pocket, has_sleeve]):
        has_front  = "front" in print_location and "pocket" not in print_location
        has_back   = "back"  in print_location
        has_pocket = "pocket" in print_location
        has_sleeve = "sleeve" in print_location

    active_zones = []
    if has_front:  active_zones.append("front")
    if has_back:   active_zones.append("back")
    if has_pocket: active_zones.append("pocket")
    if has_sleeve: active_zones.append("sleeve")

    if not active_zones:
        # Default to front if nothing detected
        active_zones = ["front"]

    if log_fn:
        log_fn(f"Active zones: {', '.join(active_zones)}", "info")

    # ── Calculate total canvas size ──────────────────────────────────────────
    # All zones use the same WIDTH (max zone width)
    zone_widths  = [cm_to_px(ZONE_SIZES_CM[z][0]) for z in active_zones]
    zone_heights = [cm_to_px(ZONE_SIZES_CM[z][1]) for z in active_zones]
    canvas_w     = max(zone_widths)

    # Total height = sum of all zones + labels + gaps between zones
    total_h = 0
    for i, z in enumerate(active_zones):
        total_h += LABEL_HEIGHT_PX + zone_heights[i]
        if i < len(active_zones) - 1:
            total_h += ZONE_GAP_PX

    if log_fn:
        log_fn(f"Total canvas: {canvas_w}x{total_h}px "
               f"({canvas_w/PX_PER_CM:.1f}x{total_h/PX_PER_CM:.1f}cm)", "info")

    # ── Render each zone and collect layers ──────────────────────────────────
    psd_layers = []
    y_offset   = 0

    for z in active_zones:
        zh = zone_heights[active_zones.index(z)]
        zw = cm_to_px(ZONE_SIZES_CM[z][0])

        # Get zone data from order
        img_file = order_data.get(f"{z}_image", "")
        text_raw = order_data.get(f"{z}_text", "")
        font_raw = order_data.get(f"{z}_fonts", "") or order_data.get("front_fonts", "")
        col_raw  = order_data.get(f"{z}_colours", "") or order_data.get("front_colours", "")

        if log_fn: log_fn(f"Rendering zone: {z.upper()}", "info")

        # ── Label layer ──────────────────────────────────────────────────────
        label_img = Image.new("RGBA", (canvas_w, LABEL_HEIGHT_PX), (0, 0, 0, 0))
        label_draw = ImageDraw.Draw(label_img)
        label_font = get_font("Arial", max(6, LABEL_HEIGHT_PX - 6))
        label_draw.text((4, 2), z, font=label_font, fill=(0, 0, 0, 255))

        psd_layers.append({
            "name":    f"label_{z}",
            "image":   label_img,
            "top":     y_offset,
            "left":    0,
            "opacity": 255,
            "visible": True,
        })
        y_offset += LABEL_HEIGHT_PX

        # ── Zone content layer ───────────────────────────────────────────────
        # Centre zone horizontally if narrower than canvas
        x_offset = (canvas_w - zw) // 2

        zone_img = render_zone(z, img_file, text_raw, font_raw, col_raw,
                               zw, zh, log_fn)

        # Layer name matches the image filename (e.g. "61450659810442-1-back")
        layer_name = img_file.replace(".jpg","").replace(".png","") if img_file else z

        psd_layers.append({
            "name":    layer_name,
            "image":   zone_img,
            "top":     y_offset,
            "left":    x_offset,
            "opacity": 255,
            "visible": True,
        })
        y_offset += zh + ZONE_GAP_PX

    # ── Write PSD ────────────────────────────────────────────────────────────
    write_psd(out_path, canvas_w, total_h, psd_layers, log_fn=log_fn)
    return True, "OK"
