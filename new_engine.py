# ─── IMAGE SERVER CONFIG ─────────────────────────────────────────────────────
IMAGE_SERVER_URL = ""   # Ask Vikesh for base URL e.g. "https://crssoft.co.uk/uploads/"

def _download_image(filename, log_fn=None):
    if not IMAGE_SERVER_URL or not filename or not filename.strip():
        return None
    import urllib.request
    url = IMAGE_SERVER_URL.rstrip("/") + "/" + filename.strip()
    try:
        tmp = os.path.join(TEMP_FOLDER, f"dl_{uuid.uuid4().hex[:8]}_{filename}")
        urllib.request.urlretrieve(url, tmp)
        img = Image.open(tmp).convert("RGBA")
        try: os.remove(tmp)
        except: pass
        return img
    except Exception as e:
        if log_fn: log_fn(f"  Could not download {filename}: {e}", "warning")
        return None

def _parse_font(fonts_raw):
    if not fonts_raw: return "Arial"
    s = fonts_raw.strip()
    if s.startswith("{"):
        import json
        try:
            d = json.loads(s)
            return d.get("NormalFont") or d.get("PremiumFont") or "Arial"
        except: pass
    return s

def _parse_colour(colours_raw):
    if not colours_raw: return "#ffffff"
    s = colours_raw.strip()
    if s.startswith("{"):
        import json
        try:
            d = json.loads(s)
            return d.get("Colour1") or "#ffffff"
        except: pass
    return s if s.startswith("#") else "#ffffff"

def _build_zone_content(zone_name, w, h, img_path, img_pil, text_lines, font_name, colour_hex, remove_bg, log_fn):
    layers = []
    src_img = None
    if img_pil:
        src_img = img_pil.convert("RGBA")
    elif img_path and os.path.isfile(img_path):
        src_img = Image.open(img_path).convert("RGBA")
    if src_img:
        ratio = min(w / src_img.width, h / src_img.height)
        nw = max(1, int(src_img.width * ratio))
        nh = max(1, int(src_img.height * ratio))
        src_img = src_img.resize((nw, nh), Image.LANCZOS)
        if remove_bg:
            bg_r, bg_g, bg_b = src_img.getpixel((4,4))[:3]
            px = src_img.load()
            for py in range(nh):
                for pxx in range(nw):
                    r,g,b,a = src_img.getpixel((pxx,py))
                    if abs(r-bg_r)<50 and abs(g-bg_g)<50 and abs(b-bg_b)<50:
                        px[pxx,py] = (r,g,b,0)
        top = (h - nh) // 2
        left = (w - nw) // 2
        layers.append({"name": "CustomerImage", "image": src_img, "top": top, "left": left, "opacity": 255, "visible": True})
        log_fn(f"  [{zone_name}] Image: {nw}x{nh}px", "info")
    if text_lines:
        r, g, b = hex_to_rgb(colour_hex)
        avail_w = int(w * 0.90)
        longest = max(text_lines, key=len)
        scratch = Image.new("RGBA", (1,1))
        draw = ImageDraw.Draw(scratch)
        lo, hi, best = 20, min(900, h // max(1, len(text_lines))), 60
        while lo <= hi:
            mid = (lo + hi) // 2
            font = get_font(font_name, mid)
            bb = draw.textbbox((0,0), longest, font=font)
            if (bb[2]-bb[0]) <= avail_w: best=mid; lo=mid+1
            else: hi=mid-1
        font = get_font(font_name, best)
        bb0 = draw.textbbox((0,0), text_lines[0], font=font)
        line_h = int((bb0[3]-bb0[1]) * 1.25)
        max_lw = max(draw.textbbox((0,0),l,font=font)[2]-draw.textbbox((0,0),l,font=font)[0] for l in text_lines)
        bw = min(max_lw+40, w)
        bh = line_h * len(text_lines) + 40
        txt_img = Image.new("RGBA", (bw, bh), (0,0,0,0))
        d2 = ImageDraw.Draw(txt_img)
        yl = 20
        for line in text_lines:
            bb = d2.textbbox((0,0), line, font=font)
            lw = bb[2]-bb[0]
            d2.text((max(0,(bw-lw)//2), yl), line, font=font, fill=(r,g,b,255))
            yl += line_h
        top = max(int(h * 0.70), h - bh - 60)
        left = max(0, (w - bw) // 2)
        layers.append({"name": "CustomerText", "image": txt_img, "top": top, "left": left, "opacity": 255, "visible": True})
        log_fn(f"  [{zone_name}] Text: '{' | '.join(text_lines[:2])}' {font_name} {colour_hex}", "info")
    return layers

def build_multizone_psd(order_id, amazon_id, zones, out_path, log_fn):
    PADDING = cm_to_px(1)
    GAP     = cm_to_px(1)
    LABEL_H = cm_to_px(1)
    canvas_w = max(z['w'] for z in zones) + PADDING * 2
    canvas_h = PADDING + sum(LABEL_H + z['h'] + GAP for z in zones) + PADDING
    log_fn(f"Canvas: {canvas_w}x{canvas_h}px | {len(zones)} zone(s)", "info")
    all_layers = []
    y_cursor = PADDING
    for zone in zones:
        zname = zone['name'].upper()
        zw, zh = zone['w'], zone['h']
        x_off = (canvas_w - zw) // 2
        label_img = Image.new("RGBA", (canvas_w, LABEL_H), (0,0,0,0))
        d = ImageDraw.Draw(label_img)
        lf_size = max(20, int(LABEL_H * 0.6))
        try:    lf = ImageFont.truetype("arial.ttf", lf_size)
        except: lf = ImageFont.load_default()
        d.text((x_off, max(0,(LABEL_H - lf_size) // 2)), zname, font=lf, fill=(0,0,0,255))
        all_layers.append({"name": f"label_{zone['name']}", "image": label_img, "top": y_cursor, "left": 0, "opacity": 255, "visible": True})
        y_cursor += LABEL_H
        zone_layers = _build_zone_content(
            zone_name=zname, w=zw, h=zh,
            img_path=zone.get('img_path',''), img_pil=zone.get('img_pil'),
            text_lines=zone.get('text_lines',[]), font_name=zone.get('font','Arial'),
            colour_hex=zone.get('colour','#ffffff'), remove_bg=zone.get('remove_bg',False),
            log_fn=log_fn)
        base = os.path.splitext(zone.get('img_filename', zone['name']))[0]
        for lyr in zone_layers:
            suffix = lyr["name"].replace("Customer","").lower()
            lyr["name"]  = f"{base}-{suffix}"
            lyr["top"]   = lyr["top"]  + y_cursor
            lyr["left"]  = lyr["left"] + x_off
            all_layers.append(lyr)
        y_cursor += zh + GAP
    write_psd(out_path, canvas_w, canvas_h, all_layers, log_fn=log_fn)
    if not os.path.isfile(out_path):
        return False, f"PSD not found: {out_path}"
    size_mb = os.path.getsize(out_path) / (1024*1024)
    log_fn(f"PSD saved: {size_mb:.1f} MB", "success")
    return True, "OK"

def run_automation(order_id, detail_id, amazon_id, order_data):
    def log(msg, level="info"):
        log_progress(order_id, msg, level)
    try:
        log("Starting automation pipeline...", "info")
        product   = order_data.get("product", "tshirt")
        font_name = _parse_font(order_data.get("font", "Arial"))
        colour    = _parse_colour(order_data.get("colour", "#ffffff"))
        remove_bg = order_data.get("remove_bg", False)
        spec      = PRODUCT_CANVAS.get(product, PRODUCT_CANVAS["tshirt"])
        zone_defs = [
            ("front",  "text",        "image_path",        "front_img_filename"),
            ("back",   "back_text",   "back_image_path",   "back_img_filename"),
            ("pocket", "pocket_text", "pocket_image_path", "pocket_img_filename"),
            ("sleeve", "sleeve_text", "sleeve_image_path", "sleeve_img_filename"),
        ]
        zones = []
        for zname, text_key, img_key, fname_key in zone_defs:
            text_raw  = order_data.get(text_key,  "") or ""
            img_path  = order_data.get(img_key,   "") or ""
            img_fname = order_data.get(fname_key, "") or ""
            text_lines = parse_texts(text_raw) if text_raw.strip() else []
            if not text_lines and not img_path and not img_fname:
                continue
            dims = spec.get(zname, spec.get("front"))
            w, h = dims
            img_pil = None
            if img_fname and not img_path:
                log(f"Downloading {zname} image: {img_fname}...", "info")
                img_pil = _download_image(img_fname, log)
            zones.append({
                "name": zname, "w": w, "h": h,
                "img_path": img_path, "img_pil": img_pil,
                "img_filename": img_fname or os.path.basename(img_path or zname),
                "text_lines": text_lines, "font": font_name,
                "colour": colour, "remove_bg": remove_bg,
            })
            log(f"Zone [{zname}]: text={text_lines} | img={img_fname or img_path or 'none'}", "info")
        if not zones:
            zone = order_data.get("zone", "front")
            dims = spec.get(zone, spec.get("front"))
            w, h = dims
            text_lines = parse_texts(order_data.get("text","")) if (order_data.get("text","") or "").strip() else []
            zones.append({
                "name": zone, "w": w, "h": h,
                "img_path": order_data.get("image_path",""), "img_pil": None,
                "img_filename": os.path.basename(order_data.get("image_path","") or zone),
                "text_lines": text_lines, "font": font_name, "colour": colour, "remove_bg": remove_bg,
            })
        log(f"Zones: {[z['name'] for z in zones]}", "info")
        today    = datetime.now().strftime("%Y-%m-%d")
        out_dir  = os.path.join(OUTPUT_FOLDER, today)
        os.makedirs(out_dir, exist_ok=True)
        safe_id  = amazon_id.replace("/", "-")
        out_path = os.path.join(out_dir, f"{safe_id}.psd")
        success, message = build_multizone_psd(
            order_id=order_id, amazon_id=amazon_id,
            zones=zones, out_path=out_path, log_fn=log)
        if not success:
            log(f"PSD failed: {message}", "error")
            flat = out_path.replace(".psd", "_FLAT.png")
            z0 = zones[0]
            _save_flat_png(z0['w'], z0['h'], z0['img_path'], z0['text_lines'], z0['font'], z0['colour'], flat)
            out_path = flat
        else:
            size_mb = os.path.getsize(out_path) / (1024*1024)
            log(f"PSD: {out_path} ({size_mb:.1f} MB)", "success")
            log(f"Zones in file: {[z['name'] for z in zones]}", "success")
        log("Updating database...", "info")
        mark_order_complete(detail_id, out_path)
        log(f"Order {amazon_id} complete!", "success")
        progress_logs[order_id].append({"done": True, "file": out_path})
    except Exception as e:
        import traceback
        log_progress(order_id, f"Pipeline error: {str(e)}", "error")
        progress_logs[order_id].append({"done": True, "error": str(e)})
