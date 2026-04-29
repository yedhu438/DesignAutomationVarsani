
content = open(r'W:\VarsaniAutomation\design_replicator.py', encoding='utf-8').read()

OLD = '''def replicate_design(img_path, text_lines, font_name, colour_hex,
                     canvas_w=CANVAS_W, canvas_h=CANVAS_H):
    """
    Replicates the crssoft preview layout:
    - Text at top (styled, centred)
    - Customer photo below with thin border
    Returns a PIL RGBA image.
    """
    has_text  = bool(text_lines)
    has_image = bool(img_path and os.path.isfile(img_path))

    # ── 1. Create transparent canvas ──────────────────────────────────────────
    canvas = Image.new("RGBA", (canvas_w, canvas_h), (255, 255, 255, 0))
    draw   = ImageDraw.Draw(canvas)

    # ── 2. Determine zones based on content ──────────────────────────────────
    if has_text and has_image:
        text_top    = int(canvas_h * 0.02)
        text_height = int(canvas_h * TEXT_ZONE_H_PCT)
        photo_top   = int(canvas_h * PHOTO_TOP_PCT)
        photo_bot   = int(canvas_h * PHOTO_BOTTOM_PCT)
        photo_left  = int(canvas_w * PHOTO_LEFT_PCT)
        photo_right = int(canvas_w * PHOTO_RIGHT_PCT)
    elif has_image and not has_text:
        # Full canvas photo
        photo_top   = int(canvas_h * 0.03)
        photo_bot   = int(canvas_h * 0.97)
        photo_left  = int(canvas_w * 0.03)
        photo_right = int(canvas_w * 0.97)
        text_top    = 0
        text_height = 0
    else:
        # Text only — centred
        text_top    = int(canvas_h * 0.30)
        text_height = int(canvas_h * 0.40)
        photo_top = photo_bot = photo_left = photo_right = 0

    # ── 3. Place customer image ───────────────────────────────────────────────
    if has_image:
        pw = photo_right - photo_left
        ph = photo_bot   - photo_top

        # Load + resize with OpenCV for quality upscaling
        cv_img = cv2.imread(img_path, cv2.IMREAD_UNCHANGED)
        if cv_img is not None:
            # Convert BGR/BGRA → RGBA
            if cv_img.shape[2] == 4:
                cv_img = cv2.cvtColor(cv_img, cv2.COLOR_BGRA2RGBA)
            else:
                cv_img = cv2.cvtColor(cv_img, cv2.COLOR_BGR2RGBA)

            # Scale to fit zone (maintain aspect ratio)
            ih, iw = cv_img.shape[:2]
            scale  = min(pw / iw, ph / ih)
            nw     = max(1, int(iw * scale))
            nh     = max(1, int(ih * scale))

            # Use LANCZOS4 for high quality resize
            resized = cv2.resize(cv_img, (nw, nh), interpolation=cv2.INTER_LANCZOS4)
            pil_photo = Image.fromarray(resized, 'RGBA')
        else:
            # Fallback to PIL
            pil_photo = Image.open(img_path).convert("RGBA")
            ih, iw    = pil_photo.height, pil_photo.width
            scale     = min(pw / iw, ph / ih)
            nw        = max(1, int(iw * scale))
            nh        = max(1, int(ih * scale))
            pil_photo = pil_photo.resize((nw, nh), Image.LANCZOS)

        # Centre photo in zone
        x_off = photo_left + (pw - nw) // 2
        y_off = photo_top  + (ph - nh) // 2

        # Draw border behind photo
        if PHOTO_BORDER_PX > 0:
            br = PHOTO_BORDER_COLOR
            border_rect = [
                x_off - PHOTO_BORDER_PX,
                y_off - PHOTO_BORDER_PX,
                x_off + nw + PHOTO_BORDER_PX,
                y_off + nh + PHOTO_BORDER_PX,
            ]
            draw.rectangle(border_rect, fill=(br[0], br[1], br[2], 255))

        # Paste photo onto canvas
        canvas.paste(pil_photo, (x_off, y_off), pil_photo)

    # ── 4. Render text ────────────────────────────────────────────────────────
    if has_text and text_lines:
        r, g, b   = hex_to_rgb(colour_hex)
        text_color = (r, g, b, 255)

        avail_w   = int(canvas_w * 0.92)
        avail_h   = text_height if text_height > 0 else int(canvas_h * 0.15)
        longest   = max(text_lines, key=len)

        # Binary search for best font size that fits width
        lo, hi, best = 18, avail_h // max(1, len(text_lines)), 72
        scratch = Image.new("RGBA", (1, 1))
        sd = ImageDraw.Draw(scratch)
        while lo <= hi:
            mid  = (lo + hi) // 2
            font = get_pil_font(font_name, mid)
            bb   = sd.textbbox((0,0), longest, font=font)
            if (bb[2]-bb[0]) <= avail_w:
                best = mid; lo = mid + 1
            else:
                hi = mid - 1

        font   = get_pil_font(font_name, best)
        bb0    = sd.textbbox((0,0), text_lines[0], font=font)
        line_h = int((bb0[3]-bb0[1]) * 1.3)
        total_text_h = line_h * len(text_lines)

        # Centre text block within text zone
        yt = text_top + max(0, (avail_h - total_text_h) // 2)

        for line in text_lines:
            bb  = draw.textbbox((0,0), line, font=font)
            lw  = bb[2] - bb[0]
            xt  = (canvas_w - lw) // 2
            # Draw subtle text shadow for legibility
            draw.text((xt+2, yt+2), line, font=font, fill=(0,0,0,60))
            draw.text((xt,   yt),   line, font=font, fill=text_color)
            yt += line_h

    return canvas'''

NEW = '''def replicate_design(img_path, text_lines, font_name, colour_hex,
                     canvas_w=CANVAS_W, canvas_h=CANVAS_H):
    """
    Replicates the EXACT crssoft preview layout measured with OpenCV:

    Layout (all coordinates measured from real preview images):
      - Customer photo fills the canvas: top=2%, bottom=95%, left=5%, right=95%
      - Text is overlaid ON TOP of the photo at the bottom: y=70%-92%
      - Text is centred horizontally
      - Shadow under text for legibility on any background

    For text-only orders (no image): text is centred in the canvas.
    """
    has_text  = bool(text_lines)
    has_image = bool(img_path and os.path.isfile(img_path))

    # ── 1. Transparent canvas ─────────────────────────────────────────────────
    canvas = Image.new("RGBA", (canvas_w, canvas_h), (255, 255, 255, 0))
    draw   = ImageDraw.Draw(canvas)

    # ── 2. Photo zone ─────────────────────────────────────────────────────────
    if has_image:
        photo_top   = int(canvas_h * PHOTO_TOP_PCT)
        photo_bot   = int(canvas_h * PHOTO_BOTTOM_PCT)
        photo_left  = int(canvas_w * PHOTO_LEFT_PCT)
        photo_right = int(canvas_w * PHOTO_RIGHT_PCT)
        pw = photo_right - photo_left
        ph = photo_bot   - photo_top

        # Load with OpenCV (LANCZOS4 = highest quality resize)
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

        # Centre photo in its zone
        x_off = photo_left + (pw - nw) // 2
        y_off = photo_top  + (ph - nh) // 2

        # Thin border behind photo
        if PHOTO_BORDER_PX > 0:
            bc = PHOTO_BORDER_COLOR
            draw.rectangle([x_off-PHOTO_BORDER_PX, y_off-PHOTO_BORDER_PX,
                            x_off+nw+PHOTO_BORDER_PX, y_off+nh+PHOTO_BORDER_PX],
                           fill=(bc[0], bc[1], bc[2], 255))

        canvas.paste(pil_photo, (x_off, y_off), pil_photo)

    # ── 3. Text — overlaid at bottom of photo (or centred if no image) ────────
    if has_text and text_lines:
        r, g, b    = hex_to_rgb(colour_hex)
        text_color = (r, g, b, 255)

        if has_image:
            # Text zone: bottom 22% of canvas (measured: 70%-92%)
            txt_zone_top = int(canvas_h * TEXT_TOP_PCT)
            txt_zone_h   = int(canvas_h * (TEXT_BOTTOM_PCT - TEXT_TOP_PCT))
        else:
            # Text only: centred vertically
            txt_zone_top = int(canvas_h * 0.25)
            txt_zone_h   = int(canvas_h * 0.50)

        avail_w = int(canvas_w * 0.90)
        longest = max(text_lines, key=len)

        # Find largest font size that fits the available width
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
            # Dark shadow for legibility on any background colour
            draw.text((xt+2, yt+2), line, font=font, fill=(0, 0, 0, 120))
            draw.text((xt,   yt),   line, font=font, fill=text_color)
            yt += line_h

    return canvas'''

if OLD in content:
    content = content.replace(OLD, NEW)
    open(r'W:\VarsaniAutomation\design_replicator.py', 'w', encoding='utf-8').write(content)
    print('Step 2 OK: replicate_design function updated')
else:
    print('ERROR step 2: function body not found')
    idx = content.find('def replicate_design')
    print('Found at index:', idx)
