"""
Fix: Groups rows by OrderID and builds merged PSDs.
Owner's rule:
  - Same design for all items → simple label ("front")
  - Different designs → label shows colour+size ("Front - White XL", "Front - White Medium")
"""

content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()

# ─── Add build_merged_psd_for_order after build_psd_for_order ────────────────
# Find the DATABASE section marker to insert before it
INSERT_BEFORE = '# ─── DATABASE ─────────────────────────────────────────────────────────────────'

NEW_FUNCTION = '''
def rows_have_same_design(rows):
    """
    Returns True if ALL rows in this order group have identical designs.
    Identical means: same FrontText, same FrontImageJSON, same FrontImage,
    same FrontFonts, same FrontColours.
    """
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
    lbl_h    = build_label_layer("front").height + cm_to_px(0.3)

    # Build zones for every row, attaching display_label based on the rule
    all_row_zones = []
    for row in rows:
        sku     = row.get("SKU") or ""
        product = detect_product(sku)
        zones   = build_zones(row, product)

        for z in zones:
            if same_design:
                z["display_label"] = z["label"].title()  # "front" → "Front"
            else:
                z["display_label"] = make_zone_label(z["label"], sku, use_sku_detail=True)
        all_row_zones.append(zones)

    # Canvas width = widest zone across all rows
    all_zones_flat = [z for zones in all_row_zones for z in zones]
    if not all_zones_flat:
        return False, "No zones in any row"

    max_zw   = max(z["w"] for z in all_zones_flat)
    canvas_w = PADDING + max_zw + PADDING

    # Canvas height = sum of all rows (each row: label + content + gap)
    def row_height(zones):
        if not zones:
            return 0
        return lbl_h + max(z["h"] for z in zones)

    canvas_h = (PADDING
                + sum(row_height(zones) for zones in all_row_zones)
                + QTY_GAP * (len(all_row_zones) - 1)
                + PADDING)

    all_layers = []
    y_cursor   = PADDING

    for row_idx, (row, zones) in enumerate(zip(rows, all_row_zones)):
        if not zones:
            continue

        # For multi-zone rows (front + back), arrange within row
        row_h = row_height(zones)
        x_cursor = PADDING
        max_zh = max(z["h"] for z in zones)

        for zone in zones:
            zw, zh = zone["w"], zone["h"]
            x_left = PADDING + (max_zw - zw) // 2
            display_label = zone.get("display_label") or zone["label"].title()

            # Label
            lbl = build_label_layer(display_label)
            all_layers.append({
                "name": f"{display_label} label",
                "image": lbl,
                "top": y_cursor,
                "left": x_left,
                "opacity": 255, "visible": True,
            })

            img_start = y_cursor + lbl_h

            # Image
            if zone["img_path"]:
                img_pil, it, il = build_image_layer(zone["img_path"], zw, zh)
                all_layers.append({
                    "name": f"{display_label} CustomerImage",
                    "image": img_pil,
                    "top": img_start + it,
                    "left": x_left + il,
                    "opacity": 255, "visible": True,
                })
            elif zone["img_filename"]:
                log(f"    WARNING image not found: {zone['img_filename']}", "WARN")

            # Text
            if zone["text_lines"]:
                txt_pil, tt, tl = build_text_layer(
                    zone["text_lines"], zone["font"], zone["colour"], zw, zh,
                    has_image=bool(zone["img_path"]))
                all_layers.append({
                    "name": f"{display_label} CustomerText",
                    "image": txt_pil,
                    "top": img_start + tt,
                    "left": x_left + tl,
                    "opacity": 255, "visible": True,
                })

        y_cursor += row_h
        if row_idx < len(all_row_zones) - 1:
            y_cursor += QTY_GAP

    write_psd(out_path, canvas_w, canvas_h, all_layers)

    if not os.path.isfile(out_path):
        return False, "PSD file not written"

    size_mb = os.path.getsize(out_path) / (1024 * 1024)
    labels  = [z.get("display_label") for zones in all_row_zones for z in zones]
    return True, f"{size_mb:.1f} MB | labels: {labels}"


'''

if INSERT_BEFORE in content:
    content = content.replace(INSERT_BEFORE, NEW_FUNCTION + INSERT_BEFORE)
    print("Step 3 OK: build_merged_psd_for_order_group inserted")
else:
    print("ERROR step 3: DATABASE marker not found")

open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)
print("Saved.")
