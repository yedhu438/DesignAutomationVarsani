
content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()

# The build_psd_for_order receives a single row at a time, but each row has
# a SKU and quantity. For multi-row orders (different SKUs) the rows are
# processed separately. So we only need to handle the within-row case:
# If quantity > 1 AND all copies share the SAME text/image → no SKU detail in label
# If quantity > 1 AND designs differ (different text per copy) → use SKU detail

# Actually looking at the Photoshop, each row IS a separate PSD row in the DB.
# The "Merge" PSD shown has TWO different rows merged:
#   row 1: MenTee_WhtM  "PROUD MUM"
#   row 2: MenTee_WhtXL  "PROUD DAD"
# These are two separate DB rows, merged into ONE PSD by the batch processor
# because they share the same OrderID.
#
# So the rule is: when building the merged PSD for an ORDER (not a single row),
# label each copy with "Front - White Medium", "Front - White XL" etc.
# But the current code processes ONE row = ONE PSD file.
#
# The fix needed: The batch processor needs to GROUP rows by OrderID, then
# for each group decide on label format:
#   - All same design → simple label
#   - Different designs → include colour+size in label

# Step 1: Update make_zone_label to be accessible (already done in fix_labels.py)
# Step 2: In build_psd_for_order, use make_zone_label with the SKU

OLD = '''        # Label strip — sits above all copies, at the canvas edge (not on the image)
        lbl = build_label_layer(zone["label"])
        all_layers.append({
            "name":    f"{zone['label']} label",
            "image":   lbl,
            "top":     y_cursor,
            "left":    x_left,
            "opacity": 255,
            "visible": True,
        })'''

NEW = '''        # Label strip — sits above all copies, at the canvas edge (not on the image)
        # Format: "Front - White XL" if designs differ, or just "front" if identical
        display_label = zone.get("display_label") or zone["label"]
        lbl = build_label_layer(display_label)
        all_layers.append({
            "name":    f"{display_label} label",
            "image":   lbl,
            "top":     y_cursor,
            "left":    x_left,
            "opacity": 255,
            "visible": True,
        })'''

if OLD in content:
    content = content.replace(OLD, NEW)
    print("Step 2a OK: label layer name updated")
else:
    print("ERROR step 2a: label strip not found")

# Step 3: Also update the CustomerImage and CustomerText layer names to use display_label
OLD2 = '''            if img_pil:
                all_layers.append({
                    "name":    f"{zone['label']} CustomerImage{suffix}",'''
NEW2 = '''            if img_pil:
                all_layers.append({
                    "name":    f"{display_label} CustomerImage{suffix}",'''

OLD3 = '''            if txt_pil:
                all_layers.append({
                    "name":    f"{zone['label']} CustomerText{suffix}",'''
NEW3 = '''            if txt_pil:
                all_layers.append({
                    "name":    f"{display_label} CustomerText{suffix}",'''

OLD4 = '''            if prev_img:
                all_layers.append({
                    "name":    f"{zone['label']} Preview Reference{suffix}",'''
NEW4 = '''            if prev_img:
                all_layers.append({
                    "name":    f"{display_label} Preview Reference{suffix}",'''

if OLD2 in content:
    content = content.replace(OLD2, NEW2)
    print("Step 2b OK: CustomerImage label updated")
else:
    print("ERROR step 2b")

if OLD3 in content:
    content = content.replace(OLD3, NEW3)
    print("Step 2c OK: CustomerText label updated")
else:
    print("ERROR step 2c")

if OLD4 in content:
    content = content.replace(OLD4, NEW4)
    print("Step 2d OK: Preview Reference label updated")
else:
    print("ERROR step 2d")

open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)
print("File saved.")
