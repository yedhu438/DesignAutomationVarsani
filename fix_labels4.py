
lines = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').readlines()

# Find line 874 (0-indexed: 873) - the start of "for i, row in enumerate"
# Find line 916 (0-indexed: 915) - the "if i % 50 == 0:" line
# We replace lines 873..915 inclusive with the new grouped loop

NEW_LOOP = '''    # Group rows by OrderID so we can build ONE merged PSD per order
    # (e.g. 2 shirts in same order → stacked vertically in one PSD)
    from collections import OrderedDict
    order_groups = OrderedDict()
    for row in orders:
        oid = row["OrderID"]
        if oid not in order_groups:
            order_groups[oid] = []
        order_groups[oid].append(row)

    total_orders = len(order_groups)
    log(f"Unique orders to process: {total_orders} (from {total} rows)")

    for i, (order_id, group_rows) in enumerate(order_groups.items(), 1):
        safe_id  = order_id.replace("/", "-")
        # Use the first row's SKU for folder/filename
        first_row = group_rows[0]
        sku       = (first_row.get("SKU") or "").replace("/", "-").replace("\\\\", "-")
        category  = detect_category(first_row.get("SKU") or "")
        cat_dir   = os.path.join(out_dir, category)
        os.makedirs(cat_dir, exist_ok=True)

        # Filename: use OrderID + SKU (if multi-item, note count)
        if len(group_rows) > 1:
            fname = f"{safe_id}_{len(group_rows)}items.psd"
        else:
            fname = f"{safe_id}_{sku}.psd"
        base_path = os.path.join(cat_dir, fname)
        out_path  = base_path
        counter   = 2
        while os.path.exists(out_path):
            out_path = base_path.replace(".psd", f"_{counter}.psd")
            counter += 1

        skus_str = " | ".join(r.get("SKU","") for r in group_rows)
        log(f"[{i}/{total_orders}] {order_id}  ({len(group_rows)} items)  |  {skus_str}")

        if dry_run:
            same = rows_have_same_design(group_rows)
            log(f"  same_design={same}", "DRY")
            for row in group_rows:
                product = detect_product(row.get("SKU") or "")
                zones   = build_zones(row, product)
                for z in zones:
                    status = "FOUND" if z["img_path"] else ("MISSING" if z["img_filename"] else "text-only")
                    label  = make_zone_label(z["label"], row.get("SKU",""), not same)
                    log(f"  [{label}]  img={z['img_filename'] or 'none'} ({status})  text={z['text_lines']}", "DRY")
            continue

        try:
            if len(group_rows) == 1:
                # Single item — use original function (handles quantity field too)
                ok, msg = build_psd_for_order(order_id, first_row, out_path)
            else:
                # Multiple items — use merged builder
                ok, msg = build_merged_psd_for_order_group(order_id, group_rows, out_path)

            if ok:
                # Mark all rows in the group as complete
                for row in group_rows:
                    mark_complete(row["idCustomOrderDetails"], out_path)
                log(f"  OK  {msg}", "OK")
                ok_count += 1
            else:
                log(f"  FAIL  {msg}", "FAIL")
                fail_count += 1
        except Exception as e:
            log(f"  ERROR  {e}", "ERROR")
            log(traceback.format_exc()[-400:], "ERROR")
            fail_count += 1

        if i % 50 == 0:
            log(f"--- Progress {i}/{total_orders}  ok={ok_count}  fail={fail_count} ---")
'''

# Lines are 0-indexed. Line 874 in the file = index 873
# Line 916 = index 915 ("if i % 50 == 0:")
# We keep lines 0..872 and 915.. (the if i%50 line onwards)

start_idx = 873   # "    for i, row in enumerate(orders, 1):"
end_idx   = 915   # "        if i % 50 == 0:"  — keep this line onwards

before = lines[:start_idx]
after  = lines[end_idx:]   # keeps "if i % 50 == 0:" and everything after

new_lines = before + [NEW_LOOP] + after
open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').writelines(new_lines)
print(f"Done. Total lines: {len(new_lines)}")
print("Verifying key lines:")
# Read back and check
content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()
if 'order_groups' in content and 'build_merged_psd_for_order_group' in content:
    print("OK: grouped loop present")
if 'rows_have_same_design' in content:
    print("OK: same_design check present")
if 'make_zone_label' in content:
    print("OK: make_zone_label present")
