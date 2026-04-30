
import re

content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()
lines   = content.splitlines(keepends=True)

# ── 1. Fix the broken code inside build_psd_for_order ─────────────────────────
# Around line 873: 'if zone.get("preview_url"):' has no body — the new loop was
# incorrectly injected as its body. We need to restore the original body.
# The new-loop block goes from line 874 to line 946 (0-indexed 873-945).
# After it, line 947 is 'y_cursor += ...' which belongs to build_psd_for_order.

# Find the bad injection start: right after 'if zone.get("preview_url"):'
BAD_START_MARKER = '        if zone.get("preview_url"):\n    # Group rows by OrderID'
# Find second bad injection: inside rows_have_same_design 
BAD_START_MARKER2 = '            (row.get("FrontText") or "").strip(),\n    # Group rows by OrderID'

# Correct body for 'if zone.get("preview_url"):'
CORRECT_PREVIEW_BODY = '''        if zone.get("preview_url"):
            pi = download_preview(zone["preview_url"])
            if pi:
                ratio = min(zw / pi.width, zh / pi.height)
                pnw = max(1, int(pi.width  * ratio))
                pnh = max(1, int(pi.height * ratio))
                prev_img = pi.resize((pnw, pnh), Image.LANCZOS)
'''

# The new loop code block (to be removed from both bad locations)
NEW_LOOP_BLOCK = '''    # Group rows by OrderID so we can build ONE merged PSD per order
    # (e.g. 2 shirts in same order \u2192 stacked vertically in one PSD)
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
                # Single item -- use original function (handles quantity field too)
                ok, msg = build_psd_for_order(order_id, first_row, out_path)
            else:
                # Multiple items -- use merged builder
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

# ── Fix 1: Restore preview_url block ─────────────────────────────────────────
# Find: 'if zone.get("preview_url"):' followed by the bad loop
# Replace with correct body + nothing after (loop goes at end of y_cursor line)

# Find the position of the bad injection 1
bad1_start = content.find('        if zone.get("preview_url"):\n    # Group rows')
bad1_end   = content.find('        y_cursor += LABEL_H', bad1_start)

if bad1_start != -1 and bad1_end != -1:
    print(f"Bad injection 1: chars {bad1_start} to {bad1_end}")
    content = content[:bad1_start] + CORRECT_PREVIEW_BODY + content[bad1_end:]
    print("Fix 1 applied: preview_url body restored")
else:
    print(f"ERROR fix1: bad1_start={bad1_start}, bad1_end={bad1_end}")

# ── Fix 2: Restore rows_have_same_design ─────────────────────────────────────
bad2_start = content.find('            (row.get("FrontText") or "").strip(),\n    # Group rows')
# End: find '        )' after the bad injection to restore the return tuple
bad2_end   = content.find('    # Canvas width', bad2_start)

if bad2_start != -1 and bad2_end != -1:
    print(f"Bad injection 2: chars {bad2_start} to {bad2_end}")
    # Restore the rest of the sig() function's return tuple
    CORRECT_SIG = '''            (row.get("FrontText") or "").strip(),
            (row.get("FrontImageJSON") or "").strip(),
            (row.get("FrontImage") or "").strip(),
            (row.get("FrontFonts") or "").strip(),
            (row.get("FrontColours") or "").strip(),
        )
    first = sig(rows[0])
    return all(sig(r) == first for r in rows)

'''
    content = content[:bad2_start] + CORRECT_SIG + content[bad2_end:]
    print("Fix 2 applied: rows_have_same_design restored")
else:
    print(f"ERROR fix2: bad2_start={bad2_start}, bad2_end={bad2_end}")

# ── Fix 3: Replace the OLD loop in run_batch with the new grouped loop ────────
old_loop_start = content.find('    for i, row in enumerate(orders, 1):')
# End at the log("DONE") line
old_loop_end   = content.find('    log("=" * 60)\n    log(f"DONE', old_loop_start)

if old_loop_start != -1 and old_loop_end != -1:
    print(f"Old loop in run_batch: chars {old_loop_start} to {old_loop_end}")
    content = content[:old_loop_start] + NEW_LOOP_BLOCK + '\n' + content[old_loop_end:]
    print("Fix 3 applied: run_batch loop replaced with grouped version")
else:
    print(f"ERROR fix3: old_loop_start={old_loop_start}, old_loop_end={old_loop_end}")

# ── Save ─────────────────────────────────────────────────────────────────────
open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)
print(f"\nSaved. Total chars: {len(content)}")

# ── Verify syntax ─────────────────────────────────────────────────────────────
import ast
try:
    ast.parse(content)
    print("SYNTAX OK")
except SyntaxError as e:
    print(f"SYNTAX ERROR: {e}")
