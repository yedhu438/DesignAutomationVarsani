
content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()

OLD_LOOP = '''    for i, row in enumerate(orders, 1):
        order_id  = row["OrderID"]
        detail_id = row["idCustomOrderDetails"]
        sku       = (row.get("SKU") or "").replace("/", "-").replace("\\\\", "-")
        safe_id   = order_id.replace("/", "-")
        category  = detect_category(row.get("SKU") or "")
        cat_dir   = os.path.join(out_dir, category)
        os.makedirs(cat_dir, exist_ok=True)
        base_path = os.path.join(cat_dir, f"{safe_id}_{sku}.psd")
        out_path  = base_path
        counter   = 2
        while os.path.exists(out_path):
            out_path = base_path.replace(".psd", f"_{counter}.psd")
            counter += 1

        log(f"[{i}/{total}] {order_id}  |  {row.get('SKU','')}  |  {row.get('PrintLocation','')}")

        if dry_run:
            product = detect_product(row.get("SKU") or "")
            zones   = build_zones(row, product)
            for z in zones:
                status = "FOUND" if z["img_path"] else ("MISSING" if z["img_filename"] else "text-only")
                log(f"  [{z['label']}]  img={z['img_filename'] or 'none'} ({status})  text={z['text_lines']}", "DRY")
            if not zones:
                log("  SKIP — no zones", "DRY")
            continue

        try:
            ok, msg = build_psd_for_order(order_id, row, out_path)
            if ok:
                mark_complete(detail_id, out_path)
                log(f"  OK  {msg}", "OK")
                ok_count += 1
            else:
                log(f"  FAIL  {msg}", "FAIL")
                fail_count += 1
        except Exception as e:
            log(f"  ERROR  {e}", "ERROR")
            log(traceback.format_exc()[-400:], "ERROR")
            fail_count += 1'''

# Check if the loop exists as-is (backslash may be escaped differently)
if '        for i, row in enumerate(orders, 1):' in content:
    print("Loop found - searching for exact match...")
else:
    print("Basic loop start found")

# Find the loop boundaries more precisely
start_marker = '    for i, row in enumerate(orders, 1):'
end_marker   = '        if i % 50 == 0:'

idx_start = content.find(start_marker)
idx_end   = content.find(end_marker)

if idx_start == -1 or idx_end == -1:
    print(f"ERROR: start={idx_start}, end={idx_end}")
else:
    old_loop = content[idx_start:idx_end]
    print(f"Found loop: chars {idx_start} to {idx_end}, length {len(old_loop)}")
    print("Preview:", repr(old_loop[:200]))
