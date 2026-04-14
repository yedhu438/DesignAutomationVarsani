
content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()

# Find the text placement line in build_text_layer and replace it
OLD_PLACEMENT = '''    bh   = img.height
    bw   = img.width
    # Text position:
    #   - has_image=True  → text at TOP (above the photo), matching real t-shirt layout
    #   - has_image=False → text centred vertically (text-only orders)
    if has_image:
        top = 10   # small top margin, sits above the customer photo
    else:
        top = max(0, (h - bh) // 2)  # centred for text-only
    left = max(0, (w - bw) // 2)
    return img, top, left'''

NEW_PLACEMENT = '''    bh   = img.height
    bw   = img.width
    # Text position — uses layout dict from preview_analyser if available,
    # otherwise falls back to simple centred (text-only) positioning.
    # NOTE: when called from build_psd_for_order, the text is composited
    # onto the canvas using copy_top offset, so 'top' here is relative
    # to the zone origin, not the whole canvas.
    left = max(0, (w - bw) // 2)
    if has_image:
        # Default: text at bottom overlay position (will be overridden by
        # preview_analyser measurements in build_psd_for_order)
        top = max(0, int(h * 0.70))
    else:
        top = max(0, (h - bh) // 2)  # centred for text-only
    return img, top, left'''

if OLD_PLACEMENT in content:
    content = content.replace(OLD_PLACEMENT, NEW_PLACEMENT)
    print("Step 2 OK: text placement updated")
else:
    # Try to find it with different whitespace
    idx = content.find('    bh   = img.height\n    bw   = img.width')
    if idx != -1:
        print(f"Found at char {idx}, checking context...")
        print(repr(content[idx:idx+500]))
    else:
        print("ERROR step 2: placement block not found")

open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)
print("Saved.")
