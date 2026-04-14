"""
integrate_text_layers.py
========================
Patches batch_processor.py to produce editable text layers
instead of rasterised pixel text.

Changes:
1. Import psd_text_layer module
2. In write_psd() - inject TySh+Txt2 blocks into layer extra data
3. In build_text_layer_call sites - attach _text_blocks to layer dict
"""

content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()
print(f"File loaded: {len(content)} chars, {content.count(chr(10))} lines")

# ── Step 1: Add import ────────────────────────────────────────────────────────
OLD_IMPORT = '''# Preview analyser — measures text/photo positions from crssoft preview images
try:
    from preview_analyser import analyse_preview, DEFAULT_LAYOUT
    PREVIEW_ANALYSER_AVAILABLE = True
except ImportError:
    PREVIEW_ANALYSER_AVAILABLE = False'''

NEW_IMPORT = '''# Preview analyser — measures text/photo positions from crssoft preview images
try:
    from preview_analyser import analyse_preview, DEFAULT_LAYOUT
    PREVIEW_ANALYSER_AVAILABLE = True
except ImportError:
    PREVIEW_ANALYSER_AVAILABLE = False

# Editable text layer writer (TySh + Txt2 PSD blocks)
try:
    from psd_text_layer import build_editable_text_tagged_blocks, resolve_ps_font_name
    EDITABLE_TEXT_AVAILABLE = True
except ImportError:
    EDITABLE_TEXT_AVAILABLE = False'''

if OLD_IMPORT in content:
    content = content.replace(OLD_IMPORT, NEW_IMPORT)
    print("Step 1 OK: import added")
else:
    # fallback: add after rembg import
    OLD_IMPORT2 = '''try:
    from rembg import remove as rembg_remove
    REMBG_AVAILABLE = True
except ImportError:
    REMBG_AVAILABLE = False'''
    NEW_IMPORT2 = OLD_IMPORT2 + '''

# Editable text layer writer (TySh + Txt2 PSD blocks)
try:
    from psd_text_layer import build_editable_text_tagged_blocks, resolve_ps_font_name
    EDITABLE_TEXT_AVAILABLE = True
except ImportError:
    EDITABLE_TEXT_AVAILABLE = False'''
    if OLD_IMPORT2 in content:
        content = content.replace(OLD_IMPORT2, NEW_IMPORT2)
        print("Step 1 OK (fallback): import added after rembg")
    else:
        print("ERROR step 1: import location not found")

# ── Step 2: Patch write_psd to inject TySh+Txt2 into layer extra ─────────────
OLD_EXTRA = '''        name_bytes = _pack_layer_name(lyr['name'])
        extra = struct.pack('>I', 0) + struct.pack('>I', 0) + name_bytes
        lr.write(struct.pack('>I', len(extra)))
        lr.write(extra)'''

NEW_EXTRA = '''        name_bytes = _pack_layer_name(lyr['name'])
        # Inject editable text blocks (TySh + Txt2) if this is a text layer
        text_tagged = lyr.get('_text_blocks', b'')
        extra = struct.pack('>I', 0) + struct.pack('>I', 0) + name_bytes + text_tagged
        # Change layer type flag if editable text blocks present
        if text_tagged:
            # Re-write flags with text layer bit — go back and patch the flags byte
            # (Layer type is encoded in the layer record flags region)
            # For now the rasterised pixel data still provides the visual
            # Photoshop will show the editable text overlay on top
            pass
        lr.write(struct.pack('>I', len(extra)))
        lr.write(extra)'''

if OLD_EXTRA in content:
    content = content.replace(OLD_EXTRA, NEW_EXTRA)
    print("Step 2 OK: write_psd extra data patched")
else:
    print("ERROR step 2: extra data block not found")
    # Show context around the area
    idx = content.find('name_bytes = _pack_layer_name')
    if idx != -1:
        print("Found at:", idx)
        print(repr(content[idx:idx+300]))

# ── Step 3: Attach _text_blocks when building CustomerText layers ─────────────
# Find the CustomerText layer append calls and add _text_blocks
OLD_TEXT_LAYER_1 = '''            if txt_pil:
                all_layers.append({
                    "name":    f"{display_label} CustomerText{suffix}",
                    "image":   txt_pil,
                    "top":     copy_top + tt,
                    "left":    x_left + tl,
                    "opacity": 255,
                    "visible": True,
                })'''

NEW_TEXT_LAYER_1 = '''            if txt_pil:
                _txt_layer = {
                    "name":    f"{display_label} CustomerText{suffix}",
                    "image":   txt_pil,
                    "top":     copy_top + tt,
                    "left":    x_left + tl,
                    "opacity": 255,
                    "visible": True,
                }
                # Attach editable text blocks if available
                if EDITABLE_TEXT_AVAILABLE and zone.get("text_lines"):
                    _joined = "\\n".join(zone["text_lines"])
                    _ps_font = resolve_ps_font_name(zone.get("font","arial"))
                    _r,_g,_b = hex_to_rgb(zone.get("colour","#ffffff"))
                    _txt_layer["_text_blocks"] = build_editable_text_tagged_blocks(
                        text=_joined,
                        font_name=_ps_font,
                        font_size_px=txt_pil.height,
                        r=_r, g=_g, b=_b,
                        px_per_cm=PX_PER_CM,
                        layer_left=x_left + tl,
                        layer_top=copy_top + tt,
                        layer_w=txt_pil.width,
                        layer_h=txt_pil.height,
                    )
                all_layers.append(_txt_layer)'''

if OLD_TEXT_LAYER_1 in content:
    content = content.replace(OLD_TEXT_LAYER_1, NEW_TEXT_LAYER_1)
    print("Step 3 OK: CustomerText layer patched in build_psd_for_order")
else:
    print("WARN step 3: CustomerText pattern not found in build_psd_for_order")
    # Count occurrences to debug
    count = content.count('CustomerText{suffix}')
    print(f"  'CustomerText{{suffix}}' appears {count} times")

# ── Step 4: Same patch for merged PSD builder ─────────────────────────────────
OLD_TEXT_LAYER_2 = '''                # Text
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
                    })'''

NEW_TEXT_LAYER_2 = '''                # Text
                if zone["text_lines"]:
                    txt_pil, tt, tl = build_text_layer(
                        zone["text_lines"], zone["font"], zone["colour"], zw, zh,
                        has_image=bool(zone["img_path"]))
                    _txt_layer2 = {
                        "name": f"{display_label} CustomerText",
                        "image": txt_pil,
                        "top": img_start + tt,
                        "left": x_left + tl,
                        "opacity": 255, "visible": True,
                    }
                    if EDITABLE_TEXT_AVAILABLE:
                        _joined2 = "\\n".join(zone["text_lines"])
                        _ps_font2 = resolve_ps_font_name(zone.get("font","arial"))
                        _r2,_g2,_b2 = hex_to_rgb(zone.get("colour","#ffffff"))
                        _txt_layer2["_text_blocks"] = build_editable_text_tagged_blocks(
                            text=_joined2,
                            font_name=_ps_font2,
                            font_size_px=txt_pil.height,
                            r=_r2, g=_g2, b=_b2,
                            px_per_cm=PX_PER_CM,
                            layer_left=x_left + tl,
                            layer_top=img_start + tt,
                            layer_w=txt_pil.width,
                            layer_h=txt_pil.height,
                        )
                    all_layers.append(_txt_layer2)'''

if OLD_TEXT_LAYER_2 in content:
    content = content.replace(OLD_TEXT_LAYER_2, NEW_TEXT_LAYER_2)
    print("Step 4 OK: CustomerText layer patched in build_merged_psd_for_order_group")
else:
    print("WARN step 4: merged PSD CustomerText pattern not found")

# ── Save and verify syntax ────────────────────────────────────────────────────
open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)
print("Saved.")

import ast
try:
    ast.parse(content)
    print("SYNTAX OK")
except SyntaxError as e:
    print(f"SYNTAX ERROR: {e}")
