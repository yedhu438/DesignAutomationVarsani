
content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()

# ── Patch 1: build_psd_for_order CustomerText layer (line ~1897) ─────────────
OLD1 = '''            if txt_pil:
                all_layers.append({
                    "name":    f"{display_label} CustomerText{suffix}",
                    "image":   txt_pil,
                    "top":     copy_top + v_off + tt,
                    "left":    x_left + tl,
                    "opacity": 255,
                    "visible": True,
                })'''

NEW1 = '''            if txt_pil:
                _tl_dict = {
                    "name":    f"{display_label} CustomerText{suffix}",
                    "image":   txt_pil,
                    "top":     copy_top + v_off + tt,
                    "left":    x_left + tl,
                    "opacity": 255,
                    "visible": True,
                }
                if EDITABLE_TEXT_AVAILABLE and zone.get("text_lines"):
                    _ps_font = resolve_ps_font_name(zone.get("font", "arial"))
                    _r, _g, _b = hex_to_rgb(zone.get("colour", "#ffffff"))
                    _tl_dict["_text_blocks"] = build_editable_text_tagged_blocks(
                        text="\\n".join(zone["text_lines"]),
                        font_name=_ps_font,
                        font_size_px=txt_pil.height,
                        r=_r, g=_g, b=_b,
                        px_per_cm=PX_PER_CM,
                        layer_left=x_left + tl,
                        layer_top=copy_top + v_off + tt,
                        layer_w=txt_pil.width,
                        layer_h=txt_pil.height,
                    )
                all_layers.append(_tl_dict)'''

if OLD1 in content:
    content = content.replace(OLD1, NEW1)
    print("Patch 1 OK: build_psd_for_order CustomerText")
else:
    print("ERROR patch 1")

# ── Patch 2: build_merged_psd_for_order_group CustomerText layer (line ~2073) ─
OLD2 = '''            if zone["_txt"]:
                all_layers.append({
                    "name": f"{display_label} CustomerText",
                    "image": zone["_txt"],
                    "top": img_start + v_off + zone["_tt"],
                    "left": x_left + zone["_tl"],
                    "opacity": 255, "visible": True,
                })'''

NEW2 = '''            if zone["_txt"]:
                _tl_dict2 = {
                    "name": f"{display_label} CustomerText",
                    "image": zone["_txt"],
                    "top": img_start + v_off + zone["_tt"],
                    "left": x_left + zone["_tl"],
                    "opacity": 255, "visible": True,
                }
                if EDITABLE_TEXT_AVAILABLE and zone.get("text_lines"):
                    _ps_font2 = resolve_ps_font_name(zone.get("font", "arial"))
                    _r2, _g2, _b2 = hex_to_rgb(zone.get("colour", "#ffffff"))
                    _tl_dict2["_text_blocks"] = build_editable_text_tagged_blocks(
                        text="\\n".join(zone["text_lines"]),
                        font_name=_ps_font2,
                        font_size_px=zone["_txt"].height,
                        r=_r2, g=_g2, b=_b2,
                        px_per_cm=PX_PER_CM,
                        layer_left=x_left + zone["_tl"],
                        layer_top=img_start + v_off + zone["_tt"],
                        layer_w=zone["_txt"].width,
                        layer_h=zone["_txt"].height,
                    )
                all_layers.append(_tl_dict2)'''

if OLD2 in content:
    content = content.replace(OLD2, NEW2)
    print("Patch 2 OK: build_merged_psd_for_order_group CustomerText")
else:
    print("ERROR patch 2")

# Save and verify
open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)

import ast
try:
    ast.parse(content)
    print("SYNTAX OK")
except SyntaxError as e:
    print(f"SYNTAX ERROR: {e}")
