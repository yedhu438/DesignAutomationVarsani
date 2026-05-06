"""
test_svg_render.py
------------------
Standalone test for the Chrome SVG glyph renderer used in batch_processor.py.
Renders "TECH C.E.O." with Refraction Ray and saves a PNG showing:
  - the composed text
  - a RED horizontal line at the computed baseline
  - a BLUE horizontal line at the bottom of each glyph's bounding box

Output: C:/Varsany/Output/test_svg_render.png
Run:
    python test_svg_render.py
"""

import sys
sys.stdout.reconfigure(encoding="utf-8")
import base64 as _b64
import os
import re
import subprocess

import numpy as np
from PIL import Image, ImageDraw

# -- CONFIG --------------------------------------------------------------------
FONT_PATH  = r"C:\Varsany\Fonts\Refraction Ray.otf"
TEST_TEXT  = "TECH C.E.O."
TRACKING   = 0.92
CANVAS_W   = 2346   # 85% of 2760px KidsTee front

OUTPUT_DIR = r"C:\Varsany\Output"
OUTPUT_PNG = os.path.join(OUTPUT_DIR, "test_svg_render.png")
TEMP_DIR   = r"C:\Varsany\Temp"

CHROME_EXE = None
for _p in [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]:
    if os.path.exists(_p):
        CHROME_EXE = _p
        break

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMP_DIR,   exist_ok=True)

# -- LOAD FONT ----------------------------------------------------------------─
from fontTools.ttLib import TTFont

ft   = TTFont(FONT_PATH)
upem = ft["head"].unitsPerEm
cmap = ft.getBestCmap()
glyph_order = ft.getGlyphOrder()
hmtx        = ft["hmtx"].metrics
svg_table   = ft.get("SVG ")

assert svg_table, "No SVG table found in font."

svg_map = {}
for doc in svg_table.docList:
    shared = (doc.startGlyphID != doc.endGlyphID)
    for gid in range(doc.startGlyphID, doc.endGlyphID + 1):
        svg_map[gid] = (doc.data, shared)

print(f"Font: {FONT_PATH}")
print(f"upem={upem}  glyphs with SVG={len(svg_map)}")

# -- GLYPH METRICS (for reference) --------------------------------------------─
print("\n-- Glyph metrics (font units, y-up) --")
for ch in sorted(set(TEST_TEXT)):
    if not ch.strip():
        continue
    gn  = cmap.get(ord(ch))
    gid = glyph_order.index(gn) if gn else None
    if gid is None:
        print(f"  '{ch}'  → NOT IN CMAP")
        continue
    adv_u = hmtx.get(gn, (upem // 3, 0))[0]
    entry = svg_map.get(gid)
    if entry:
        svg_raw, _ = entry
        # Find the first <image> element in the SVG to read y/height
        m = re.search(r'<image\s[^>]*>', svg_raw[:800])
        coords = dict(re.findall(r'(\w+)="(-?[\d.]+)"', m.group(0))) if m else {}
        print(f"  '{ch}'  glyph={gn:<20} adv={adv_u:>4}u  "
              f"img_x={coords.get('x','?')} img_y={coords.get('y','?')} "
              f"w={coords.get('width','?')} h={coords.get('height','?')}")
    else:
        print(f"  '{ch}'  glyph={gn}  (no SVG entry)  adv={adv_u}u")

# -- RENDER --------------------------------------------------------------------
print("\n-- Rendering --")

# Same sizing logic as batch_processor
max_chars = max(len(l) for l in TEST_TEXT.split("\n") if l)
target_w  = int(CANVAS_W / max(1, max_chars))
glyph_h   = max(800, min(2400, target_w * 4))

# Square viewBox: 0 -850 1000 1000
vb_top      = -850
vb_h_render = 1000   # == upem → square viewBox
scale       = glyph_h / upem   # glyph_h / 1000

print(f"max_chars={max_chars}  target_w={target_w}  glyph_h={glyph_h}")
print(f"vb_top={vb_top}  vb_h_render={vb_h_render}  scale={scale:.4f}")

# Compute baseline pixel position in each square glyph element
_vb_top_abs   = 850          # = -vb_top
baseline_y_px = int(_vb_top_abs / vb_h_render * glyph_h)   # 0.85 * glyph_h
sf            = glyph_h / upem
print(f"baseline_y_px={baseline_y_px}  sf={sf:.4f}")

# Prepare SVG for each unique character
all_chars = sorted(set(ch for ch in TEST_TEXT if ch.strip()))
char_svg  = {}

for ch in all_chars:
    gn = cmap.get(ord(ch))
    if not gn:
        continue
    try:
        gid = glyph_order.index(gn)
    except ValueError:
        continue
    entry = svg_map.get(gid)
    if not entry:
        continue
    svg_raw, shared = entry
    adv_units = hmtx.get(gn, (upem // 3, 0))[0]
    adv_px    = max(1, int(adv_units * scale))    # advance width for crop/composition
    h_px      = glyph_h

    svg = svg_raw
    vb_attr = f'viewBox="0 {vb_top} {upem} {vb_h_render}"'
    if 'viewBox' in svg:
        svg = re.sub(r'viewBox="[^"]*"', vb_attr, svg)
    else:
        svg = svg.replace('<svg', f'<svg {vb_attr}', 1)

    # Square element: h_px × h_px (== glyph_h × glyph_h)
    svg = re.sub(r'\s+preserveAspectRatio="[^"]*"', '', svg)
    svg = re.sub(r'\s+width="[^"]*"',  '', svg)
    svg = re.sub(r'\s+height="[^"]*"', '', svg)
    svg = svg.replace('<svg', f'<svg width="{h_px}px" height="{h_px}px"', 1)

    if shared:
        hide = (f'<style>g{{display:none}}'
                f'#glyph{gid}{{display:inline}}'
                f'#glyph\\.{gid}{{display:inline}}</style>')
        svg = re.sub(r'(<svg[^>]*>)', r'\1' + hide, svg, count=1)

    char_svg[ch] = (adv_px, h_px, svg)

# Build collage HTML — each element is square h_px × h_px
collage_w = sum(v[1] for v in char_svg.values()) + 10
collage_h = glyph_h + 10
items_html = ""
x_pos = {}
cx = 0
for ch, (adv_px, h_px, svg_str) in char_svg.items():
    b64 = _b64.b64encode(svg_str.encode()).decode()
    items_html += (
        f'<img style="position:absolute;left:{cx}px;top:0;'
        f'width:{h_px}px;height:{h_px}px;display:block" '
        f'src="data:image/svg+xml;base64,{b64}">\n'
    )
    x_pos[ch] = cx
    cx += h_px   # square element width

html_src = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
*{{margin:0;padding:0}}html,body{{background:#ffffff;overflow:hidden}}
.c{{position:relative;width:{collage_w}px;height:{collage_h}px}}
</style></head><body>
<div class="c">{items_html}</div>
</body></html>"""

html_path = os.path.join(TEMP_DIR, "test_collage.html")
png_path  = os.path.join(TEMP_DIR, "test_collage.png")

with open(html_path, "w", encoding="utf-8") as f:
    f.write(html_src)
if os.path.exists(png_path):
    os.remove(png_path)

cmd = [
    CHROME_EXE, "--headless", "--no-sandbox",
    "--disable-gpu", "--disable-extensions",
    "--no-first-run", "--disable-sync",
    f"--screenshot={png_path}",
    f"--window-size={collage_w},{collage_h}",
    "file:///" + html_path.replace("\\", "/"),
]
print(f"\nRunning Chrome: window={collage_w}×{collage_h}")
subprocess.run(cmd, capture_output=True, timeout=40)

collage_img = Image.open(png_path).convert("RGBA")
dpr = collage_img.width / collage_w if collage_w > 0 else 1.0
print(f"Collage PNG: {collage_img.size}  dpr={dpr:.2f}")

# -- CROP GLYPHS --------------------------------------------------------------─
print("\n-- Per-glyph crop results --")
glyph_imgs_raw = {}
glyph_above_bl = {}
glyph_below_bl = {}

for ch, (adv_px, h_px, _) in char_svg.items():
    x0   = int(x_pos[ch] * dpr)
    x1   = int((x_pos[ch] + adv_px) * dpr)   # advance-width crop
    y1   = int(h_px * dpr)
    crop = collage_img.crop((x0, 0, min(x1, collage_img.width),
                             min(y1, collage_img.height)))
    if dpr != 1.0:
        crop = crop.resize((adv_px, h_px), Image.LANCZOS)
    arr   = np.array(crop)
    white = (arr[:, :, 0] > 240) & (arr[:, :, 1] > 240) & (arr[:, :, 2] > 240)
    arr[white, 3] = 0
    img_rgba = Image.fromarray(arr)
    bbox = img_rgba.getbbox()
    if bbox and arr[:, :, 3].max() > 0:
        glyph_imgs_raw[ch] = img_rgba.crop(bbox)
        glyph_above_bl[ch] = bbox[3] - bbox[1]   # crop height → bottom-align
        glyph_below_bl[ch] = 0
    else:
        glyph_imgs_raw[ch] = None
        glyph_above_bl[ch] = 0
        glyph_below_bl[ch] = 0
    print(f"  '{ch}'  adv_px={adv_px:>4}  "
          f"bbox={bbox}  "
          f"crop_size={str(glyph_imgs_raw[ch].size) if glyph_imgs_raw[ch] else 'None':>14}  "
          f"above_bl={glyph_above_bl[ch]:>4}  below_bl={glyph_below_bl[ch]:>3}")

# -- LINE METRICS --------------------------------------------------------------
max_above = max((v for v in glyph_above_bl.values() if v > 0), default=int(glyph_h * 0.625))
max_below = max((v for v in glyph_below_bl.values() if v > 0), default=0)
line_h    = max_above + max_below

print(f"\nmax_above={max_above}  max_below={max_below}  line_h={line_h}")

# -- COMPOSE LINE --------------------------------------------------------------
space_em = hmtx.get('space', hmtx.get('uni0020', (upem // 3, 0)))[0]
space_w  = max(4, int(space_em * sf * TRACKING))
x = 0
parts = []
for ch in TEST_TEXT:
    gimg = glyph_imgs_raw.get(ch)
    if ch in char_svg:
        gn    = cmap.get(ord(ch))
        adv_u = hmtx.get(gn, (upem // 3, 0))[0] if gn else upem // 3
        adv   = max(1, int(adv_u * sf * TRACKING))
        parts.append((gimg, x, glyph_above_bl.get(ch, max_above)))
        x += adv
    else:
        x += space_w

line_img = Image.new("RGBA", (x, line_h), (0, 0, 0, 0))
for gimg, gx, above in parts:
    if gimg is not None:
        gy = max_above - above
        line_img.paste(gimg, (gx, max(0, gy)), gimg)

# Trim horizontal
bbox2 = line_img.getbbox()
if bbox2:
    line_img = line_img.crop(bbox2)

# Scale to canvas width
ratio   = CANVAS_W / line_img.width
result  = line_img.resize((CANVAS_W, max(1, int(line_img.height * ratio))), Image.LANCZOS)

# -- DRAW BASELINE MARKER ------------------------------------------------------
# Red line at computed baseline position; blue line at actual letter bottom.
out = result.copy().convert("RGBA")
draw = ImageDraw.Draw(out)

# Baseline position in the result image
# In line_img: baseline is at y = max_above (letters hang from top, period sits at max_above)
# After crop (bbox2 trims top blank space), shift down by bbox2[1]
baseline_in_result = int((max_above - (bbox2[1] if bbox2 else 0)) * ratio)
letter_bottom      = int((line_h - (bbox2[1] if bbox2 else 0)) * ratio)

draw.line([(0, baseline_in_result), (out.width, baseline_in_result)],
          fill=(255, 0, 0, 200), width=3)   # RED = baseline (bottom of caps)
draw.line([(0, letter_bottom),      (out.width, letter_bottom)],
          fill=(0, 80, 255, 180), width=2)  # BLUE = max_below bottom

# White background for visibility
final = Image.new("RGBA", out.size, (200, 200, 200, 255))
final.paste(out, (0, 0), out)

final.save(OUTPUT_PNG)
print(f"\nSaved → {OUTPUT_PNG}  ({final.size})")
print(f"Baseline in result image: y={baseline_in_result}  (red line)")
print(f"Letter bottom:            y={letter_bottom}  (blue line)")
