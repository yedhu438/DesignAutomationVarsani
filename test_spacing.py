"""Render 'LIKE' with Reflection Font and save collage + glyph crops for inspection."""
import os, sys, re, subprocess, base64
import numpy as np
from PIL import Image
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
from fontTools.ttLib import TTFont
import batch_processor as bp

OUT_DIR = r'C:\Varsany\Output\spacing_test'
os.makedirs(OUT_DIR, exist_ok=True)

font_path = r'C:\Varsany\Fonts\Refraction Ray.otf'
ft        = TTFont(font_path)
upem      = ft['head'].unitsPerEm
hmtx      = ft['hmtx'].metrics
cmap      = ft.getBestCmap()
svg_t     = ft['SVG '].docList

svg_map = {}
for doc_entry in svg_t:
    raw, s, e = doc_entry
    txt = raw.decode('utf-8') if isinstance(raw, (bytes, bytearray)) else raw
    shared = (s != e)
    for gid in range(s, e + 1):
        svg_map[gid] = (txt, shared)
glyph_order = ft.getGlyphOrder()

CHROME_EXE  = bp.CHROME_EXE
TEMP_FOLDER = bp.TEMP_FOLDER

# Replicate production collage exactly
test_chars = 'LIKE'
canvas_w   = 3600
glyph_h    = max(800, min(2400, int(canvas_w * 0.85 / len(test_chars)) * 4))
scale      = glyph_h / upem

print(f'canvas_w={canvas_w}, glyph_h={glyph_h}, scale={scale:.3f}')

char_svg   = {}
items_html = ''
x_pos      = {}
cx         = 0

for ch in test_chars:
    cp    = ord(ch)
    gname = cmap.get(cp)
    if not gname: continue
    gid   = glyph_order.index(gname)
    entry = svg_map.get(gid)
    if not entry: continue
    svg_raw, shared = entry
    adv_u  = hmtx.get(gname, (upem // 3, 0))[0]
    adv_px = max(1, int(adv_u * scale))
    h_px   = glyph_h

    svg = svg_raw
    vb_attr = f'viewBox="0 -850 {upem} 1000"'
    svg = re.sub(r'viewBox="[^"]*"', vb_attr, svg) if 'viewBox' in svg else svg.replace('<svg', f'<svg {vb_attr}', 1)
    svg = re.sub(r'\s+preserveAspectRatio="[^"]*"', '', svg)
    svg = re.sub(r'\s+width="[^"]*"',  '', svg)
    svg = re.sub(r'\s+height="[^"]*"', '', svg)
    svg = svg.replace('<svg', f'<svg width="{h_px}px" height="{h_px}px"', 1)
    if shared:
        hide = (f'<style>g{{display:none}}'
                f'#glyph{gid}{{display:inline}}'
                f'#glyph\\.{gid}{{display:inline}}</style>')
        svg = re.sub(r'(<svg[^>]*>)', r'\g<1>' + hide, svg, count=1)

    char_svg[ch] = (adv_px, h_px, svg)
    svg_b64 = base64.b64encode(svg.encode('utf-8')).decode('ascii')
    items_html += (
        f'<img style="position:absolute;left:{cx}px;top:0;'
        f'width:{h_px}px;height:{h_px}px;display:block" '
        f'src="data:image/svg+xml;base64,{svg_b64}">\n'
    )
    x_pos[ch] = cx
    cx += h_px  # same as production: advance by h_px, not adv_px

collage_w = cx + 10
collage_h = glyph_h + 10
html_src = (
    f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    f'*{{margin:0;padding:0}}html,body{{background:#ffffff;overflow:hidden}}'
    f'.c{{position:relative;width:{collage_w}px;height:{collage_h}px}}'
    f'</style></head><body>\n'
    f'<div class="c">{items_html}</div></body></html>'
)

html_path = os.path.join(TEMP_FOLDER, 'test_collage.html')
png_path  = os.path.join(TEMP_FOLDER, 'test_collage.png')
with open(html_path, 'w', encoding='utf-8') as fh:
    fh.write(html_src)

html_url = 'file:///' + html_path.replace('\\', '/')
cmd = [
    CHROME_EXE, '--headless', '--no-sandbox', '--disable-gpu',
    '--disable-extensions', '--no-first-run', '--disable-sync',
    f'--screenshot={png_path}',
    f'--window-size={collage_w},{collage_h}',
    html_url,
]
subprocess.run(cmd, capture_output=True, timeout=30)

if not os.path.exists(png_path):
    print('Chrome failed')
    sys.exit(1)

# Save the raw collage
collage_img = Image.open(png_path).convert('RGBA')
collage_img.save(os.path.join(OUT_DIR, 'collage_raw.png'))
print(f'Collage saved: {collage_img.size}')

# Now crop each glyph exactly as production does
arr = np.array(collage_img)
white = (arr[:, :, 0] > 240) & (arr[:, :, 1] > 240) & (arr[:, :, 2] > 240)
arr[white, 3] = 0
collage_clean = Image.fromarray(arr)

print(f'\n{"Ch":3} {"adv_px":7} {"xpos_col":9} {"lsb_px":7} {"rsb_px":7} {"art_w":7}')
print('-' * 50)
for ch, (adv_px, h_px, _) in char_svg.items():
    x0 = x_pos[ch]
    x1 = x_pos[ch] + adv_px
    crop = arr[:, x0:x1, :]
    col_alpha = crop[:, :, 3].max(axis=0)
    vis = list(np.where(col_alpha > 0)[0])
    if not vis:
        print(f'{ch}: NO VISIBLE')
        continue
    lsb = vis[0]
    rsb = adv_px - vis[-1] - 1
    art_w = vis[-1] - vis[0] + 1
    print(f'{ch:3} {adv_px:7} {x0:9} {lsb:7} {rsb:7} {art_w:7}')
    # Save individual crop
    crop_img = Image.fromarray(crop)
    crop_img.save(os.path.join(OUT_DIR, f'crop_{ch}.png'))

# Now compose "LIKE" at tracking=0.92 and tracking=1.0 and save both
for tracking_val in [1.0, 0.92, 0.90, 0.85]:
    line_h = glyph_h
    parts = []
    x = 0
    for ch, (adv_px, h_px, _) in char_svg.items():
        x0c = x_pos[ch]
        x1c = x_pos[ch] + adv_px
        crop = Image.fromarray(arr[:, x0c:x1c, :])
        # vertical tight crop
        ca = np.array(crop)
        row_alpha = ca[:, :, 3].max(axis=1)
        vis_rows = np.where(row_alpha > 0)[0]
        if len(vis_rows) > 0:
            yt, yb = int(vis_rows[0]), int(vis_rows[-1]) + 1
            crop = crop.crop((0, yt, crop.width, yb))
            above = yb - yt
        else:
            above = 0
        adv = max(1, int(adv_px * tracking_val))
        parts.append((crop, x, above))
        x += adv

    max_above = max(p[2] for p in parts if p[2] > 0)
    line_w = x
    line_img = Image.new('RGBA', (line_w, max_above), (0, 0, 0, 0))
    for gimg, gx, above in parts:
        if gimg is not None:
            gy = max_above - above
            line_img.paste(gimg, (gx, max(0, gy)), gimg)

    out = line_img.crop(line_img.getbbox())
    out.save(os.path.join(OUT_DIR, f'LIKE_tracking_{str(tracking_val).replace(".", "_")}.png'))
    print(f'tracking={tracking_val}: output width={out.width}px')

print(f'\nAll files saved to {OUT_DIR}')
