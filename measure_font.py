"""Measure actual Chrome-rendered pixel extents for Refraction Ray A-Z."""
import os, sys, re, subprocess
import numpy as np
from PIL import Image
sys.path.insert(0, os.path.dirname(__file__))
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
from fontTools.ttLib import TTFont
import batch_processor as bp

font_path = r'C:\Varsany\Fonts\Refraction Ray.otf'
ft        = TTFont(font_path)
upem      = ft['head'].unitsPerEm
hmtx      = ft['hmtx'].metrics
cmap      = ft.getBestCmap()
svg_t     = ft['SVG '].docList

svg_map = {}
for doc_entry in svg_t:
    svg_raw_bytes, start_id, end_id = doc_entry
    svg_raw = (svg_raw_bytes.decode('utf-8')
               if isinstance(svg_raw_bytes, (bytes, bytearray))
               else svg_raw_bytes)
    shared = (start_id != end_id)
    for gid in range(start_id, end_id + 1):
        svg_map[gid] = (svg_raw, shared)
glyph_order = ft.getGlyphOrder()

test_chars      = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
glyph_h         = 800
scale           = glyph_h / upem
CHROME_EXE      = bp.CHROME_EXE
TEMP_FOLDER     = bp.TEMP_FOLDER

svgs_html       = []
x_positions     = {}
expected_widths = {}
x_cursor        = 0

for ch in test_chars:
    cp    = ord(ch)
    gname = cmap.get(cp)
    if not gname:
        continue
    gid   = glyph_order.index(gname)
    entry = svg_map.get(gid)
    if not entry:
        continue
    svg_raw, shared = entry
    adv_u  = hmtx.get(gname, (upem // 3, 0))[0]
    adv_px = max(1, int(adv_u * scale))
    h_px   = glyph_h

    svg = svg_raw
    vb_attr = f'viewBox="0 -850 {upem} 1000"'
    svg = (re.sub(r'viewBox="[^"]*"', vb_attr, svg)
           if 'viewBox' in svg
           else svg.replace('<svg', f'<svg {vb_attr}', 1))
    svg = re.sub(r'\s+preserveAspectRatio="[^"]*"', '', svg)
    svg = re.sub(r'\s+width="[^"]*"',  '', svg)
    svg = re.sub(r'\s+height="[^"]*"', '', svg)
    svg = svg.replace('<svg', f'<svg width="{h_px}px" height="{h_px}px"', 1)
    if shared:
        hide = (f'<style>g{{display:none}}'
                f'#glyph{gid}{{display:inline}}'
                f'#glyph\\.{gid}{{display:inline}}</style>')
        svg = re.sub(r'(<svg[^>]*>)', r'\g<1>' + hide, svg, count=1)

    x_positions[ch]     = x_cursor
    expected_widths[ch] = adv_px
    svgs_html.append(
        f'<div style="display:inline-block;width:{adv_px}px;height:{h_px}px;'
        f'overflow:hidden;position:relative">{svg}</div>'
    )
    x_cursor += adv_px

total_w = x_cursor
html = (
    '<!DOCTYPE html><html><head><meta charset="utf-8">'
    f'<style>*{{margin:0;padding:0}} body{{background:white;width:{total_w}px;white-space:nowrap}}</style>'
    f'</head><body>{"".join(svgs_html)}</body></html>'
)

_pid      = os.getpid()
html_path = os.path.join(TEMP_FOLDER, f'measure_{_pid}.html')
png_path  = os.path.join(TEMP_FOLDER, f'measure_{_pid}.png')
with open(html_path, 'w', encoding='utf-8') as fh:
    fh.write(html)

html_url = 'file:///' + html_path.replace('\\', '/')
cmd = [
    CHROME_EXE, '--headless', '--no-sandbox', '--disable-gpu',
    '--disable-extensions', '--no-first-run', '--disable-sync',
    f'--screenshot={png_path}',
    f'--window-size={total_w},{glyph_h}',
    html_url,
]
subprocess.run(cmd, capture_output=True, timeout=30)

if not os.path.exists(png_path):
    print('Chrome render failed')
    sys.exit(1)

img = Image.open(png_path).convert('RGBA')
arr = np.array(img)
white = (arr[:, :, 0] > 240) & (arr[:, :, 1] > 240) & (arr[:, :, 2] > 240)
arr[white, 3] = 0

print(f"{'Ch':3} {'adv_px':7} {'LSB_px':7} {'RSB_px':7} {'art_w_px':9} {'fill%':7}")
print('-' * 50)
ratios = []
for ch in test_chars:
    if ch not in x_positions:
        continue
    x0 = x_positions[ch]
    x1 = x0 + expected_widths[ch]
    ca = arr[:, x0:x1, :]
    col_alpha = ca[:, :, 3].max(axis=0)
    vis_cols  = list(np.where(col_alpha > 0)[0])
    if not vis_cols:
        print(f'{ch}: NO VISIBLE PIXELS')
        continue
    left  = vis_cols[0]
    right = vis_cols[-1] + 1
    art_w = right - left
    adv   = expected_widths[ch]
    lsb   = left
    rsb   = adv - right
    fill  = art_w / adv
    ratios.append(fill)
    print(f'{ch:3} {adv:7} {lsb:7} {rsb:7} {art_w:9} {fill*100:6.1f}%')

if ratios:
    avg = sum(ratios) / len(ratios)
    print()
    print(f'Mean pixel fill:    {avg:.4f}')
    print(f'Tight tracking (letters touch): {avg:.3f}')
    print(f'Slight gap (5% art width):      {round(avg + 0.05*(1-avg), 3):.3f}')
