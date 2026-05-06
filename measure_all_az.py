"""Measure real pixel fill for all A-Z glyphs using production collage method."""
import os, sys, re, subprocess, base64
import numpy as np
from PIL import Image
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
from fontTools.ttLib import TTFont
import batch_processor as bp

FONT_PATH = r'C:\Varsany\Fonts\Refraction Ray.otf'
ft        = TTFont(FONT_PATH)
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

test_chars  = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
glyph_h     = 800   # small for speed
scale       = glyph_h / upem

char_svg    = {}
items_html  = ''
x_pos       = {}
cx          = 0

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
    svg = re.sub(r'\s+width="[^"]*"', '', svg)
    svg = re.sub(r'\s+height="[^"]*"', '', svg)
    svg = svg.replace('<svg', f'<svg width="{h_px}px" height="{h_px}px"', 1)
    if shared:
        hide = (f'<style>g{{display:none}}'
                f'#glyph{gid}{{display:inline}}'
                f'#glyph\\.{gid}{{display:inline}}</style>')
        svg = re.sub(r'(<svg[^>]*>)', r'\g<1>' + hide, svg, count=1)

    char_svg[ch] = (adv_px, h_px)
    svg_b64 = base64.b64encode(svg.encode('utf-8')).decode('ascii')
    items_html += (
        f'<img style="position:absolute;left:{cx}px;top:0;'
        f'width:{h_px}px;height:{h_px}px;display:block" '
        f'src="data:image/svg+xml;base64,{svg_b64}">\n'
    )
    x_pos[ch] = cx
    cx += h_px  # production: advance by h_px

collage_w = cx + 10
html_src  = (
    f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    f'*{{margin:0;padding:0}}html,body{{background:#ffffff;overflow:hidden}}'
    f'.c{{position:relative;width:{collage_w}px;height:{glyph_h + 10}px}}'
    f'</style></head><body><div class="c">{items_html}</div></body></html>'
)
html_path = os.path.join(TEMP_FOLDER, 'az_collage.html')
png_path  = os.path.join(TEMP_FOLDER, 'az_collage.png')
with open(html_path, 'w', encoding='utf-8') as fh:
    fh.write(html_src)

html_url = 'file:///' + html_path.replace('\\', '/')
cmd = [
    CHROME_EXE, '--headless', '--no-sandbox', '--disable-gpu',
    '--disable-extensions', '--no-first-run', '--disable-sync',
    f'--screenshot={png_path}',
    f'--window-size={collage_w},{glyph_h + 10}',
    html_url,
]
subprocess.run(cmd, capture_output=True, timeout=30)
if not os.path.exists(png_path):
    print('Chrome failed'); sys.exit(1)

img = Image.open(png_path).convert('RGBA')
arr = np.array(img)
white = (arr[:, :, 0] > 240) & (arr[:, :, 1] > 240) & (arr[:, :, 2] > 240)
arr[white, 3] = 0

print(f"{'Ch':3} {'adv':7} {'lsb':6} {'rsb':6} {'art_w':7} {'fill%':7}")
print('-' * 45)
fills       = []
lsb_dict    = {}
art_end_dict= {}
adv_dict    = {}

for ch in test_chars:
    if ch not in x_pos: continue
    adv_px, h_px = char_svg[ch]
    xp   = x_pos[ch]
    crop = arr[:, xp:xp + adv_px, :]
    col_alpha = crop[:, :, 3].max(axis=0)
    vis = list(np.where(col_alpha > 0)[0])
    if not vis:
        print(f'{ch}: no visible pixels')
        continue
    lsb   = vis[0]
    rsb   = adv_px - vis[-1] - 1
    art_w = vis[-1] - vis[0] + 1
    fill  = art_w / adv_px
    fills.append(fill)
    lsb_dict[ch]     = lsb
    art_end_dict[ch] = lsb + art_w
    adv_dict[ch]     = adv_px
    print(f'{ch:3} {adv_px:7} {lsb:6} {rsb:6} {art_w:7} {fill*100:6.1f}%')

print()
avg_fill = sum(fills) / len(fills) if fills else 0
print(f'Mean pixel fill: {avg_fill:.4f}')
print(f'Touch tracking:  {avg_fill:.3f}')

# Compute min safe tracking for every adjacent pair in alphabet
max_touch = 0.0
for i, ch1 in enumerate(test_chars[:-1]):
    ch2 = test_chars[i + 1]
    if ch1 not in art_end_dict or ch2 not in lsb_dict: continue
    # Safe T: art_end(ch1) <= adv(ch1)*T + lsb(ch2)
    t_safe = (art_end_dict[ch1] - lsb_dict[ch2]) / adv_dict[ch1]
    if t_safe > max_touch:
        max_touch = t_safe
        worst_pair = (ch1, ch2)

print(f'Min safe tracking (no overlap): {max_touch:.3f}  (pair: {worst_pair[0]}-{worst_pair[1]})')
print(f'Recommended tracking (+5px gap): {max_touch + 5/max(v for v in adv_dict.values()):.3f}')
