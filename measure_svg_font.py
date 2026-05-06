"""Measure actual Chrome pixel fill for any SVG font A-Z using production collage method."""
import os, sys, re, subprocess, base64, argparse
import numpy as np
from PIL import Image
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))
from fontTools.ttLib import TTFont
import batch_processor as bp

def measure(font_path):
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

    test_chars  = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
    glyph_h     = 800
    scale       = glyph_h / upem
    vb_top      = -850

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
        vb_attr = f'viewBox="0 {vb_top} {upem} 1000"'
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

        char_svg[ch] = adv_px
        svg_b64 = base64.b64encode(svg.encode('utf-8')).decode('ascii')
        items_html += (
            f'<img style="position:absolute;left:{cx}px;top:0;'
            f'width:{h_px}px;height:{h_px}px;display:block" '
            f'src="data:image/svg+xml;base64,{svg_b64}">\n'
        )
        x_pos[ch] = cx
        cx += h_px

    collage_w = cx + 10
    html_src  = (
        f'<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
        f'*{{margin:0;padding:0}}html,body{{background:#ffffff;overflow:hidden}}'
        f'.c{{position:relative;width:{collage_w}px;height:{glyph_h+10}px}}'
        f'</style></head><body><div class="c">{items_html}</div></body></html>'
    )
    html_path = os.path.join(bp.TEMP_FOLDER, 'svg_measure.html')
    png_path  = os.path.join(bp.TEMP_FOLDER, 'svg_measure.png')
    with open(html_path, 'w', encoding='utf-8') as fh:
        fh.write(html_src)

    html_url = 'file:///' + html_path.replace('\\', '/')
    cmd = [
        bp.CHROME_EXE, '--headless', '--no-sandbox', '--disable-gpu',
        '--disable-extensions', '--no-first-run', '--disable-sync',
        f'--screenshot={png_path}',
        f'--window-size={collage_w},{glyph_h+10}',
        html_url,
    ]
    subprocess.run(cmd, capture_output=True, timeout=30)
    if not os.path.exists(png_path):
        print('Chrome failed'); return

    img = Image.open(png_path).convert('RGBA')
    arr = np.array(img)
    white = (arr[:, :, 0] > 240) & (arr[:, :, 1] > 240) & (arr[:, :, 2] > 240)
    arr[white, 3] = 0

    baseline_y_px = int(abs(vb_top) / 1000 * glyph_h)  # = 680 at h=800

    print(f"  {'Ch':3} {'adv':5} {'lsb':5} {'rsb':5} {'art_w':6} {'fill%':6}  below_bl")
    print('  ' + '-'*55)
    fills       = []
    lsb_dict    = {}
    art_end_dict= {}
    adv_dict    = {}

    for ch in test_chars:
        if ch not in x_pos: continue
        adv_px = char_svg[ch]
        xp     = x_pos[ch]
        crop   = arr[:, xp:xp + adv_px, :]
        col_alpha = crop[:, :, 3].max(axis=0)
        row_alpha = crop[:, :, 3].max(axis=1)
        vis_cols  = list(np.where(col_alpha > 0)[0])
        vis_rows  = list(np.where(row_alpha > 0)[0])
        if not vis_cols or not vis_rows:
            continue
        lsb   = vis_cols[0]
        rsb   = adv_px - vis_cols[-1] - 1
        art_w = vis_cols[-1] - vis_cols[0] + 1
        fill  = art_w / adv_px
        # descender check: does the glyph extend below baseline?
        yb    = vis_rows[-1]
        below_bl = max(0, yb - baseline_y_px)
        fills.append(fill)
        lsb_dict[ch]     = lsb
        art_end_dict[ch] = lsb + art_w
        adv_dict[ch]     = adv_px
        flag = ' <<<' if below_bl > 5 else ''
        print(f'  {ch:3} {adv_px:5} {lsb:5} {rsb:5} {art_w:6} {fill*100:5.1f}%  {below_bl:4}px{flag}')

    if fills:
        avg = sum(fills) / len(fills)
        max_touch = max(
            (art_end_dict[ch1] - lsb_dict.get(test_chars[i+1], 0)) / adv_dict[ch1]
            for i, ch1 in enumerate(test_chars[:-1])
            if ch1 in art_end_dict and test_chars[i+1] in lsb_dict
        ) if len(fills) > 1 else avg
        print(f'\n  Mean pixel fill:          {avg:.4f}')
        print(f'  Touch tracking (avg gap=0): {avg:.3f}')
        print(f'  Min safe tracking:          {max_touch:.3f}')
        print(f'  Recommended (tiny gap):     {max(avg, max_touch) + 0.01:.3f}')


parser = argparse.ArgumentParser()
parser.add_argument('font_key', nargs='?', default=None)
args = parser.parse_args()

import batch_processor as bp2
keys = [args.font_key] if args.font_key else ['smartkids', 'cozywinter']
for key in keys:
    path = bp2.FONT_INDEX.get(key)
    if not path:
        print(f'{key}: not found'); continue
    print(f'\n=== {key} ({os.path.basename(path)}) ===')
    measure(path)
