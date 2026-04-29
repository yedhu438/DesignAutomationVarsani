# Quick test: render sample text with each premium font and save PNG results.
import os, sys, subprocess, base64, re
from PIL import Image
import numpy as np

FONT_FOLDER = r"C:\Varsany\Fonts"
OUT_FOLDER  = r"C:\Varsany\Output\premium_font_test"
TEMP_FOLDER = r"C:\Varsany\Temp"
CANVAS_W    = 1200   # small canvas for fast testing
TEST_TEXT   = ["HELLO", "WORLD"]

os.makedirs(OUT_FOLDER, exist_ok=True)
os.makedirs(TEMP_FOLDER, exist_ok=True)

# Find Chrome
CHROME_EXE = None
for p in [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]:
    if os.path.exists(p):
        CHROME_EXE = p
        break

print(f"Chrome: {CHROME_EXE or 'NOT FOUND'}")
if not CHROME_EXE:
    sys.exit("Chrome/Edge not found — install Chrome first")

# ── Strategy A: SVG glyph collage (preserves font colours) ────────────────────
def render_collage(font_path, text_lines, canvas_w):
    try:
        from fontTools.ttLib import TTFont
    except ImportError:
        return None, "fontTools not installed"

    try:
        ft   = TTFont(font_path)
        svgt = ft.get('SVG ')
        cmap = ft.getBestCmap()
        gord = ft.getGlyphOrder()
        upem = ft['head'].unitsPerEm
        asc  = ft['hhea'].ascent
        desc = ft['hhea'].descent
        hmtx = ft['hmtx'].metrics
    except Exception as e:
        return None, f"fontTools load error: {e}"

    if not svgt:
        return None, "no SVG table in font"

    svg_map = {}
    for doc in svgt.docList:
        for gid in range(doc.startGlyphID, doc.endGlyphID + 1):
            svg_map[gid] = (doc.data, doc.startGlyphID != doc.endGlyphID)

    max_chars  = max(len(l) for l in text_lines if l)
    target_w   = int(canvas_w * 0.85 / max(1, max_chars))
    glyph_h    = max(600, min(2400, target_w * 2))
    vb_h       = asc * 2
    scale      = glyph_h / vb_h

    all_chars = sorted(set(ch for l in text_lines for ch in l if ch.strip()))
    char_svg  = {}
    x_pos     = {}

    for ch in all_chars:
        cp   = ord(ch)
        gn   = cmap.get(cp)
        if not gn:
            continue
        try:
            gid = gord.index(gn)
        except ValueError:
            continue
        entry = svg_map.get(gid)
        if not entry:
            continue
        svg_raw, shared = entry
        adv   = max(1, int(hmtx.get(gn, (upem//3, 0))[0] * scale))
        svg   = svg_raw
        vb    = f'viewBox="0 {-asc} {upem} {vb_h}"'
        svg   = re.sub(r'viewBox="[^"]*"', vb, svg) if 'viewBox' in svg else svg.replace('<svg', f'<svg {vb}', 1)
        svg   = re.sub(r'\s+width="[^"]*"', '', svg)
        svg   = re.sub(r'\s+height="[^"]*"', '', svg)
        svg   = svg.replace('<svg', f'<svg width="{adv}px" height="{glyph_h}px"', 1)
        if shared:
            hide = f'<style>g{{display:none}}#glyph{gid}{{display:inline}}#glyph\\.{gid}{{display:inline}}</style>'
            svg  = re.sub(r'(<svg[^>]*>)', r'\1' + hide, svg, count=1)
        char_svg[ch] = (adv, glyph_h, svg)

    if not char_svg:
        return None, "no glyphs found in SVG table"

    cw = sum(v[0] for v in char_svg.values()) + 10
    ch_h = max(v[1] for v in char_svg.values()) + 10
    items = ""
    cx = 0
    for ch, (adv, h, svg_str) in char_svg.items():
        b64 = base64.b64encode(svg_str.encode()).decode()
        items += f'<img style="position:absolute;left:{cx}px;top:0;width:{adv}px;height:{h}px" src="data:image/svg+xml;base64,{b64}">\n'
        x_pos[ch] = cx
        cx += adv

    html = f"""<!DOCTYPE html><html><head><style>
*{{margin:0;padding:0}}html,body{{background:#ffffff;overflow:hidden}}
.c{{position:relative;width:{cw}px;height:{ch_h}px}}
</style></head><body><div class="c">{items}</div></body></html>"""

    hp = os.path.join(TEMP_FOLDER, "test_collage.html")
    pp = os.path.join(TEMP_FOLDER, "test_collage.png")
    with open(hp, "w", encoding="utf-8") as f:
        f.write(html)
    if os.path.exists(pp):
        os.remove(pp)

    subprocess.run([
        CHROME_EXE, "--headless", "--no-sandbox", "--disable-gpu",
        "--no-first-run", "--disable-sync", "--disable-extensions",
        f"--screenshot={pp}", f"--window-size={cw},{ch_h}",
        "file:///" + hp.replace("\\", "/"),
    ], capture_output=True, timeout=40)

    if not os.path.exists(pp):
        return None, "Chrome did not produce screenshot"

    img = Image.open(pp).convert("RGBA")
    arr = np.array(img)
    white = (arr[:,:,0]>240) & (arr[:,:,1]>240) & (arr[:,:,2]>240)
    arr[white, 3] = 0
    if arr[:,:,3].max() == 0:
        return None, "all pixels transparent after white-key"

    # Crop individual glyphs and compose lines
    dpr = img.width / cw if cw > 0 else 1.0
    glyph_imgs = {}
    for ch, (adv, h, _) in char_svg.items():
        x0 = int(x_pos[ch] * dpr)
        x1 = int((x_pos[ch] + adv) * dpr)
        y1 = int(h * dpr)
        crop = Image.fromarray(arr).crop((x0, 0, min(x1, img.width), min(y1, img.height)))
        if dpr != 1.0:
            crop = crop.resize((adv, h), Image.LANCZOS)
        ca = np.array(crop)
        w2 = (ca[:,:,0]>240)&(ca[:,:,1]>240)&(ca[:,:,2]>240)
        ca[w2,3] = 0
        glyph_imgs[ch] = Image.fromarray(ca) if ca[:,:,3].max() > 0 else None

    line_imgs = []
    for line in text_lines:
        x = 0
        parts = []
        for ch in line:
            adv = char_svg[ch][0] if ch in char_svg else int((upem//3)*scale)
            parts.append((glyph_imgs.get(ch), x, adv))
            x += adv
        if x <= 0:
            continue
        li = Image.new("RGBA", (x, glyph_h), (0,0,0,0))
        for gi, gx, _ in parts:
            if gi:
                li.paste(gi, (gx, 0), gi)
        line_imgs.append(li)

    if not line_imgs:
        return None, "no line images composed"

    total_h = glyph_h * len(line_imgs)
    result  = Image.new("RGBA", (canvas_w, total_h), (0,0,0,0))
    for i, li in enumerate(line_imgs):
        cx2 = max(0, (canvas_w - li.width) // 2)
        result.paste(li, (cx2, i * glyph_h), li)

    bbox = result.getbbox()
    if not bbox:
        return None, "blank result"
    result = result.crop(bbox)

    avail = int(canvas_w * 0.85)
    if result.width != avail:
        ratio  = avail / result.width
        result = result.resize((avail, max(1, int(result.height * ratio))), Image.LANCZOS)

    return result, "OK (SVG collage)"


# ── Strategy B: CSS @font-face with base64 font ───────────────────────────────
def render_css(font_path, text_lines, canvas_w):
    max_chars   = max(len(l) for l in text_lines if l)
    font_px     = max(80, min(500, int(canvas_w * 0.85 / max(1, max_chars))))
    line_height = int(font_px * 1.25)
    total_h     = line_height * len(text_lines) + font_px
    bg = "#00FE00"

    with open(font_path, "rb") as fh:
        fb64 = base64.b64encode(fh.read()).decode()
    fmt = "opentype" if font_path.lower().endswith(".otf") else "truetype"

    import html as _html
    lines_html = "\n".join(
        f'<div class="tl">{_html.escape(l) if l.strip() else "&nbsp;"}</div>'
        for l in text_lines
    )

    html_src = f"""<!DOCTYPE html><html><head><meta charset="utf-8"><style>
*{{margin:0;padding:0;box-sizing:border-box}}
html,body{{background:{bg};width:{canvas_w}px}}
@font-face{{font-family:'PF';src:url('data:font/{fmt};base64,{fb64}') format('{fmt}')}}
.tl{{font-family:'PF',sans-serif;font-size:{font_px}px;
     line-height:{line_height}px;text-align:center;
     width:{canvas_w}px;white-space:pre;overflow:hidden}}
</style></head><body>{lines_html}</body></html>"""

    hp = os.path.join(TEMP_FOLDER, "test_css.html")
    pp = os.path.join(TEMP_FOLDER, "test_css.png")
    with open(hp, "w", encoding="utf-8") as f:
        f.write(html_src)
    if os.path.exists(pp):
        os.remove(pp)

    subprocess.run([
        CHROME_EXE, "--headless", "--no-sandbox", "--disable-gpu",
        "--no-first-run", "--disable-sync", "--disable-extensions",
        f"--screenshot={pp}", f"--window-size={canvas_w},{max(200, total_h)}",
        "file:///" + hp.replace("\\", "/"),
    ], capture_output=True, timeout=30)

    if not os.path.exists(pp):
        return None, "Chrome did not produce screenshot"

    img = Image.open(pp).convert("RGBA")
    arr = np.array(img)
    lime = (arr[:,:,0] < 30) & (arr[:,:,1] > 240) & (arr[:,:,2] < 30)
    arr[lime, 3] = 0
    if arr[:,:,3].max() == 0:
        return None, "all pixels transparent after lime-key (font may not have loaded)"

    result = Image.fromarray(arr)
    bbox = result.getbbox()
    return (result.crop(bbox), "OK (CSS @font-face)") if bbox else (None, "blank after crop")


# ── Run tests ─────────────────────────────────────────────────────────────────
fonts = [f for f in os.listdir(FONT_FOLDER) if f.lower().endswith((".otf", ".ttf"))]
print(f"\nTesting {len(fonts)} fonts -> output: {OUT_FOLDER}\n")
print(f"{'Font':<35} {'SVG table':<12} {'Strategy A':<25} {'Strategy B':<25}")
print("-" * 100)

for fname in sorted(fonts):
    path = os.path.join(FONT_FOLDER, fname)
    name = os.path.splitext(fname)[0]

    # Check SVG table
    try:
        from fontTools.ttLib import TTFont
        ft  = TTFont(path)
        has_svg = "YES" if ft.get('SVG ') else "no"
    except Exception as e:
        has_svg = f"err:{e}"

    img_a, msg_a = render_collage(path, TEST_TEXT, CANVAS_W)
    img_b, msg_b = render_css(path, TEST_TEXT, CANVAS_W)

    best = img_a or img_b
    if best:
        out_png = os.path.join(OUT_FOLDER, f"{name}.png")
        # White background for visibility
        bg_img = Image.new("RGBA", (CANVAS_W, best.height + 40), (200, 200, 200, 255))
        cx = max(0, (CANVAS_W - best.width) // 2)
        bg_img.paste(best, (cx, 20), best)
        bg_img.convert("RGB").save(out_png)
        status = f"SAVED -> {name}.png"
    else:
        status = "FAILED"

    print(f"{fname:<35} {has_svg:<12} {msg_a:<25} {msg_b:<25}  {status}")

print(f"\nDone. Check {OUT_FOLDER} for PNG previews.")
