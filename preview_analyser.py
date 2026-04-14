"""
preview_analyser.py
===================
Analyses crssoft.co.uk preview images using OpenCV to extract the EXACT
text and photo placement for each order.

Instead of using hardcoded layout constants (TEXT_TOP_PCT = 0.70 etc.),
the batch processor calls analyse_preview() to measure the actual layout
from each order's own preview image — giving pixel-perfect placement.

Usage:
    from preview_analyser import analyse_preview, DEFAULT_LAYOUT

    layout = analyse_preview(preview_path, text_colour_hex)
    # Returns dict with all placement percentages, or DEFAULT_LAYOUT on failure

    # Use in PSD builder:
    txt_y_start = int(canvas_h * layout['text_y_start'])
    txt_y_end   = int(canvas_h * layout['text_y_end'])
    photo_top   = int(canvas_h * layout['photo_top'])
    ...
"""

import cv2
import numpy as np
from PIL import Image

# ─── DEFAULT LAYOUT ───────────────────────────────────────────────────────────
# Fallback when preview image not available or analysis fails.
# Measured from real crssoft previews in earlier sessions.
DEFAULT_LAYOUT = {
    "text_y_start":   0.70,   # text starts 70% down canvas
    "text_y_end":     0.92,   # text ends at 92%
    "text_x_start":   0.05,   # text left margin ~5%
    "text_x_end":     0.95,   # text right margin ~95%
    "text_x_centre":  0.50,   # centred
    "text_width":     0.90,   # spans 90% of canvas width
    "text_height":    0.20,   # ~20% of canvas height
    "photo_top":      0.02,   # photo starts near top
    "photo_bottom":   0.95,   # photo fills to 95%
    "photo_left":     0.05,
    "photo_right":    0.95,
    "text_above_photo": False,  # text is overlaid at bottom
    "text_overlaps_photo": True,
    "source": "default",
}


# ─── COLOUR DETECTION HELPERS ─────────────────────────────────────────────────

def hex_to_bgr(hex_col):
    """Convert #rrggbb to (B, G, R) for OpenCV."""
    h = hex_col.lstrip('#')
    if len(h) == 3:
        h = ''.join(c*2 for c in h)
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return (b, g, r)
    except:
        return (255, 255, 255)


def build_text_mask(img_bgr, text_colour_hex, tolerance=50):
    """
    Create a binary mask of pixels that match the text colour.
    Handles special cases: white, black, and any hex colour.
    Returns a boolean mask (h x w).
    """
    h_col = text_colour_hex.lstrip('#')
    if len(h_col) == 3:
        h_col = ''.join(c*2 for c in h_col)

    try:
        r, g, b = int(h_col[0:2],16), int(h_col[2:4],16), int(h_col[4:6],16)
    except:
        r, g, b = 255, 255, 255

    # Special case: white text
    if r > 220 and g > 220 and b > 220:
        mask = (
            (img_bgr[:,:,2].astype(int) > 210) &
            (img_bgr[:,:,1].astype(int) > 210) &
            (img_bgr[:,:,0].astype(int) > 210)
        )
        return mask

    # Special case: black text
    if r < 30 and g < 30 and b < 30:
        mask = (
            (img_bgr[:,:,2].astype(int) < 40) &
            (img_bgr[:,:,1].astype(int) < 40) &
            (img_bgr[:,:,0].astype(int) < 40)
        )
        return mask

    # General colour: use inRange with tolerance
    lower = np.array([max(0, b-tolerance), max(0, g-tolerance), max(0, r-tolerance)])
    upper = np.array([min(255, b+tolerance), min(255, g+tolerance), min(255, r+tolerance)])
    cv_mask = cv2.inRange(img_bgr, lower, upper)
    return cv_mask > 0


def find_text_region(mask):
    """
    Find the bounding box of the text pixels in the mask.
    Applies morphological cleaning to remove noise.
    Returns (y_min, y_max, x_min, x_max) or None if not found.
    """
    # Clean up noise with morphological operations
    kernel = np.ones((3, 3), np.uint8)
    cleaned = cv2.morphologyEx(mask.astype(np.uint8), cv2.MORPH_CLOSE, kernel)

    # Find connected components to pick the largest text block
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(cleaned, connectivity=8)

    if num_labels < 2:  # only background
        return None

    # Filter components by size (ignore tiny noise, pick the main text region)
    min_area = mask.shape[0] * mask.shape[1] * 0.0005  # at least 0.05% of canvas
    valid = [(stats[i, cv2.CC_STAT_AREA], i) for i in range(1, num_labels)
             if stats[i, cv2.CC_STAT_AREA] > min_area]

    if not valid:
        return None

    # Combine all valid text components into one bounding box
    y_pixels = np.where(np.isin(labels, [i for _, i in valid]))[0]
    x_pixels = np.where(np.isin(labels, [i for _, i in valid]))[1]

    if len(y_pixels) < 50:
        return None

    return (y_pixels.min(), y_pixels.max(), x_pixels.min(), x_pixels.max())


def find_photo_region(img_bgr):
    """
    Find the customer photo rectangle using edge detection + contour analysis.
    Returns (y_top, y_bottom, x_left, x_right) as pixels, or None.
    """
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 40, 120)

    # Dilate edges slightly to connect nearby boundaries
    kernel = np.ones((3, 3), np.uint8)
    edges = cv2.dilate(edges, kernel, iterations=1)

    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    h, w = img_bgr.shape[:2]
    min_area = h * w * 0.05  # photo must be at least 5% of canvas

    candidates = []
    for c in contours:
        area = cv2.contourArea(c)
        if area < min_area:
            continue
        x, y, cw, ch = cv2.boundingRect(c)
        # Photo should be reasonably rectangular and central-ish
        aspect = cw / max(ch, 1)
        if 0.2 < aspect < 5.0:  # not too thin or too wide
            candidates.append((area, x, y, cw, ch))

    if not candidates:
        return None

    candidates.sort(reverse=True)
    _, x, y, cw, ch = candidates[0]
    return (y, y+ch, x, x+cw)


# ─── MAIN ANALYSIS FUNCTION ───────────────────────────────────────────────────

def analyse_preview(preview_path, text_colour_hex="#ffffff", verbose=False):
    """
    Analyse a crssoft preview image to extract text and photo placement.

    Args:
        preview_path:     Full path to the preview JPG/PNG file
        text_colour_hex:  Text colour from the database (e.g. "#ffffff")
        verbose:          Print debug info if True

    Returns:
        dict with layout percentages, or DEFAULT_LAYOUT if analysis fails
    """
    try:
        img_bgr = cv2.imread(preview_path, cv2.IMREAD_COLOR)
        if img_bgr is None:
            if verbose: print(f"  [preview_analyser] Cannot read: {preview_path}")
            return DEFAULT_LAYOUT.copy()

        h, w = img_bgr.shape[:2]
        if verbose: print(f"  [preview_analyser] Canvas: {w}x{h}, colour: {text_colour_hex}")

        result = DEFAULT_LAYOUT.copy()
        result["source"] = "analysed"
        result["preview_path"] = preview_path
        result["canvas_w"] = w
        result["canvas_h"] = h

        # ── 1. Find text region ────────────────────────────────────────────────
        text_found = False
        for tol in [50, 70, 90]:  # progressively wider tolerance
            mask = build_text_mask(img_bgr, text_colour_hex, tolerance=tol)
            text_box = find_text_region(mask)
            if text_box:
                ty_min, ty_max, tx_min, tx_max = text_box
                result["text_y_start"]   = round(ty_min / h, 4)
                result["text_y_end"]     = round(ty_max / h, 4)
                result["text_x_start"]   = round(tx_min / w, 4)
                result["text_x_end"]     = round(tx_max / w, 4)
                result["text_x_centre"]  = round((tx_min + tx_max) / 2 / w, 4)
                result["text_width"]     = round((tx_max - tx_min) / w, 4)
                result["text_height"]    = round((ty_max - ty_min) / h, 4)
                text_found = True
                if verbose:
                    print(f"  [preview_analyser] Text Y: {result['text_y_start']*100:.1f}% - {result['text_y_end']*100:.1f}%")
                    print(f"  [preview_analyser] Text X: {result['text_x_start']*100:.1f}% - {result['text_x_end']*100:.1f}%")
                break

        if not text_found and verbose:
            print(f"  [preview_analyser] WARNING: Text colour {text_colour_hex} not detected")

        # ── 2. Find photo region ───────────────────────────────────────────────
        photo_box = find_photo_region(img_bgr)
        if photo_box:
            py_top, py_bottom, px_left, px_right = photo_box
            result["photo_top"]    = round(py_top    / h, 4)
            result["photo_bottom"] = round(py_bottom / h, 4)
            result["photo_left"]   = round(px_left   / w, 4)
            result["photo_right"]  = round(px_right  / w, 4)
            if verbose:
                print(f"  [preview_analyser] Photo Y: {result['photo_top']*100:.1f}% - {result['photo_bottom']*100:.1f}%")
                print(f"  [preview_analyser] Photo X: {result['photo_left']*100:.1f}% - {result['photo_right']*100:.1f}%")

        # ── 3. Determine text/photo relationship ──────────────────────────────
        if text_found and photo_box:
            photo_mid_y = (result["photo_top"] + result["photo_bottom"]) / 2
            result["text_above_photo"]    = result["text_y_end"] < photo_mid_y
            result["text_overlaps_photo"] = (
                result["text_y_start"] < result["photo_bottom"] and
                result["text_y_end"]   > result["photo_top"]
            )
            if verbose:
                pos = "ABOVE" if result["text_above_photo"] else "BELOW/OVERLAPPING"
                print(f"  [preview_analyser] Text is {pos} photo midpoint")

        return result

    except Exception as e:
        if verbose:
            print(f"  [preview_analyser] ERROR: {e}")
        layout = DEFAULT_LAYOUT.copy()
        layout["source"] = f"error: {e}"
        return layout


# ─── BATCH LEARNING: CACHE LAYOUT RULES PER PRODUCT TYPE ─────────────────────

def build_layout_cache(orders_with_previews, find_image_fn, output_json_path):
    """
    Analyse a batch of orders and build a layout cache per product type.
    Saves results to JSON for reuse — so each preview is only analysed once.

    This allows the batch processor to learn:
      adulttshirt  -> text at 69-91% (from 500 real orders)
      babyvest     -> text at [different position]
      hoodie       -> text at [different position]

    Args:
        orders_with_previews: list of dicts with keys: product, preview_filename, colour
        find_image_fn:        the find_image() function from batch_processor
        output_json_path:     where to save the learned cache
    """
    import json, os
    from collections import defaultdict

    cache = defaultdict(list)

    for order in orders_with_previews:
        preview_path = find_image_fn(order.get('preview_filename', ''))
        if not preview_path:
            continue
        layout = analyse_preview(preview_path, order.get('colour', '#ffffff'))
        if layout['source'] == 'analysed':
            product = order.get('product', 'default')
            cache[product].append({
                'text_y_start':  layout['text_y_start'],
                'text_y_end':    layout['text_y_end'],
                'text_x_centre': layout['text_x_centre'],
                'text_width':    layout['text_width'],
                'text_height':   layout['text_height'],
                'photo_top':     layout['photo_top'],
                'photo_bottom':  layout['photo_bottom'],
                'photo_left':    layout['photo_left'],
                'photo_right':   layout['photo_right'],
                'text_above':    layout['text_above_photo'],
            })

    # Average per product type
    averaged = {}
    for product, layouts in cache.items():
        n = len(layouts)
        averaged[product] = {
            'text_y_start':  round(sum(l['text_y_start']  for l in layouts)/n, 4),
            'text_y_end':    round(sum(l['text_y_end']    for l in layouts)/n, 4),
            'text_x_centre': round(sum(l['text_x_centre'] for l in layouts)/n, 4),
            'text_height':   round(sum(l['text_height']   for l in layouts)/n, 4),
            'photo_top':     round(sum(l['photo_top']     for l in layouts)/n, 4),
            'photo_bottom':  round(sum(l['photo_bottom']  for l in layouts)/n, 4),
            'photo_left':    round(sum(l['photo_left']    for l in layouts)/n, 4),
            'photo_right':   round(sum(l['photo_right']   for l in layouts)/n, 4),
            'sample_count':  n,
        }
        print(f"  {product}: {n} samples -> text Y {averaged[product]['text_y_start']*100:.1f}%-{averaged[product]['text_y_end']*100:.1f}%")

    os.makedirs(os.path.dirname(output_json_path), exist_ok=True)
    with open(output_json_path, 'w') as f:
        json.dump(averaged, f, indent=2)
    print(f"Layout cache saved: {output_json_path}")
    return averaged


# ─── CLI TEST ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys, json

    if len(sys.argv) < 2:
        print("Usage: python preview_analyser.py <preview_image_path> [text_colour_hex]")
        print("Example: python preview_analyser.py W:/images/Feb-Image/60883356077922-frontpreview.jpg #ffff00")
        sys.exit(1)

    path   = sys.argv[1]
    colour = sys.argv[2] if len(sys.argv) > 2 else "#ffffff"

    print(f"Analysing: {path}")
    print(f"Text colour: {colour}")
    print()

    layout = analyse_preview(path, colour, verbose=True)

    print()
    print("=== RESULT ===")
    print(json.dumps(layout, indent=2, default=str))
