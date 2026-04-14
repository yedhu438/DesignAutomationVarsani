"""
Test: Men's T-Shirt Multi-Zone PSD
Run: python test_men_tshirt.py
"""
import sys
sys.path.insert(0, r"W:\VarsaniAutomation")
from men_tshirt_engine import build_men_tshirt_psd

def log(msg, level="info"):
    symbols = {"info": "  ", "success": "✓ ", "warning": "⚠ ", "error": "✗ "}
    print(f"{symbols.get(level,'  ')}{msg}")

# ── Test 1: Front + Back (text only) ────────────────────────────────────────
print("\n=== TEST 1: Front + Back (text) ===")
order = {
    "print_location": "Front + Back",
    "front_text":    "GO MOMMY, GO!",
    "front_fonts":   '{"NormalFont":"Arial","PremiumFont":"No"}',
    "front_colours": '{"Colour1":"#ff0000"}',
    "back_text":     "Kitchen\nDisco",
    "back_fonts":    '{"NormalFont":"Arial","PremiumFont":"No"}',
    "back_colours":  '{"Colour1":"#ffa500"}',
}
ok, msg = build_men_tshirt_psd(order, r"C:\Varsany\Output\test_front_back.psd", log)
print(f"Result: {msg}\n")

# ── Test 2: Front Pocket + Back (image filenames) ────────────────────────────
print("=== TEST 2: Front Pocket + Back ===")
order2 = {
    "print_location": "Front Pocket + Back",
    "pocket_image":   "60863796419722-pocket.jpg",
    "back_image":     "60863796419722-back.jpg",
}
ok2, msg2 = build_men_tshirt_psd(order2, r"C:\Varsany\Output\test_pocket_back.psd", log)
print(f"Result: {msg2}\n")

# ── Test 3: All 4 zones ──────────────────────────────────────────────────────
print("=== TEST 3: All 4 zones ===")
order3 = {
    "print_location": "Front + Back + Sleeve",
    "front_text":    "FRONT TEXT",
    "front_fonts":   "Arial",
    "front_colours": "#ffffff",
    "back_text":     "BACK TEXT\nLine 2",
    "back_fonts":    "Arial",
    "back_colours":  "#000000",
    "pocket_text":   "PKT",
    "pocket_fonts":  "Arial",
    "pocket_colours":"#ffffff",
    "sleeve_text":   "SLEEVE",
    "sleeve_fonts":  "Arial",
    "sleeve_colours":"#ffffff",
}
ok3, msg3 = build_men_tshirt_psd(order3, r"C:\Varsany\Output\test_all_zones.psd", log)
print(f"Result: {msg3}\n")

import os
for f in ["test_front_back.psd","test_pocket_back.psd","test_all_zones.psd"]:
    path = f"C:\\Varsany\\Output\\{f}"
    if os.path.exists(path):
        mb = os.path.getsize(path)/1024/1024
        print(f"✓ {f} — {mb:.2f} MB")
    else:
        print(f"✗ {f} — NOT FOUND")
