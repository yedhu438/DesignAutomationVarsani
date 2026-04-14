"""Download Google Fonts to W:\\fonts"""
import urllib.request, os

FONTS_DIR = r"W:\fonts"
os.makedirs(FONTS_DIR, exist_ok=True)

# Direct download URLs from Google Fonts
FONTS = {
    "Bebas Neue.ttf":       "https://fonts.gstatic.com/s/bebasneuepro/v3/B21c45Yl52mSBvpBGb2N1Yb4pCk8tIj_w0iA.ttf",
    "Chewy.ttf":            "https://fonts.gstatic.com/s/chewy/v18/uU9PCBU3RoKS2Zhgw9k.ttf",
    "Permanent Marker.ttf": "https://fonts.gstatic.com/s/permanentmarker/v16/Fh4uPib9Iyv2ucM6pGhO3LQNwl8Xog.ttf",
    "Ultra.ttf":            "https://fonts.gstatic.com/s/ultra/v20/zOLy4prXmrtY-tT6yMOl.ttf",
}

headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

for filename, url in FONTS.items():
    out = os.path.join(FONTS_DIR, filename)
    if os.path.exists(out):
        print(f"Exists: {filename}")
        continue
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
        with open(out, 'wb') as f:
            f.write(data)
        print(f"OK: {filename} ({len(data)//1024} KB)")
    except Exception as e:
        print(f"FAIL: {filename} - {e}")

print("\nDone. Fonts in W:\\fonts:")
for f in sorted(os.listdir(FONTS_DIR)):
    if f.endswith(('.ttf','.otf')):
        size = os.path.getsize(os.path.join(FONTS_DIR, f)) // 1024
        print(f"  {f} ({size} KB)")
