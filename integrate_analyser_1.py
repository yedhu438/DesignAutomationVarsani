
content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()

# 1. Add import at the top (after existing imports)
OLD_IMPORT = 'try:\n    from rembg import remove as rembg_remove\n    REMBG_AVAILABLE = True\nexcept ImportError:\n    REMBG_AVAILABLE = False'

NEW_IMPORT = '''try:
    from rembg import remove as rembg_remove
    REMBG_AVAILABLE = True
except ImportError:
    REMBG_AVAILABLE = False

# Preview analyser — measures text/photo positions from crssoft preview images
try:
    from preview_analyser import analyse_preview, DEFAULT_LAYOUT
    PREVIEW_ANALYSER_AVAILABLE = True
except ImportError:
    PREVIEW_ANALYSER_AVAILABLE = False
    DEFAULT_LAYOUT = {
        "text_y_start": 0.70, "text_y_end": 0.92,
        "text_x_centre": 0.50, "text_width": 0.90,
        "photo_top": 0.02, "photo_bottom": 0.95,
        "photo_left": 0.05, "photo_right": 0.95,
        "text_above_photo": False,
    }'''

if OLD_IMPORT in content:
    content = content.replace(OLD_IMPORT, NEW_IMPORT)
    print("Step 1 OK: import added")
else:
    print("ERROR step 1: import block not found")

open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)
print("Saved.")
