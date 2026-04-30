
content = open(r'W:\VarsaniAutomation\design_replicator.py', encoding='utf-8').read()

OLD = '''# ─── LAYOUT CONSTANTS ─────────────────────────────────────────────────────────
# These match the proportions observed in the crssoft preview exactly
CANVAS_W    = 900    # px (at 150 DPI equivalent for preview)
CANVAS_H    = 900

# Text zone: top 18% of canvas
TEXT_ZONE_H_PCT   = 0.18

# Photo zone: 18% to 92% vertically, 5% margin on each side
PHOTO_TOP_PCT     = 0.18
PHOTO_BOTTOM_PCT  = 0.94
PHOTO_LEFT_PCT    = 0.05
PHOTO_RIGHT_PCT   = 0.95

# Photo border
PHOTO_BORDER_PX   = 3
PHOTO_BORDER_COLOR = (220, 220, 220)   # light grey

# Background
BG_COLOR = (255, 255, 255, 0)   # transparent'''

NEW = '''# ─── LAYOUT CONSTANTS ─────────────────────────────────────────────────────────
# Measured precisely from crssoft preview images using OpenCV analysis:
#
#   Photo fills the full canvas (top=2%, bottom=95%, left=11%, right=82%)
#   Text is overlaid ON TOP of the photo at the bottom (69%-91% from top)
#   This means text is INSIDE the photo, at the bottom — like a caption overlay
#
CANVAS_W    = 900
CANVAS_H    = 900

# Photo fills almost the entire canvas
PHOTO_TOP_PCT    = 0.02
PHOTO_BOTTOM_PCT = 0.95
PHOTO_LEFT_PCT   = 0.05
PHOTO_RIGHT_PCT  = 0.95

# Text overlaid at bottom of photo (measured from crssoft: 69% to 91%)
TEXT_TOP_PCT     = 0.70    # text starts 70% down the canvas
TEXT_BOTTOM_PCT  = 0.92    # text ends at 92%

# Photo border
PHOTO_BORDER_PX    = 3
PHOTO_BORDER_COLOR = (220, 220, 220)

BG_COLOR = (255, 255, 255, 0)'''

if OLD in content:
    content = content.replace(OLD, NEW)
    open(r'W:\VarsaniAutomation\design_replicator.py', 'w', encoding='utf-8').write(content)
    print('Step 1 OK: layout constants updated')
else:
    print('ERROR step 1: old block not found')
