# Varsany Print Automation — Project Knowledge Base

> Business: Varsany / Fullymerched — Amazon custom print-on-demand
> Last updated: 2026-03-20

---

## 1. Business Overview

**Company:** Varsany (sold on Amazon UK as CrystalsRus)
**Store:** amazon.co.uk/stores/Varsany
**Business model:** Customer customises garment on Amazon → order placed →
design team creates print file → printing department prints and ships

**Products sold:**
- Hoodies, sweatshirts, zip sweatshirts
- T-shirts (adult, kids, polo)
- Baby vests, slippers, swimsuits, legsuits
- Gymnastics towel ponchos, tote bags, track pants, leggings

**Print type:** DTF (Direct to Film) — all products personalised per order

---

## 2. Current Manual Process (Before Automation)

1. Customer places order on Amazon with customisation (image, text, font, colour)
2. Order data saved to cloud database (Azure/AWS SQL Server)
3. Designer opens order page at `crssoft.co.uk/Order/CustomOrder`
4. Designer manually upscales image in Topaz Labs website
5. Designer opens Photoshop, creates canvas for correct product/zone
6. Designer types customer text, applies font and colour manually
7. Designer positions elements matching the Amazon preview image
8. Designer saves as [OrderID].psd to Synology Drive
9. File syncs from India to UK via Synology + SpeedFusion/Peplink router
10. Printing department opens PSD, flattens, and prints

**Time per order:** 10-15 minutes manually
**Target with automation:** 30-60 seconds automatically

---

## 3. Database Structure

**Database name:** dbAmazonCustomOrders
**Type:** Microsoft SQL Server
**Hosted:** Azure / AWS cloud server
**Local backup:** .bak file restored to localhost\SQLEXPRESS for testing
**Connection (local):**
```
Server=localhost\SQLEXPRESS; Database=dbAmazonCustomOrders;
Trusted_Connection=yes; TrustServerCertificate=yes;
```
**Connection (live):**
```
Server=your-server.database.windows.net;
Database=dbAmazonCustomOrders;
UID=automation_user; PWD=your_password;
TrustServerCertificate=yes;
```

### tblCustomOrder — Key Columns
| Column | Type | Description |
|---|---|---|
| idCustomOrder | uniqueidentifier PK | Primary key |
| OrderID | nvarchar | Amazon order ID e.g. 205-6487629-5805162 |
| SKU | nvarchar | Encodes colour+size e.g. MenTee_BlkM |
| Quantity | int | Number of copies ordered |
| ItemType | nvarchar | Product type (DTF etc.) |
| IsShipped | bit | Whether shipped |
| Notes | nvarchar | Used for automation error messages |

### tblCustomOrderDetails — Key Columns
| Column | Type | Description |
|---|---|---|
| idCustomOrderDetails | uniqueidentifier PK | Primary key |
| idCustomOrder | uniqueidentifier FK | Links to tblCustomOrder |
| PrintLocation | nvarchar | e.g. "Front + Back + Sleeve" |
| IsFrontLocation | bit | Front zone active |
| IsBackLocation | bit | Back zone active |
| IsPocketLocation | bit | Pocket zone active |
| IsSleeveLocation | bit | Sleeve zone active |
| FrontImage | nvarchar(500) | URL to customer uploaded image |
| FrontText | nvarchar(1000) | Customer text (newline separated) |
| FrontFonts | nvarchar(200) | Font name |
| FrontColours | nvarchar(200) | Hex colour e.g. #ffffff |
| FrontPreviewImage | nvarchar(500) | Amazon preview image URL |
| FrontPremiumFont | nvarchar | Yes/No premium font flag |
| IsFrontPSDDownload | bit | Whether PSD generated |
| (same columns for Back, Pocket, Sleeve) | | |
| IsOrderProcess | bit | Whether order processed |
| IsDesignComplete | bit | Whether design complete |
| ProcessBy | nvarchar | Set to 'AutomationScript' by automation |
| AdditionalPSD | nvarchar(500) | Output file path |
| CustomizationJSON | nvarchar | Full customization data |

### New Columns to Add (run on both local and live DB)
```sql
USE dbAmazonCustomOrders;
ALTER TABLE tblCustomOrderDetails ADD
    IsFrontBgRemove  bit NULL DEFAULT 0,
    IsBackBgRemove   bit NULL DEFAULT 0,
    IsPocketBgRemove bit NULL DEFAULT 0,
    IsSleeveBgRemove bit NULL DEFAULT 0,
    QCStatus         nvarchar(20)  NULL DEFAULT 'pending',
    QCNotes          nvarchar(500) NULL,
    IsComplexOrder   bit           NULL DEFAULT 0,
    OutputFilePath   nvarchar(500) NULL;

-- Local DB only (live DB uses short URLs):
ALTER TABLE tblCustomOrderDetails ALTER COLUMN FrontImage  nvarchar(500);
ALTER TABLE tblCustomOrderDetails ALTER COLUMN BackImage   nvarchar(500);
ALTER TABLE tblCustomOrderDetails ALTER COLUMN PocketImage nvarchar(500);
ALTER TABLE tblCustomOrderDetails ALTER COLUMN SleeveImage nvarchar(500);
ALTER TABLE tblCustomOrderDetails ALTER COLUMN FrontText   nvarchar(1000);
ALTER TABLE tblCustomOrderDetails ALTER COLUMN BackText    nvarchar(1000);
ALTER TABLE tblCustomOrderDetails ALTER COLUMN SleeveText  nvarchar(1000);
ALTER TABLE tblCustomOrderDetails ALTER COLUMN PocketText  nvarchar(1000);
```

---

## 4. Print Specifications

**Resolution:** 320 pixels/centimetre = 812.8 DPI
**Colour mode:** CMYK
**ICC Profile:** U.S. Web Coated (SWOP) v2
**Background:** Always transparent
**Format:** PSD with layers (printing team flattens before printing)
**Large files:** Auto-save as PSB if file exceeds 2GB

### Canvas Sizes (formula: cm × 320 = pixels)
| Product | Zone | Width px | Height px |
|---|---|---|---|
| Hoodie / T-shirt | Front | 9600 | 9600 |
| Hoodie / T-shirt | Back | 9600 | 14400 |
| Hoodie / T-shirt | Sleeve | 4800 | 9600 |
| Hoodie / T-shirt | Pocket | 9600 | 9600 |
| Kids T-shirt | Front/Back | 7360 | 9600 |
| Baby Vest | Front | 4800 | 5440 |
| Tote Bag | Front | 8886 | 8886 |
| Tote Bag | Back | 8886 | 18941 |
| Slippers | Front | 3654 | 2240 |
| Long T-shirt | Front/Back | 9514 | 11354 |

---

## 5. Amazon Print Location Combos

| Combo | Zones | Price |
|---|---|---|
| Front Pocket ONLY | pocket | base |
| Front ONLY | front | +INR 245.67 |
| Back ONLY | back | +INR 245.67 |
| Front + Back | front, back | +INR 491.34 |
| Front Pocket + Back | pocket, back | +INR 491.34 |
| Front + Sleeve | front, sleeve | +INR 368.51 |
| Front + Back + Sleeve | front, back, sleeve | +INR 614.18 |

---

## 6. File Storage & Synology Drive

**Drive letter:** Z:
**Drive name:** Vector Designs
**Path:** Z:\Drive DTF Orders\1. Amazon DTF\
**Sync:** India Synology → UK Synology via SpeedFusion/Peplink B One router

### Recommended Output Structure
```
Z:\Drive DTF Orders\1. Amazon DTF\Automation Output\
└── 2026-03-20\
    ├── 205-6487629-5805162_front.psd
    ├── 205-6487629-5805162_sleeve.psd
    └── ...
```
Script auto-creates dated subfolder per day.

---

## 7. Design Rules (from designer)

### Background Removal Rule
Remove background when background colour matches garment colour:
- White background + white t-shirt → REMOVE
- Black background + black hoodie → REMOVE
- White background + black hoodie → KEEP (white rectangle is intentional)
- Use IsFrontBgRemove flag in database (set by designer on order page)

### Font Sizing
- Auto-fit: largest font where all text lines fit within canvas width and height
- Short text (e.g. "I Love AI") → ~217pt
- Long text → smaller, still fills canvas nicely

### Text Line Breaks
- Customer types line breaks on Amazon → stored as \n in database
- Parse: `text.split("\n")`

### Zone Labels
Small black text in top-left corner of every PSD:
- "front", "back", "sleeve", "pocket", "pocket left", "pocket right"
- Front zone sometimes uses SKU code e.g. "Ylw34" (Yellow size 3-4)

### Quantity > 1
Stack copies vertically on one canvas. Quantity 3 = canvas height × 3.

### Photo Collage (multiple images)
- 1 image → full canvas, 2 → side by side, 4 → 2×2, 6 → 2×3 grid
- Text (if any) goes below collage

### Screenshot Border Removal
Some customers upload screenshots with black letterbox borders.
Detect: scan edges, if 85%+ pixels are dark → auto-crop borders.

### Amazon Position Data
Amazon does NOT share X/Y coordinates of where customer dragged elements.
Only the preview image (FrontPreviewImage) shows the correct position.
Use preview image as reference for crop/position decisions.

---

## 8. PSD Templates

**Total:** 28 templates
**Organisation:** One PSD per product type + zone
**Naming:** {product}_{zone}.psd e.g. hoodie_front.psd, tshirt_sleeve.psd

### Template Map
```python
TEMPLATE_MAP = {
    ("hoodie",    "front"):  "hoodie_front.psd",
    ("hoodie",    "back"):   "hoodie_back.psd",
    ("hoodie",    "sleeve"): "hoodie_sleeve.psd",
    ("hoodie",    "pocket"): "hoodie_pocket.psd",
    ("tshirt",    "front"):  "tshirt_front.psd",
    ("tshirt",    "back"):   "tshirt_back.psd",
    ("tshirt",    "sleeve"): "tshirt_sleeve.psd",
    ("tshirt",    "pocket"): "tshirt_pocket.psd",
    ("kidstshirt","front"):  "kidstshirt_front.psd",
    ("kidstshirt","back"):   "kidstshirt_back.psd",
    ("babyvest",  "front"):  "babyvest_front.psd",
    ("totebag",   "front"):  "totebag_front.psd",
    ("totebag",   "back"):   "totebag_back.psd",
    ("slipper",   "front"):  "slipper_front.psd",
    ("swimsuit",  "front"):  "swimsuit_front.psd",
    ("leggings",  "front"):  "leggings_front.psd",
    ("polo",      "front"):  "polo_front.psd",
    ("polo",      "back"):   "polo_back.psd",
    ("polo",      "pocket"): "polo_pocket.psd",
    ("poncho",    "front"):  "poncho_front.psd",
}
```

---

## 9. Premium Fonts

- Database field: FrontPremiumFont / BackPremiumFont (Yes/No)
- When Yes: colour ignored, font has built-in texture (camo, glitter etc.)
- Font .ttf file must be in C:\Varsany\Fonts\ matching exact name in DB
- If font file found → auto-process
- If font file missing → flag to designer with email/Slack alert

---

## 10. Automation Pipeline

**Script:** varsany_automation.py
**Language:** Python 3.10+
**Poll interval:** Every 30 seconds

### Dependencies
```bash
pip install pyodbc psd-tools Pillow python-dotenv rembg flask
# Also install: GIMP (for layered PSD), Real-ESRGAN (for upscaling)
winget install GIMP.GIMP
```

### Processing Steps Per Order
1. Read order from tblCustomOrderDetails (IsDesignComplete=0)
2. Check premium font → flag if .ttf missing
3. Check complexity → flag if screenshot borders, low res etc.
4. For each active zone:
   a. Download image from URL
   b. Remove screenshot borders (detect black letterbox)
   c. Remove background if Is[Zone]BgRemove=1
   d. Upscale 4x via Real-ESRGAN
   e. Open correct PSD template from TEMPLATE_MAP
   f. Place image on canvas
   g. Render text with auto-sized font and correct colour
   h. Add zone label (top-left black text)
   i. Assemble layered PSD via GIMP headless
   j. Save as [OrderID]_[zone].psd to Synology
5. Update DB: IsDesignComplete=1, ProcessBy='AutomationScript'
6. Send Slack/email alert if any zone flagged

---

## 11. Complexity Flags (auto-flagged to designer)

| Code | Reason |
|---|---|
| screenshot_border | Black letterbox borders detected |
| low_resolution | Image < 500px even after upscaling |
| bg_removal_uncertain | Confidence < 80% |
| too_many_photos | More than 6 photos in collage |
| text_overflow | Text too long for canvas |
| unknown_product | Product not in PRODUCT_CANVAS |
| processing_error | Any script error |
| premium_font | Font .ttf file not installed |

---

## 12. Layered PSD Generation

**Problem:** Pure Python libraries (psd-tools, ImageMagick, Pillow)
cannot create proper multi-layer PSDs.

**Solution: GIMP in headless mode**
- Free, open source, runs on Windows and Linux server
- No GUI opens — completely background process
- Python generates Script-Fu commands, calls GIMP silently
- Output: proper PSD with named editable layers

```bash
# Install
winget install GIMP.GIMP        # Windows
apt install gimp                # Linux server

# Python runs GIMP headlessly:
gimp --no-interface --batch "(script-fu-command)" --batch "(gimp-quit 0)"
```

**Layer structure in output PSD:**
- CustomerImage (pixel layer — image/graphic)
- CustomerText  (pixel layer — rendered text)
- ZoneLabel     (pixel layer — black zone identifier)

**Fallback:** If GIMP unavailable → save separate PNG per layer

---

## 13. Background Removal

**Library:** rembg (model: u2netp — 3x faster than default)
**Trigger:** Designer sets IsFrontBgRemove=1 on order page

**Decision logic (3 layers):**
1. Manual flag (IsFrontBgRemove) — highest priority
2. Background colour matches garment colour from SKU
3. Automated corner analysis (all 4 corners same dark/light colour)

**SKU colour map:**
```python
colour_map = {
    "blk": (20, 20, 20),
    "wht": (255, 255, 255),
    "nvy": (31, 40, 80),
    "red": (200, 30, 30),
    "ylw": (255, 220, 0),
    "pnk": (255, 150, 180),
    "gry": (150, 150, 150),
}
```

---

## 14. QC System

1. Script generates PNG thumbnail of output
2. Compares to FrontPreviewImage (Amazon reference)
3. Confidence score 0-100%
4. ≥85% → auto-move to print folder
5. <85% → alert design team

**Alert contains:** Order ID, reason, thumbnail, link to order page
**Channels:** Email + Slack webhook
**DB flags:** QCStatus, QCNotes, IsComplexOrder

---

## 15. Prototype Web App

**File:** C:\Varsany\prototype_app.py
**Framework:** Flask — runs at http://localhost:5000
**Database:** Local SQLEXPRESS (dbAmazonCustomOrders)

**Run:**
```bash
python prototype_app.py
```

**Features:**
- Customisation form (product, zone, image upload, text, font, colour)
- Live automation progress log (right panel)
- Order history dashboard (bottom)

---

## 16. Folder Structure

```
C:\Varsany\
├── prototype_app.py         Flask prototype
├── varsany_automation.py    Production automation script
├── varsany_automation.log   Runtime log
├── .env                     Secrets (EMAIL_USER, EMAIL_PASS, SLACK_WEBHOOK)
├── Uploads\                 Customer uploaded images
├── Output\                  Finished PSD files (date-organised)
│   └── 2026-03-20\
├── Fonts\                   TTF files (names must match DB font names exactly)
├── Templates\               28 PSD template files
├── Temp\                    Temp files during processing
└── realesrgan\
    └── realesrgan-ncnn-vulkan.exe
```

---

## 17. Developer Tasks Pending

1. Add 8 new columns to tblCustomOrderDetails (SQL in section 3)
2. Add "Remove Background: Yes/No" dropdown on order page for each zone
3. Create automation_user SQL login (SELECT + UPDATE only)
4. Open firewall port 1433 for automation PC IP
5. Confirm whether FrontPreviewImage is populated on live database
6. Video call with Dhruv to reorganise Synology folder structure

---

## 18. Key People

| Person | Role |
|---|---|
| Yedhu | Project owner — Yedhu@fullymerched.com |
| Dhruv | IT / Synology / systems |
| Nimesh | Designer (Photoshop) |
| India team | Design + processing |
| UK team | Printing department |

---

## 19. Important Technical Notes

- **72 PPI display issue:** Photoshop shows 72ppi when opening output but
  pixel dimensions (9600×9600) are correct — this is display metadata only
- **File naming:** [OrderID]_[zone].psd e.g. 205-6487629-5805162_front.psd
- **FrontTextJSON:** Currently empty — Amazon does not share drag position
- **FrontPreviewImage:** May not be populated — confirm with developer
- **PSB:** Auto-triggered when estimated size >2GB
- **CMYK:** PIL converts via image.convert("CMYK")
- **Wrong template:** Was a daily problem for printing team —
  automation fixes this permanently via TEMPLATE_MAP
- **Antigravity:** Uses Google account login only —
  Yedhu@fullymerched.com (Microsoft) won't work — use personal Gmail

---

## 20. Questions Still Open

1. Exact 28 PSD template filenames (ask designer for folder listing)
2. Is FrontPreviewImage populated on live database?
3. What layer names are inside the PSD templates?
4. Video call with Dhruv re: Synology folder structure
5. Font files list in C:\Varsany\Fonts\ on designer PC
6. Does printing team prefer all-in-one file or separate per zone for quantity orders?
