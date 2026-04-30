
content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()

OLD = (
    'def cm_to_px(cm): return int(round(cm * PX_PER_CM))\n'
    '\n'
    'PRODUCT_CANVAS = {\n'
    '    "polo":       {"front": (cm_to_px(30), cm_to_px(30)), "back":   (cm_to_px(30), cm_to_px(45)), "pocket": (cm_to_px(15), cm_to_px(20))},\n'
    '    "hoodie":     {"front": (cm_to_px(30), cm_to_px(30)), "back":   (cm_to_px(30), cm_to_px(45)), "sleeve": (cm_to_px(15), cm_to_px(30)), "pocket": (cm_to_px(15), cm_to_px(20))},\n'
    '    "tshirt":     {"front": (cm_to_px(30), cm_to_px(30)), "back":   (cm_to_px(30), cm_to_px(45)), "sleeve": (cm_to_px(15), cm_to_px(30)), "pocket": (cm_to_px(15), cm_to_px(20))},\n'
    '    "kidstshirt": {"front": (cm_to_px(23), cm_to_px(30)), "back":   (cm_to_px(23), cm_to_px(30))},\n'
    '    "totebag":    {"front": (cm_to_px(27), cm_to_px(27)), "back":   (cm_to_px(27), cm_to_px(59))},\n'
    '    "slipper":    {"front": (cm_to_px(11), cm_to_px(7))},\n'
    '    "babyvest":   {"front": (cm_to_px(15), cm_to_px(17))},\n'
    '    "default":    {"front": (cm_to_px(30), cm_to_px(30)), "back":   (cm_to_px(30), cm_to_px(45)), "pocket": (cm_to_px(15), cm_to_px(20)), "sleeve": (cm_to_px(15), cm_to_px(30))},\n'
    '}'
)

NEW = '''def cm_to_px(cm): return int(round(cm * PX_PER_CM))

# ─── CANVAS SIZES — from owner Canvases.xlsx ──────────────────────────────────
PRODUCT_CANVAS = {
    # T-shirts
    "adulttshirt":    {"front": (cm_to_px(30), cm_to_px(30)), "back": (cm_to_px(30), cm_to_px(30)), "pocket": (cm_to_px(9),  cm_to_px(7))},
    "kidstshirt":     {"front": (cm_to_px(23), cm_to_px(30)), "back": (cm_to_px(23), cm_to_px(30)), "pocket": (cm_to_px(9),  cm_to_px(7))},
    # Hoodies
    "adulthoodie":    {"front": (cm_to_px(25), cm_to_px(25)), "back": (cm_to_px(25), cm_to_px(25)), "pocket": (cm_to_px(9),  cm_to_px(7)), "sleeve": (cm_to_px(9), cm_to_px(7))},
    "kidshoodie":     {"front": (cm_to_px(23), cm_to_px(20)), "back": (cm_to_px(23), cm_to_px(20)), "pocket": (cm_to_px(9),  cm_to_px(7))},
    # Bags
    "totebag":        {"front": (cm_to_px(28), cm_to_px(28)), "back": (cm_to_px(28), cm_to_px(28))},
    "backpack":       {"front": (cm_to_px(18), cm_to_px(12))},
    "makeupbag":      {"front": (cm_to_px(23), cm_to_px(14))},
    "shoebag":        {"front": (cm_to_px(23), cm_to_px(14))},
    "shoebag2":       {"front": (cm_to_px(14), cm_to_px(14))},
    "stringbag":      {"front": (cm_to_px(22), cm_to_px(24))},
    "knittingbag":    {"front": (cm_to_px(25), cm_to_px(21))},
    # Accessories
    "buckethat":      {"front": (cm_to_px(18), cm_to_px(5))},
    "beanie":         {"front": (cm_to_px(9),  cm_to_px(4))},
    "socks":          {"front": (cm_to_px(6),  cm_to_px(12))},
    "seatbelt":       {"front": (cm_to_px(18), cm_to_px(4))},
    # Baby / Kids
    "babyvest":       {"front": (cm_to_px(15), cm_to_px(17))},
    "sleepsuit":      {"front": (cm_to_px(13), cm_to_px(18))},
    "hodieblanket":   {"front": (cm_to_px(17), cm_to_px(5))},
    # Home / Other
    "cushion":        {"front": (cm_to_px(30), cm_to_px(30))},
    "memorialplaque": {"front": (cm_to_px(13), cm_to_px(8))},
    "golftowel":      {"front": (cm_to_px(17), cm_to_px(17))},
    "golfcase":       {"front": (cm_to_px(15), cm_to_px(6))},
    "slipper":        {"front": (cm_to_px(6),  cm_to_px(6))},
    # Default fallback
    "default":        {"front": (cm_to_px(30), cm_to_px(30)), "back": (cm_to_px(30), cm_to_px(30)), "pocket": (cm_to_px(9), cm_to_px(7))},
}

# ─── SKU PREFIX → PRODUCT KEY — from owner Canvases.xlsx ─────────────────────
SKU_MAP = [
    # Adult T-shirt
    ("MenTee_",                       "adulttshirt"),
    ("AnyTxtOverSizeTee_",            "adulttshirt"),
    ("WmnTee_",                       "adulttshirt"),
    ("PoloTee_",                      "adulttshirt"),
    ("AdultPoloTee_",                 "adulttshirt"),
    ("SignLan01_Tee_",                "adulttshirt"),
    ("Custom04_Tee_",                 "adulttshirt"),
    ("LegendSince",                   "adulttshirt"),
    ("AnyTxt",                        "adulttshirt"),
    # Kids T-shirt
    ("KidsTee_",                      "kidstshirt"),
    ("SLan01KidsTee_",                "kidstshirt"),
    ("PerSingleLetter01KidsTee_",     "kidstshirt"),
    ("FootballKids",                  "kidstshirt"),
    ("67BdayT02Kid",                  "kidstshirt"),
    # Adult Hoodie
    ("AnyTxtAdultHood_",              "adulthoodie"),
    ("MenHood_",                      "adulthoodie"),
    ("HandStand",                     "adulthoodie"),
    ("SplitGirl",                     "adulthoodie"),
    ("FballN",                        "adulthoodie"),
    ("NewFball",                      "adulthoodie"),
    # Kids Hoodie
    ("AnyTxtKidsHood_",               "kidshoodie"),
    ("KidsHood_",                     "kidshoodie"),
    # Tote Bag
    ("AnyTxtTote_",                   "totebag"),
    ("Tote",                          "totebag"),
    # Backpack
    ("AnyTxtBckpck_",                 "backpack"),
    ("BckPack",                       "backpack"),
    ("Name01",                        "backpack"),
    # Baby Vest
    ("AnyTxtBabyVest_",               "babyvest"),
    ("BabyVest",                      "babyvest"),
    # Bucket Hat
    ("AnyTextHat_",                   "buckethat"),
    # Beanie
    ("AnytxtBeanie_",                 "beanie"),
    # Make Up Bag
    ("AnyTxtMakUp_",                  "makeupbag"),
    # Hoodie Blanket
    ("AnyTxtBlanketHood_",            "hodieblanket"),
    # Shoe Bag Sports
    ("AnyTxtShoeB_",                  "shoebag"),
    # Slipper
    ("AnyTxtSlip",                    "slipper"),
    # Socks
    ("AnyTxtSocks",                   "socks"),
    # Cushion
    ("PCushion",                      "cushion"),
    # Gym / Swim
    ("GymLeo",                        "default"),
    ("SwimSuit",                      "default"),
]'''

if OLD in content:
    content = content.replace(OLD, NEW)
    open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)
    print('SUCCESS: PRODUCT_CANVAS updated with correct sizes from owner file')
else:
    print('ERROR: could not find old block')
    # show what the file has around that area
    idx = content.find('PRODUCT_CANVAS')
    print('Found PRODUCT_CANVAS at index:', idx)
    print(repr(content[idx:idx+300]))
