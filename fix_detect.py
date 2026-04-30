
content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()

OLD = (
    'def detect_product(sku):\n'
    '    if not sku:\n'
    '        return "default"\n'
    '    s = sku.lower()\n'
    '    if "polo" in s:\n'
    '        return "polo"\n'
    '    if "kidstee" in s or "kidstshirt" in s or "kidstee" in s:\n'
    '        return "kidstshirt"\n'
    '    if "hood" in s:\n'
    '        return "hoodie"\n'
    '    if "tote" in s:\n'
    '        return "totebag"\n'
    '    if "slipper" in s:\n'
    '        return "slipper"\n'
    '    if "baby" in s or "vest" in s:\n'
    '        return "babyvest"\n'
    '    return "tshirt"'
)

NEW = (
    'def detect_product(sku):\n'
    '    """Map SKU to product key using SKU_MAP from owner canvas file.\n'
    '    Tries each prefix in order — first match wins.\n'
    '    Falls back to keyword matching, then default.\n'
    '    """\n'
    '    if not sku:\n'
    '        return "default"\n'
    '    # Direct prefix match from owner SKU_MAP\n'
    '    for prefix, product_key in SKU_MAP:\n'
    '        if sku.startswith(prefix):\n'
    '            return product_key\n'
    '    # Keyword fallback for edge cases\n'
    '    s = sku.lower()\n'
    '    if "kidstee" in s:          return "kidstshirt"\n'
    '    if "kidshoo" in s:          return "kidshoodie"\n'
    '    if "hood" in s:             return "adulthoodie"\n'
    '    if "tote" in s:             return "totebag"\n'
    '    if "slipper" in s:          return "slipper"\n'
    '    if "baby" in s:             return "babyvest"\n'
    '    if "vest" in s:             return "babyvest"\n'
    '    if "backpack" in s or "bckpck" in s: return "backpack"\n'
    '    if "beanie" in s:           return "beanie"\n'
    '    if "hat" in s:              return "buckethat"\n'
    '    if "tee" in s or "polo" in s: return "adulttshirt"\n'
    '    return "default"'
)

if OLD in content:
    content = content.replace(OLD, NEW)
    open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)
    print('SUCCESS: detect_product updated to use SKU_MAP')
else:
    print('ERROR: old detect_product block not found')
    idx = content.find('def detect_product')
    print(repr(content[idx:idx+400]))
