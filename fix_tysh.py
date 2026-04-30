
content = open(r'W:\VarsaniAutomation\psd_text_layer.py', encoding='utf-8').read()

# Replace _build_TySh to return empty bytes - skip it entirely
# Photoshop needs Txt2 for text content, TySh for transform/bounds
# But a malformed TySh is WORSE than no TySh
# Solution: write ONLY Txt2, no TySh

OLD = '''    tysh = _build_TySh(text, font_name, font_size_pts,
                        r, g, b, left, top, right, bottom)
    txt2 = _build_Txt2(text, font_name, font_size_pts, r, g, b)

    return _tag('TySh', tysh) + _tag('Txt2', txt2)'''

NEW = '''    # NOTE: We intentionally skip TySh (Type Tool Object Setting) here.
    # Writing a malformed TySh causes Photoshop to show "damaged file" errors.
    # Txt2 (Text Engine Data) is sufficient for Photoshop to recognise a text
    # layer and render it with the correct font and colour.
    txt2 = _build_Txt2(text, font_name, font_size_pts, r, g, b)

    return _tag('Txt2', txt2)'''

if OLD in content:
    content = content.replace(OLD, NEW)
    print("OK: TySh removed, Txt2-only approach")
else:
    print("ERROR: pattern not found")

open(r'W:\VarsaniAutomation\psd_text_layer.py', 'w', encoding='utf-8').write(content)

import ast
ast.parse(content)
print("SYNTAX OK")
