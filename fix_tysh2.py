
content = open(r'W:\VarsaniAutomation\psd_text_layer.py', encoding='utf-8').read()

# Fix 1: Change _Desc.encode() to use Unicode string for class_id
OLD_ENCODE = '''    def encode(self):
        buf = io.BytesIO()
        buf.write(_id_str(self.class_id))
        buf.write(struct.pack('>I', 0))   # class name length = 0
        buf.write(struct.pack('>I', len(self._items)))'''

NEW_ENCODE = '''    def encode(self):
        buf = io.BytesIO()
        # Class ID: Unicode string format (4-byte char count + UTF-16BE)
        # This is DIFFERENT from item keys which use ASCII id_str format
        cid = self.class_id.encode('utf-16-be') if isinstance(self.class_id, str) else self.class_id
        char_count = len(self.class_id) if isinstance(self.class_id, str) else len(self.class_id)//2
        buf.write(struct.pack('>I', char_count))
        buf.write(cid)
        # Class name: empty unicode string (length = 0)
        buf.write(struct.pack('>I', 0))
        buf.write(struct.pack('>I', len(self._items)))'''

if OLD_ENCODE in content:
    content = content.replace(OLD_ENCODE, NEW_ENCODE)
    print("Fix 1 OK: class ID now uses Unicode string encoding")
else:
    print("ERROR fix 1: pattern not found")

# Fix 2: Restore TySh alongside Txt2 now that encoding is correct
OLD_RETURN = '''    # NOTE: We intentionally skip TySh (Type Tool Object Setting) here.
    # Writing a malformed TySh causes Photoshop to show "damaged file" errors.
    # Txt2 (Text Engine Data) is sufficient for Photoshop to recognise a text
    # layer and render it with the correct font and colour.
    txt2 = _build_Txt2(text, font_name, font_size_pts, r, g, b)

    return _tag('Txt2', txt2)'''

NEW_RETURN = '''    tysh = _build_TySh(text, font_name, font_size_pts,
                        r, g, b, left, top, right, bottom)
    txt2 = _build_Txt2(text, font_name, font_size_pts, r, g, b)

    return _tag('TySh', tysh) + _tag('Txt2', txt2)'''

if OLD_RETURN in content:
    content = content.replace(OLD_RETURN, NEW_RETURN)
    print("Fix 2 OK: TySh restored with correct encoding")
else:
    print("ERROR fix 2: pattern not found")

open(r'W:\VarsaniAutomation\psd_text_layer.py', 'w', encoding='utf-8').write(content)

import ast
ast.parse(content)
print("SYNTAX OK")
