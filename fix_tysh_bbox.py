
content = open(r'W:\VarsaniAutomation\psd_text_layer.py', encoding='utf-8').read()

# In _build_TySh, after writing the warp descriptor, add the 4*8 bounding box
OLD = '''    buf.write(warp2.encode())

    return buf.getvalue()'''

NEW = '''    buf.write(warp2.encode())

    # Per Adobe PSD spec: TySh ends with 4 x 8-byte doubles (left, top, right, bottom)
    buf.write(struct.pack('>4d', float(left), float(top), float(right), float(bottom)))

    return buf.getvalue()'''

if OLD in content:
    content = content.replace(OLD, NEW)
    open(r'W:\VarsaniAutomation\psd_text_layer.py', 'w', encoding='utf-8').write(content)
    print("OK: added 4x8 bounding box to end of TySh")
    import ast; ast.parse(content); print("SYNTAX OK")
else:
    print("ERROR: pattern not found")
