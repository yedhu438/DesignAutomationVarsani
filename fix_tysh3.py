
content = open(r'W:\VarsaniAutomation\psd_text_layer.py', encoding='utf-8').read()

OLD = '''    def encode(self):
        buf = io.BytesIO()
        buf.write(_id_str(self.class_id))
        buf.write(struct.pack('>I', 0))   # class name = empty unicode string
        buf.write(struct.pack('>I', len(self._items)))'''

NEW = '''    def encode(self):
        buf = io.BytesIO()
        # Class ID must use Unicode string format (4-byte char count + UTF-16BE)
        # NOT the same as item keys which use ASCII id_str (len=0 means 4 chars)
        cid_str = self.class_id if isinstance(self.class_id, str) else self.class_id.decode('ascii')
        buf.write(struct.pack('>I', len(cid_str)))
        buf.write(cid_str.encode('utf-16-be'))
        # Class name: empty unicode string
        buf.write(struct.pack('>I', 0))
        buf.write(struct.pack('>I', len(self._items)))'''

if OLD in content:
    content = content.replace(OLD, NEW)
    open(r'W:\VarsaniAutomation\psd_text_layer.py', 'w', encoding='utf-8').write(content)
    print("OK: class ID encoding fixed")
    import ast; ast.parse(content); print("SYNTAX OK")
else:
    print("ERROR: not found")
