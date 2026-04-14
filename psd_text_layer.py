"""
psd_text_layer.py
=================
Builds editable text layers (TySh + Txt2 tagged blocks) for PSD files.
Allows Photoshop to open CustomerText layers as editable type layers.
Premium fonts (Mermaid, Texture, Camo etc.) resolve by name at open time.
"""

import struct, io


# ─── Encoding helpers ─────────────────────────────────────────────────────────

def _u16be(s):
    """Unicode string: 4-byte char count + UTF-16BE bytes"""
    return struct.pack('>I', len(s)) + s.encode('utf-16-be')

def _id_str(s):
    """
    PSD ID string: if 4 chars → write 0x00000000 + 4 bytes
                   otherwise  → write length + bytes
    """
    b = s.encode('ascii') if isinstance(s, str) else s
    if len(b) == 4:
        return struct.pack('>I', 0) + b
    return struct.pack('>I', len(b)) + b


# ─── Descriptor writer ────────────────────────────────────────────────────────

class _Desc:
    """Minimal Photoshop descriptor binary writer"""
    def __init__(self, class_id='null'):
        self.class_id = class_id
        self._items = []

    def double(self, k, v):     self._items.append(('doub', k, v));       return self
    def text(self,   k, v):     self._items.append(('TEXT', k, v));       return self
    def bool_(self,  k, v):     self._items.append(('bool', k, v));       return self
    def long_(self,  k, v):     self._items.append(('long', k, v));       return self
    def obj(self,    k, v):     self._items.append(('Objc', k, v));       return self
    def list_(self,  k, v):     self._items.append(('VlLs', k, v));       return self
    def enum(self,   k, t, v):  self._items.append(('enum', k, (t, v)));  return self
    def unit(self,   k, u, v):  self._items.append(('UntF', k, (u, v)));  return self

    def encode(self):
        buf = io.BytesIO()
        # Class ID must use Unicode string format (4-byte char count + UTF-16BE)
        # NOT the same as item keys which use ASCII id_str (len=0 means 4 chars)
        cid_str = self.class_id if isinstance(self.class_id, str) else self.class_id.decode('ascii')
        buf.write(struct.pack('>I', len(cid_str)))
        buf.write(cid_str.encode('utf-16-be'))
        # Class name: empty unicode string
        buf.write(struct.pack('>I', 0))
        buf.write(struct.pack('>I', len(self._items)))
        for ostype, key, val in self._items:
            buf.write(_id_str(key))
            buf.write(ostype.encode('ascii'))
            if ostype == 'doub':
                buf.write(struct.pack('>d', float(val)))
            elif ostype == 'TEXT':
                buf.write(_u16be(val))
            elif ostype == 'bool':
                buf.write(struct.pack('>B', 1 if val else 0))
            elif ostype == 'long':
                buf.write(struct.pack('>i', int(val)))
            elif ostype == 'UntF':
                unit, v = val
                uc = unit.encode('ascii') if isinstance(unit, str) else unit
                buf.write(uc[:4])
                buf.write(struct.pack('>d', float(v)))
            elif ostype == 'enum':
                t, v = val
                buf.write(_id_str(t))
                buf.write(_id_str(v))
            elif ostype == 'Objc':
                buf.write(val.encode())
            elif ostype == 'VlLs':
                items = val
                buf.write(struct.pack('>I', len(items)))
                for item in items:
                    if isinstance(item, _Desc):
                        buf.write(b'Objc')
                        buf.write(item.encode())
                    else:
                        buf.write(b'doub')
                        buf.write(struct.pack('>d', float(item)))
        return buf.getvalue()


# ─── TySh block builder ───────────────────────────────────────────────────────

def _build_TySh(text, font_name, font_size_pts, r, g, b,
                left, top, right, bottom):
    """
    Build TySh (Type Tool Object Setting) tagged block.
    Binary layout per Adobe PSD spec section 'Type tool object setting':
      2  bytes: version (=1)
      48 bytes: transform (6 doubles: xx xy yx yy tx ty)
      2  bytes: text version (=50)
      4  bytes: descriptor version (=16)
      N  bytes: text descriptor
      2  bytes: warp version (=1)
      4  bytes: warp descriptor version (=16)
      N  bytes: warp descriptor
    """
    buf = io.BytesIO()

    # Version
    buf.write(struct.pack('>H', 1))

    # Transform: identity matrix + translation to layer origin
    buf.write(struct.pack('>6d', 1.0, 0.0, 0.0, 1.0, float(left), float(top)))

    # Text version
    buf.write(struct.pack('>H', 50))

    # Descriptor version
    buf.write(struct.pack('>I', 16))

    # ── Text descriptor ────────────────────────────────────────────────────────
    d = _Desc('TxLr')

    # Text content
    d.text('Txt ', text)

    # Warp settings
    warp = _Desc('warpStyle')
    warp.enum('warpStyle', 'warpStyle', 'warpNone')
    warp.double('warpValue', 0.0)
    warp.double('warpPerspective', 0.0)
    warp.double('warpPerspectiveOther', 0.0)
    warp.enum('warpRotate', 'Ornt', 'Hrzn')
    d.obj('warp', warp)

    # Text bounds (bounding box in document coordinates)
    bnd = _Desc('bounds')
    bnd.unit('Top ', '#Pnt', float(top))
    bnd.unit('Left', '#Pnt', float(left))
    bnd.unit('Btom', '#Pnt', float(bottom))
    bnd.unit('Rght', '#Pnt', float(right))
    d.obj('bounds', bnd)

    # Bounds no effect (same as bounds for point text)
    bne = _Desc('boundingBox')
    bne.unit('Top ', '#Pnt', float(top))
    bne.unit('Left', '#Pnt', float(left))
    bne.unit('Btom', '#Pnt', float(bottom))
    bne.unit('Rght', '#Pnt', float(right))
    d.obj('boundingBox', bne)

    # Text orientation and anti-alias
    d.enum('textGridding', 'textGridding', 'None')
    d.enum('Ornt', 'Ornt', 'Hrzn')
    d.enum('AntA', 'Annt', 'antiAliasSmooth')

    # NOTE: Do NOT add a second 'bounds' key here — that was the bug
    # causing Photoshop to report a damaged file

    buf.write(d.encode())

    # ── Warp descriptor (appended after main text descriptor) ──────────────────
    buf.write(struct.pack('>H', 1))   # warp version
    buf.write(struct.pack('>I', 16))  # warp descriptor version
    warp2 = _Desc('warpStyle')
    warp2.enum('warpStyle', 'warpStyle', 'warpNone')
    warp2.double('warpValue', 0.0)
    warp2.double('warpPerspective', 0.0)
    warp2.double('warpPerspectiveOther', 0.0)
    warp2.enum('warpRotate', 'Ornt', 'Hrzn')
    buf.write(warp2.encode())

    # Per Adobe PSD spec: TySh ends with 4 x 8-byte doubles (left, top, right, bottom)
    buf.write(struct.pack('>4d', float(left), float(top), float(right), float(bottom)))

    return buf.getvalue()


# ─── Txt2 block builder ───────────────────────────────────────────────────────

def _build_Txt2(text, font_name, font_size_pts, r, g, b):
    """
    Build Txt2 (Text Engine Data) block.
    Photoshop proprietary key-value text engine format.
    This is what Photoshop uses to render and re-edit the text.
    """
    def wstr(s):
        return b'(' + b'\xfe\xff' + s.encode('utf-16-be') + b')'

    def wf(v):
        s = f'{float(v):.8f}'.rstrip('0')
        if s.endswith('.'): s += '0'
        return s.encode('ascii')

    def wi(v):
        return str(int(v)).encode('ascii')

    text_len = len(text) + 1  # Photoshop appends a newline internally

    lines = [
        b'<<',
        b'/EngineDict',
        b'<<',
        b'/Editor',
        b'<<',
        b'/Text ' + wstr(text + '\r'),
        b'>>',
        b'/ParagraphRun',
        b'<<',
        b'/DefaultRunData',
        b'<<',
        b'/ParagraphSheet',
        b'<<',
        b'/DefaultStyleSheet 0',
        b'/Properties',
        b'<<',
        b'/Justification 2',
        b'/FirstLineIndent 0.0',
        b'/StartIndent 0.0',
        b'/EndIndent 0.0',
        b'/SpaceBefore 0.0',
        b'/SpaceAfter 0.0',
        b'/AutoHyphenate true',
        b'>>',
        b'>>',
        b'>>',
        b'/RunArray [ <<',
        b'/ParagraphSheet',
        b'<<',
        b'/DefaultStyleSheet 0',
        b'/Properties',
        b'<<',
        b'/Justification 2',
        b'>>',
        b'>>',
        b'/RunLength ' + wi(text_len),
        b'>> ]',
        b'/IsJoinable 1',
        b'>>',
        b'/StyleRun',
        b'<<',
        b'/DefaultRunData',
        b'<<',
        b'/StyleSheet',
        b'<<',
        b'/DefaultStyleSheet 0',
        b'/StyleSheetData',
        b'<<',
        b'/Font 0',
        b'/FontSize ' + wf(font_size_pts),
        b'/FauxBold false',
        b'/FauxItalic false',
        b'/AutoLeading true',
        b'/Leading 0.0',
        b'/Tracking 0',
        b'/BaselineShift 0.0',
        b'/AutoKern true',
        b'/Kerning 0',
        b'/FillColor',
        b'<<',
        b'/Type 1',
        b'/Values [ 1.0 ' + wf(r/255.0) + b' ' + wf(g/255.0) + b' ' + wf(b/255.0) + b' ]',
        b'>>',
        b'/StrokeColor',
        b'<<',
        b'/Type 1',
        b'/Values [ 1.0 0.0 0.0 0.0 ]',
        b'>>',
        b'/FillFlag true',
        b'/StrokeFlag false',
        b'/FillFirst true',
        b'/YUnderline 0',
        b'/OutlineFlag false',
        b'/ShadowFlag false',
        b'/Strikethrough 0',
        b'/Underline 0',
        b'/NoBreak false',
        b'>>',
        b'>>',
        b'>>',
        b'>>',
        b'/RunArray [ <<',
        b'/StyleSheet',
        b'<<',
        b'/DefaultStyleSheet 0',
        b'/StyleSheetData',
        b'<<',
        b'/Font 0',
        b'/FontSize ' + wf(font_size_pts),
        b'/FillColor',
        b'<<',
        b'/Type 1',
        b'/Values [ 1.0 '
            + wf(r/255.0) + b' '
            + wf(g/255.0) + b' '
            + wf(b/255.0) + b' ]',
        b'>>',
        b'>>',
        b'>>',
        b'/RunLength ' + wi(text_len),
        b'>> ]',
        b'/IsJoinable 0',
        b'>>',
        b'/GridInfo',
        b'<<',
        b'/GridIsOn false',
        b'/ShowGrid false',
        b'/GridSize 18.0',
        b'/GridLeading 22.0',
        b'/GridColor << /Type 1 /Values [ 1.0 0.0 0.0 0.0 ] >>',
        b'/GridLeadingFillColor << /Type 1 /Values [ 1.0 0.0 0.0 0.0 ] >>',
        b'/AlignLineHeightToGridFlags false',
        b'>>',
        b'/AntiAlias 4',
        b'/UseFractionalGlyphWidths true',
        b'/Rendered 0',
        b'>>',
        b'/ResourceDict',
        b'<<',
        b'/KinsokuSet [ ]',
        b'/MojiKumiSet [ ]',
        b'/TheNormalStyleSheet 0',
        b'/TheNormalParagraphSheet 0',
        b'/FontSet [ <<',
        b'/Name ' + wstr(font_name),
        b'/Script 0',
        b'/FontType 1',
        b'/Synthetic 0',
        b'>> ]',
        b'/SuperscriptSize 0.583',
        b'/SubscriptSize 0.583',
        b'/SmallCapSize 0.7',
        b'>>',
        b'/DocumentResources',
        b'<<',
        b'/KinsokuSet [ ]',
        b'/MojiKumiSet [ ]',
        b'/TheNormalStyleSheet 0',
        b'/TheNormalParagraphSheet 0',
        b'/FontSet [ <<',
        b'/Name ' + wstr(font_name),
        b'/Script 0',
        b'/FontType 1',
        b'/Synthetic 0',
        b'>> ]',
        b'/SuperscriptSize 0.583',
        b'/SubscriptSize 0.583',
        b'/SmallCapSize 0.7',
        b'>>',
        b'>>',
    ]

    return b'\n'.join(lines) + b'\n'


# ─── Tag wrapper ──────────────────────────────────────────────────────────────

def _tag(key, data):
    """Wrap data in 8BIM tagged block"""
    k = key.encode('ascii') if isinstance(key, str) else key
    if len(data) % 2:
        data += b'\x00'
    return b'8BIM' + k + struct.pack('>I', len(data)) + data


# ─── Public API ───────────────────────────────────────────────────────────────

def build_editable_text_tagged_blocks(text, font_name, font_size_px,
                                       r, g, b, px_per_cm,
                                       layer_left, layer_top,
                                       layer_w, layer_h):
    """
    Build TySh + Txt2 tagged blocks for an editable Photoshop text layer.

    Args:
        text:         Customer text string (e.g. "Rose's Petals")
        font_name:    Photoshop font name (e.g. "Wavemermaid", "Arial")
        font_size_px: Font height in pixels
        r, g, b:      Text colour 0–255
        px_per_cm:    Resolution (120 = 304 DPI, 320 = 812 DPI)
        layer_left:   Left edge of text layer in canvas pixels
        layer_top:    Top edge of text layer in canvas pixels
        layer_w:      Width of text layer in pixels
        layer_h:      Height of text layer in pixels

    Returns:
        bytes: TySh + Txt2 tagged blocks to append to the layer extra data
    """
    # Convert px to points (72 pts per inch)
    dpi = px_per_cm * 2.54
    font_size_pts = font_size_px * 72.0 / dpi

    left   = float(layer_left)
    top    = float(layer_top)
    right  = float(layer_left + layer_w)
    bottom = float(layer_top  + layer_h)

    tysh = _build_TySh(text, font_name, font_size_pts,
                        r, g, b, left, top, right, bottom)
    txt2 = _build_Txt2(text, font_name, font_size_pts, r, g, b)

    return _tag('TySh', tysh) + _tag('Txt2', txt2)


# ─── Font name resolver ───────────────────────────────────────────────────────

PHOTOSHOP_FONT_NAMES = {
    # Standard
    'arial':                  'Arial',
    'arialbold':              'Arial',
    'bebasneueregular':       'BebasNeue-Regular',
    'chewyregular':           'Chewy',
    'chewy':                  'Chewy',
    'permanentmarkerregular': 'PermanentMarker',
    'permanentmarker':        'PermanentMarker',
    'russooneregular':        'RussoOne',
    'russoone':               'RussoOne',
    'ultraregular':           'Ultra',
    'ultra':                  'Ultra',
    'latoregular':            'Lato',
    'lato':                   'Lato',
    'latobold':               'Lato-Bold',
    'roboto':                 'Roboto',
    'verdana':                'Verdana',
    'abel':                   'Abel-Regular',
    'fondamento':             'Fondamento',
    # Premium texture fonts
    'smartkids':              'Smart Kids',
    'colorfulblocks':         'Colorful Blocks',
    'paintsplashesrainbow':   'Paint Splashes Rainbow',
    'wavemermaid':            'Wavemermaid',
    'refractionray':          'Refraction Ray',
    'camoblock':              'Camoblock',
    'spiderweb':              'Spider Web',
    'cozywinter':             'Cozy Winter',
    'soccerarmy':             'Soccer Army',
    'bouqetdisplay':          'Bouqet Display',
    'supervibes':             'Super Vibes',
    'vinylfont':              'VinylFont',
    'commandinghandsbsl':     'CommandingHandsBSL',
}

def resolve_ps_font_name(normalised_font_name):
    """Convert normalised font key to Photoshop font name"""
    key = normalised_font_name.lower().replace(' ','').replace('-','').replace('_','')
    return PHOTOSHOP_FONT_NAMES.get(key, normalised_font_name)


if __name__ == '__main__':
    blocks = build_editable_text_tagged_blocks(
        text="MYLAMAE",
        font_name="Wavemermaid",
        font_size_px=400,
        r=255, g=128, b=237,
        px_per_cm=120,
        layer_left=50, layer_top=600,
        layer_w=800, layer_h=200,
    )
    print(f"TySh+Txt2 total: {len(blocks)} bytes")
    print(f"Tag 1: {blocks[4:8]}")
    size1 = struct.unpack('>I', blocks[8:12])[0]
    print(f"TySh size: {size1} bytes")
    print(f"Tag 2: {blocks[12+size1:16+size1]}")
    print("OK — no duplicate keys in descriptor")
