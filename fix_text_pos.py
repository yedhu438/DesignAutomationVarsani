
content = open(r'W:\VarsaniAutomation\batch_processor.py', encoding='utf-8').read()

OLD = '    top  = max(int(h * 0.70), h - bh - 60) if has_image else (h - bh) // 2'

NEW = '''    # Text position:
    #   - has_image=True  → text at TOP (above the photo), matching real t-shirt layout
    #   - has_image=False → text centred vertically (text-only orders)
    if has_image:
        top = 10   # small top margin, sits above the customer photo
    else:
        top = max(0, (h - bh) // 2)  # centred for text-only'''

if OLD in content:
    content = content.replace(OLD, NEW)
    open(r'W:\VarsaniAutomation\batch_processor.py', 'w', encoding='utf-8').write(content)
    print('SUCCESS: text placement fixed — text now goes to TOP when image present')
else:
    print('ERROR: line not found')
    idx = content.find('has_image')
    print(repr(content[idx-50:idx+200]))
