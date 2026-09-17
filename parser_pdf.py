"""PDF parser for Hygiene Kitchens factory quotes."""
import pdfplumber, re

STOP = ['total quantity','total cost','sub total','grand total','tax (',
        'transportation cost','installation cost','terms','warranty',
        'delivery','payment terms','validity','authorised','nivedhitha',
        'remarks','scope of supply','please feel','looking forward',
        '# description','thank you for reaching']

def is_stop(t): return any(s in (t or '').lower() for s in STOP)
def to_f(s):
    try: return float(str(s).replace(',','').strip())
    except: return None

def trailing_nums(text):
    tokens = text.split(); nums = []
    while tokens:
        n = to_f(tokens[-1])
        if n is not None: nums.insert(0,n); tokens.pop()
        else: break
    return ' '.join(tokens), nums

def is_hyginox_pdf(lines):
    text = ' '.join(lines[:30]).lower()
    return 'price rs' in text and 'w x h' in text

def parse_hyginox_pdf(path):
    quote_no, date_str, client, project = '', '', 'MELANGE', 'KITCHEN'
    rows = []
    STOP2 = ['total cost','sub total','grand total','gst','transportation',
             'installation cost','terms','warranty','delivery','payment',
             'validity','bank details','thanking','yours faithfully',
             'please feel','ss with ss hairline']
    def is_stop2(t): return any(s in (t or '').lower() for s in STOP2)
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            m = re.search(r'Quotation No\.\s*([\w]+)', text)
            if m: quote_no = m.group(1)
            m = re.search(r'Date:\s*([\d\-./ ]+)', text)
            if m: date_str = m.group(1).strip()
            m = re.search(r'M/s\.\s*(.+)', text)
            if m: client = m.group(1).strip()
            for table in page.extract_tables():
                for row in table:
                    if not row or len(row) < 3: continue
                    sl    = str(row[0] or '').strip()
                    desc  = str(row[1] or '').replace('\n',' ').strip()
                    wh    = str(row[-2] or '').strip()
                    price = to_f(str(row[-1] or ''))
                    if not sl or not desc or not price: continue
                    if is_stop2(desc): continue
                    if not re.match(r'^\d+[a-zA-Z]?$', sl): continue
                    is_acc = bool(re.search(r'[a-zA-Z]$', sl))
                    rows.append(dict(sl=sl, desc=desc, wh=wh,
                                     amount=price, is_acc=is_acc,
                                     w=None,d=None,h=None,qty=None,rate=None))
    return quote_no, date_str, client, project, rows

def parse_factory_pdf(path):
    lines = []
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            lines.extend((p.extract_text() or '').split('\n'))

    if is_hyginox_pdf(lines):
        return parse_hyginox_pdf(path)

    quote_no, date_str, client, project = '', '', 'MELANGE', 'KITCHEN'
    for l in lines:
        m = re.search(r'Quote\s+([\w\-]+)', l)
        if m: quote_no = m.group(1)
        m = re.search(r'Date\s+([\d\-./]+)', l)
        if m: date_str = m.group(1)
        m = re.search(r'To:\s*(.+)', l, re.I)
        if m: client = m.group(1).strip()
        m = re.search(r'Project Name[:\s]+(.+)', l, re.I)
        if m: project = m.group(1).strip()

    has_text = any(l.strip() for l in lines)
    if not has_text:
        try:
            import pymupdf, pytesseract, io
            from PIL import Image as PILImage
            doc = pymupdf.open(path)
            lines = []
            for page in doc:
                pix = page.get_pixmap(matrix=pymupdf.Matrix(2.5,2.5))
                img = PILImage.open(io.BytesIO(pix.tobytes('png')))
                lines.extend(pytesseract.image_to_string(img, config='--psm 6').split('\n'))
            doc.close()
        except: pass

    items, cur = [], None
    for raw in lines:
        line = raw.strip()
        if not line or is_stop(line): continue
        m = re.match(r'^(\d+[a-zA-Z]?)\s+(.+)', line)
        if m:
            if cur: items.append(cur)
            sl = m.group(1)
            desc, nums = trailing_nums(m.group(2))
            is_acc = bool(re.search(r'[a-zA-Z]$', sl))
            cur = dict(sl=sl, desc=desc, nums=nums, is_acc=is_acc)
        elif cur and re.match(r'^[A-Za-z(]', line):
            cur['desc'] += ' ' + line
    if cur: items.append(cur)

    for r in items:
        nums = r.pop('nums')
        r.update(w=None,d=None,h=None,qty=None,rate=None,amount=None,wh=None)
        if not r['is_acc'] and len(nums) >= 6:
            r['w'],r['d'],r['h'],r['qty'],r['rate'],r['amount'] = \
                nums[-6],nums[-5],nums[-4],nums[-3],nums[-2],nums[-1]
            r['wh'] = f"{int(r['w'])} x {int(r['h'])}" if r['w'] and r['h'] else ''
        elif len(nums) >= 3:
            r['qty'],r['rate'],r['amount'] = nums[-3],nums[-2],nums[-1]
            q = r['qty']
            r['wh'] = f"{int(q)} Nos" if q and q==int(q) else ''
        elif len(nums) == 2:
            r['rate'],r['amount'] = nums[0],nums[1]

    return quote_no, date_str, client, project, items
