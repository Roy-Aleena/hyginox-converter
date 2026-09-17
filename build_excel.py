"""
Excel builder — Solana Eco Homes branding
Columns: # | Description | W | D | H | Qty | Rate | Amount
10% applied to Rate; Amount = new Rate × Qty
"""
import openpyxl, re, tempfile
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from pathlib import Path

MARKUP = 0.10

TERMS = [
    ('b','Terms and Conditions'),
    ('n',' Your order is accepted subject to the following terms and conditions:'),
    ('n','1. Warranty: Standard warranty is applicable for lifetime against any manufacturing defects and rust, subject to proper maintenance and care of the kitchen unit by the user. Use of any kind of abrasives, acids or chemicals like chlorine, bleaching powder, etc. are harmful to stainless steel. The warranty stands void in instance of damage due to the above or any other un-recommended products for cleaning or any other purpose by the user.'),
    ('n','2. Transportation and Installation: The price quoted is inclusive of transportation and Installation charges. However, additional cost incurred towards civil works such as granite works, plumbing, electrical works, demolitions of any existing structures, etc. are to be borne by customer. Additional charges will be applicable for any kind of modifications or additions after acceptance of the final drawing and placing the order.'),
    ('n','3. Taxes and duties: The price quoted above is inclusive of all applicable govt. taxes and duties.'),
    ('n','4. Delivery Schedule: The minimum delivery period will be between 30 to 40 days from the date of placing your confirmed order. Under any Circumstances, the kitchen manufactured against your confirmed order and drawing will not be taken back, altered or exchanged.'),
    ('n','5. Payment Terms: Advance of 50% of the bill value is payable while placing the order: Before dispatch of material from factory, 40%: balance 10% after installation. All payments to be made by cheque/D.D. drawn in favour of Solana Eco Homes, payable at Bangalore.'),
    ('n','6. Validity of offer: The price quoted is valid for a period of 30 days only.'),
    ('n','7. Drawings and sketches: The drawings and sketches supplied along with this quote are indicative only. The original product and its colour and fittings may vary from the drawing.'),
    ('_',''),
    ('b','Our Bank Details:'),
    ('n','Bank: South Indian Bank'),
    ('n','Branch: Carmelaram'),
    ('n','Account Name: Solana Eco Homes'),
    ('n','Account Number: 0883073000000090'),
    ('n','IFSC Code: SIBL0000883'),
    ('_',''),
    ('n','Please feel free to contact us for any further details or clarifications. Looking forward to your patronage at all times.'),
    ('_',''),
    ('n','Thanking you'),
    ('n','Yours Faithfully'),
    ('n','For Solana Eco Homes.'),
    ('_',''),
    ('n','____________________________'),
    ('n','Roy Joseph'),
    ('n','Authorised Signatory'),
]

def build(items, quote_no, date_str, client, project, logo_path=None):
    # ── Borders ───────────────────────────────────────────────────────────────
    THIN  = Side(style='thin',   color='000000')
    THICK = Side(style='medium', color='000000')
    NO    = Side(style=None)
    def bdr(t=NO,b=NO,l=NO,r=NO): return Border(top=t,bottom=b,left=l,right=r)
    def all_thin(): return bdr(THIN,THIN,THIN,THIN)

    # ── Fills ─────────────────────────────────────────────────────────────────
    DARK_FILL  = PatternFill('solid', fgColor='1F2D3D')
    WHITE_FILL = PatternFill('solid', fgColor='FFFFFF')
    LGRAY_FILL = PatternFill('solid', fgColor='F9F9F9')
    GREEN_FILL = PatternFill('solid', fgColor='D5E8D4')

    # ── Fonts ─────────────────────────────────────────────────────────────────
    def F(bold=False, size=10, color='000000', italic=False):
        return Font(name='Calibri', bold=bold, size=size, color=color, italic=italic)

    # ── Alignments ────────────────────────────────────────────────────────────
    AC  = Alignment(horizontal='center', vertical='center')
    AL  = Alignment(horizontal='left',   vertical='center')
    AR  = Alignment(horizontal='right',  vertical='center')
    ALW = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    ATW = Alignment(horizontal='left',   vertical='top',    wrap_text=True)

    # ── Workbook ──────────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Hyginox Quote'
    ws.sheet_view.showGridLines = False

    # Columns: A(#) B(Description-wide) C(W) D(D) E(H) F(Qty) G(Rate) H(Amount)
    for col, w in [('A',8),('B',52),('C',7),('D',7),('E',7),
                   ('F',7),('G',14),('H',14)]:
        ws.column_dimensions[col].width = w

    def rh(r,h): ws.row_dimensions[r].height = h

    # ── ROWS 1-5: Company header ──────────────────────────────────────────────
    rh(1,20); rh(2,14); rh(3,14); rh(4,14); rh(5,14)

    ws.merge_cells('A1:E1')
    ws['A1'] = 'Solana Eco Homes'
    ws['A1'].font = F(bold=True, size=20)
    ws['A1'].alignment = AL

    for row, txt in [
        (2,'#7, Sampath Building, Opp. HDFC Bank,  Dommasandra,'),
        (3,'Bangalore 560125.'),
        (4,'Phone - 8861314554.'),
        (5,'Email:  solanabangalore@gmail.com'),
    ]:
        ws.merge_cells(f'A{row}:E{row}')
        ws[f'A{row}'] = txt
        ws[f'A{row}'].font = F(bold=True, size=9)
        ws[f'A{row}'].alignment = AL

    # Outer border around header block A1:H5
    for row in range(1,6):
        for col in range(1,9):
            t = THIN if row==1 else NO
            b = THIN if row==5 else NO
            l = THIN if col==1 else NO
            r = THIN if col==8 else NO
            if any([t!=NO,b!=NO,l!=NO,r!=NO]):
                ws.cell(row,col).border = bdr(t,b,l,r)

    # Logo (F1:H5)
    if logo_path and Path(logo_path).exists():
        try:
            from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
            img = XLImage(str(logo_path))
            img.width=200; img.height=88
            from_m = AnchorMarker(col=5, colOff=114300, row=0, rowOff=114300)
            to_m   = AnchorMarker(col=7, colOff=914400, row=4, rowOff=914400)
            anchor = TwoCellAnchor(editAs='twoCell')
            anchor._from = from_m; anchor.to = to_m
            img.anchor = anchor
            ws.add_image(img)
        except: pass

    # ── ROW 6: Title bar ──────────────────────────────────────────────────────
    rh(6,20)
    ws.merge_cells('A6:H6')
    ws['A6'] = 'Quotation Hyginox Signature Stainless Steel Modular Kitchen'
    ws['A6'].font = F(bold=True, size=11)
    ws['A6'].alignment = AC
    for col in range(1,9):
        ws.cell(6,col).border = bdr(THIN,THIN, THIN if col==1 else NO, THIN if col==8 else NO)

    # ── ROW 7: Intro ──────────────────────────────────────────────────────────
    rh(7,22)
    ws.merge_cells('A7:H7')
    ws['A7'] = ('Thank you for your valuable enquiry and expressing interest in Hyginox Steel '
                'Modular Kitchen. Furnished below our best offer for the same.')
    ws['A7'].font = F(size=8)
    ws['A7'].alignment = ALW
    for col in range(1,9):
        ws.cell(7,col).border = bdr(NO,THIN, THIN if col==1 else NO, THIN if col==8 else NO)

    # ── ROWS 8-9: Client block ────────────────────────────────────────────────
    rh(8,16); rh(9,16)

    def meta(r, c1, c2, val, bold=False):
        ws.merge_cells(f'{c1}{r}:{c2}{r}')
        c = ws[f'{c1}{r}']
        c.value=val; c.font=F(bold=bold,size=10); c.alignment=AL
        from openpyxl.utils import column_index_from_string
        for col in range(column_index_from_string(c1), column_index_from_string(c2)+1):
            ws.cell(r,col).border = bdr(THIN,THIN,
                THIN if col==column_index_from_string(c1) else NO,
                THIN if col==column_index_from_string(c2) else NO)

    meta(8,'A','C', f'M/s. {client}', bold=True)
    meta(8,'D','E', 'GSTN 29AFVPJ8463N1Z9')
    meta(8,'F','H', f'Quotation No. {quote_no}', bold=True)
    meta(9,'A','C', f'Project Name: {project}')
    meta(9,'D','E', f'Date: {date_str}')
    meta(9,'F','H', 'Mode of Payment  RTGS/ Cheque')

    # ── ROW 10: Table header ──────────────────────────────────────────────────
    rh(10,24)
    headers = [('A','#'),('B','Description'),('C','W'),('D','D'),
               ('E','H'),('F','Qty'),('G','Rate'),('H','Amount')]
    for col_l, txt in headers:
        c = ws[f'{col_l}10']
        c.value=txt; c.fill=DARK_FILL
        c.font=F(bold=True, size=10, color='FFFFFF')
        c.alignment=AC; c.border=all_thin()

    # ── DATA ROWS ─────────────────────────────────────────────────────────────
    ROW = 11
    total_mod = total_acc = 0.0

    def fmt_n(v):
        if v is None: return ''
        return str(int(v)) if float(v)==int(float(v)) else f'{float(v):.1f}'

    for r in items:
        is_acc   = r['is_acc']
        orig_rate = r.get('rate') or 0.0
        qty       = r.get('qty')  or (1.0 if is_acc else 1.0)
        new_rate  = round(orig_rate * (1 + MARKUP), 2)
        new_amt   = round(new_rate * qty, 2)
        fill      = LGRAY_FILL if ROW % 2 == 0 else WHITE_FILL

        # Calculate row height from description length
        chars_per_line = 70
        lines = max(1, -(-len(r['desc']) // chars_per_line))
        rh(ROW, max(18 if is_acc else 24, lines * (13 if is_acc else 15)))

        # Sl. No.
        sl_val = int(r['sl']) if r['sl'].isdigit() else r['sl']
        ws.cell(ROW,1,sl_val).font      = F(size=9, bold=not is_acc, italic=is_acc)
        ws.cell(ROW,1).alignment = AC; ws.cell(ROW,1).fill=fill; ws.cell(ROW,1).border=all_thin()

        # Description
        ws.cell(ROW,2,r['desc']).font      = F(size=8.5, italic=is_acc)
        ws.cell(ROW,2).alignment = ALW; ws.cell(ROW,2).fill=fill; ws.cell(ROW,2).border=all_thin()

        # W, D, H — store as numbers to avoid green triangle warnings
        for col, val in [(3, r.get('w')), (4, r.get('d')), (5, r.get('h'))]:
            num_val = int(val) if val and float(val)==int(float(val)) else (float(val) if val else None)
            ws.cell(ROW,col, num_val)
            ws.cell(ROW,col).font=F(size=9); ws.cell(ROW,col).alignment=AC
            ws.cell(ROW,col).fill=fill; ws.cell(ROW,col).border=all_thin()

        # Qty — store as number
        qty_val = int(qty) if qty and float(qty)==int(float(qty)) else (float(qty) if qty else None)
        ws.cell(ROW,6, qty_val)
        ws.cell(ROW,6).font=F(size=9); ws.cell(ROW,6).alignment=AC
        ws.cell(ROW,6).fill=fill; ws.cell(ROW,6).border=all_thin()

        # Rate (10% markup)
        ws.cell(ROW,7, new_rate)
        ws.cell(ROW,7).font=F(size=9); ws.cell(ROW,7).alignment=AR
        ws.cell(ROW,7).number_format='#,##0.00'
        ws.cell(ROW,7).fill=fill; ws.cell(ROW,7).border=all_thin()

        # Amount = new Rate × Qty
        ws.cell(ROW,8, new_amt)
        ws.cell(ROW,8).font=F(size=9); ws.cell(ROW,8).alignment=AR
        ws.cell(ROW,8).number_format='#,##0.00'
        ws.cell(ROW,8).fill=fill; ws.cell(ROW,8).border=all_thin()

        if is_acc: total_acc += new_amt
        else:      total_mod += new_amt
        ROW += 1

    # ── TOTALS ────────────────────────────────────────────────────────────────
    INSTALL   = 25000.0
    sub_total = total_mod + total_acc + INSTALL
    gst_amt   = round(sub_total * 0.18, 2)
    grand     = round(sub_total + gst_amt, 2)

    ROW += 1
    for label, val, bold, fill in [
        ('Total cost for Modules',               total_mod,  False, WHITE_FILL),
        ('Total cost for Accessories',            total_acc,  False, WHITE_FILL),
        ('Installation cost',                    INSTALL,    False, WHITE_FILL),
        ('Transportation (Actual not included)', 0.0,        False, WHITE_FILL),
        ('Sub Total',                            sub_total,  True,  LGRAY_FILL),
        ('GST 18%',                              gst_amt,    False, WHITE_FILL),
        ('Grand Total Hyginox Stainless Steel Kitchen', grand, True, GREEN_FILL),
    ]:
        rh(ROW, 16)
        ws.merge_cells(f'A{ROW}:G{ROW}')
        for col in range(1,8):
            ws.cell(ROW,col).fill=fill
            ws.cell(ROW,col).border=bdr(THIN,THIN,
                THIN if col==1 else NO, THIN if col==7 else NO)
        ws[f'A{ROW}'].value=label
        ws[f'A{ROW}'].font=F(bold=bold,size=10)
        ws[f'A{ROW}'].alignment=AL
        c = ws.cell(ROW,8,val)
        c.font=F(bold=bold,size=10); c.alignment=AR
        c.number_format='#,##0.00'; c.fill=fill; c.border=all_thin()
        ROW += 1

    # ── SS note ───────────────────────────────────────────────────────────────
    ROW += 1
    ws.merge_cells(f'A{ROW}:H{ROW}')
    ws.cell(ROW,1,'SS WITH SS HAIRLINE FINISH DOORS').font=F(bold=True,size=9)
    rh(ROW,14); ROW += 2

    # ── Terms ─────────────────────────────────────────────────────────────────
    for kind, text in TERMS:
        ws.merge_cells(f'A{ROW}:H{ROW}')
        c = ws.cell(ROW,1,text)
        c.font = F(bold=(kind=='b'), size=9)
        c.alignment = ATW
        if kind=='_': rh(ROW,7)
        else:
            lines = max(1,-(-len(text)//130))
            rh(ROW, max(14, lines*16))
        ROW += 1

    # ── Page setup ────────────────────────────────────────────────────────────
    ws.freeze_panes = 'A11'
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.print_area = f'A1:H{ROW}'

    out = tempfile.NamedTemporaryFile(delete=False, suffix='_HYGINOX.xlsx')
    wb.save(out.name)
    return out.name, grand
