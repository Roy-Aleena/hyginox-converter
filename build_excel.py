"""
Excel builder that matches Solana PDF exactly.
Columns: A(Sl.No) | B:G merged(Description) | H(W x H) | I(Price RS.)
Full borders on every cell, matching the PDF layout.
"""
import openpyxl, re, tempfile
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from pathlib import Path

MARKUP = 0.10

TERMS = [
    ('b','Terms and Conditions'),
    ('n',' Your order is accepted subject to the following terms and conditions:'),
    ('n','1. Warranty: Standard warranty is applicable for lifetime against any manufacturing defects and rust, subject to proper maintenance and care of the kitchen unit by the user. Use of any kind of abrasives, acids or chemicals like chlorine, bleaching powder, etc. are harmful to stainless steel. The warranty stands void in instance of damage due to the above or any other un-recommended products for cleaning or any other purpose by the user.'),
    ('n','2. Transportation and Installation: The price quoted is inclusive of transportation and Installation charges. However, additional cost incurred towards civil works such as granite works, plumbing, electrical works, demolitions of any existing structures, etc. are to be borne by customer. Additional charges will be applicable for any kind of modifications or additions after acceptance of the final drawing and placing the order.'),
    ('n','3. Taxes and duties: The price quoted above is inclusive of all applicable govt. taxes and duties.'),
    ('n','4. Delivery Schedule: The minimum delivery period will be between 30 to 40 days from the date of placing your confirmed order. Under any Circumstances, the kitchen manufactured against your confirmed order and drawing will not be taken back, altered or exchanged.'),
    ('n','5. Payment Terms:  Advance of 50% of the bill value is payable while placing the order: Before dispatch of material from factory, 40%: balance 10% after installation. All payments to be made by cheque/D.D. drawn in favour of Solana Eco Homes, payable at Bangalore.'),
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
    # ── Sides & borders ───────────────────────────────────────────────────────
    THIN   = Side(style='thin',   color='000000')
    THICK  = Side(style='medium', color='000000')
    NO     = Side(style=None)

    def bdr(t=NO,b=NO,l=NO,r=NO):
        return Border(top=t, bottom=b, left=l, right=r)
    def all_thin():  return bdr(THIN,THIN,THIN,THIN)
    def all_thick(): return bdr(THICK,THICK,THICK,THICK)
    def box(thick=False):
        s = THICK if thick else THIN
        return bdr(s,s,s,s)

    # Fills
    DARK_FILL  = PatternFill('solid', fgColor='1F2D3D')   # table header
    WHITE_FILL = PatternFill('solid', fgColor='FFFFFF')
    LGRAY_FILL = PatternFill('solid', fgColor='F9F9F9')
    GREEN_FILL = PatternFill('solid', fgColor='D5E8D4')
    LTBLUE     = PatternFill('solid', fgColor='EBF3FB')   # alt rows

    # Fonts
    def F(bold=False, size=10, color='000000', italic=False, name='Calibri'):
        return Font(name=name, bold=bold, size=size, color=color, italic=italic)

    # Alignments
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

    # Column widths: A | B:G (description) | H (W x H) | I (Price)
    widths = {'A':9, 'B':8, 'C':10, 'D':12, 'E':20, 'F':12, 'G':10,
              'H':16, 'I':15}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    def rh(r, h): ws.row_dimensions[r].height = h

    # Helper: write merged cell with style
    def mwrite(cell_range, value, font=None, fill=None, align=None, border=None):
        ws.merge_cells(cell_range)
        cr = openpyxl.utils.cell.range_boundaries(cell_range)
        # cr = (min_col, min_row, max_col, max_row)
        c = ws.cell(cr[1], cr[0], value)
        if font:   c.font      = font
        if fill:   c.fill      = fill
        if align:  c.alignment = align
        if border: c.border    = border
        # Apply border to all cells in range for proper display
        if border:
            from openpyxl.utils import get_column_letter
            for row in range(cr[1], cr[3]+1):
                for col in range(cr[0], cr[2]+1):
                    cell = ws.cell(row, col)
                    # Determine which sides to show
                    t = border.top    if row == cr[1] else NO
                    b = border.bottom if row == cr[3] else NO
                    l = border.left   if col == cr[0] else NO
                    r = border.right  if col == cr[2] else NO
                    cell.border = bdr(t,b,l,r)
        return c

    # ── ROW 1-5: Company header ───────────────────────────────────────────────
    rh(1,20); rh(2,14); rh(3,14); rh(4,14); rh(5,14)

    # Company name + address
    ws.merge_cells('A1:G1')
    ws['A1'] = 'Solana Eco Homes'
    ws['A1'].font = F(bold=True, size=20, name='Arial')
    ws['A1'].alignment = AL
    # Address lines
    for r, txt in [(2,'#7, Sampath Building, Opp. HDFC Bank,  Dommasandra,'),
                   (3,'Bangalore 560125.'),
                   (4,'Phone - 8861314554.'),
                   (5,'Email:  solanabangalore@gmail.com')]:
        ws.merge_cells(f'A{r}:G{r}')
        ws[f'A{r}'] = txt
        ws[f'A{r}'].font = F(bold=True, size=9)
        ws[f'A{r}'].alignment = AL

    # Outer border around header block
    for row in range(1,6):
        for col in range(1, 10):
            c = ws.cell(row, col)
            t = THIN if row == 1 else NO
            b = THIN if row == 5 else NO
            l = THIN if col == 1 else NO
            r = THIN if col == 9 else NO
            if any([t!=NO, b!=NO, l!=NO, r!=NO]):
                c.border = bdr(t,b,l,r)

    # Logo — top right, spanning H1:I5
    ws.column_dimensions['H'].width = 20
    ws.column_dimensions['I'].width = 20
    rh(1,20); rh(2,16); rh(3,16); rh(4,16); rh(5,16)

    if logo_path and Path(logo_path).exists():
        try:
            from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
            img = XLImage(str(logo_path))
            # Logo aspect ratio ≈ 764:216 ≈ 3.5:1  →  width 220, height 62
            img.width  = 220
            img.height = 62
            # Anchor: from H1 to I5 (col index 0-based: H=7, I=8; row 0-based)
            from_m = AnchorMarker(col=7, colOff=114300, row=0, rowOff=114300)
            to_m   = AnchorMarker(col=8, colOff=914400, row=4, rowOff=914400)
            anchor = TwoCellAnchor(editAs='twoCell')
            anchor._from = from_m
            anchor.to    = to_m
            img.anchor   = anchor
            ws.add_image(img)
        except Exception as exc:
            pass  # silent fail, logo not critical

    # ── ROW 6: Title bar ──────────────────────────────────────────────────────
    rh(6, 20)
    ws.merge_cells('A6:I6')
    ws['A6'] = 'Quotation Hyginox Signature Stainless Steel Modular Kitchen'
    ws['A6'].font      = F(bold=True, size=11)
    ws['A6'].alignment = AC
    ws['A6'].fill      = WHITE_FILL
    for col in range(1, 10):
        t = THIN if True else NO
        b = THIN
        l = THIN if col==1 else NO
        r = THIN if col==9 else NO
        ws.cell(6, col).border = bdr(THIN, THIN, l, r)

    # ── ROW 7: Intro ──────────────────────────────────────────────────────────
    rh(7, 22)
    ws.merge_cells('A7:I7')
    ws['A7'] = ('Thank you for your valuable enquiry and expressing inerest in Hyginox Steel '
                'Modular Kitchen. Furnished below our best offer for the same.')
    ws['A7'].font = F(size=8)
    ws['A7'].alignment = ALW
    for col in range(1,10):
        ws.cell(7,col).border = bdr(NO, THIN, THIN if col==1 else NO, THIN if col==9 else NO)

    # ── ROW 8-9: Client block ─────────────────────────────────────────────────
    rh(8, 16); rh(9, 16)

    # Row 8: M/s | GSTN | Quotation No.
    ws.merge_cells('A8:B8')
    ws['A8'] = f'M/s. {client}'
    ws['A8'].font = F(bold=True, size=10)
    ws['A8'].alignment = AL
    for col in range(1,3):
        ws.cell(8,col).border = bdr(THIN,NO,THIN if col==1 else NO, NO)

    ws.merge_cells('C8:E8')
    ws['C8'] = 'GSTN 29AFVPJ8463N1Z9'
    ws['C8'].font = F(size=9); ws['C8'].alignment = AL
    for col in range(3,6):
        ws.cell(8,col).border = bdr(THIN,NO, THIN if col==3 else NO, NO)

    ws.merge_cells('F8:I8')
    ws['F8'] = f'Quotation No. {quote_no}'
    ws['F8'].font = F(size=9); ws['F8'].alignment = AL
    for col in range(6,10):
        ws.cell(8,col).border = bdr(THIN,NO, THIN if col==6 else NO, THIN if col==9 else NO)

    # Row 9: Infinity | Date | Mode of Payment
    ws.merge_cells('A9:B9')
    ws['A9'] = 'Infinity'
    ws['A9'].font = F(bold=True, size=10); ws['A9'].alignment = AL
    for col in range(1,3):
        ws.cell(9,col).border = bdr(NO,THIN,THIN if col==1 else NO, NO)

    ws.merge_cells('C9:E9')
    ws['C9'] = f'Date: {date_str}'
    ws['C9'].font = F(size=9); ws['C9'].alignment = AL
    for col in range(3,6):
        ws.cell(9,col).border = bdr(NO,THIN, THIN if col==3 else NO, NO)

    ws.merge_cells('F9:I9')
    ws['F9'] = 'Mode of Payment  RTGS/ Cheque'
    ws['F9'].font = F(size=9); ws['F9'].alignment = AL
    for col in range(6,10):
        ws.cell(9,col).border = bdr(NO,THIN, THIN if col==6 else NO, THIN if col==9 else NO)

    # ── ROW 10: Table header ──────────────────────────────────────────────────
    rh(10, 26)
    for col in range(1,10):
        ws.cell(10,col).fill = DARK_FILL
        ws.cell(10,col).border = all_thin()

    ws['A10'] = 'Sl. No.'
    ws['A10'].font = F(bold=True, size=11, color='FFFFFF')
    ws['A10'].alignment = AC

    ws.merge_cells('B10:G10')
    ws['B10'] = 'Description of Cabinets'
    ws['B10'].font = F(bold=True, size=11, color='FFFFFF')
    ws['B10'].alignment = AC
    ws['B10'].fill = DARK_FILL

    ws['H10'] = 'W x H'
    ws['H10'].font = F(bold=True, size=11, color='FFFFFF')
    ws['H10'].alignment = AC

    ws['I10'] = 'Price RS.'
    ws['I10'].font = F(bold=True, size=11, color='FFFFFF')
    ws['I10'].alignment = AC

    # Apply fill to merged header cells
    for col in range(2,8):
        ws.cell(10,col).fill = DARK_FILL
        ws.cell(10,col).border = bdr(THIN,THIN, THIN if col==2 else NO, THIN if col==7 else NO)

    # ── DATA ROWS ─────────────────────────────────────────────────────────────
    ROW = 11
    total_mod = total_acc = 0.0

    for idx, r in enumerate(items):
        factory_amt = r['amount'] or 0.0
        sol_price   = round(factory_amt * (1 + MARKUP), 2)
        wh          = r.get('wh') or ''
        is_acc      = r['is_acc']
        fill        = LGRAY_FILL if idx % 2 == 0 else WHITE_FILL

        # Calculate row height based on description length
        # Description col B:G combined width ≈ 72 chars per line
        chars_per_line = 95
        desc_len = len(r['desc'])
        lines_needed = max(1, -(-desc_len // chars_per_line))  # ceiling div
        if is_acc:
            row_h = max(18, lines_needed * 14)
        else:
            row_h = max(28, lines_needed * 16)

        # A: Sl. No. — store pure numbers as int to avoid "stored as text" warning
        sl_val = int(r['sl']) if r['sl'].isdigit() else r['sl']
        ws.cell(ROW,1, sl_val)
        ws.cell(ROW,1).font      = F(size=9, bold=not is_acc)
        ws.cell(ROW,1).alignment = AC
        ws.cell(ROW,1).fill      = fill
        ws.cell(ROW,1).border    = all_thin()

        # B:G Description (merged)
        ws.merge_cells(f'B{ROW}:G{ROW}')
        ws.cell(ROW,2, r['desc'])
        ws.cell(ROW,2).font      = F(size=8.5, italic=is_acc)
        ws.cell(ROW,2).alignment = ALW
        ws.cell(ROW,2).fill      = fill
        for col in range(2,8):
            ws.cell(ROW,col).fill   = fill
            ws.cell(ROW,col).border = bdr(
                THIN, THIN,
                THIN if col==2 else NO,
                THIN if col==7 else NO
            )

        # H: W x H
        ws.cell(ROW,8, wh)
        ws.cell(ROW,8).font      = F(size=9)
        ws.cell(ROW,8).alignment = AC
        ws.cell(ROW,8).fill      = fill
        ws.cell(ROW,8).border    = all_thin()

        # I: Price RS.
        ws.cell(ROW,9, sol_price)
        ws.cell(ROW,9).font          = F(size=9)
        ws.cell(ROW,9).alignment     = AR
        ws.cell(ROW,9).number_format = '#,##0.00'
        ws.cell(ROW,9).fill          = fill
        ws.cell(ROW,9).border        = all_thin()

        rh(ROW, row_h)
        if is_acc: total_acc += sol_price
        else:      total_mod += sol_price
        ROW += 1

    # ── TOTALS SECTION ────────────────────────────────────────────────────────
    INSTALL   = 25000.0
    sub_total = total_mod + total_acc + INSTALL
    gst_amt   = round(sub_total * 0.18, 2)
    grand     = round(sub_total + gst_amt, 2)

    tot_rows = [
        ('Total cost for Modules',               total_mod,  False, WHITE_FILL),
        ('Total cost for Accessories',            total_acc,  False, WHITE_FILL),
        ('Installation cost',                    INSTALL,    False, WHITE_FILL),
        ('Transportation (Actual not included)', 0.0,        False, WHITE_FILL),
        ('Sub Total',                            sub_total,  True,  LGRAY_FILL),
        ('GST 18%',                              gst_amt,    False, WHITE_FILL),
        ('Grand Total Hyginox Stainless Steel Kitchen', grand, True, GREEN_FILL),
    ]
    for label, val, bold, fill in tot_rows:
        rh(ROW, 16)
        ws.merge_cells(f'A{ROW}:H{ROW}')
        ws.cell(ROW,1, label)
        ws.cell(ROW,1).font      = F(bold=bold, size=10)
        ws.cell(ROW,1).alignment = AL
        ws.cell(ROW,1).fill      = fill
        for col in range(1,9):
            ws.cell(ROW,col).fill   = fill
            ws.cell(ROW,col).border = bdr(
                THIN, THIN,
                THIN if col==1 else NO,
                THIN if col==8 else NO
            )
        ws.cell(ROW,9, val)
        ws.cell(ROW,9).font          = F(bold=bold, size=10)
        ws.cell(ROW,9).alignment     = AR
        ws.cell(ROW,9).number_format = '#,##0.00'
        ws.cell(ROW,9).fill          = fill
        ws.cell(ROW,9).border        = all_thin()
        ROW += 1

    # SS note
    ROW += 1
    ws.merge_cells(f'A{ROW}:I{ROW}')
    ws.cell(ROW,1,'SS WITH SS HAIRLINE FINISH DOORS')
    ws.cell(ROW,1).font = F(bold=True, size=9)
    rh(ROW,14); ROW += 2

    # ── TERMS ─────────────────────────────────────────────────────────────────
    for kind, text in TERMS:
        ws.merge_cells(f'A{ROW}:I{ROW}')
        ws.cell(ROW,1,text)
        ws.cell(ROW,1).font      = F(bold=(kind=='b'), size=9)
        ws.cell(ROW,1).alignment = ATW
        if kind == '_':
            rh(ROW, 8)
        else:
            # Calculate lines needed at ~130 chars per line in merged col A:I
            chars_per_line = 130
            lines = max(1, -(-len(text) // chars_per_line))
            rh(ROW, max(16, lines * 16))
        ROW += 1

    # ── Page setup ────────────────────────────────────────────────────────────
    ws.freeze_panes = 'A11'
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.print_area = f'A1:I{ROW}'

    out = tempfile.NamedTemporaryFile(delete=False, suffix='_HYGINOX.xlsx')
    wb.save(out.name)
    return out.name, grand

