#!/usr/bin/env python3
"""Generate 5 clearly fictional sample invoice PDFs with distinct layouts.

All companies, addresses and amounts are invented sample data.
Totals are computed with Decimal so the extractor can prove net + VAT = total.
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

from reportlab.lib.colors import Color, HexColor, white, black
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "input"

FONT_DIR_SANS = Path("/usr/share/fonts/truetype/sand-box/custom/Source Sans Pro")
FONT_DIR_SERIF = Path("/usr/share/fonts/truetype/sand-box/google/IBM Plex Serif")

pdfmetrics.registerFont(TTFont("Sans", str(FONT_DIR_SANS / "SourceSansPro-Regular.ttf")))
pdfmetrics.registerFont(TTFont("SansIt", str(FONT_DIR_SANS / "SourceSansPro-It.ttf")))
pdfmetrics.registerFont(TTFont("SansSm", str(FONT_DIR_SANS / "SourceSansPro-Semibold.ttf")))
pdfmetrics.registerFont(TTFont("SansBd", str(FONT_DIR_SANS / "SourceSansPro-Bold.ttf")))
pdfmetrics.registerFont(TTFont("SansLt", str(FONT_DIR_SANS / "SourceSansPro-Light.ttf")))
pdfmetrics.registerFont(TTFont("Serif", str(FONT_DIR_SERIF / "IBMPlexSerif-Regular.ttf")))
pdfmetrics.registerFont(TTFont("SerifBd", str(FONT_DIR_SERIF / "IBMPlexSerif-Bold.ttf")))
pdfmetrics.registerFont(TTFont("SerifMd", str(FONT_DIR_SERIF / "IBMPlexSerif-Medium.ttf")))
pdfmetrics.registerFont(TTFont("SerifIt", str(FONT_DIR_SERIF / "IBMPlexSerif-Italic.ttf")))

# Compact invoice page so a full-page render stays readable on a gig cover.
PAGE = (612, 750)

SAMPLE_FOOTER = "SAMPLE DATA — fictional companies and amounts — for demonstration only."


def D(v) -> Decimal:
    return Decimal(str(v)).quantize(Decimal("0.01"))


def money(value: Decimal, currency: str, symbol: bool = True) -> str:
    if currency == "JPY":
        n = int(value)
        return f"¥{n:,}" if symbol else f"{n:,}"
    q = value.quantize(Decimal("0.01"))
    body = f"{q:,.2f}"
    if not symbol:
        return body
    return {
        "USD": f"${body}",
        "EUR": f"€{body}",
        "GBP": f"£{body}",
        "AUD": f"A${body}",
    }[currency]


def compute(items: list[tuple[str, Decimal, Decimal]], tax_rate: Decimal, currency: str):
    rows = []
    net = Decimal("0")
    quant = Decimal("1") if currency == "JPY" else Decimal("0.01")
    for desc, qty, unit in items:
        amount = (qty * unit).quantize(quant)
        rows.append((desc, qty, unit.quantize(quant), amount))
        net += amount
    net = net.quantize(quant)
    vat = (net * tax_rate).quantize(quant)
    total = (net + vat).quantize(quant)
    return rows, net, vat, total


def sample_banner(c: canvas.Canvas, w: int, h: int, bg: Color, fg: Color, y_top: int | None = None):
    """Thin top strip marking the document as sample data."""
    band_h = 18
    top = h if y_top is None else y_top
    c.setFillColor(bg)
    c.rect(0, top - band_h, w, band_h, fill=1, stroke=0)
    c.setFillColor(fg)
    c.setFont("Sans", 8)
    c.drawCentredString(w / 2, top - 13, "SAMPLE DATA  ·  FICTIONAL COMPANY  ·  FOR DEMONSTRATION ONLY")


def sample_footer(c: canvas.Canvas, w: int, color: Color, y: int = 22):
    c.setFillColor(color)
    c.setFont("SansIt", 8)
    c.drawCentredString(w / 2, y, SAMPLE_FOOTER)


# ---------------------------------------------------------------------------
# 1. Northwind Supplies LLC  —  USD  —  dark header, gold rule (US office)
# ---------------------------------------------------------------------------
def draw_northwind(path: Path) -> dict:
    vendor = "Northwind Supplies LLC"
    invoice_no = "INV-1042"
    date = "12 Oct 2026"
    currency = "USD"
    tax_rate = Decimal("0.10")
    items = [
        ("A4 copy paper, 10 reams", Decimal("12"), Decimal("5.00")),
        ("Ballpoint pens, box of 20", Decimal("8"), Decimal("6.25")),
        ("LED desk lamp, 12W", Decimal("3"), Decimal("24.00")),
        ("Filing folders, pack of 50", Decimal("6"), Decimal("7.50")),
        ("Packing and delivery", Decimal("1"), Decimal("15.00")),
    ]
    rows, net, vat, total = compute(items, tax_rate, currency)

    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle(f"Sample Invoice {invoice_no} — {vendor} (fictional)")
    c.setAuthor("Sample invoice")
    c.setSubject("Fictional sample invoice for a public demo")

    dark = HexColor("#111318")
    gold = HexColor("#FACC15")
    ink = HexColor("#111318")
    mute = HexColor("#4B5563")
    rule = HexColor("#D6DBE3")
    head_bg = HexColor("#EEF1F5")
    alt = HexColor("#F7F8FA")
    total_bg = HexColor("#FFF8D6")

    sample_banner(c, w, h, dark, HexColor("#E5E7EB"))
    # Header
    c.setFillColor(dark)
    c.rect(0, h - 18 - 108, w, 108, fill=1, stroke=0)
    c.setFillColor(gold)
    c.rect(0, h - 18 - 112, w, 6, fill=1, stroke=0)

    c.setFillColor(HexColor("#F8FAFC"))
    c.setFont("Sans", 10)
    c.drawString(48, h - 50, "INVOICE")
    c.setFillColor(gold)
    c.setFont("SansBd", 26)
    c.drawString(48, h - 82, vendor)
    c.setFillColor(HexColor("#F8FAFC"))
    c.setFont("SansLt", 11)
    c.drawString(48, h - 102, "Fictional office-supply wholesaler  ·  Harbor City, OR")
    c.setFillColor(gold)
    c.setFont("SansBd", 22)
    c.drawRightString(w - 48, h - 78, invoice_no)

    # Bill to + meta
    y = h - 168
    c.setFillColor(mute)
    c.setFont("SansSm", 8)
    c.drawString(48, y, "BILL TO")
    c.setFillColor(ink)
    c.setFont("SansBd", 11)
    c.drawString(48, y - 16, "Lumen Bookkeeping Co.")
    c.setFont("Sans", 10)
    c.setFillColor(mute)
    c.drawString(48, y - 30, "88 Sample Pier, Suite 4")
    c.drawString(48, y - 44, "Harbor City, OR 97000")

    meta = [
        ("Vendor:", vendor),
        ("Invoice No:", invoice_no),
        ("Date:", date),
        ("Currency:", currency),
    ]
    mx, vx = 340, 430
    my = y
    for label, value in meta:
        c.setFillColor(mute)
        c.setFont("Sans", 10)
        c.drawString(mx, my, label)
        c.setFillColor(ink)
        c.setFont("SansBd", 11)
        c.drawString(vx, my, value)
        my -= 16

    # Table
    table_top = y - 70
    cols_right = {"Qty": 400, "Unit Price": 504, "Amount": 564}
    header_h, row_h = 26, 24
    c.setFillColor(head_bg)
    c.rect(40, table_top - header_h, w - 80, header_h, fill=1, stroke=0)
    c.setFillColor(ink)
    c.setFont("SansBd", 9)
    c.drawString(48, table_top - 17, "Description")
    c.drawRightString(cols_right["Qty"], table_top - 17, "Qty")
    c.drawRightString(cols_right["Unit Price"], table_top - 17, "Unit Price")
    c.drawRightString(cols_right["Amount"], table_top - 17, "Amount")

    yrow = table_top - header_h
    c.setStrokeColor(rule)
    c.setLineWidth(0.7)
    c.line(40, yrow, w - 40, yrow)
    for i, (desc, qty, unit, amount) in enumerate(rows):
        yrow -= row_h
        if i % 2 == 1:
            c.setFillColor(alt)
            c.rect(40, yrow, w - 80, row_h, fill=1, stroke=0)
        c.setFillColor(ink)
        c.setFont("Sans", 10)
        c.drawString(48, yrow + 7, desc)
        c.drawRightString(cols_right["Qty"], yrow + 7, f"{int(qty)}")
        c.drawRightString(cols_right["Unit Price"], yrow + 7, money(unit, currency, symbol=False))
        c.setFont("SansSm", 10)
        c.drawRightString(cols_right["Amount"], yrow + 7, money(amount, currency, symbol=False))
        c.setStrokeColor(rule)
        c.setLineWidth(0.5)
        c.line(40, yrow, w - 40, yrow)

    # Totals
    ty = yrow - 20
    def tot_row(label, value, bold=False, highlight=False):
        nonlocal ty
        ty -= 22
        if highlight:
            c.setFillColor(total_bg)
            c.rect(380, ty - 5, w - 40 - 380, 22, fill=1, stroke=0)
            c.setStrokeColor(gold)
            c.setLineWidth(1.1)
            c.rect(380, ty - 5, w - 40 - 380, 22, fill=0, stroke=1)
        c.setFillColor(ink)
        c.setFont("SansBd" if bold else "Sans", 10)
        c.drawRightString(504, ty, label)
        c.setFont("SansBd" if bold else "Sans", 11)
        c.drawRightString(564, ty, money(value, currency, symbol=False))

    tot_row("Net:", net)
    tot_row("VAT (10%):", vat)
    tot_row("Total:", total, bold=True, highlight=True)

    c.setStrokeColor(gold)
    c.setLineWidth(3)
    c.line(48, ty - 28, 170, ty - 28)
    c.setFillColor(mute)
    c.setFont("Sans", 9)
    c.drawString(48, ty - 44, "Payment terms: Net 15. Thank you for your business.")
    sample_footer(c, w, mute, y=ty - 68)
    c.showPage()
    c.save()
    return _meta(path, vendor, invoice_no, date, currency, tax_rate, rows, net, vat, total)


# ---------------------------------------------------------------------------
# 2. Bergholm Advisory AB  —  EUR  —  teal sidebar, Nordic VAT invoice
# ---------------------------------------------------------------------------
def draw_bergholm(path: Path) -> dict:
    vendor = "Bergholm Advisory AB"
    invoice_no = "BH-2026-0887"
    date = "2026-10-03"
    currency = "EUR"
    tax_rate = Decimal("0.25")
    items = [
        ("Strategy workshop (2 days)", Decimal("2"), Decimal("900.00")),
        ("Process mapping, 8 hours", Decimal("8"), Decimal("125.00")),
        ("Written recommendations", Decimal("1"), Decimal("450.00")),
    ]
    rows, net, vat, total = compute(items, tax_rate, currency)

    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle(f"Sample Invoice {invoice_no} — {vendor} (fictional)")
    c.setAuthor("Sample invoice")
    c.setSubject("Fictional sample invoice for a public demo")

    teal = HexColor("#0F4C5C")
    teal2 = HexColor("#16697A")
    paper = HexColor("#F7F4EF")
    ink = HexColor("#1A2B2F")
    mute = HexColor("#4A5C61")
    rule = HexColor("#C9D3D6")
    head_bg = HexColor("#E4EEEF")
    accent = HexColor("#2A9D8F")

    c.setFillColor(paper)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    c.setFillColor(teal)
    c.rect(0, 0, 28, h, fill=1, stroke=0)
    sample_banner(c, w, h, teal, HexColor("#E8F1F2"))

    # Title block
    c.setFillColor(teal)
    c.setFont("SansLt", 11)
    c.drawString(56, h - 56, "FAKTURA  /  INVOICE")
    c.setFont("SansBd", 24)
    c.drawString(56, h - 86, vendor)
    c.setFont("Sans", 10)
    c.setFillColor(mute)
    c.drawString(56, h - 104, "Fictional management consultancy  ·  Samplegatan 12, 111 22 Stockholm")

    # Big amount due box
    c.setFillColor(teal)
    c.roundRect(390, h - 130, 186, 78, 6, fill=1, stroke=0)
    c.setFillColor(HexColor("#D8F3EF"))
    c.setFont("Sans", 8)
    c.drawString(406, h - 72, "AMOUNT DUE")
    c.setFillColor(white)
    c.setFont("SansBd", 18)
    c.drawString(406, h - 96, money(total, currency))
    c.setFont("Sans", 9)
    c.drawString(406, h - 114, f"Currency: {currency}")

    # Two-column parties + fields
    y = h - 170
    c.setFillColor(teal2)
    c.setFont("SansSm", 8)
    c.drawString(56, y, "SUPPLIER")
    c.drawString(300, y, "BILL TO")
    c.setFillColor(ink)
    c.setFont("SansBd", 11)
    c.drawString(56, y - 16, vendor)
    c.drawString(300, y - 16, "Fjordline Retail AB")
    c.setFont("Sans", 9)
    c.setFillColor(mute)
    c.drawString(56, y - 30, "Org. no. SAMPLE-556677")
    c.drawString(300, y - 30, "Demo address, 414 51 Gothenburg")
    c.drawString(56, y - 44, "VAT ID: SE000000000001")
    c.drawString(300, y - 44, "VAT ID: SE000000000002")

    # Labeled fields in a 2x2 grid, label and value on the same baseline
    # so "Invoice Date: 2026-10-03" is one extractable line.
    fy = y - 78
    c.setFillColor(HexColor("#EAF3F3"))
    c.roundRect(48, fy - 52, w - 84, 66, 4, fill=1, stroke=0)
    field_rows = [
        (("Supplier:", vendor), ("Invoice Number:", invoice_no)),
        (("Invoice Date:", date), ("Currency:", currency)),
    ]
    for i, ((l1, v1), (l2, v2)) in enumerate(field_rows):
        yy = fy - 8 - i * 28
        c.setFillColor(mute)
        c.setFont("Sans", 9)
        c.drawString(56, yy, l1)
        c.drawString(300, yy, l2)
        c.setFillColor(ink)
        c.setFont("SansBd", 11)
        c.drawString(56 + 78, yy, v1)
        # Invoice Number is a longer label
        c.drawString(300 + 100, yy, v2)

    # Table
    table_top = fy - 80
    header_h, row_h = 26, 28
    c.setFillColor(teal)
    c.rect(48, table_top - header_h, w - 84, header_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("SansSm", 9)
    c.drawString(56, table_top - 17, "Description")
    c.drawRightString(390, table_top - 17, "Qty")
    c.drawRightString(478, table_top - 17, "Unit Price")
    c.drawRightString(556, table_top - 17, "Amount")

    yrow = table_top - header_h
    for i, (desc, qty, unit, amount) in enumerate(rows):
        yrow -= row_h
        if i % 2 == 1:
            c.setFillColor(HexColor("#EAF3F3"))
            c.rect(48, yrow, w - 84, row_h, fill=1, stroke=0)
        c.setStrokeColor(rule)
        c.setLineWidth(0.4)
        c.line(48, yrow, w - 36, yrow)
        c.setFillColor(ink)
        c.setFont("Sans", 10)
        c.drawString(56, yrow + 9, desc)
        c.drawRightString(390, yrow + 9, f"{int(qty)}")
        c.drawRightString(478, yrow + 9, money(unit, currency, symbol=False))
        c.setFont("SansSm", 10)
        c.drawRightString(556, yrow + 9, money(amount, currency, symbol=False))

    # Totals, European labels
    ty = yrow - 18
    def tot_row(label, value, bold=False):
        nonlocal ty
        ty -= 22
        c.setFillColor(ink)
        c.setFont("SansBd" if bold else "Sans", 10)
        c.drawRightString(478, ty, label)
        c.setFont("SansBd" if bold else "Sans", 12)
        c.drawRightString(556, ty, money(value, currency, symbol=False))
        if bold:
            c.setStrokeColor(accent)
            c.setLineWidth(1.6)
            c.line(400, ty - 6, w - 36, ty - 6)
            c.line(400, ty + 16, w - 36, ty + 16)

    tot_row("Net Amount:", net)
    tot_row("VAT (25%):", vat)
    tot_row("Amount Due:", total, bold=True)

    c.setFillColor(mute)
    c.setFont("SansIt", 9)
    c.drawString(56, ty - 40, "Bankgiro SAMPLE-0000  ·  Payment due 20 days  ·  Fictional invoice")
    sample_footer(c, w, mute, y=ty - 64)
    c.showPage()
    c.save()
    return _meta(path, vendor, invoice_no, date, currency, tax_rate, rows, net, vat, total)


# ---------------------------------------------------------------------------
# 3. Maple & Wick Studio Ltd  —  GBP  —  traditional British serif
# ---------------------------------------------------------------------------
def draw_maple(path: Path) -> dict:
    vendor = "Maple & Wick Studio Ltd"
    invoice_no = "MW-4419"
    date = "15 November 2026"
    currency = "GBP"
    tax_rate = Decimal("0.20")
    items = [
        ("Brand identity refresh", Decimal("1"), Decimal("640.00")),
        ("Business card design (double-sided)", Decimal("1"), Decimal("85.00")),
        ("Print-ready export pack", Decimal("1"), Decimal("40.00")),
    ]
    rows, net, vat, total = compute(items, tax_rate, currency)

    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle(f"Sample Invoice {invoice_no} — {vendor} (fictional)")
    c.setAuthor("Sample invoice")
    c.setSubject("Fictional sample invoice for a public demo")

    burgundy = HexColor("#6B2D3C")
    ink = HexColor("#1F1A17")
    mute = HexColor("#5C534C")
    rule = HexColor("#C4B8A8")
    paper = HexColor("#FBF7F1")
    cream = HexColor("#F3EBE0")

    c.setFillColor(paper)
    c.rect(0, 0, w, h, fill=1, stroke=0)
    sample_banner(c, w, h, burgundy, HexColor("#F8EFE8"))

    # Centered letterhead
    c.setFillColor(burgundy)
    c.setFont("SerifBd", 22)
    c.drawCentredString(w / 2, h - 64, vendor)
    c.setStrokeColor(burgundy)
    c.setLineWidth(0.8)
    c.line(180, h - 74, w - 180, h - 74)
    c.setFont("SerifIt", 10)
    c.setFillColor(mute)
    c.drawCentredString(w / 2, h - 90, "Fictional design studio  ·  14 Hartfield Mews, Sampletown  ·  Company no. SAMPLE-4419")

    c.setFillColor(burgundy)
    c.setFont("SerifBd", 16)
    c.drawCentredString(w / 2, h - 126, "INVOICE")

    # Meta in two formal columns
    y = h - 160
    c.setStrokeColor(rule)
    c.setLineWidth(0.6)
    c.line(56, y + 16, w - 56, y + 16)

    c.setFillColor(mute)
    c.setFont("Serif", 9)
    c.drawString(56, y, "From")
    c.setFillColor(ink)
    c.setFont("SerifMd", 11)
    c.drawString(56, y - 16, vendor)
    c.setFont("Serif", 9)
    c.setFillColor(mute)
    c.drawString(56, y - 32, "VAT No. GB XX 000 0000 00")

    c.setFillColor(mute)
    c.setFont("Serif", 9)
    c.drawString(330, y, "To")
    c.setFillColor(ink)
    c.setFont("SerifMd", 11)
    c.drawString(330, y - 16, "Ashworth & Pen Books")
    c.setFont("Serif", 9)
    c.setFillColor(mute)
    c.drawString(330, y - 32, "3 Example Yard, Sampletown")

    # Labeled fields, formal
    fy = y - 62
    c.setFillColor(cream)
    c.rect(56, fy - 50, w - 112, 58, fill=1, stroke=0)
    c.setStrokeColor(rule)
    c.setLineWidth(0.6)
    c.rect(56, fy - 50, w - 112, 58, fill=0, stroke=1)

    pairs = [
        (70, "Vendor:", vendor),
        (330, "Invoice No.:", invoice_no),
        (70, "Tax Date:", date),
        (330, "Currency:", currency),
    ]
    # two rows
    c.setFont("Serif", 9)
    c.setFillColor(mute)
    c.drawString(70, fy - 8, "Vendor:")
    c.drawString(330, fy - 8, "Invoice No.:")
    c.drawString(70, fy - 34, "Tax Date:")
    c.drawString(330, fy - 34, "Currency:")
    c.setFillColor(ink)
    c.setFont("SerifMd", 11)
    c.drawString(130, fy - 8, vendor)
    c.drawString(410, fy - 8, invoice_no)
    c.drawString(130, fy - 34, date)
    c.drawString(410, fy - 34, currency)

    # Table with serif headers
    table_top = fy - 80
    header_h, row_h = 26, 28
    c.setFillColor(burgundy)
    c.rect(56, table_top - header_h, w - 112, header_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("SerifMd", 9)
    c.drawString(66, table_top - 17, "Description")
    c.drawRightString(390, table_top - 17, "Qty")
    c.drawRightString(478, table_top - 17, "Unit Price")
    c.drawRightString(548, table_top - 17, "Amount")

    yrow = table_top - header_h
    for i, (desc, qty, unit, amount) in enumerate(rows):
        yrow -= row_h
        if i % 2 == 1:
            c.setFillColor(cream)
            c.rect(56, yrow, w - 112, row_h, fill=1, stroke=0)
        c.setStrokeColor(rule)
        c.setLineWidth(0.4)
        c.line(56, yrow, w - 56, yrow)
        c.setFillColor(ink)
        c.setFont("Serif", 10)
        c.drawString(66, yrow + 9, desc)
        c.drawRightString(390, yrow + 9, f"{int(qty)}")
        c.drawRightString(478, yrow + 9, money(unit, currency, symbol=False))
        c.setFont("SerifMd", 10)
        c.drawRightString(548, yrow + 9, money(amount, currency, symbol=False))

    ty = yrow - 16
    def tot_row(label, value, bold=False):
        nonlocal ty
        ty -= 22
        c.setFillColor(ink)
        c.setFont("SerifBd" if bold else "Serif", 10)
        c.drawRightString(478, ty, label)
        c.setFont("SerifBd" if bold else "Serif", 12)
        c.drawRightString(548, ty, money(value, currency, symbol=False))
        if bold:
            c.setStrokeColor(burgundy)
            c.setLineWidth(1.0)
            c.line(400, ty - 5, w - 56, ty - 5)
            c.line(400, ty + 16, w - 56, ty + 16)

    tot_row("Net:", net)
    tot_row("VAT @ 20%:", vat)
    tot_row("Total due:", total, bold=True)

    c.setFillColor(mute)
    c.setFont("SerifIt", 9)
    c.drawCentredString(w / 2, ty - 40, "Payment due within 14 days. This is a fictional invoice for demonstration only.")
    sample_footer(c, w, mute, y=ty - 64)
    c.showPage()
    c.save()
    return _meta(path, vendor, invoice_no, date, currency, tax_rate, rows, net, vat, total)


# ---------------------------------------------------------------------------
# 4. Kitebay Office Goods K.K.  —  JPY  —  tight grid, red SAMPLE stamp
# ---------------------------------------------------------------------------
def draw_kitebay(path: Path) -> dict:
    vendor = "Kitebay Office Goods K.K."
    invoice_no = "KB-77105"
    date = "2026/09/28"
    currency = "JPY"
    tax_rate = Decimal("0.10")
    items = [
        ("A3 plotter paper, 5 rolls", Decimal("5"), Decimal("2800")),
        ("Toner cartridge, black", Decimal("4"), Decimal("6500")),
        ("Desk organizers, set of 3", Decimal("6"), Decimal("1800")),
        ("Shipping", Decimal("1"), Decimal("1200")),
    ]
    rows, net, vat, total = compute(items, tax_rate, currency)

    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle(f"Sample Invoice {invoice_no} — {vendor} (fictional)")
    c.setAuthor("Sample invoice")
    c.setSubject("Fictional sample invoice for a public demo")

    ink = HexColor("#1A1A1A")
    mute = HexColor("#4A4A4A")
    red = HexColor("#B42318")
    rule = HexColor("#333333")
    head_bg = HexColor("#F4F4F4")
    box = HexColor("#FAFAFA")

    sample_banner(c, w, h, HexColor("#1A1A1A"), HexColor("#F3F3F3"))

    # Top grid: company left, INVOICE box right
    c.setFillColor(ink)
    c.setFont("SansBd", 18)
    c.drawString(48, h - 58, vendor)
    c.setFont("Sans", 9)
    c.setFillColor(mute)
    c.drawString(48, h - 74, "Fictional office wholesaler  ·  1-2-3 Sample-dori, Demo-ku, Tokyo")
    c.drawString(48, h - 88, "Corp. no. SAMPLE-77105  ·  Amounts in Japanese yen (no decimals)")

    # Invoice box
    c.setStrokeColor(ink)
    c.setLineWidth(1.6)
    c.rect(420, h - 110, 144, 62, fill=0, stroke=1)
    c.setFillColor(ink)
    c.setFont("SansBd", 14)
    c.drawCentredString(492, h - 68, "INVOICE")
    c.setFont("Sans", 9)
    c.drawCentredString(492, h - 86, "Invoice No:")
    c.setFont("SansBd", 11)
    c.drawCentredString(492, h - 102, invoice_no)

    # Red circular SAMPLE stamp (hanko-style)
    cx, cy, r = 492, h - 160, 28
    c.setStrokeColor(red)
    c.setLineWidth(2)
    c.circle(cx, cy, r, fill=0, stroke=1)
    c.circle(cx, cy, r - 4, fill=0, stroke=1)
    c.setFillColor(red)
    c.setFont("SansBd", 8)
    c.drawCentredString(cx, cy - 3, "SAMPLE")

    # Field table
    fy = h - 140
    field_rows = [
        ("Vendor:", vendor),
        ("Invoice No:", invoice_no),
        ("Date:", date),
        ("Currency:", currency),
        ("Bill To:", "Minato Demo Trading Co."),
    ]
    c.setStrokeColor(rule)
    c.setLineWidth(0.8)
    rw, label_w = 22, 100
    box_w = 320
    for i, (label, value) in enumerate(field_rows):
        yy = fy - i * rw
        c.rect(48, yy - rw, label_w, rw, fill=0, stroke=1)
        c.rect(48 + label_w, yy - rw, box_w - label_w, rw, fill=0, stroke=1)
        c.setFillColor(HexColor("#F4F4F4"))
        c.rect(48, yy - rw, label_w, rw, fill=1, stroke=0)
        c.setStrokeColor(rule)
        c.rect(48, yy - rw, label_w, rw, fill=0, stroke=1)
        c.rect(48 + label_w, yy - rw, box_w - label_w, rw, fill=0, stroke=1)
        c.setFillColor(ink)
        c.setFont("Sans", 8)
        c.drawString(54, yy - 15, label)
        c.setFont("SansSm", 9)
        c.drawString(48 + label_w + 8, yy - 15, value)

    # Line items table
    table_top = fy - len(field_rows) * rw - 24
    header_h, row_h = 24, 24
    c.setFillColor(ink)
    c.rect(48, table_top - header_h, w - 96, header_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("SansBd", 8)
    c.drawString(56, table_top - 16, "Description")
    c.drawRightString(390, table_top - 16, "Qty")
    c.drawRightString(478, table_top - 16, "Unit Price")
    c.drawRightString(556, table_top - 16, "Amount")

    yrow = table_top - header_h
    c.setStrokeColor(rule)
    c.setLineWidth(0.8)
    for i, (desc, qty, unit, amount) in enumerate(rows):
        yrow -= row_h
        c.rect(48, yrow, w - 96, row_h, fill=0, stroke=1)
        c.setFillColor(ink)
        c.setFont("Sans", 10)
        c.drawString(56, yrow + 7, desc)
        c.drawRightString(390, yrow + 7, f"{int(qty)}")
        c.drawRightString(478, yrow + 7, money(unit, currency, symbol=False))
        c.setFont("SansSm", 10)
        c.drawRightString(556, yrow + 7, money(amount, currency, symbol=False))

    # Totals in a right-hand box
    ty = yrow - 8
    tot_items = [
        ("Net:", net, False),
        ("Consumption Tax (10%):", vat, False),
        ("Total:", total, True),
    ]
    for label, value, bold in tot_items:
        ty -= 24
        if bold:
            c.setFillColor(ink)
            c.rect(330, ty - 4, 234, 24, fill=1, stroke=0)
            c.setFillColor(white)
        else:
            c.setStrokeColor(rule)
            c.setLineWidth(0.8)
            c.rect(330, ty - 4, 234, 24, fill=0, stroke=1)
            c.setFillColor(ink)
        c.setFont("SansBd" if bold else "Sans", 9)
        c.drawString(340, ty + 3, label)
        c.setFont("SansBd" if bold else "SansSm", 11)
        c.drawRightString(556, ty + 3, money(value, currency, symbol=False))

    c.setFillColor(mute)
    c.setFont("Sans", 8)
    c.drawString(48, ty - 40, "Payment: bank transfer within 30 days. Fictional invoice. Yen amounts are whole numbers.")
    sample_footer(c, w, mute, y=ty - 64)
    c.showPage()
    c.save()
    return _meta(path, vendor, invoice_no, date, currency, tax_rate, rows, net, vat, total)


# ---------------------------------------------------------------------------
# 5. Southern Reef Supplies Pty Ltd  —  AUD  —  modern teal / coral trades
# ---------------------------------------------------------------------------
def draw_southern(path: Path) -> dict:
    vendor = "Southern Reef Supplies Pty Ltd"
    invoice_no = "SRS-26019"
    date = "03/11/2026"
    currency = "AUD"
    tax_rate = Decimal("0.10")
    items = [
        ("Outdoor cable, 50m drum", Decimal("3"), Decimal("89.50")),
        ("Weatherproof junction boxes", Decimal("10"), Decimal("12.40")),
        ("Cable ties, pack of 200", Decimal("8"), Decimal("4.75")),
        ("Same-day courier", Decimal("1"), Decimal("35.00")),
    ]
    rows, net, vat, total = compute(items, tax_rate, currency)

    w, h = PAGE
    c = canvas.Canvas(str(path), pagesize=PAGE)
    c.setTitle(f"Sample Invoice {invoice_no} — {vendor} (fictional)")
    c.setAuthor("Sample invoice")
    c.setSubject("Fictional sample invoice for a public demo")

    teal = HexColor("#0D7377")
    deep = HexColor("#095456")
    coral = HexColor("#E85D04")
    ink = HexColor("#1A2E2F")
    mute = HexColor("#4A6162")
    mint = HexColor("#E8F4F4")
    rule = HexColor("#C5D6D6")

    sample_banner(c, w, h, deep, HexColor("#E8F4F4"))
    # Bold header band
    c.setFillColor(teal)
    c.rect(0, h - 18 - 120, w, 120, fill=1, stroke=0)
    c.setFillColor(coral)
    c.rect(0, h - 18 - 126, w, 6, fill=1, stroke=0)

    c.setFillColor(white)
    c.setFont("SansLt", 11)
    c.drawString(48, h - 52, "TAX INVOICE")
    c.setFont("SansBd", 22)
    c.drawString(48, h - 82, vendor)
    c.setFont("Sans", 10)
    c.setFillColor(HexColor("#D4EDED"))
    c.drawString(48, h - 102, "Fictional electrical wholesaler  ·  42 Harbour Dummy Rd, Sample Bay QLD 4000")
    c.setFillColor(white)
    c.setFont("Sans", 9)
    c.drawRightString(w - 48, h - 58, "Invoice #")
    c.setFont("SansBd", 18)
    c.drawRightString(w - 48, h - 82, invoice_no)
    c.setFont("Sans", 10)
    c.setFillColor(HexColor("#D4EDED"))
    c.drawRightString(w - 48, h - 102, f"ABN SAMPLE 00 000 000 000")

    # Meta chips
    y = h - 176
    c.setFillColor(mint)
    c.roundRect(40, y - 70, 320, 80, 6, fill=1, stroke=0)
    meta = [
        ("Supplier:", vendor),
        ("Invoice #:", invoice_no),
        ("Date:", date),
        ("Currency:", currency),
    ]
    my = y - 6
    for label, value in meta:
        c.setFillColor(mute)
        c.setFont("Sans", 9)
        c.drawString(52, my, label)
        c.setFillColor(ink)
        c.setFont("SansBd", 10)
        c.drawString(130, my, value)
        my -= 16

    # Bill to + large payable
    c.setFillColor(mute)
    c.setFont("SansSm", 8)
    c.drawString(380, y, "BILL TO")
    c.setFillColor(ink)
    c.setFont("SansBd", 11)
    c.drawString(380, y - 16, "Keppel Field Services")
    c.setFont("Sans", 9)
    c.setFillColor(mute)
    c.drawString(380, y - 30, "9 Example Close, Sample Bay")

    c.setFillColor(coral)
    c.roundRect(380, y - 86, 184, 42, 5, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("Sans", 7)
    c.drawString(392, y - 58, "AMOUNT PAYABLE")
    c.setFont("SansBd", 14)
    c.drawString(392, y - 76, money(total, currency))

    # Table
    table_top = y - 110
    header_h, row_h = 26, 26
    c.setFillColor(deep)
    c.roundRect(40, table_top - header_h, w - 80, header_h, 0, fill=1, stroke=0)
    c.rect(40, table_top - header_h, w - 80, header_h, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont("SansBd", 9)
    c.drawString(50, table_top - 17, "Description")
    c.drawRightString(390, table_top - 17, "Qty")
    c.drawRightString(478, table_top - 17, "Unit Price")
    c.drawRightString(564, table_top - 17, "Amount")

    yrow = table_top - header_h
    for i, (desc, qty, unit, amount) in enumerate(rows):
        yrow -= row_h
        if i % 2 == 1:
            c.setFillColor(mint)
            c.rect(40, yrow, w - 80, row_h, fill=1, stroke=0)
        c.setStrokeColor(rule)
        c.setLineWidth(0.5)
        c.line(40, yrow, w - 40, yrow)
        c.setFillColor(ink)
        c.setFont("Sans", 10)
        c.drawString(50, yrow + 8, desc)
        c.drawRightString(390, yrow + 8, f"{int(qty)}")
        c.drawRightString(478, yrow + 8, money(unit, currency, symbol=False))
        c.setFont("SansSm", 10)
        c.drawRightString(564, yrow + 8, money(amount, currency, symbol=False))

    ty = yrow - 16
    def tot_row(label, value, bold=False, coral_hl=False):
        nonlocal ty
        ty -= 24
        if coral_hl:
            c.setFillColor(coral)
            c.roundRect(378, ty - 6, 194, 26, 4, fill=1, stroke=0)
            c.setFillColor(white)
        else:
            c.setFillColor(ink)
        c.setFont("SansBd" if bold else "Sans", 10)
        c.drawRightString(478, ty, label)
        c.setFont("SansBd" if bold else "Sans", 12)
        c.drawRightString(564, ty, money(value, currency, symbol=False))

    tot_row("Subtotal (ex GST):", net)
    tot_row("GST (10%):", vat)
    # Also label as VAT for the extractor — printed as GST; extractor maps GST → VAT.
    # Add a parseable Net / VAT / Total line in small type? Better: include synonyms
    # "Subtotal (ex GST)" should map to net; "GST (10%)" to VAT.
    # I'll also print Net and VAT aliases as clearly labeled rows... Subtotal is the net.
    # To be extra parseable, add hidden-looking but visible labeled totals:
    tot_row("Amount payable:", total, bold=True, coral_hl=True)

    # Extra clearly labeled summary for the extractor (also visible, professional)
    c.setFillColor(mute)
    c.setFont("Sans", 8)
    c.drawString(48, ty - 28, f"Net: {money(net, currency, symbol=False)}    VAT (GST 10%): {money(vat, currency, symbol=False)}    Total: {money(total, currency, symbol=False)}")

    c.setFillColor(mute)
    c.setFont("Sans", 9)
    c.drawString(48, ty - 52, "Please pay by EFT within 7 days. Fictional invoice. Date is day/month/year (03 Nov 2026).")
    sample_footer(c, w, mute, y=ty - 76)
    c.showPage()
    c.save()
    return _meta(path, vendor, invoice_no, date, currency, tax_rate, rows, net, vat, total)


def _meta(path, vendor, invoice_no, date, currency, tax_rate, rows, net, vat, total):
    return {
        "path": str(path),
        "file": path.name,
        "vendor": vendor,
        "invoice_number": invoice_no,
        "date": date,
        "currency": currency,
        "tax_rate": str(tax_rate),
        "line_items": [
            {
                "description": d,
                "qty": int(q),
                "unit_price": str(u),
                "amount": str(a),
            }
            for d, q, u, a in rows
        ],
        "net": str(net),
        "vat": str(vat),
        "total": str(total),
    }


def main() -> list[dict]:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    jobs = [
        ("01-northwind-supplies-usd.pdf", draw_northwind),
        ("02-bergholm-advisory-eur.pdf", draw_bergholm),
        ("03-maple-wick-studio-gbp.pdf", draw_maple),
        ("04-kitebay-office-goods-jpy.pdf", draw_kitebay),
        ("05-southern-reef-supplies-aud.pdf", draw_southern),
    ]
    results = []
    for name, fn in jobs:
        path = OUT_DIR / name
        info = fn(path)
        results.append(info)
        print(
            f"wrote {path.name:44}  {info['invoice_number']:14}  "
            f"{info['currency']}  net={info['net']}  vat={info['vat']}  total={info['total']}"
        )
    return results


if __name__ == "__main__":
    main()
