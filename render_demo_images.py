#!/usr/bin/env python3
"""Render demo images from the real invoice PDFs and the extracted workbook.

Numbers on every image come from output/extracted_invoices.xlsx or from a
direct render of the input PDFs. Nothing is typed in by hand.
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import numpy as np
from openpyxl import load_workbook
from pdf2image import convert_from_path
from PIL import Image, ImageDraw, ImageFilter, ImageFont

ROOT = Path(__file__).resolve().parent
INPUT_DIR = ROOT / "input"
XLSX = ROOT / "output" / "extracted_invoices.xlsx"
IMG_DIR = ROOT / "images"
PDF_RENDER_DIR = IMG_DIR / "pdf-renders"

FONT_DIR = Path("/usr/share/fonts/truetype/sand-box/custom/Source Sans Pro")
FONTS = {
    "reg": FONT_DIR / "SourceSansPro-Regular.ttf",
    "it": FONT_DIR / "SourceSansPro-It.ttf",
    "sm": FONT_DIR / "SourceSansPro-Semibold.ttf",
    "bd": FONT_DIR / "SourceSansPro-Bold.ttf",
    "lt": FONT_DIR / "SourceSansPro-Light.ttf",
}

# Cover palette (matches the invoice gold/ink language)
INK = (17, 19, 24)
INK_2 = (14, 16, 21)
GOLD = (250, 204, 21)
GOLD_SOFT = (250, 204, 21, 40)
WHITE = (248, 250, 252)
MUTE = (156, 163, 175)
GRID = (214, 219, 227)
ROW_ALT = (247, 248, 250)
HEAD = (17, 19, 24)
GREEN = (16, 185, 129)
GREEN_BG = (209, 250, 229)
GREEN_FG = (6, 95, 70)
RED_BG = (254, 202, 202)
RED_FG = (127, 29, 29)
SHEET_GRID = (226, 227, 227)
SHEET_HEAD = (248, 249, 250)
SHEET_GREEN = (24, 128, 56)
HIGHLIGHT = (255, 248, 214)


def F(weight: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONTS[weight]), size)


def format_amount(value, currency: str) -> str:
    if value is None:
        return ""
    if currency == "JPY":
        return f"{int(round(float(value))):,}"
    return f"{Decimal(str(value)).quantize(Decimal('0.01')):,.2f}"


def with_symbol(value, currency: str) -> str:
    body = format_amount(value, currency)
    return {
        "USD": f"${body}",
        "EUR": f"€{body}",
        "GBP": f"£{body}",
        "JPY": f"¥{body}",
        "AUD": f"A${body}",
    }.get(currency, body)


def load_workbook_rows():
    wb = load_workbook(XLSX, data_only=True)
    inv = []
    ws = wb["Invoices"]
    headers = [c.value for c in ws[1]]
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if not row[0]:
            continue
        inv.append(dict(zip(headers, row)))
    items = []
    ws2 = wb["Line items"]
    headers2 = [c.value for c in ws2[1]]
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, values_only=True):
        if not row[0]:
            continue
        items.append(dict(zip(headers2, row)))
    review = []
    ws3 = wb["Review"]
    headers3 = [c.value for c in ws3[1]]
    for row in ws3.iter_rows(min_row=2, max_row=1 + len(inv), values_only=True):
        review.append(dict(zip(headers3, row)))
    summary = ws3.cell(2 + len(inv), 2).value
    overall = ws3.cell(2 + len(inv), 15).value
    return {
        "wb": wb,
        "invoices": inv,
        "items": items,
        "review": review,
        "summary": summary,
        "overall": overall,
    }


def render_pdfs(dpi: int = 200) -> list[Image.Image]:
    PDF_RENDER_DIR.mkdir(parents=True, exist_ok=True)
    images = []
    for pdf in sorted(INPUT_DIR.glob("*.pdf")):
        page = convert_from_path(str(pdf), dpi=dpi)[0].convert("RGB")
        dest = PDF_RENDER_DIR / f"{pdf.stem}.png"
        # Also keep a 160-dpi preview written earlier; overwrite with sharper copy.
        page.save(dest, "PNG", optimize=True)
        images.append(page)
    return images


def crop_bottom_whitespace(im: Image.Image, pad: int = 36) -> Image.Image:
    """Trim empty space under the footer, keeping full width (sidebars stay)."""
    arr = np.array(im)
    gray = arr.mean(axis=2)
    h, w = gray.shape
    core = gray[:, int(w * 0.10) : int(w * 0.96)]
    mask = core < 242
    rows = np.where(mask.any(axis=1))[0]
    if len(rows) == 0:
        return im
    bottom = min(h, int(rows[-1]) + pad)
    return im.crop((0, 0, w, bottom))


def rounded_mask(size, radius: int) -> Image.Image:
    mask = Image.new("L", size, 0)
    d = ImageDraw.Draw(mask)
    d.rounded_rectangle((0, 0, size[0] - 1, size[1] - 1), radius=radius, fill=255)
    return mask


def apply_round(im: Image.Image, radius: int) -> Image.Image:
    im = im.convert("RGBA")
    mask = rounded_mask(im.size, radius)
    out = Image.new("RGBA", im.size, (0, 0, 0, 0))
    out.paste(im, (0, 0), mask)
    return out


def drop_shadow(
    card: Image.Image,
    blur: int = 22,
    offset: tuple[int, int] = (0, 14),
    opacity: int = 160,
    pad: int | None = None,
) -> Image.Image:
    """Return the card with a soft drop shadow, padded so the shadow fits."""
    pad = blur * 2 + max(abs(offset[0]), abs(offset[1])) if pad is None else pad
    w, h = card.size
    canvas = Image.new("RGBA", (w + pad * 2, h + pad * 2), (0, 0, 0, 0))
    shadow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    # Use the card's alpha as the shadow shape.
    alpha = card.split()[-1] if card.mode == "RGBA" else Image.new("L", card.size, 255)
    sh = Image.new("L", canvas.size, 0)
    sh.paste(alpha, (pad + offset[0], pad + offset[1]))
    sh = sh.point(lambda p: int(p * opacity / 255))
    sh = sh.filter(ImageFilter.GaussianBlur(blur))
    shadow.putalpha(sh)
    canvas = Image.alpha_composite(canvas, shadow)
    canvas.paste(card, (pad, pad), card if card.mode == "RGBA" else None)
    return canvas


def scale_to_width(im: Image.Image, width: int) -> Image.Image:
    h = int(im.height * (width / im.width))
    return im.resize((width, h), Image.Resampling.LANCZOS)


def scale_to_height(im: Image.Image, height: int) -> Image.Image:
    w = int(im.width * (height / im.height))
    return im.resize((w, height), Image.Resampling.LANCZOS)


def bg_dark(size: tuple[int, int]) -> Image.Image:
    """Dark canvas with a faint grid and a soft gold glow in the centre."""
    w, h = size
    im = Image.new("RGB", size, INK_2)
    px = np.array(im)
    # Subtle vertical vignette
    yy = np.linspace(0, 1, h)[:, None]
    xx = np.linspace(0, 1, w)[None, :]
    lift = (0.10 * (1 - np.abs(yy - 0.45) * 1.3)).clip(0, 1)
    px = np.clip(px + (lift * 12)[:, :, None], 0, 255).astype(np.uint8)
    im = Image.fromarray(px, "RGB")
    d = ImageDraw.Draw(im, "RGBA")
    # Faint dot grid
    for x in range(28, w, 28):
        for y in range(28, h, 28):
            d.point((x, y), fill=(255, 255, 255, 18))
    # Soft gold glow behind the centre (arrow area)
    glow = Image.new("RGBA", size, (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    cx, cy = w // 2, int(h * 0.55)
    for r, a in [(220, 18), (140, 28), (80, 40)]:
        gd.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(250, 204, 21, a))
    glow = glow.filter(ImageFilter.GaussianBlur(40))
    im = Image.alpha_composite(im.convert("RGBA"), glow).convert("RGB")
    return im


def draw_chevron(canvas: Image.Image, cx: int, cy: int, size: int = 28) -> None:
    """Gold arrow with a soft glow, pointing right."""
    glow = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    gd.ellipse((cx - 36, cy - 36, cx + 36, cy + 36), fill=(250, 204, 21, 70))
    glow = glow.filter(ImageFilter.GaussianBlur(12))
    base = canvas.convert("RGBA")
    base = Image.alpha_composite(base, glow)
    d = ImageDraw.Draw(base)
    w, h = size, int(size * 1.35)
    pts = [
        (cx - w // 3, cy - h // 2),
        (cx + w // 2, cy),
        (cx - w // 3, cy + h // 2),
        (cx - w // 3 + 10, cy),
    ]
    d.polygon(pts, fill=GOLD)
    canvas.paste(base.convert(canvas.mode), (0, 0))


def stack_invoices(pages: list[Image.Image], front_width: int) -> Image.Image:
    """Front invoice with a thin sliver of a second invoice peeking on the left."""
    front = apply_round(scale_to_width(pages[0], front_width), 14)
    sliver_src = crop_bottom_whitespace(pages[1], pad=28)
    sliver = apply_round(scale_to_width(sliver_src, front_width), 14)
    # Align bottoms; offset left so only a strip of the second invoice shows.
    offset = 22
    pad = 50
    width = front.width + pad * 2 + offset
    height = front.height + pad * 2
    canvas = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    sliver_s = drop_shadow(sliver, blur=16, offset=(0, 10), opacity=110)
    front_s = drop_shadow(front, blur=22, offset=(0, 14), opacity=170)
    # drop_shadow adds its own internal pad; line the cards up by their unpadded size
    # Place sliver first, then the front slightly to the right.
    canvas.alpha_composite(sliver_s, (pad - offset - 8, pad - 8))
    canvas.alpha_composite(front_s, (pad, pad - 4))
    return canvas


def invoices_card(data: dict, width: int, height: int) -> Image.Image:
    """Clean rounded spreadsheet card of the Invoices sheet (real values)."""
    im = Image.new("RGB", (width, height), WHITE)
    d = ImageDraw.Draw(im)
    # Dark header bar with the five invoice numbers
    d.rectangle((0, 0, width, 52), fill=HEAD)
    d.text((16, 8), "extracted_invoices.xlsx", font=F("reg", 11), fill=(200, 204, 210))
    d.text((16, 26), "Invoices  ·  5 sample PDFs", font=F("bd", 14), fill=WHITE)

    rows = data["invoices"]
    # Columns: Vendor, Invoice no., Date, Net, VAT, Total
    cols = [
        ("Vendor", "Vendor", 0.27, "left"),
        ("Invoice no.", "Invoice number", 0.16, "left"),
        ("Date", "Date (on invoice)", 0.18, "left"),
        ("Net", "Net", 0.13, "right"),
        ("VAT", "VAT", 0.12, "right"),
        ("Total", "Total", 0.14, "right"),
    ]
    # Header row
    y0 = 52
    row_h = 32
    head_h = 28
    d.rectangle((0, y0, width, y0 + head_h), fill=HEAD)
    x = 0
    col_xs = []
    usable = width
    for title, key, frac, align in cols:
        cw = int(usable * frac)
        col_xs.append((x, cw, key, align))
        if align == "left":
            d.text((x + 10, y0 + 7), title, font=F("sm", 11), fill=WHITE)
        else:
            tw = d.textlength(title, font=F("sm", 11))
            d.text((x + cw - 10 - tw, y0 + 7), title, font=F("sm", 11), fill=WHITE)
        x += cw

    y = y0 + head_h
    for i, inv in enumerate(rows):
        bg = HIGHLIGHT if i == 0 else (ROW_ALT if i % 2 else WHITE)
        d.rectangle((0, y, width, y + row_h), fill=bg)
        # hairline
        d.line((0, y + row_h, width, y + row_h), fill=GRID, width=1)
        cur = inv["Currency"]
        values = {
            "Vendor": inv["Vendor"],
            "Invoice number": inv["Invoice number"],
            "Date (on invoice)": inv["Date (on invoice)"],
            "Net": with_symbol(inv["Net"], cur),
            "VAT": with_symbol(inv["VAT"], cur),
            "Total": with_symbol(inv["Total"], cur),
        }
        for x, cw, key, align in col_xs:
            text = str(values[key])
            font = F("bd", 12) if key == "Total" else F("reg", 12)
            fill = INK
            if align == "left":
                # Truncate vendor if needed
                while d.textlength(text, font=font) > cw - 16 and len(text) > 4:
                    text = text[:-2] + "…"
                d.text((x + 10, y + 8), text, font=font, fill=fill)
            else:
                tw = d.textlength(text, font=font)
                d.text((x + cw - 10 - tw, y + 8), text, font=font, fill=fill)
        y += row_h

    # Green footer with the real Review summary
    footer_h = height - y
    if footer_h < 36:
        footer_h = 36
        # shouldn't happen
    d.rectangle((0, y, width, height), fill=GREEN_BG)
    tick = "✓  "
    n = len(rows)
    overall = data["overall"] or ""
    label = f"{tick}{n} of {n} invoices passed every check  ·  {overall}"
    d.text((16, y + (height - y) // 2 - 8), label, font=F("bd", 13), fill=GREEN_FG)
    return apply_round(im, 16)


def fiverr_cover(pages: list[Image.Image], data: dict, path: Path) -> Path:
    W, H = 1280, 769
    im = bg_dark((W, H))
    d = ImageDraw.Draw(im)

    # Title
    title_y = 28
    # "PDF invoices" in gold, rest in white, one line as specified.
    t1 = "PDF invoices"
    t2 = " to Excel / Google Sheets"
    f_title = F("bd", 42)
    # Center the whole title
    w1 = d.textlength(t1, font=f_title)
    w2 = d.textlength(t2, font=f_title)
    tx = (W - (w1 + w2)) // 2
    d.text((tx, title_y), t1, font=f_title, fill=GOLD)
    d.text((tx + w1, title_y), t2, font=f_title, fill=WHITE)
    sub = "Five fictional sample invoices  ·  one checked workbook"
    f_sub = F("reg", 15)
    sw = d.textlength(sub, font=f_sub)
    d.text(((W - sw) // 2, title_y + 50), sub, font=f_sub, fill=MUTE)

    # Content band
    content_top = 100
    content_bot = H - 28
    band_h = content_bot - content_top

    # Left stack
    cropped_front = crop_bottom_whitespace(pages[0], pad=40)
    # Target height for the stack roughly band_h
    front_width = 500
    stack = stack_invoices([cropped_front] + [crop_bottom_whitespace(p, 40) for p in pages[1:]], front_width)
    # Scale stack to fit the band
    max_h = band_h + 40  # allow shadow pad
    if stack.height > max_h:
        stack = scale_to_height(stack, max_h)
    # Place left
    sx = 10
    sy = content_top + (band_h - stack.height) // 2 - 8
    im.paste(stack, (sx, sy), stack)

    # Right spreadsheet card
    card_w, card_h = 620, 300
    # Height: header 52 + head 28 + 5*32 + footer ~48 = 52+28+160+48 = 288
    card_h = 52 + 28 + 32 * len(data["invoices"]) + 48
    card = invoices_card(data, card_w, card_h)
    card_s = drop_shadow(card, blur=22, offset=(0, 14), opacity=170)
    # Place right, vertically centered in the band
    cx = W - card_s.width - 8
    cy = content_top + (band_h - card_s.height) // 2 - 8
    im.paste(card_s, (cx, cy), card_s)

    # Arrow between the two
    # Estimate visual centres
    left_visual_right = sx + stack.width - 40
    right_visual_left = cx + 40
    arrow_x = (left_visual_right + right_visual_left) // 2
    arrow_y = content_top + band_h // 2 + 8
    draw_chevron(im, arrow_x, arrow_y, size=32)

    im = im.convert("RGB")
    path.parent.mkdir(parents=True, exist_ok=True)
    im.save(path, "PNG", optimize=True)
    print(f"wrote {path} {im.size}")
    return path


# ---------------------------------------------------------------------------
# Google Sheets-style screenshots, painted from the live workbook
# ---------------------------------------------------------------------------
def cell_rgb(cell, default=(255, 255, 255)):
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return default
    fg = fill.fgColor
    rgb = getattr(fg, "rgb", None)
    if not rgb or rgb == "00000000":
        return default
    s = str(rgb)
    if len(s) == 8:
        s = s[2:]
    if len(s) != 6:
        return default
    try:
        return tuple(int(s[i : i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return default


def font_color(cell, default=(17, 19, 24)):
    color = getattr(cell.font, "color", None)
    rgb = getattr(color, "rgb", None) if color else None
    if not rgb:
        return default
    s = str(rgb)
    if len(s) == 8:
        s = s[2:]
    if len(s) != 6:
        return default
    try:
        return tuple(int(s[i : i + 2], 16) for i in (0, 2, 4))
    except ValueError:
        return default


def display_value(cell) -> str:
    v = cell.value
    if v is None:
        return ""
    nf = cell.number_format or "General"
    if isinstance(v, (int, float, Decimal)) and nf and nf != "General":
        if nf == "#,##0":
            return f"{int(round(float(v))):,}"
        if nf == "#,##0.00":
            return f"{Decimal(str(v)).quantize(Decimal('0.01')):,.2f}"
    return str(v)


def sheets_screenshot(
    wb,
    sheet_name: str,
    path: Path,
    max_data_rows: int | None = None,
    extra_note_rows: int = 0,
    width: int = 1600,
) -> Path:
    ws = wb[sheet_name]
    # Determine used range: up to last interesting row (skip trailing empty)
    last_row = ws.max_row
    last_col = ws.max_column
    # For Review, include the summary and a couple of note rows.
    if sheet_name == "Review":
        last_row = 1
        for r in range(2, ws.max_row + 1):
            vals = [ws.cell(r, c).value for c in range(1, last_col + 1)]
            if all(v is None for v in vals):
                continue
            if vals[0] == "Checks":
                break
            last_row = r
    elif max_data_rows:
        last_row = min(ws.max_row, 1 + max_data_rows)

    n_rows = last_row
    n_cols = last_col

    # Column widths from the worksheet, converted to pixels.
    col_px = []
    for i in range(1, n_cols + 1):
        letter = ws.cell(1, i).column_letter
        w = ws.column_dimensions[letter].width or 12
        col_px.append(max(64, int(w * 8.2)))

    row_h = 34
    head_h = 44 if sheet_name == "Review" else 34
    row_num_w = 48
    col_head_h = 28
    tab_h = 44
    title_h = 52
    pad = 28

    table_w = row_num_w + sum(col_px)
    table_h = col_head_h + head_h + row_h * (n_rows - 1)
    # Scale columns to fill a target width if slightly under/over
    inner_w = width - pad * 2
    if table_w != inner_w:
        scale = (inner_w - row_num_w) / sum(col_px)
        col_px = [max(56, int(c * scale)) for c in col_px]
        # Fix rounding
        col_px[-1] += (inner_w - row_num_w) - sum(col_px)
        table_w = row_num_w + sum(col_px)

    height = pad + title_h + table_h + tab_h + pad
    im = Image.new("RGB", (width, height), (241, 243, 244))
    d = ImageDraw.Draw(im)

    # Window / toolbar
    d.rectangle((0, 0, width, pad + title_h), fill=(255, 255, 255))
    d.line((0, pad + title_h, width, pad + title_h), fill=SHEET_GRID, width=1)
    # Sheets-like green mark + filename
    d.rounded_rectangle((pad, pad + 10, pad + 22, pad + 32), 3, fill=SHEET_GREEN)
    d.text((pad + 5, pad + 11), "S", font=F("bd", 14), fill=WHITE)
    d.text((pad + 32, pad + 10), "extracted_invoices.xlsx", font=F("sm", 18), fill=INK)
    d.text((pad + 280, pad + 14), "Google Sheets-style preview  ·  values from the live workbook", font=F("reg", 12), fill=(107, 114, 128))

    ox, oy = pad, pad + title_h
    # Column letters
    d.rectangle((ox, oy, ox + table_w, oy + col_head_h), fill=SHEET_HEAD)
    d.rectangle((ox, oy, ox + row_num_w, oy + col_head_h), fill=SHEET_HEAD)
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    x = ox + row_num_w
    for i, cw in enumerate(col_px):
        d.line((x, oy, x, oy + col_head_h + table_h - col_head_h), fill=SHEET_GRID, width=1)
        lab = letters[i]
        tw = d.textlength(lab, font=F("sm", 11))
        d.text((x + (cw - tw) / 2, oy + 7), lab, font=F("sm", 11), fill=(95, 99, 104))
        x += cw
    d.line((ox, oy + col_head_h, ox + table_w, oy + col_head_h), fill=SHEET_GRID, width=1)
    d.line((ox + row_num_w, oy, ox + row_num_w, oy + col_head_h + (table_h - col_head_h)), fill=SHEET_GRID, width=1)

    # Data rows
    y = oy + col_head_h
    for r in range(1, n_rows + 1):
        this_h = head_h if r == 1 else row_h
        # Row number
        d.rectangle((ox, y, ox + row_num_w, y + this_h), fill=SHEET_HEAD)
        rn = str(r)
        tw = d.textlength(rn, font=F("reg", 11))
        d.text((ox + (row_num_w - tw) / 2, y + this_h / 2 - 8), rn, font=F("reg", 11), fill=(95, 99, 104))
        d.line((ox, y + this_h, ox + table_w, y + this_h), fill=SHEET_GRID, width=1)

        x = ox + row_num_w
        for c, cw in enumerate(col_px, 1):
            cell = ws.cell(r, c)
            text = display_value(cell)
            bg = cell_rgb(cell, default=None)
            if bg is None:
                bg = HEAD if r == 1 else ((247, 248, 250) if r % 2 == 0 else (255, 255, 255))
            # For header we already set HEAD in workbook
            d.rectangle((x, y, x + cw, y + this_h), fill=bg)
            d.line((x, y, x, y + this_h), fill=SHEET_GRID, width=1)

            fill = font_color(cell, default=WHITE if r == 1 else INK)
            # Header wrapping
            font = F("sm", 11) if r == 1 else (F("bd", 12) if (cell.font and cell.font.bold) else F("reg", 13))
            if r == 1:
                font = F("sm", 11)
                fill = WHITE
            # Alignment
            align = "left"
            if r == 1:
                align = "center"
            elif getattr(cell.alignment, "horizontal", None) == "right":
                align = "right"
            elif getattr(cell.alignment, "horizontal", None) == "center":
                align = "center"
            elif isinstance(cell.value, (int, float, Decimal)):
                align = "right"

            # Truncate
            max_w = cw - 14
            display = text
            if r != 1:
                while display and d.textlength(display, font=font) > max_w:
                    display = display[:-2] + "…"
            else:
                # wrap header into two lines if needed
                if d.textlength(display, font=font) > max_w and " " in display:
                    parts = display.rsplit(" ", 1)
                    # fall through to two-line draw
                    mid_x = x + cw / 2
                    t1, t2 = parts[0], parts[1]
                    # If still long, just shrink
                    d.text((x + (cw - d.textlength(t1, font=font)) / 2, y + 4), t1, font=font, fill=fill)
                    d.text((x + (cw - d.textlength(t2, font=font)) / 2, y + 20), t2, font=font, fill=fill)
                    x += cw
                    continue

            tw = d.textlength(display, font=font)
            if align == "right":
                tx = x + cw - 8 - tw
            elif align == "center":
                tx = x + (cw - tw) / 2
            else:
                tx = x + 8
            ty = y + this_h / 2 - 8
            d.text((tx, ty), display, font=font, fill=fill)
            x += cw
        y += this_h

    # Bottom tab bar (Google Sheets style)
    tab_y = y
    d.rectangle((0, tab_y, width, height), fill=(255, 255, 255))
    d.line((0, tab_y, width, tab_y), fill=SHEET_GRID, width=1)
    tabs = ["Invoices", "Line items", "Review"]
    tx = pad + 12
    for name in tabs:
        active = name == sheet_name
        tw = d.textlength(name, font=F("sm" if active else "reg", 14))
        tab_w = tw + 28
        if active:
            d.rectangle((tx, tab_y + tab_h - 3, tx + tab_w, tab_y + tab_h), fill=SHEET_GREEN)
            d.text((tx + 14, tab_y + 10), name, font=F("sm", 14), fill=SHEET_GREEN)
        else:
            d.text((tx + 14, tab_y + 10), name, font=F("reg", 14), fill=(95, 99, 104))
        tx += tab_w + 8

    # Outer white card margin already the background; add a subtle border around the grid
    d.rectangle((ox, oy, ox + table_w, tab_y), outline=SHEET_GRID, width=1)

    path.parent.mkdir(parents=True, exist_ok=True)
    im.save(path, "PNG", optimize=True)
    print(f"wrote {path} {im.size}")
    return path


def upwork_portfolio(pages: list[Image.Image], data: dict, shots: dict[str, Image.Image], path: Path) -> Path:
    W, H = 1600, 1200
    im = bg_dark((W, H))
    d = ImageDraw.Draw(im)

    t1, t2 = "PDF invoices", " to Excel / Google Sheets"
    f_title = F("bd", 42)
    w1 = d.textlength(t1, font=f_title)
    w2 = d.textlength(t2, font=f_title)
    tx = (W - (w1 + w2)) // 2
    d.text((tx, 22), t1, font=f_title, fill=GOLD)
    d.text((tx + w1, 22), t2, font=f_title, fill=WHITE)
    sub = "Five fictional sample invoices in  ·  one checked workbook out"
    f_sub = F("reg", 16)
    sw = d.textlength(sub, font=f_sub)
    d.text(((W - sw) // 2, 70), sub, font=f_sub, fill=MUTE)

    # Five invoice thumbnails, tightly cropped
    thumbs_top = 102
    thumbs_h = 250
    gap = 14
    side = 28
    usable = W - side * 2
    tw = (usable - gap * 4) // 5
    x = side
    labels = [f"{inv['Currency']}  ·  {inv['Invoice number']}" for inv in data["invoices"]]
    for page, lab in zip(pages, labels):
        cropped = crop_bottom_whitespace(page, pad=28)
        thumb = scale_to_width(cropped, tw - 2)
        max_th = thumbs_h - 22
        if thumb.height > max_th:
            thumb = scale_to_height(thumb, max_th)
        thumb = apply_round(thumb.convert("RGB"), 10)
        sh = drop_shadow(thumb, blur=12, offset=(0, 7), opacity=150)
        px = x + (tw - sh.width) // 2
        py = thumbs_top + 4
        im.paste(sh, (px, py - 16), sh)
        lw = d.textlength(lab, font=F("sm", 12))
        d.text((x + (tw - lw) / 2, thumbs_top + thumbs_h - 2), lab, font=F("sm", 12), fill=(210, 214, 220))
        x += tw + gap

    # Full-width stacked sheets so the cells stay readable
    sheet_top = thumbs_top + thumbs_h + 22
    sheet_bot = H - 20
    avail_h = sheet_bot - sheet_top
    gap_s = 14
    # Give Review a little more of the remaining height (more columns).
    h1 = int(avail_h * 0.40) - gap_s // 2
    h2 = avail_h - h1 - gap_s
    max_w = W - 36

    def fit_shot(shot: Image.Image, max_w: int, max_h: int) -> Image.Image:
        s = shot.convert("RGB")
        scale = min(max_w / s.width, max_h / s.height)
        nw, nh = int(s.width * scale), int(s.height * scale)
        s = s.resize((nw, nh), Image.Resampling.LANCZOS)
        s = apply_round(s, 12)
        return drop_shadow(s, blur=14, offset=(0, 8), opacity=150)

    top = fit_shot(shots["invoices"], max_w, h1 + 30)
    bot = fit_shot(shots["review"], max_w, h2 + 30)
    im.paste(top, ((W - top.width) // 2, sheet_top - 16), top)
    im.paste(bot, ((W - bot.width) // 2, sheet_top - 16 + top.height - 20 + gap_s), bot)

    im = im.convert("RGB")
    path.parent.mkdir(parents=True, exist_ok=True)
    im.save(path, "PNG", optimize=True)
    print(f"wrote {path} {im.size}")
    return path



def main() -> None:
    IMG_DIR.mkdir(parents=True, exist_ok=True)
    data = load_workbook_rows()
    # Sanity: 5 invoices, all review rows PASS
    assert len(data["invoices"]) == 5, data["invoices"]
    assert data["overall"] == "ALL PASSED", data["overall"]
    pages = [crop_bottom_whitespace(p, pad=28) for p in render_pdfs(dpi=200)]

    fiverr = fiverr_cover(pages, data, IMG_DIR / "fiverr-gig-cover-1280x769.png")

    # Screenshots from the live workbook (before the portfolio, so we can nest them)
    inv_shot = sheets_screenshot(
        data["wb"], "Invoices", IMG_DIR / "sheet-invoices.png", max_data_rows=5, width=1600
    )
    items_shot = sheets_screenshot(
        data["wb"], "Line items", IMG_DIR / "sheet-line-items.png", max_data_rows=19, width=1600
    )
    review_shot = sheets_screenshot(
        data["wb"],
        "Review",
        IMG_DIR / "sheet-review.png",
        extra_note_rows=0,
        width=1840,
    )

    upwork_portfolio(
        pages,
        data,
        {
            "invoices": Image.open(inv_shot),
            "review": Image.open(review_shot),
        },
        IMG_DIR / "upwork-portfolio-1600x1200.png",
    )

    # Confirm exact pixel sizes for the two specified images
    for p, size in (
        (fiverr, (1280, 769)),
        (IMG_DIR / "upwork-portfolio-1600x1200.png", (1600, 1200)),
    ):
        im = Image.open(p)
        assert im.size == size, f"{p} is {im.size}, expected {size}"
        print(f"ok  {p.name} {im.size}")


if __name__ == "__main__":
    main()
