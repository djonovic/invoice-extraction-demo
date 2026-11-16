#!/usr/bin/env python3
"""Extract invoice fields from sample PDFs with pdfplumber and regex.

Reads every PDF in input/ (or paths passed on the CLI) and writes a workbook
with three sheets:

  Invoices    — one row per document (vendor, number, date, net, VAT, total)
  Line items  — every line item, keyed to its invoice number
  Review      — checks: net + VAT = total, date parsed, duplicate invoice numbers

An LLM is not used. Totals are compared with Decimal arithmetic.
"""
from __future__ import annotations

import argparse
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pdfplumber
from dateutil import parser as date_parser
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parent
INPUT_DIR = ROOT / "input"
OUTPUT_XLSX = ROOT / "output" / "extracted_invoices.xlsx"

# Currency-aware money: optional symbol, optional thousands separators, optional cents.
MONEY_RE = re.compile(
    r"(?:A\$|\$|€|£|¥)?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})|[0-9]+(?:\.[0-9]{2})|[0-9]{1,3}(?:,[0-9]{3})+|[0-9]+)"
)

DATE_FORMATS = [
    "%d %b %Y",
    "%d %B %Y",
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%d.%m.%Y",
    "%d/%m/%Y",
    "%m/%d/%Y",
]


def to_decimal(raw: str, currency: str | None = None) -> Decimal:
    s = raw.replace("A$", "").replace("$", "").replace("€", "").replace("£", "").replace("¥", "")
    s = s.replace(" ", "").replace(",", "")
    try:
        value = Decimal(s)
    except InvalidOperation as exc:
        raise ValueError(f"not a number: {raw!r}") from exc
    if currency == "JPY":
        return value.quantize(Decimal("1"))
    return value.quantize(Decimal("0.01"))


def money_str(value: Decimal, currency: str) -> str:
    if currency == "JPY":
        return f"{int(value):,}"
    return f"{value.quantize(Decimal('0.01')):,.2f}"


def extract_text_and_tables(pdf_path: Path) -> tuple[str, list]:
    with pdfplumber.open(pdf_path) as pdf:
        texts = []
        tables = []
        for page in pdf.pages:
            texts.append(page.extract_text(layout=True) or page.extract_text() or "")
            try:
                tables.extend(page.extract_tables() or [])
            except Exception:
                pass
    return "\n".join(texts), tables


OTHER_LABEL = r"(?:Invoice\s+Number|Invoice\s+No\.?|Invoice\s*#|Tax Date|Invoice Date|Currency|Bill To|Supplier|Vendor|Date|From|To|Net|Total|VAT)\s*:"

# Typical invoice numbers contain letters, a hyphen and at least one digit: INV-1042, BH-2026-0887.
INVOICE_ID = r"([A-Z]{1,8}-[A-Z0-9][A-Z0-9\-/]*\d[A-Z0-9\-/]*)"


def first_match(text: str, patterns: list[str]) -> str | None:
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE | re.M)
        if m:
            value = m.group(1).strip()
            # Stop at a trailing second label on the same line.
            value = re.split(r"\s{2,}" + OTHER_LABEL, value)[0]
            value = re.split(r"\s{2,}Currency:\s*", value)[0]
            value = value.strip(" :-")
            if value and not re.fullmatch(OTHER_LABEL.rstrip(":").strip(), value, re.I):
                return value
    return None


def label_then_value(text: str, labels: list[str], value_pat: str) -> str | None:
    """Match 'Label: value' on one line, or 'Label:' with the value on the next line."""
    label = r"(?:%s)" % "|".join(labels)
    same = re.search(rf"(?:^|\s){label}\s*:\s*({value_pat})", text, re.I | re.M)
    if same:
        return same.group(1).strip()
    nxt = re.search(
        rf"(?:^|\s){label}\s*:\s*$\s*^\s*({value_pat})",
        text,
        re.I | re.M,
    )
    if nxt:
        return nxt.group(1).strip()
    return None


def parse_fields(text: str) -> dict:
    vendor = first_match(
        text,
        [
            r"Vendor:\s*(.+)$",
            r"Supplier:\s*(.+)$",
        ],
    )
    invoice_no = label_then_value(
        text,
        [r"Invoice\s+Number", r"Invoice\s+No\.?", r"Invoice\s*#"],
        r"[A-Z]{1,8}-[A-Z0-9][A-Z0-9\-/]*\d[A-Z0-9\-/]*",
    )
    if not invoice_no:
        near = re.search(
            rf"(?:Invoice\s+(?:Number|No\.?)|Invoice\s*#).{{0,80}}?{INVOICE_ID}",
            text,
            re.I | re.S,
        )
        if near:
            invoice_no = near.group(1)
    date_raw = None
    for lab in [r"Tax Date", r"Invoice Date", r"Date"]:
        m = re.search(rf"(?:^|\s){lab}:\s*(.+)$", text, re.I | re.M)
        if not m:
            continue
        cand = m.group(1).strip()
        cand = re.split(r"\s{2,}" + OTHER_LABEL, cand)[0].strip()
        cand = re.split(r"\s{2,}Currency:", cand)[0].strip()
        if cand and not re.match(OTHER_LABEL, cand, re.I):
            date_raw = cand
            break
        # Two-line or two-column meta: the value sits on the next line.
        m2 = re.search(rf"(?:^|\s){lab}:\s*(?:{OTHER_LABEL})?.*$\s*^\s*(\S.+)$", text, re.I | re.M)
        if m2:
            cand = m2.group(1).strip()
            cand = re.split(r"\s{2,}", cand)[0].strip()
            if cand and not re.match(OTHER_LABEL, cand, re.I):
                date_raw = cand
                break
    currency = first_match(text, [r"Currency:\s*([A-Z]{3})"])

    if not vendor or not invoice_no or not date_raw or not currency:
        missing = [
            name
            for name, val in [
                ("vendor", vendor),
                ("invoice_number", invoice_no),
                ("date", date_raw),
                ("currency", currency),
            ]
            if not val
        ]
        raise ValueError(f"missing fields {missing}")

    # Totals: prefer explicit Net / VAT / Total labels, then synonyms.
    net = find_amount(
        text,
        [
            r"Net Amount:\s*{m}",
            r"(?<![A-Za-z])Net:\s*{m}",
            r"Subtotal\s*\(ex GST\):\s*{m}",
            r"Subtotal:\s*{m}",
        ],
        currency,
    )
    vat = find_amount(
        text,
        [
            r"VAT\s*\((?:GST\s*)?[^)]*\)\s*:?\s*{m}",
            r"VAT\s*@\s*\d+%\s*:?\s*{m}",
            r"VAT\s*\(\d+%\)\s*:?\s*{m}",
            r"Consumption Tax\s*\(\d+%\)\s*:?\s*{m}",
            r"GST\s*\(\d+%\)\s*:?\s*{m}",
            r"(?<![A-Za-z])VAT:\s*{m}",
            r"(?<![A-Za-z])Tax\s*\(\d+%\)\s*:?\s*{m}",
        ],
        currency,
    )
    total = find_amount(
        text,
        [
            r"Amount Payable:\s*{m}",
            r"Amount Due:\s*{m}",
            r"Total due:\s*{m}",
            r"(?<!Sub)(?<!Amount )(?<!Line items )Total:\s*{m}",
            r"(?<!Sub)Total:\s*{m}",
        ],
        currency,
    )
    return {
        "vendor": vendor,
        "invoice_number": invoice_no,
        "date_raw": date_raw,
        "currency": currency,
        "net": net,
        "vat": vat,
        "total": total,
    }


def find_amount(text: str, patterns: list[str], currency: str) -> Decimal:
    money = MONEY_RE.pattern
    # MONEY_RE has a capturing group; wrap carefully.
    for pat in patterns:
        # {m} is replaced with a currency-aware money pattern that captures the number.
        filled = pat.format(m=r"(?:A\$|\$|€|£|¥)?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})?|[0-9]+(?:\.[0-9]{2})|[0-9]+)")
        m = re.search(filled, text, re.IGNORECASE | re.M)
        if m:
            return to_decimal(m.group(1), currency)
    raise ValueError(f"could not find amount with {patterns[0]}")


def parse_date(raw: str) -> tuple[datetime | None, bool]:
    cleaned = raw.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(cleaned, fmt), True
        except ValueError:
            continue
    try:
        # dayfirst=True so 03/11/2026 is 3 November (AU / EU), not 11 March.
        dt = date_parser.parse(cleaned, dayfirst=True, yearfirst=False)
        return dt, True
    except (ValueError, OverflowError, TypeError):
        return None, False


def looks_like_header(cell: str) -> bool:
    s = re.sub(r"\s+", " ", (cell or "")).strip().lower()
    return s in {"description", "qty", "quantity", "unit price", "unit", "amount", "total"}


def parse_line_items(tables: list, text: str, currency: str) -> list[dict]:
    items = []
    for table in tables:
        if not table or len(table) < 2:
            continue
        header = [re.sub(r"\s+", " ", (c or "")).strip().lower() for c in table[0]]
        if "description" not in header:
            # Sometimes the header is not the first row, or cells are merged oddly.
            if not any(looks_like_header(c) for c in table[0]):
                continue
        # Map columns
        def col(*names):
            for n in names:
                if n in header:
                    return header.index(n)
            return None

        i_desc = col("description")
        i_qty = col("qty", "quantity")
        i_unit = col("unit price", "unit")
        i_amt = col("amount")
        if i_desc is None or i_amt is None:
            continue
        for row in table[1:]:
            if not row or not any(row):
                continue
            desc = re.sub(r"\s+", " ", (row[i_desc] or "")).strip()
            if not desc or looks_like_header(desc):
                continue
            # Skip total rows that leaked into the table.
            if re.search(r"^(net|vat|tax|gst|subtotal|total|amount due|amount payable|consumption tax)\b", desc, re.I):
                continue
            try:
                qty_raw = row[i_qty] if i_qty is not None else "1"
                unit_raw = row[i_unit] if i_unit is not None else None
                amt_raw = row[i_amt]
                if not amt_raw:
                    continue
                qty = Decimal(re.sub(r"[^\d.]", "", str(qty_raw) or "1") or "1")
                amount = to_decimal(str(amt_raw), currency)
                if unit_raw:
                    unit = to_decimal(str(unit_raw), currency)
                else:
                    unit = (amount / qty).quantize(Decimal("1") if currency == "JPY" else Decimal("0.01"))
            except (InvalidOperation, ValueError, ZeroDivisionError):
                continue
            items.append(
                {
                    "description": desc,
                    "qty": int(qty),
                    "unit_price": unit,
                    "amount": amount,
                }
            )
        if items:
            return items

    # Fallback: regex on layout text between a Description header and a totals label.
    header = re.search(r"Description\s+Qty\s+Unit Price\s+Amount", text, re.I)
    tail = re.search(r"(Net Amount:|Net:|Subtotal(?:\s*\(ex GST\))?:|Amount Due:)", text, re.I)
    if not header:
        return items
    block = text[header.end() : tail.start() if tail else len(text)]
    money = r"(?:A\$|\$|€|£|¥)?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})?|[0-9]+(?:\.[0-9]{2})|[0-9]+)"
    item_re = re.compile(rf"^(.+?)\s{{2,}}(\d+)\s+{money}\s+{money}\s*$", re.M)
    for m in item_re.finditer(block):
        desc = re.sub(r"\s+", " ", m.group(1)).strip()
        if not desc or re.search(r"^(vendor|invoice|date|currency|bill to|from|to|supplier)\b", desc, re.I):
            continue
        items.append(
            {
                "description": desc,
                "qty": int(m.group(2)),
                "unit_price": to_decimal(m.group(3), currency),
                "amount": to_decimal(m.group(4), currency),
            }
        )
    return items


def extract_pdf(pdf_path: Path) -> dict:
    text, tables = extract_text_and_tables(pdf_path)
    fields = parse_fields(text)
    items = parse_line_items(tables, text, fields["currency"])
    parsed_date, date_ok = parse_date(fields["date_raw"])
    line_sum = sum((it["amount"] for it in items), Decimal("0"))
    currency = fields["currency"]
    if currency == "JPY":
        line_sum = line_sum.quantize(Decimal("1"))
    else:
        line_sum = line_sum.quantize(Decimal("0.01"))

    net, vat, total = fields["net"], fields["vat"], fields["total"]
    totals_ok = (net + vat) == total
    lines_ok = (line_sum == net) if items else False

    return {
        "source_pdf": pdf_path.name,
        "vendor": fields["vendor"],
        "invoice_number": fields["invoice_number"],
        "date_raw": fields["date_raw"],
        "date_parsed": parsed_date.date().isoformat() if parsed_date else None,
        "date_ok": date_ok,
        "currency": currency,
        "net": net,
        "vat": vat,
        "total": total,
        "totals_ok": totals_ok,
        "line_items": items,
        "line_items_sum": line_sum,
        "line_items_sum_ok": lines_ok,
        "text": text,
    }


def write_xlsx(invoices: list[dict], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    thin = Border(
        left=Side(style="thin", color="D6DBE3"),
        right=Side(style="thin", color="D6DBE3"),
        top=Side(style="thin", color="D6DBE3"),
        bottom=Side(style="thin", color="D6DBE3"),
    )
    head_fill = PatternFill("solid", fgColor="111318")
    head_font = Font(name="Calibri", bold=True, color="F8FAFC", size=11)
    body = Font(name="Calibri", size=11)
    bold = Font(name="Calibri", bold=True, size=11)
    alt = PatternFill("solid", fgColor="F7F8FA")
    green = PatternFill("solid", fgColor="D1FAE5")
    green_font = Font(name="Calibri", bold=True, color="065F46", size=11)
    red = PatternFill("solid", fgColor="FECACA")
    red_font = Font(name="Calibri", bold=True, color="7F1D1D", size=11)
    pass_fill = PatternFill("solid", fgColor="10B981")
    fail_fill = PatternFill("solid", fgColor="EF4444")
    white_bold = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    wrap = Alignment(wrap_text=True, vertical="center")
    left = Alignment(vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    center = Alignment(horizontal="center", vertical="center")

    def style_header(ws, row, ncols):
        ws.row_dimensions[row].height = 22
        for c in range(1, ncols + 1):
            cell = ws.cell(row, c)
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(horizontal="center" if c > 1 else "left", vertical="center")
            cell.border = thin

    def style_body(cell, row_i, align="left"):
        cell.font = body
        cell.border = thin
        cell.alignment = {"left": left, "right": right, "center": center}[align]
        if row_i % 2 == 1:
            cell.fill = alt

    # ----- Invoices -----
    ws = wb.active
    ws.title = "Invoices"
    headers = [
        "Vendor",
        "Invoice number",
        "Date (on invoice)",
        "Date parsed",
        "Currency",
        "Net",
        "VAT",
        "Total",
        "Source PDF",
    ]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)
    style_header(ws, 1, len(headers))

    for i, inv in enumerate(invoices):
        r = i + 2
        values = [
            inv["vendor"],
            inv["invoice_number"],
            inv["date_raw"],
            inv["date_parsed"] or "",
            inv["currency"],
            inv["net"],
            inv["vat"],
            inv["total"],
            inv["source_pdf"],
        ]
        for c, val in enumerate(values, 1):
            cell = ws.cell(r, c, val)
            if c in (6, 7, 8):
                if inv["currency"] == "JPY":
                    cell.number_format = "#,##0"
                else:
                    cell.number_format = "#,##0.00"
                style_body(cell, i, "right")
            elif c == 5:
                style_body(cell, i, "center")
            else:
                style_body(cell, i, "left")
        ws.row_dimensions[r].height = 20

    widths = [34, 16, 20, 14, 10, 12, 12, 12, 38]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{1 + len(invoices)}"
    ws.sheet_properties.tabColor = "111318"

    # ----- Line items -----
    ws2 = wb.create_sheet("Line items")
    headers2 = [
        "Invoice number",
        "Vendor",
        "Description",
        "Qty",
        "Unit price",
        "Amount",
        "Currency",
    ]
    for c, h in enumerate(headers2, 1):
        ws2.cell(1, c, h)
    style_header(ws2, 1, len(headers2))

    r = 2
    row_i = 0
    for inv in invoices:
        for it in inv["line_items"]:
            vals = [
                inv["invoice_number"],
                inv["vendor"],
                it["description"],
                it["qty"],
                it["unit_price"],
                it["amount"],
                inv["currency"],
            ]
            for c, val in enumerate(vals, 1):
                cell = ws2.cell(r, c, val)
                if c in (5, 6):
                    cell.number_format = "#,##0" if inv["currency"] == "JPY" else "#,##0.00"
                    style_body(cell, row_i, "right")
                elif c in (4, 7):
                    style_body(cell, row_i, "center")
                else:
                    style_body(cell, row_i, "left")
            ws2.row_dimensions[r].height = 20
            r += 1
            row_i += 1

    widths2 = [16, 34, 40, 8, 12, 12, 10]
    for i, width in enumerate(widths2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = width
    ws2.freeze_panes = "A2"
    last_item_row = max(r - 1, 1)
    if last_item_row > 1:
        ws2.auto_filter.ref = f"A1:G{last_item_row}"
    ws2.sheet_properties.tabColor = "16697A"

    # ----- Review -----
    # Duplicates across the batch
    counts: dict[str, int] = {}
    for inv in invoices:
        key = inv["invoice_number"].strip().upper()
        counts[key] = counts.get(key, 0) + 1

    ws3 = wb.create_sheet("Review")
    headers3 = [
        "Invoice number",
        "Vendor",
        "Date (on invoice)",
        "Date parsed",
        "Date OK",
        "Currency",
        "Net",
        "VAT",
        "Total",
        "Net + VAT",
        "Net + VAT = Total",
        "Line items sum",
        "Line items = Net",
        "Duplicate number",
        "Overall",
    ]
    for c, h in enumerate(headers3, 1):
        ws3.cell(1, c, h)
    style_header(ws3, 1, len(headers3))

    def flag_cell(cell, ok: bool, yes="PASS", no="FAIL"):
        cell.value = yes if ok else no
        cell.fill = green if ok else red
        cell.font = green_font if ok else red_font
        cell.alignment = center
        cell.border = thin

    for i, inv in enumerate(invoices):
        r = i + 2
        key = inv["invoice_number"].strip().upper()
        is_dup = counts[key] > 1
        net_plus_vat = inv["net"] + inv["vat"]
        overall = inv["totals_ok"] and inv["date_ok"] and not is_dup and inv["line_items_sum_ok"]

        ws3.cell(r, 1, inv["invoice_number"])
        ws3.cell(r, 2, inv["vendor"])
        ws3.cell(r, 3, inv["date_raw"])
        ws3.cell(r, 4, inv["date_parsed"] or "")
        flag_cell(ws3.cell(r, 5), inv["date_ok"])
        ws3.cell(r, 6, inv["currency"])
        ws3.cell(r, 7, inv["net"])
        ws3.cell(r, 8, inv["vat"])
        ws3.cell(r, 9, inv["total"])
        ws3.cell(r, 10, net_plus_vat)
        flag_cell(ws3.cell(r, 11), inv["totals_ok"])
        ws3.cell(r, 12, inv["line_items_sum"])
        flag_cell(ws3.cell(r, 13), inv["line_items_sum_ok"])
        flag_cell(ws3.cell(r, 14), not is_dup, yes="No", no="YES")
        overall_cell = ws3.cell(r, 15, "PASS" if overall else "FAIL")
        overall_cell.fill = pass_fill if overall else fail_fill
        overall_cell.font = white_bold
        overall_cell.alignment = center
        overall_cell.border = thin

        for c in range(1, 16):
            cell = ws3.cell(r, c)
            cell.border = thin
            if c in (7, 8, 9, 10, 12):
                cell.number_format = "#,##0" if inv["currency"] == "JPY" else "#,##0.00"
                cell.alignment = right
                cell.font = body
            elif c in (5, 6, 11, 13, 14, 15):
                cell.alignment = center
                if c == 6:
                    cell.font = body
            else:
                cell.alignment = left
                cell.font = body
        ws3.row_dimensions[r].height = 22

    # Summary row
    summary_row = 2 + len(invoices)
    n = len(invoices)
    n_pass = sum(
        1
        for inv in invoices
        if inv["totals_ok"]
        and inv["date_ok"]
        and inv["line_items_sum_ok"]
        and counts[inv["invoice_number"].strip().upper()] == 1
    )
    n_dup = sum(1 for v in counts.values() if v > 1)
    ws3.cell(summary_row, 1, "Summary")
    ws3.cell(summary_row, 2, f"{n_pass} of {n} invoices passed every check")
    ws3.cell(summary_row, 15, "ALL PASSED" if n_pass == n else "ISSUES FOUND")
    for c in range(1, 16):
        cell = ws3.cell(summary_row, c)
        cell.font = bold
        cell.fill = green if n_pass == n else red
        cell.alignment = left if c < 15 else center
        cell.border = thin
    ws3.cell(summary_row, 15).font = white_bold
    ws3.cell(summary_row, 15).fill = pass_fill if n_pass == n else fail_fill
    ws3.row_dimensions[summary_row].height = 22

    # Note
    note_row = summary_row + 2
    ws3.cell(note_row, 1, "Checks")
    ws3.cell(note_row, 1).font = bold
    notes = [
        "Date OK — the date printed on the invoice was parsed to ISO (YYYY-MM-DD). Day-first is used for numeric dates (03/11/2026 = 3 November 2026).",
        "Net + VAT = Total — Decimal addition, exact match. JPY values have no cents.",
        "Line items = Net — the sum of extracted line-item amounts equals the invoice net.",
        "Duplicate number — FAIL if the same invoice number appears more than once in this batch.",
        "All companies and amounts in the source PDFs are fictional sample data.",
    ]
    for j, note in enumerate(notes):
        ws3.cell(note_row + 1 + j, 1, note)
        ws3.merge_cells(start_row=note_row + 1 + j, start_column=1, end_row=note_row + 1 + j, end_column=15)
        ws3.cell(note_row + 1 + j, 1).font = Font(name="Calibri", size=10, italic=True, color="4B5563")

    widths3 = [16, 32, 20, 14, 10, 10, 12, 12, 12, 12, 18, 14, 16, 18, 12]
    for i, width in enumerate(widths3, 1):
        ws3.column_dimensions[get_column_letter(i)].width = width
    ws3.freeze_panes = "A2"
    ws3.auto_filter.ref = f"A1:O{1 + len(invoices)}"
    ws3.sheet_properties.tabColor = "10B981"
    ws3.row_dimensions[1].height = 28
    for col in range(1, 16):
        ws3.cell(1, col).alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")

    ws.page_setup.orientation = "landscape"
    ws2.page_setup.orientation = "landscape"
    ws3.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1

    wb.save(path)
    return path


def main() -> list[dict]:
    p = argparse.ArgumentParser(description="Extract fictional sample invoices to a checked workbook")
    p.add_argument("pdfs", nargs="*", type=Path, help="PDF files (default: input/*.pdf)")
    p.add_argument("--out", type=Path, default=OUTPUT_XLSX)
    args = p.parse_args()
    pdfs = [p.resolve() for p in args.pdfs] if args.pdfs else sorted(INPUT_DIR.glob("*.pdf"))
    if not pdfs:
        raise SystemExit(f"no PDFs found in {INPUT_DIR}")

    invoices = []
    errors = []
    for pdf in pdfs:
        try:
            inv = extract_pdf(pdf)
            invoices.append(inv)
            status = "OK" if inv["totals_ok"] and inv["date_ok"] and inv["line_items_sum_ok"] else "CHECK"
            print(
                f"{status:5} {pdf.name:44}  {inv['invoice_number']:14}  "
                f"{inv['currency']}  net={inv['net']}  vat={inv['vat']}  total={inv['total']}  "
                f"items={len(inv['line_items'])}  date={inv['date_parsed']}"
            )
            if not inv["line_items"]:
                print(f"      WARNING: no line items extracted from {pdf.name}")
            if not inv["totals_ok"]:
                print(f"      WARNING: net+VAT != total ({inv['net']} + {inv['vat']} != {inv['total']})")
            if not inv["line_items_sum_ok"]:
                print(f"      WARNING: line items sum {inv['line_items_sum']} != net {inv['net']}")
        except Exception as exc:
            errors.append((pdf, exc))
            print(f"FAIL  {pdf.name}: {exc}")

    if not invoices:
        raise SystemExit("no invoices extracted")

    out = write_xlsx(invoices, args.out.resolve())
    print(f"wrote {out}")
    if errors:
        raise SystemExit(f"{len(errors)} PDF(s) failed")
    return invoices


if __name__ == "__main__":
    main()
