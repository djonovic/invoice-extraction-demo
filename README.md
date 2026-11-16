Live demo: https://djonovic.github.io/invoice-extraction-demo/

This is a demo on fictional sample invoices, not client work.

# Invoice extraction demo

Python pipeline that turns sample invoice PDFs into a checked Excel workbook. Built as a before/after demo for Pasko Djonovic’s Fiverr gig and Upwork portfolio.

All companies, addresses and amounts are fictional sample data, clearly marked on every PDF.

## What it does

1. `generate_sample_invoices.py` writes five one-page invoices with different layouts, vendors and currencies (USD, EUR, GBP, JPY, AUD). Each invoice has vendor, invoice number, date, line items, net, VAT and total. Totals are computed with `Decimal` so net + VAT equals total exactly.
2. `extract_invoices.py` reads every PDF with pdfplumber and regex (no LLM, no API spend). It writes `output/extracted_invoices.xlsx` with three sheets:
   - **Invoices** — one row per document (vendor, number, original date, parsed ISO date, currency, net, VAT, total).
   - **Line items** — every line, keyed to its invoice number.
   - **Review** — flags for net + VAT = total, date parsed, line items summing to net, and duplicate invoice numbers.
3. `render_demo_images.py` builds the gig cover, Upwork portfolio image and sheet screenshots from the real PDF pages and the live workbook. Numbers on the images are not typed in by hand.

In this run all five invoices passed every Review check. Dates such as `03/11/2026` are parsed day-first (3 November 2026). Yen amounts have no cents.

## Upwork portfolio entry (3 lines)

I extracted five fictional sample invoices (USD, EUR, GBP, JPY and AUD) from PDF into one Excel workbook with Python. The file has an Invoices sheet, a Line items sheet, and a Review tab that checks net + VAT = total, parses dates to ISO, and flags duplicate invoice numbers. All five invoices passed every check; the companies and amounts are sample data only.

## Regenerate

```bash
cd /workspace/demos/invoice-extraction
python3 generate_sample_invoices.py
python3 extract_invoices.py
python3 render_demo_images.py
```

Requires poppler (`pdftoppm`) for PDF rendering, plus the packages in `requirements.txt`.

## Output files

| Path | What |
| --- | --- |
| `input/01-northwind-supplies-usd.pdf` | US office-supply invoice, USD, 10% tax |
| `input/02-bergholm-advisory-eur.pdf` | Nordic consulting invoice, EUR, 25% VAT |
| `input/03-maple-wick-studio-gbp.pdf` | UK design invoice, GBP, 20% VAT |
| `input/04-kitebay-office-goods-jpy.pdf` | Japanese-style invoice, JPY, 10% consumption tax |
| `input/05-southern-reef-supplies-aud.pdf` | Australian tax invoice, AUD, 10% GST |
| `output/extracted_invoices.xlsx` | Workbook with Invoices, Line items and Review |
| `images/fiverr-gig-cover-1280x769.png` | Fiverr gig cover (invoice PDF → spreadsheet) |
| `images/upwork-portfolio-1600x1200.png` | Upwork portfolio image (all 5 PDFs + both sheets) |
| `images/sheet-invoices.png` | Screenshot of the Invoices sheet |
| `images/sheet-line-items.png` | Screenshot of the Line items sheet |
| `images/sheet-review.png` | Screenshot of the Review tab |
