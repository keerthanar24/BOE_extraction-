# Bill of Entry — PDF to Excel extraction

Reads a customs Bill of Entry PDF and writes one Excel workbook per invoice,
with a row per line item and the duty figures already broken out.

## Usage

```bash
pip install -r requirements.txt
python -m boe_extraction.cli path/to/boe.pdf -o output/
```

```
FBA15M6L9KGF01_courier_cbe_xiv.pdf: CBE-XIV, BE CBEXIV_DEL_2026-2027_2808_10570, 1 invoice(s), 4 line item(s)
  output/BOE__FBA15M6L9KGF01__extracted.xlsx  (4 rows, assessable 104,220.00, duty 45,814.00)
```

## Which forms it reads

The form is identified from the text on page 1 — nothing depends on the file
name.

| Document type | How it is read | Cost per document |
| --- | --- | --- |
| CBE-XIV — ECCS courier BOE | Direct parse of the PDF text layer | Free |
| CBE-XIII — ECCS courier BOE | Direct parse of the PDF text layer | Free |
| ICEGATE BOE (Parts I–III) | Direct parse of the PDF text layer | Free |

All three paths are deterministic: the same input always produces the same
output, and no AI model is called.

## What comes out

`Sheet1` carries one row per line item across the 20 established columns:

Item Number · HS Code · Description · Quantity · Unit of Measure · Unit Price ·
Assessable Value · BCD Rate & Amount · SWS Rate & Amount · IGST Rate & Amount ·
AIDC Rate & Amount · Compensation Cess Rate & Amount · Duty Amount ·
Exchange Rate

A second sheet, `BOE Header`, carries the header and summary side of the
document — BE number and date, port code, importer, IEC, GSTIN, AD code,
country of origin and consignment, exchange rate — plus the invoice's own
number, date, supplier, value, and its assessable and duty totals.

A document carrying several invoices produces one workbook per invoice.

## How the PDFs are read

Both form families stamp rotated text over the page: the section labels running
up the left margin, and the diagonal status watermark. pdfplumber returns those
glyphs interleaved with the real cell text, which corrupts words — `LAPTOP
COVER` comes back as `LAPTOSP COVER`, and `X002KD14CB` as `X002KD14ECB`. Every
rotated glyph carries a non-zero `matrix[1]`, so `pdf_text.upright_page` drops
them before the text is laid out.

Beyond that the two families need different readers:

- **Courier (`cbe_grid.py`, `courier_parser.py`)** — the CBE forms are a rigid
  two-column label/value grid. Reading them line by line goes wrong as soon as
  a value itself starts with a capitalised word: `Unit of Measure : PCS
  Quantity : 15` looks like an empty value followed by a label `PCS Quantity`.
  So the two colon columns are located on the page and every word is placed
  relative to them, with labels and values rejoined across their hyphenated
  line wraps.
- **ICEGATE (`standard_parser.py`)** — Part II lists each invoice's items with
  description, unit price and quantity; Part III repeats them with the
  assessable value and the duty grid. The two are joined on the invoice serial
  and item serial numbers. The duty grid is read by column position, not by
  splitting on whitespace: empty cells collapse, so `Rate 15 10 18 0 0` only
  says which duties were charged once each figure sits under its own heading.

## Verifying a run

`tests/test_extraction.py` runs both sample documents end to end and checks the
extracted rows against the totals the document prints for itself — for the
ICEGATE sample, the nine rows add up to the assessable value, BCD, SWS, IGST
and total duty printed in its own Part I summary.

```bash
python -m pytest tests -q
```

## Limits worth knowing

- The form is identified from the text on page 1. A scanned or image-only first
  page has no text layer, is not recognised, and is rejected rather than
  guessed at.
- On some page layouts the underlying table reader drops the BCD cell.
  `LineItem.backfill_bcd` recovers it from the assessable value and the total
  duty, so a blank in the source does not become a zero in the output — but it
  is the one figure worth spot-checking on an unfamiliar document format.
- It is a script run on demand, not a scheduled job or a screen.
