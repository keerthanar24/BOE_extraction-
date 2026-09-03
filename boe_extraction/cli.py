"""Command line: boe-extract <pdf> [<pdf> ...] [-o OUTPUT_DIR]"""

import argparse
import sys
from pathlib import Path

from .excel_writer import output_paths, safe_name, write_highlighted, write_invoice
from .extract import extract, extract_with_highlights


def build_parser():
    parser = argparse.ArgumentParser(
        prog="boe-extract",
        description="Extract Bill of Entry line items from PDF into Excel.")
    parser.add_argument("pdfs", nargs="+", type=Path, help="Bill of Entry PDF(s)")
    parser.add_argument("-o", "--output-dir", type=Path, default=Path("."),
                        help="where to write the workbooks (default: current directory)")
    parser.add_argument("--highlights", action="store_true",
                        help="extract what a reviewer highlighted on the PDF "
                             "into one workbook per document")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    args.output_dir.mkdir(parents=True, exist_ok=True)

    failures = 0
    for pdf_path in args.pdfs:
        try:
            if args.highlights:
                boe, fields, highlights = extract_with_highlights(pdf_path)
            else:
                boe, fields, highlights = extract(pdf_path), None, None
        except Exception as error:  # a bad document must not stop the batch
            print(f"{pdf_path}: {error}", file=sys.stderr)
            failures += 1
            continue

        items = boe.all_items()
        print(f"{pdf_path.name}: {boe.form_type}, BE {boe.be_number}, "
              f"{len(boe.invoices)} invoice(s), {len(items)} line item(s)")

        if args.highlights:
            name = safe_name(boe.be_number or pdf_path.stem)
            out_path = args.output_dir / f"BOE__{name}__highlighted.xlsx"
            write_highlighted(boe, fields, highlights, out_path)
            print(f"  {out_path}  ({len(highlights)} highlight(s), "
                  f"{len(fields)} field(s), {len(items)} line item(s))")
            continue

        for invoice, out_path in output_paths(boe, args.output_dir, pdf_path.stem):
            write_invoice(boe, invoice, out_path)
            total = round(sum(i.assessable_value or 0 for i in invoice.items), 2)
            duty = round(sum(i.duty_amount or 0 for i in invoice.items), 2)
            print(f"  {out_path}  ({len(invoice.items)} rows, "
                  f"assessable {total:,.2f}, duty {duty:,.2f})")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
