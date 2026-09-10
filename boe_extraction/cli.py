"""Command line: boe-extract <pdf> [<pdf> ...] [-o OUTPUT_DIR]"""

import argparse
import sys
from pathlib import Path

from .excel_writer import (output_paths, safe_name, write_combined,
                           write_all_fields, write_highlighted, write_invoice,
                           write_mandatory)
from .extract import (document_fields, extract, extract_document,
                      extract_with_highlights, verify_document)
from .verify import report


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
    parser.add_argument("--combined", metavar="FILE", type=Path,
                        help="write every document given to a single workbook "
                             "instead of one per invoice")
    parser.add_argument("--mandatory", metavar="FILE", type=Path,
                        help="write the highlighted fields as the columns of "
                             "the extract, one row per document and per item")
    parser.add_argument("--all-fields", action="store_true",
                        help="write every field each document carries, "
                             "whether highlighted or not")
    parser.add_argument("--verify", action="store_true",
                        help="check the extraction against the counts and "
                             "totals each document declares, and write nothing")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    args.output_dir.mkdir(parents=True, exist_ok=True)

    if args.verify:
        return _verify(args)

    if args.all_fields:
        args.output_dir.mkdir(parents=True, exist_ok=True)
        for pdf_path in args.pdfs:
            boe, rows = document_fields(pdf_path)
            name = safe_name(boe.be_number or pdf_path.stem)
            out = args.output_dir / f"BOE__{name}__all_fields.xlsx"
            write_all_fields(boe, rows, out)
            print(f"{pdf_path.name}: {boe.form_type}, {len(rows)} document field(s), "
                  f"{len(boe.all_items())} line item(s)")
            print(f"  {out}")
        return 0

    if args.combined or args.mandatory:
        return _combined(args)

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
        if not items:
            print(f"{pdf_path}: the form was recognised but no line items were "
                  f"extracted", file=sys.stderr)
            failures += 1

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


def _verify(args):
    """Check each document against what it says about itself."""
    failures = 0
    for pdf_path in args.pdfs:
        try:
            boe, checks = verify_document(pdf_path)
        except Exception as error:  # a bad document must not stop the batch
            print(f"{pdf_path}: {error}", file=sys.stderr)
            failures += 1
            continue

        items = boe.all_items()
        print(f"{pdf_path.name}: {boe.form_type}, BE {boe.be_number}, "
              f"{len(boe.invoices)} invoice(s), {len(items)} line item(s)")
        lines, passed = report(checks, boe.form_type)
        for line in lines:
            print(line)
        print("  => " + ("all checks passed" if passed else "CHECKS FAILED"))
        if not passed:
            failures += 1
    return 1 if failures else 0


def _combined(args):
    """Write every document given into one workbook."""
    documents, failures = [], 0
    for pdf_path in args.pdfs:
        try:
            document = extract_document(pdf_path, with_highlights=True)
        except Exception as error:  # a bad document must not stop the batch
            print(f"{pdf_path}: {error}", file=sys.stderr)
            failures += 1
            continue
        documents.append(document)
        items = document.boe.all_items()
        print(f"{pdf_path.name}: {document.boe.form_type}, "
              f"BE {document.boe.be_number}, {len(document.boe.invoices)} invoice(s), "
              f"{len(items)} line item(s), {len(document.highlights)} highlight(s)")

    if not documents:
        return 1
    for out_path, writer in ((args.combined, write_combined),
                             (args.mandatory, write_mandatory)):
        if out_path is None:
            continue
        out_path.parent.mkdir(parents=True, exist_ok=True)
        writer(documents, out_path)
        print(f"  {out_path}  ({len(documents)} document(s))")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
