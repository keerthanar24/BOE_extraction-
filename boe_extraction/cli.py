"""Command line: boe-extract <pdf> [<pdf> ...] [-o OUTPUT_DIR]"""

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook

from .excel_writer import (output_paths, safe_name, write_all_fields,
                           write_highlighted, write_invoice,
                           write_invoice_highlighted, write_mandatory,
                           write_schema, write_schema_combined,
                           write_workbook)
from .extract import (document_fields, extract, extract_document,
                      extract_with_highlights, schema_extract, verify_document)
from .schema import FAMILIES, FAMILY
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
    parser.add_argument("--schema", action="store_true",
                        help="write exactly the fields the schema names, one "
                             "workbook per document")
    parser.add_argument("--combined", metavar="FILE", type=Path,
                        help="with --schema, write every document given to "
                             "this one workbook instead of one per family")
    parser.add_argument("--workbook", metavar="FILE", type=Path,
                        help="write every document given to one workbook, "
                             "each on its own sheets rather than sharing them")
    parser.add_argument("--mandatory", action="store_true",
                        help="write the highlighted fields as the columns of "
                             "the extract, one workbook per document")
    parser.add_argument("--per-invoice", action="store_true",
                        help="with --highlights or --mandatory, write one "
                             "workbook per invoice rather than one per document")
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

    if args.schema:
        return _schema(args)

    if args.workbook:
        return _workbook(args)

    if args.mandatory:
        return _mandatory(args)

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
            if args.per_invoice:
                for invoice in boe.invoices:
                    label = safe_name(invoice.number or name)
                    out_path = args.output_dir / f"BOE__{name}__{label}__highlighted.xlsx"
                    write_invoice_highlighted(boe, invoice, fields, highlights, out_path)
                    print(f"  {out_path}  ({len(invoice.items)} line item(s))")
                continue
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


def _schema(args):
    """Exactly the named fields, per document or all in one workbook."""
    extracts, failures = [], 0
    for pdf_path in args.pdfs:
        try:
            boe, fields, items = schema_extract(pdf_path)
        except Exception as error:  # a bad document must not stop the batch
            print(f"{pdf_path}: {error}", file=sys.stderr)
            failures += 1
            continue

        extracts.append((boe, fields, items))
        blank = sum(1 for _, _, value in fields if not value)
        print(f"{pdf_path.name}: {boe.form_type}, BE {boe.be_number}, "
              f"{len(fields)} field(s), {blank} left blank on the form, "
              f"{len(items)} line item(s)")
        if not items:
            print(f"{pdf_path}: no line items were extracted", file=sys.stderr)
            failures += 1

    if not extracts:
        return 1

    # One workbook per family by default -- the cargo bills in one file, the
    # courier bills in another -- and each document on its own sheets within.
    groups = ([(args.combined, extracts)] if args.combined else
              _by_family(args, extracts))
    for out_path, group in groups:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        write_schema_combined(group, out_path)
        print(f"  {out_path}  ({len(group)} document(s))")
        for name in load_workbook(out_path).sheetnames:
            print(f"    {name}")
    return 1 if failures else 0


def _by_family(args, extracts):
    """Each family of bill of entry, with the workbook it is written to."""
    groups = []
    for family in FAMILIES:
        group = [e for e in extracts if FAMILY.get(e[0].form_type) == family]
        if group:
            groups.append((args.output_dir / f"{family}_BOE_extract.xlsx", group))
    return groups


def _workbook(args):
    """One workbook for the run, with a sheet group per document."""
    documents, failures = [], 0
    for pdf_path in args.pdfs:
        try:
            document = extract_document(pdf_path, with_highlights=True)
        except Exception as error:  # a bad document must not stop the batch
            print(f"{pdf_path}: {error}", file=sys.stderr)
            failures += 1
            continue
        documents.append(document)
        boe = document.boe
        print(f"{pdf_path.name}: {boe.form_type}, BE {boe.be_number}, "
              f"{len(boe.invoices)} invoice(s), {len(boe.all_items())} line "
              f"item(s), {len(document.highlights)} highlight(s)")

    if not documents:
        return 1
    args.workbook.parent.mkdir(parents=True, exist_ok=True)
    write_workbook(documents, args.workbook, args.per_invoice)
    print(f"  {args.workbook}")
    for name in load_workbook(args.workbook).sheetnames:
        print(f"    {name}")
    return 1 if failures else 0


def _mandatory(args):
    """The highlighted fields as columns, one workbook per document."""
    failures = 0
    for pdf_path in args.pdfs:
        try:
            document = extract_document(pdf_path, with_highlights=True)
        except Exception as error:  # a bad document must not stop the batch
            print(f"{pdf_path}: {error}", file=sys.stderr)
            failures += 1
            continue

        boe = document.boe
        items = boe.all_items()
        print(f"{pdf_path.name}: {boe.form_type}, BE {boe.be_number}, "
              f"{len(boe.invoices)} invoice(s), {len(items)} line item(s), "
              f"{len(document.highlights)} highlight(s)")
        name = safe_name(boe.be_number or pdf_path.stem)
        for invoice, out_path in _mandatory_paths(args, boe, name):
            write_mandatory(document, out_path, invoice)
            covered = invoice.items if invoice is not None else items
            print(f"  {out_path}  ({len(covered)} line item(s))")
    return 1 if failures else 0


def _mandatory_paths(args, boe, name):
    """Where each mandatory workbook goes: one per document, or per invoice."""
    if not args.per_invoice:
        return [(None, args.output_dir / f"BOE__{name}__mandatory.xlsx")]
    return [(invoice,
             args.output_dir
             / f"BOE__{name}__{safe_name(invoice.number or name)}__mandatory.xlsx")
            for invoice in boe.invoices]


if __name__ == "__main__":
    sys.exit(main())
