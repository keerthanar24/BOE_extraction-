"""Writes extracted line items to Excel, one workbook per invoice."""

import re

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from .model import ITEM_COLUMNS

# Column A repeats the include flag carried by the existing extract format.
FLAG_HEADER = "TRUE"

COLUMN_WIDTHS = {
    "Description": 48,
    "HS Code": 12,
    "Assessable Value": 16,
    "Unit of Measure": 15,
    "CMPNSTRY Amount": 17,
    "CMPNSTRY Rate": 15,
}
DEFAULT_WIDTH = 13

MONEY = "#,##0.00"
MONEY_COLUMNS = {"Unit Price", "Assessable Value", "BCD Amount", "SWS Amount",
                 "IGST Amount", "AIDC Amount", "CMPNSTRY Amount", "Duty Amount"}

HEADER_LABELS = {
    "form_type": "Form Type",
    "source_file": "Source File",
    "be_number": "BE Number",
    "be_date": "BE Date",
    "be_type": "BE Type",
    "port_code": "Port Code",
    "importer_name": "Importer",
    "iec": "IEC",
    "gstin": "GSTIN",
    "ad_code": "AD Code",
    "country_of_origin": "Country of Origin",
    "country_of_consignment": "Country of Consignment",
    "exchange_rate": "Exchange Rate",
    "currency": "Currency",
}


def safe_name(text, fallback="invoice"):
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", (text or "").strip()).strip("_")
    return cleaned or fallback


def _write_items(sheet, items):
    sheet.title = "Sheet1"
    sheet.append([FLAG_HEADER] + [title for _, title in ITEM_COLUMNS])
    for item in items:
        row = [True]
        for field, _ in ITEM_COLUMNS:
            row.append(getattr(item, field))
        sheet.append(row)

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for index, (_, title) in enumerate(ITEM_COLUMNS, start=2):
        letter = get_column_letter(index)
        sheet.column_dimensions[letter].width = COLUMN_WIDTHS.get(title, DEFAULT_WIDTH)
        if title in MONEY_COLUMNS:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row=row, column=index).number_format = MONEY
    sheet.column_dimensions["A"].width = 7
    sheet.freeze_panes = "B2"


def _write_header(sheet, boe, invoice):
    sheet.append(["Field", "Value"])
    for field, value in boe.header_fields():
        sheet.append([HEADER_LABELS.get(field, field), value])
    sheet.append([])
    sheet.append(["Invoice Number", invoice.number])
    sheet.append(["Invoice Date", invoice.date])
    sheet.append(["Supplier", invoice.supplier])
    sheet.append(["Invoice Value", invoice.invoice_value])
    sheet.append(["Invoice Currency", invoice.currency])
    sheet.append(["Line Items", len(invoice.items)])
    sheet.append(["Total Assessable Value",
                  round(sum(i.assessable_value or 0 for i in invoice.items), 2)])
    sheet.append(["Total Duty", round(sum(i.duty_amount or 0 for i in invoice.items), 2)])

    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.column_dimensions["A"].width = 26
    sheet.column_dimensions["B"].width = 60


def write_invoice(boe, invoice, path):
    """One workbook: the invoice's line items, plus the BOE header behind them."""
    workbook = Workbook()
    _write_items(workbook.active, invoice.items)
    _write_header(workbook.create_sheet("BOE Header"), boe, invoice)
    workbook.save(path)
    return path


def _write_highlighted(sheet, fields):
    sheet.append(["Page", "Field", "Value"])
    for field in fields:
        sheet.append([field.page + 1, field.label, field.value])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.column_dimensions["A"].width = 7
    sheet.column_dimensions["B"].width = 38
    sheet.column_dimensions["C"].width = 70
    sheet.freeze_panes = "A2"


def _write_raw_highlights(sheet, highlights):
    sheet.append(["Page", "Highlighted text"])
    for highlight in highlights:
        sheet.append([highlight.page + 1, highlight.text])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.column_dimensions["A"].width = 7
    sheet.column_dimensions["B"].width = 110
    sheet.freeze_panes = "A2"


def _write_all_items(sheet, boe):
    sheet.append(["Invoice"] + [title for _, title in ITEM_COLUMNS])
    for invoice in boe.invoices:
        for item in invoice.items:
            sheet.append([invoice.number]
                         + [getattr(item, field) for field, _ in ITEM_COLUMNS])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.column_dimensions["A"].width = 18
    for index, (_, title) in enumerate(ITEM_COLUMNS, start=2):
        letter = get_column_letter(index)
        sheet.column_dimensions[letter].width = COLUMN_WIDTHS.get(title, DEFAULT_WIDTH)
        if title in MONEY_COLUMNS:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row=row, column=index).number_format = MONEY
    sheet.freeze_panes = "B2"


def _highlighted_item_labels(boe, fields):
    """Highlighted labels that name a per-item field.

    A field highlighted on one item is a request for that field on every item,
    so any highlighted label the items themselves carry is reported for all of
    them.
    """
    items = boe.all_items()
    # Keep a label only where at least one item actually carries a value, so a
    # document-level label that happens to share a name adds no empty column.
    carried = {label for item in items for label, value in item.details.items() if value}
    labels = []
    for field in fields:
        if field.label in carried and field.label not in labels:
            labels.append(field.label)
    return labels


def _write_item_details(sheet, boe, labels):
    sheet.append(["Invoice", "Item Number"] + labels)
    for invoice in boe.invoices:
        for item in invoice.items:
            sheet.append([invoice.number, item.item_number]
                         + [item.details.get(label, "") for label in labels])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.column_dimensions["A"].width = 18
    sheet.column_dimensions["B"].width = 12
    for index in range(3, len(labels) + 3):
        sheet.column_dimensions[get_column_letter(index)].width = 28
    sheet.freeze_panes = "C2"


def write_highlighted(boe, fields, highlights, path):
    """One workbook holding everything the reviewer highlighted.

    The resolved fields are the working sheet; the raw sheet records the
    highlighted text verbatim so nothing is lost to a mis-resolution.
    """
    workbook = Workbook()
    resolved = workbook.active
    resolved.title = "Highlighted Fields"
    _write_highlighted(resolved, fields)
    _write_all_items(workbook.create_sheet("Line Items"), boe)
    labels = _highlighted_item_labels(boe, fields)
    if labels:
        _write_item_details(workbook.create_sheet("Item Details"), boe, labels)
    _write_raw_highlights(workbook.create_sheet("Highlights (raw)"), highlights)
    workbook.save(path)
    return path


def _totals(items, field):
    return round(sum(getattr(i, field) or 0 for i in items), 2)


def _write_documents(sheet, documents):
    sheet.append(["Document", "Form Type", "BE Number", "BE Date", "Port Code",
                  "Importer", "IEC", "GSTIN", "Currency", "Exchange Rate",
                  "Invoices", "Line Items", "Assessable Value", "Total Duty",
                  "Highlights"])
    for document in documents:
        boe, items = document.boe, document.boe.all_items()
        sheet.append([document.name, boe.form_type, boe.be_number, boe.be_date,
                      boe.port_code, boe.importer_name, boe.iec, boe.gstin,
                      boe.currency, boe.exchange_rate, len(boe.invoices),
                      len(items), _totals(items, "assessable_value"),
                      _totals(items, "duty_amount"), len(document.highlights)])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for index in range(1, 16):
        sheet.column_dimensions[get_column_letter(index)].width = 20
    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["F"].width = 28
    for row in range(2, sheet.max_row + 1):
        for index in (13, 14):
            sheet.cell(row=row, column=index).number_format = MONEY
    sheet.freeze_panes = "B2"


def _write_invoices(sheet, documents):
    sheet.append(["Document", "Invoice Number", "Invoice Date", "Supplier",
                  "Invoice Value", "Currency", "Exchange Rate", "Line Items",
                  "Assessable Value", "Total Duty"])
    for document in documents:
        for invoice in document.boe.invoices:
            sheet.append([document.name, invoice.number, invoice.date,
                          invoice.supplier, invoice.invoice_value,
                          invoice.currency, invoice.exchange_rate,
                          len(invoice.items),
                          _totals(invoice.items, "assessable_value"),
                          _totals(invoice.items, "duty_amount")])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for index in range(1, 11):
        sheet.column_dimensions[get_column_letter(index)].width = 18
    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["D"].width = 30
    for row in range(2, sheet.max_row + 1):
        for index in (5, 9, 10):
            sheet.cell(row=row, column=index).number_format = MONEY
    sheet.freeze_panes = "B2"


def _write_combined_fields(sheet, documents):
    sheet.append(["Document", "Page", "Field", "Value"])
    for document in documents:
        for field in document.fields:
            sheet.append([document.name, field.page + 1, field.label, field.value])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["B"].width = 7
    sheet.column_dimensions["C"].width = 38
    sheet.column_dimensions["D"].width = 66
    sheet.freeze_panes = "C2"


def _write_combined_items(sheet, documents):
    sheet.append(["Document", "Invoice"] + [title for _, title in ITEM_COLUMNS])
    for document in documents:
        for invoice in document.boe.invoices:
            for item in invoice.items:
                sheet.append([document.name, invoice.number]
                             + [getattr(item, field) for field, _ in ITEM_COLUMNS])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["B"].width = 18
    for index, (_, title) in enumerate(ITEM_COLUMNS, start=3):
        letter = get_column_letter(index)
        sheet.column_dimensions[letter].width = COLUMN_WIDTHS.get(title, DEFAULT_WIDTH)
        if title in MONEY_COLUMNS:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row=row, column=index).number_format = MONEY
    sheet.freeze_panes = "C2"


def _write_combined_details(sheet, documents):
    columns = []
    for document in documents:
        for label in _highlighted_item_labels(document.boe, document.fields):
            if label not in columns:
                columns.append(label)
    if not columns:
        return False
    sheet.append(["Document", "Invoice", "Item Number"] + columns)
    for document in documents:
        for invoice in document.boe.invoices:
            for item in invoice.items:
                sheet.append([document.name, invoice.number, item.item_number]
                             + [item.details.get(label, "") for label in columns])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["C"].width = 12
    for index in range(4, len(columns) + 4):
        sheet.column_dimensions[get_column_letter(index)].width = 28
    sheet.freeze_panes = "D2"
    return True


def _write_combined_raw(sheet, documents):
    sheet.append(["Document", "Page", "Highlighted text"])
    for document in documents:
        for highlight in document.highlights:
            sheet.append([document.name, highlight.page + 1, highlight.text])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.column_dimensions["A"].width = 34
    sheet.column_dimensions["B"].width = 7
    sheet.column_dimensions["C"].width = 110
    sheet.freeze_panes = "C2"


def write_combined(documents, path):
    """One workbook covering every document in a run.

    Each sheet carries a Document column, so several bills of entry -- and
    several document types -- sit side by side in the same file.
    """
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Documents"
    _write_documents(summary, documents)
    _write_invoices(workbook.create_sheet("Invoices"), documents)
    _write_combined_items(workbook.create_sheet("Line Items"), documents)
    if any(d.fields for d in documents):
        _write_combined_fields(workbook.create_sheet("Highlighted Fields"), documents)
        details = workbook.create_sheet("Item Details")
        if not _write_combined_details(details, documents):
            workbook.remove(details)
        _write_combined_raw(workbook.create_sheet("Highlights (raw)"), documents)
    workbook.save(path)
    return path


def output_paths(boe, output_dir, stem=None):
    """The workbook path for each invoice in the document."""
    paths = []
    for index, invoice in enumerate(boe.invoices, start=1):
        name = safe_name(invoice.number or stem or f"invoice_{index}")
        paths.append((invoice, output_dir / f"BOE__{name}__extracted.xlsx"))
    return paths
