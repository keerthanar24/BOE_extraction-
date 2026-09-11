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
    sheet.append(["Page", "Section", "Field", "Value"])
    for field in fields:
        sheet.append([field.page + 1, field.section, field.label, field.value])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for column, width in zip("ABCD", (7, 46, 38, 70)):
        sheet.column_dimensions[column].width = width
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
    # Keep a label only where items actually carry a value under it. Where
    # there are several items it must appear on more than one: a block that
    # falls after the last item heading -- the payment details at the end of a
    # courier form -- lands in that item alone and is not a per-item field.
    counts = {}
    for item in items:
        for label, value in item.details.items():
            if value:
                counts[label] = counts.get(label, 0) + 1
    threshold = 2 if len(items) > 1 else 1
    carried = {label for label, seen in counts.items() if seen >= threshold}
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
    # Every per-item field is reported on Line Items, so a separate sheet of
    # them only repeats it.
    _write_all_items(workbook.create_sheet("Line Items"), boe)
    _write_raw_highlights(workbook.create_sheet("Highlights (raw)"), highlights)
    workbook.save(path)
    return path


def _totals(items, field):
    return round(sum(getattr(i, field) or 0 for i in items), 2)


def _write_document_invoices(sheet, document, invoice=None):
    """The invoices of one document, with what each one totals.

    Part II's per-invoice figures are reported here rather than among the
    document's fields, where only the first invoice's could ever show.
    """
    sheet.append(["BE No", "Invoice Number", "Invoice Date", "Supplier",
                  "Invoice Value", "Currency", "Exchange Rate", "Line Items",
                  "Assessable Value", "Total Duty"])
    for one in ([invoice] if invoice is not None else document.boe.invoices):
        sheet.append([document.boe.be_number, one.number, one.date,
                      one.supplier, one.invoice_value, one.currency,
                      one.exchange_rate, len(one.items),
                      _totals(one.items, "assessable_value"),
                      _totals(one.items, "duty_amount")])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for index in range(1, 11):
        sheet.column_dimensions[get_column_letter(index)].width = 18
    sheet.column_dimensions["A"].width = 34   # the courier BE numbers are long
    sheet.column_dimensions["D"].width = 30
    for row in range(2, sheet.max_row + 1):
        for index in (5, 9, 10):
            sheet.cell(row=row, column=index).number_format = MONEY
    sheet.freeze_panes = "C2"


def _write_document_items(sheet, document, invoice=None):
    """The line items of one document, or of one invoice of it."""
    invoices = [invoice] if invoice is not None else document.boe.invoices
    named = invoice is None and len(document.boe.invoices) > 1
    lead = ["Invoice"] if named else []
    sheet.append(lead + [title for _, title in ITEM_COLUMNS])
    for one in invoices:
        for item in one.items:
            sheet.append(([one.number] if named else [])
                         + [getattr(item, field) for field, _ in ITEM_COLUMNS])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    first = 2 if named else 1
    if named:
        sheet.column_dimensions["A"].width = 18
    for index, (_, title) in enumerate(ITEM_COLUMNS, start=first):
        letter = get_column_letter(index)
        sheet.column_dimensions[letter].width = COLUMN_WIDTHS.get(title, DEFAULT_WIDTH)
        if title in MONEY_COLUMNS:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row=row, column=index).number_format = MONEY
    sheet.freeze_panes = f"{get_column_letter(first + 1)}2"


def _write_document_raw(sheet, document):
    sheet.append(["Page", "Highlighted text"])
    for highlight in document.highlights:
        sheet.append([highlight.page + 1, highlight.text])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.column_dimensions["A"].width = 7
    sheet.column_dimensions["B"].width = 110
    sheet.freeze_panes = "B2"


def _split_levels(document):
    """The document's highlighted fields, split into per-item and per-document."""
    item_labels = set(_highlighted_item_labels(document.boe, document.fields))
    document_fields, item_fields = [], []
    for field in document.fields:
        (item_fields if field.label in item_labels else document_fields).append(field)
    return document_fields, item_fields


def _column_order(fields):
    """Every field name the document carries, in the order they were read."""
    order = []
    for field in fields:
        if field.key not in order:
            order.append(field.key)
    return order


def _style_wide(sheet, fixed_widths, first_data_column):
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for index, width in enumerate(fixed_widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for index in range(first_data_column, sheet.max_column + 1):
        sheet.column_dimensions[get_column_letter(index)].width = 26
    sheet.freeze_panes = f"{get_column_letter(first_data_column)}2"
    sheet.row_dimensions[1].height = 46


def _write_mandatory_document(sheet, document, fields, columns):
    sheet.append(["Form Type", "BE Number"] + columns)
    values = {field.key: field.value for field in fields}
    sheet.append([document.boe.form_type, document.boe.be_number]
                 + [values.get(column, "") for column in columns])
    _style_wide(sheet, [14, 30], 3)


def _write_mandatory_items(sheet, invoices, columns, labels_by_key, named):
    lead = (["Invoice"] if named else []) + ["Item Number"]
    sheet.append(lead + columns)
    for invoice in invoices:
        for item in invoice.items:
            sheet.append(([invoice.number] if named else []) + [item.item_number]
                         + [item.details.get(labels_by_key[c], "")
                            for c in columns])
    _style_wide(sheet, [18, 12] if named else [12], len(lead) + 1)


def _write_checklist(sheet, fields, document_columns, item_columns,
                     labels_by_key, items):
    """Which mandatory fields came out filled, and which did not."""
    sheet.append(["Field", "Level", "Records", "Filled"])
    values = {field.key: field.value for field in fields}
    for column in document_columns:
        sheet.append([column, "Document", 1, 1 if values.get(column) else 0])
    for column in item_columns:
        label = labels_by_key[column]
        sheet.append([column, "Line item", len(items),
                      sum(1 for i in items if i.details.get(label))])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for index, width in enumerate([56, 12, 12, 10], start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    sheet.freeze_panes = "A2"


def write_mandatory(document, path, invoice=None):
    """The highlighted fields of ONE document as a schema: one column each.

    The highlights are the required-field list, so they become the columns of
    the extract rather than a name/value listing. Where a bill of entry
    carries several invoices, pass the one this workbook covers.
    """
    document_fields, item_fields = _split_levels(document)
    document_columns = _column_order(document_fields)
    item_columns = _column_order(item_fields)
    labels_by_key = {field.key: field.label for field in item_fields}

    invoices = [invoice] if invoice is not None else document.boe.invoices
    named = invoice is None and len(document.boe.invoices) > 1
    items = [item for one in invoices for item in one.items]

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Mandatory Fields"
    _write_mandatory_document(sheet, document, document_fields, document_columns)
    if item_columns:
        _write_mandatory_items(workbook.create_sheet("Mandatory Item Fields"),
                               invoices, item_columns, labels_by_key, named)
    _write_document_items(workbook.create_sheet("Line Items"), document, invoice)
    _write_checklist(workbook.create_sheet("Field Checklist"), document_fields,
                     document_columns, item_columns, labels_by_key, items)
    _write_document_raw(workbook.create_sheet("Highlights (raw)"), document)
    workbook.save(path)
    return path


# A field's label starts with a capital. Table fragments start with a digit,
# and the declaration paragraphs with "(i)" or a lower-case word.
JUNK_LABEL = re.compile(r"^(?:[^A-Z]|Sr\.No|Note$|Port :)")


def _real_labels(items):
    """The per-item labels that name a field, in the order the form prints.

    A label is kept even where every item leaves it blank: an empty column
    still says the form asks for that field.
    """
    labels = []
    for item in items:
        for label in item.details:
            if label not in labels and len(label) <= 40 and not JUNK_LABEL.match(label):
                labels.append(label)
    return labels


def _write_invoice_items(sheet, invoice):
    sheet.append([title for _, title in ITEM_COLUMNS])
    for item in invoice.items:
        sheet.append([getattr(item, name) for name, _ in ITEM_COLUMNS])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for index, (_, title) in enumerate(ITEM_COLUMNS, start=1):
        letter = get_column_letter(index)
        sheet.column_dimensions[letter].width = COLUMN_WIDTHS.get(title, DEFAULT_WIDTH)
        if title in MONEY_COLUMNS:
            for row in range(2, sheet.max_row + 1):
                sheet.cell(row=row, column=index).number_format = MONEY
    sheet.freeze_panes = "A2"


def write_invoice_highlighted(boe, invoice, fields, highlights, path):
    """One workbook for a single invoice of a document.

    A bill carrying several invoices is reported one invoice at a time, so its
    line items are not merged into a single sheet.
    """
    workbook = Workbook()
    resolved = workbook.active
    resolved.title = "Highlighted Fields"
    _write_highlighted(resolved, fields)
    _write_invoice_items(workbook.create_sheet("Line Items"), invoice)
    _write_raw_highlights(workbook.create_sheet("Highlights (raw)"), highlights)
    workbook.save(path)
    return path


def write_all_fields(boe, document_fields, path):
    """Every field the document carries, whether highlighted or not."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Document Fields"
    sheet.append(["Section", "Field", "Value"])
    for section, label, value in document_fields:
        sheet.append([section, label, value])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for column, width in zip("ABC", (34, 38, 70)):
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"

    items = boe.all_items()
    extra = [l for l in _real_labels(items)
             if l not in {t for _, t in ITEM_COLUMNS}]
    detail = workbook.create_sheet("Item Fields")
    detail.append(["Invoice"] + [t for _, t in ITEM_COLUMNS] + extra)
    for invoice in boe.invoices:
        for item in invoice.items:
            detail.append([invoice.number]
                          + [getattr(item, name) for name, _ in ITEM_COLUMNS]
                          + [item.details.get(l, "") for l in extra])
    for cell in detail[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
    detail.column_dimensions["A"].width = 18
    for index in range(2, detail.max_column + 1):
        detail.column_dimensions[get_column_letter(index)].width = 22
    detail.freeze_panes = "B2"
    workbook.save(path)
    return path


# Excel caps a sheet name at 31 characters, and " Highlights" is the longest
# suffix a group uses.
SHEET_NAME_LIMIT = 31
SHEET_SUFFIXES = (" Invoices", " Fields", " Line Items", " Highlights")


# What the reader calls each form: the ICEGATE form is the cargo bill of
# entry, and the two ECCS forms are the courier ones.
FORM_LABELS = {"ICEGATE BOE": "Cargo BOE", "CBE-XIV": "Courier CBE-XIV",
               "CBE-XIII": "Courier CBE-XIII"}


def _sheet_safe(text, fallback):
    """A sheet name Excel accepts: no : \\ / ? * [ ] and no leading quote."""
    cleaned = re.sub(r"[:\\/?*\[\]]+", " ", (text or "").strip())
    return re.sub(r"\s{2,}", " ", cleaned).strip(" '") or fallback


def _unit_label(document, invoice, taken):
    """A short name for the sheets of one document, or of one of its invoices.

    A document's sheets are named for the form it is, not for anything inside
    it: an invoice number on the tab reads as though the invoices had been
    split apart. Only sheets that really do cover a single invoice take its
    number, and a BE number tells two bills of the same form apart.
    """
    room = SHEET_NAME_LIMIT - max(len(s) for s in SHEET_SUFFIXES)
    form = FORM_LABELS.get(document.boe.form_type, document.boe.form_type)
    base = _sheet_safe(invoice.number if invoice is not None else form,
                       document.name)[:room].strip()
    # The BE number tells two bills of the same form apart, but only where it
    # fits whole: a truncated one names nothing.
    label = base
    for candidate in (base, f"{base} {document.boe.be_number}"):
        candidate = _sheet_safe(candidate, base).strip()
        if len(candidate) <= room and candidate not in taken:
            label = candidate
            break
    index = 1
    while label in taken:
        index += 1
        suffix = f" {index}"
        label = base[:room - len(suffix)].strip() + suffix
    taken.add(label)
    return label


def _units(documents, per_invoice):
    """Each document, or each of its invoices, with the label its sheets take."""
    taken, units = set(), []
    for document in documents:
        invoices = document.boe.invoices if per_invoice else [None]
        for invoice in invoices:
            units.append((document, invoice, _unit_label(document, invoice, taken)))
    return units


def write_workbook(documents, path, per_invoice=False):
    """One workbook for a run, with every document on its own sheets.

    Nothing is combined: a sheet only ever holds one bill of entry, and with
    per_invoice only one invoice of it. The sheet name says which.
    """
    workbook = Workbook()
    workbook.remove(workbook.active)
    for document, invoice, label in _units(documents, per_invoice):
        _write_document_invoices(workbook.create_sheet(f"{label} Invoices"),
                                 document, invoice)
        _write_highlighted(workbook.create_sheet(f"{label} Fields"),
                           document.fields)
        items = workbook.create_sheet(f"{label} Line Items")
        if invoice is None:
            _write_document_items(items, document)
        else:
            _write_invoice_items(items, invoice)
        _write_raw_highlights(workbook.create_sheet(f"{label} Highlights"),
                              document.highlights)
    workbook.save(path)
    return path


def _write_schema_fields(sheet, fields):
    sheet.append(["Section", "Field", "Value"])
    for section, label, value in fields:
        sheet.append([section, label, value])
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    # A field a bill answers once per invoice reports a line per invoice, so
    # the value column wraps rather than showing only the first.
    for row in range(2, sheet.max_row + 1):
        sheet.cell(row=row, column=3).alignment = Alignment(
            vertical="top", wrap_text=True)
    for column, width in zip("ABC", (34, 40, 70)):
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "C2"


def _write_schema_items(lines, items):
    from .schema import ITEM_FIELDS

    lines.append(list(ITEM_FIELDS))
    for row in items:
        lines.append(row)
    for cell in lines[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for index, title in enumerate(ITEM_FIELDS, start=1):
        letter = get_column_letter(index)
        lines.column_dimensions[letter].width = COLUMN_WIDTHS.get(title, DEFAULT_WIDTH)
        if title in MONEY_COLUMNS:
            for row in range(2, lines.max_row + 1):
                lines.cell(row=row, column=index).number_format = MONEY
    lines.freeze_panes = "C2"
    lines.row_dimensions[1].height = 46


def write_schema(boe, fields, items, path):
    """One document's schema fields and line items: two sheets, nothing else.

    Every field the schema names gets a row or a column, whether or not this
    document filled it, and nothing the schema does not name is written.
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Document Fields"
    _write_schema_fields(sheet, fields)
    _write_schema_items(workbook.create_sheet("Line Items"), items)
    workbook.save(path)
    return path


def write_schema_combined(extracts, path):
    """Every document of a run in one workbook, each on its own sheets.

    Combined into one file, never into one sheet: a sheet holds a single bill
    of entry, and its name says which.
    """
    workbook = Workbook()
    workbook.remove(workbook.active)
    taken = set()
    for boe, fields, items in extracts:
        label = _unit_label(_Named(boe), None, taken)
        # "Fields", not "Document Fields": with the form name in front, the
        # longer word costs more than Excel's 31 characters allow.
        _write_schema_fields(workbook.create_sheet(f"{label} Fields"), fields)
        _write_schema_items(workbook.create_sheet(f"{label} Line Items"), items)
    workbook.save(path)
    return path


class _Named:
    """What _unit_label needs of a document: its bill of entry and a name."""

    __slots__ = ("boe", "name")

    def __init__(self, boe):
        self.boe = boe
        self.name = boe.source_file or boe.be_number or "document"


def output_paths(boe, output_dir, stem=None):
    """The workbook path for each invoice in the document."""
    paths = []
    for index, invoice in enumerate(boe.invoices, start=1):
        name = safe_name(invoice.number or stem or f"invoice_{index}")
        paths.append((invoice, output_dir / f"BOE__{name}__extracted.xlsx"))
    return paths
