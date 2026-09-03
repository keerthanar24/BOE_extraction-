"""Parser for ECCS Courier Bill of Entry forms (CBE-XIV and CBE-XIII).

These forms carry every figure in the PDF's own text layer, so the whole
document is read deterministically -- no AI model, no per-document cost, and
the same input always produces the same output.
"""

import re

from .cbe_grid import Marker, document_lines, read_entries
from .model import BillOfEntry, Invoice, LineItem

DUTY_HEADS = {
    "BCD": "bcd",
    "AIDC": "aidc",
    "SWSRCHRG": "sws",
    "SWS": "sws",
    "IGST": "igst",
    "CMPNSTRY": "cess",
}

DUTY_ROW = re.compile(
    r"^\s*\d+\s+([A-Za-z ]+?)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$")

INVOICE_MARKER = re.compile(r"^Details\s+Of\s+Invoice\s*-\s*(\d+)$", re.IGNORECASE)
INVOICE_ITEMS_MARKER = re.compile(
    r"^DETAILS OF ITEM\(S\) IN INVOICE\s*-\s*(\d+)", re.IGNORECASE)
ITEM_MARKER = re.compile(r"^Details\s+Of\s+Item\s*-\s*(\d+)$", re.IGNORECASE)


def matches(first_page_text):
    """The CBE form type this page announces, or None."""
    flat = re.sub(r"[\s.]", "", first_page_text).upper()
    if "CBE-XIV" in flat or "COURIERBILLOFENTRY-XIV" in flat:
        return "CBE-XIV"
    if "CBE-XIII" in flat or "COURIERBILLOFENTRY-XIII" in flat:
        return "CBE-XIII"
    return None


def _num(value):
    if not value:
        return None
    value = value.replace(",", "").strip()
    if not value or value.upper() == "N/A":
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _clean(value):
    value = (value or "").strip()
    return "" if value.upper() == "N/A" else value


class _Section:
    """A run of grid cells, addressable by label."""

    def __init__(self, entries):
        self.entries = entries
        self.by_label = {}
        for entry in entries:
            label = getattr(entry, "label", "")
            if label and label not in self.by_label:
                self.by_label[label] = entry.value

    def get(self, *labels):
        for label in labels:
            if label in self.by_label:
                value = _clean(self.by_label[label])
                if value:
                    return value
        return ""

    def number(self, *labels):
        return _num(self.get(*labels))

    def span(self):
        """The (page, top) bounds this section covers."""
        first, last = self.entries[0], self.entries[-1]
        return (first.page, first.top), (last.page, last.top)


def _sections(entries, pattern):
    """Split entries at each marker matching ``pattern``.

    Yields (marker number, entries up to the next marker of the same kind).
    """
    starts = [i for i, e in enumerate(entries)
              if isinstance(e, Marker) and pattern.match(e.text)]
    for position, start in enumerate(starts):
        end = starts[position + 1] if position + 1 < len(starts) else len(entries)
        number = int(pattern.match(entries[start].text).group(1))
        yield number, entries[start:end]


def _duties_in(lines, start, end):
    """The DUTY DETAILS grid falling between two document positions."""
    duties = {}
    for page, top, text in lines:
        if (page, top) < start or (page, top) >= end:
            continue
        match = DUTY_ROW.match(text)
        if not match:
            continue
        head = DUTY_HEADS.get(re.sub(r"\s", "", match.group(1)).upper())
        if head:
            duties[head] = (_num(match.group(2)), _num(match.group(5)))
    return duties


def _parse_item(section, number, lines, next_start):
    item = LineItem(item_number=number)
    item.description = section.get("Item Description", "General Description")
    item.hs_code = section.get("CTSH", "CETSH", "RITC")
    item.quantity = section.number("Quantity")
    item.unit_of_measure = section.get("Unit of Measure")
    item.unit_price = section.number("Unit Price")
    item.exchange_rate = section.number("Rate Of Exchange")
    item.assessable_value = section.number("Assessable Value")

    start, _ = section.span()
    for head, (rate, amount) in _duties_in(lines, start, next_start).items():
        setattr(item, f"{head}_rate", rate)
        setattr(item, f"{head}_amount", amount)

    item.details = {label: _clean(value) for label, value in section.by_label.items()}
    item.backfill_bcd()
    item.duty_amount = item.total_duty()
    return item


def _parse_invoice(header, items_entries, lines, section_end):
    invoice = Invoice()
    invoice.number = header.get("Invoice Number")
    invoice.date = header.get("Date of Invoice")
    invoice.currency = header.get("Currency")
    invoice.invoice_value = header.number("Invoice Value")

    # The supplier is the first "Name" under the SUPPLIER DETAILS heading.
    for position, entry in enumerate(header.entries):
        if isinstance(entry, Marker) and entry.text.upper().startswith("SUPPLIER DETAILS"):
            invoice.supplier = _Section(header.entries[position:position + 4]).get("Name")
            break

    item_sections = list(_sections(items_entries, ITEM_MARKER))
    for position, (number, entries) in enumerate(item_sections):
        if position + 1 < len(item_sections):
            following = item_sections[position + 1][1][0]
            next_start = (following.page, following.top)
        else:
            next_start = section_end
        item = _parse_item(_Section(entries), number, lines, next_start)
        if item.exchange_rate is not None:
            invoice.exchange_rate = item.exchange_rate
        invoice.items.append(item)
    return invoice


def parse(pdf, form_type, source_file=""):
    entries = read_entries(pdf)
    lines = document_lines(pdf)
    document = _Section(entries)

    boe = BillOfEntry(form_type=form_type, source_file=source_file)
    boe.be_number = document.get("CBEXIV Number", "CBEXIII Number", "BOE Number")
    boe.be_date = document.get("BOE Date")
    boe.be_type = document.get("Type Of BOE")
    boe.iec = document.get("Import export Code")
    boe.gstin = document.get("KYC ID")
    boe.ad_code = document.get("Authorised Dealer Code Of Bank")
    boe.country_of_origin = document.get("Country of Origin")
    boe.country_of_consignment = document.get("Country of Consignment")

    for position, entry in enumerate(entries):
        if isinstance(entry, Marker) and entry.text.upper().startswith("PARTICULARS OF THE IMPORTER"):
            boe.importer_name = _Section(entries[position:position + 6]).get("Name")
            break

    headers = dict(_sections(entries, INVOICE_MARKER))
    item_blocks = list(_sections(entries, INVOICE_ITEMS_MARKER))
    end_of_document = (len(pdf.pages), 0)

    for position, (number, block) in enumerate(item_blocks):
        section_end = end_of_document
        if position + 1 < len(item_blocks):
            following = item_blocks[position + 1][1][0]
            section_end = (following.page, following.top)
        header = _Section(headers.get(number, []))
        boe.invoices.append(_parse_invoice(header, block, lines, section_end))

    rates = [i.exchange_rate for i in boe.invoices if i.exchange_rate]
    if rates:
        boe.exchange_rate = rates[0]
    currencies = [i.currency for i in boe.invoices if i.currency]
    if currencies:
        boe.currency = currencies[0]
    return boe
