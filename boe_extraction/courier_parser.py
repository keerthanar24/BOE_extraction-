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
    "CMPNSTRY": "cmpnstry",
    "ADD": "add",
}

# "Sr.No. Duty Head Ad Valorem Specific Rate Duty Forgone Duty Amount"
DUTY_ROW_FULL = re.compile(
    r"^\s*\d+\s+([A-Za-z ]+?)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$")
# "Sr.No. Notification Number Serial Number of Notification"
NOTIFICATION_ROW = re.compile(
    r"^\s*\d+\s+(\d{3}/\d{4})\s+(\S+)\s*$")

DUTY_ROW = re.compile(
    r"^\s*\d+\s+([A-Za-z ]+?)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$")

INVOICE_MARKER = re.compile(r"^Details\s+Of\s+Invoice\s*-\s*(\d+)$", re.IGNORECASE)
INVOICE_ITEMS_MARKER = re.compile(
    r"^DETAILS OF ITEM\(S\) IN INVOICE\s*-\s*(\d+)", re.IGNORECASE)
ITEM_MARKER = re.compile(r"^Details\s+Of\s+Item\s*-\s*(\d+)$", re.IGNORECASE)
# CBE-XIII lists its items flat, each opened by a bare "ITEM :".
XIII_ITEM_MARKER = re.compile(r"^ITEM\s*:$", re.IGNORECASE)

# The two forms name the same field differently.
DESCRIPTION = ("Item Description", "Description of Goods", "General Description")
HS_CODE = ("CTSH", "CETSH", "RITC")
EXCHANGE_RATE = ("Rate Of Exchange", "Rate of Exchange")


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

    Yields (marker number, entries up to the next marker of the same kind). A
    marker that carries no number of its own -- CBE-XIII opens every item with
    a bare "ITEM :" -- is numbered by its position.
    """
    starts = [i for i, e in enumerate(entries)
              if isinstance(e, Marker) and pattern.match(e.text)]
    for position, start in enumerate(starts):
        end = starts[position + 1] if position + 1 < len(starts) else len(entries)
        match = pattern.match(entries[start].text)
        number = int(match.group(1)) if match.groups() else position + 1
        yield number, entries[start:end]


def _duties_in(lines, start, end):
    """The DUTY DETAILS grid falling between two document positions.

    Each row is "Sr.No. Head AdValorem SpecificRate DutyForgone DutyAmount",
    so the specific rate is read alongside the ad valorem one.
    """
    duties = {}
    for page, top, text in lines:
        if (page, top) < start or (page, top) >= end:
            continue
        match = DUTY_ROW_FULL.match(text)
        if not match:
            continue
        head = DUTY_HEADS.get(re.sub(r"\s", "", match.group(1)).upper())
        if head:
            duties[head] = (_num(match.group(2)), _num(match.group(5)),
                            _num(match.group(3)))
    return duties


def _notifications_in(lines, start, end):
    """The NOTIFICATION USED FOR THE ITEM table, as (numbers, serials)."""
    numbers, serials = [], []
    for page, top, text in lines:
        if (page, top) < start or (page, top) >= end:
            continue
        match = NOTIFICATION_ROW.match(text)
        if match:
            numbers.append(match.group(1))
            serials.append(match.group(2))
    return "\n".join(numbers), "\n".join(serials)


def _parse_item(section, number, lines, next_start):
    item = LineItem(item_number=number)
    item.description = section.get(*DESCRIPTION)
    item.hs_code = section.get(*HS_CODE)
    item.quantity = section.number("Quantity")
    item.unit_of_measure = section.get("Unit of Measure")
    item.unit_price = section.number("Unit Price")
    item.exchange_rate = section.number(*EXCHANGE_RATE)
    item.assessable_value = section.number("Assessable Value")

    start, _ = section.span()
    for head, (rate, amount, specific) in _duties_in(lines, start, next_start).items():
        setattr(item, f"{head}_rate", rate)
        setattr(item, f"{head}_amount", amount)
        if head == "bcd":
            item.bcd_specific_rate = specific
    item.notification_number, item.notification_serial = _notifications_in(
        lines, start, next_start)

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


def _parse_xiii(entries, lines, document, boe, end_of_document):
    """Items of a CBE-XIII, which lists them flat rather than under invoices.

    There is no invoice section on this form: each item names the invoice it
    belongs to, so the invoices are grouped from the items.
    """
    supplier = document.get("Name of Consignor")
    sections = list(_sections(entries, XIII_ITEM_MARKER))
    invoices = {}
    for position, (_, block) in enumerate(sections):
        if position + 1 < len(sections):
            following = sections[position + 1][1][0]
            next_start = (following.page, following.top)
        else:
            next_start = end_of_document
        section = _Section(block)
        number = sum(len(i.items) for i in invoices.values()) + 1
        item = _parse_item(section, number, lines, next_start)

        key = section.get("Invoice Number")
        invoice = invoices.get(key)
        if invoice is None:
            invoice = Invoice(number=key, supplier=supplier,
                              currency=section.get("Currency of Invoice"),
                              invoice_value=0.0)
            invoices[key] = invoice
        invoice.items.append(item)
        # The form gives each item's share of the invoice, not the total.
        invoice.invoice_value = round(
            (invoice.invoice_value or 0) + (section.number("Invoice Value") or 0), 2)
        if item.exchange_rate is not None:
            invoice.exchange_rate = item.exchange_rate
    return list(invoices.values())


def parse(pdf, form_type, source_file=""):
    entries = read_entries(pdf)
    lines = document_lines(pdf)
    document = _Section(entries)

    boe = BillOfEntry(form_type=form_type, source_file=source_file)
    boe.be_number = document.get("CBEXIV Number", "CBE-XIII Number",
                                 "CBEXIII Number", "BOE Number")
    boe.be_date = document.get("BOE Date")
    boe.be_type = document.get("Type Of BOE")
    boe.iec = document.get("Import export Code", "Import Export Code")
    boe.gstin = document.get("KYC ID")
    boe.ad_code = document.get("Authorised Dealer Code Of Bank", "AD Code")
    boe.country_of_origin = document.get("Country of Origin")
    boe.country_of_consignment = document.get("Country of Consignment",
                                              "Country of Exportation")

    for position, entry in enumerate(entries):
        if isinstance(entry, Marker) and entry.text.upper().startswith("PARTICULARS OF THE IMPORTER"):
            boe.importer_name = _Section(entries[position:position + 6]).get("Name")
            break

    end_of_document = (len(pdf.pages), 0)
    if form_type == "CBE-XIII":
        boe.importer_name = document.get("Name of Consignee")
        boe.invoices = _parse_xiii(entries, lines, document, boe, end_of_document)
        return _finish(boe)

    headers = dict(_sections(entries, INVOICE_MARKER))
    item_blocks = list(_sections(entries, INVOICE_ITEMS_MARKER))

    for position, (number, block) in enumerate(item_blocks):
        section_end = end_of_document
        if position + 1 < len(item_blocks):
            following = item_blocks[position + 1][1][0]
            section_end = (following.page, following.top)
        header = _Section(headers.get(number, []))
        boe.invoices.append(_parse_invoice(header, block, lines, section_end))
    return _finish(boe)


def _finish(boe):
    """Lift the exchange rate and currency the invoices agree on."""
    rates = [i.exchange_rate for i in boe.invoices if i.exchange_rate]
    if rates:
        boe.exchange_rate = rates[0]
    currencies = [i.currency for i in boe.invoices if i.currency]
    if currencies:
        boe.currency = currencies[0]
    return boe
