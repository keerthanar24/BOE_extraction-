"""Parser for the standard ICEGATE Bill of Entry (Parts I-III).

Part II lists each invoice's items with description, unit price and quantity.
Part III repeats the items and carries the assessable value and the duty grid.
The two are joined on the invoice serial number and item serial number that
Part III prints on every item block.

The duty grid is read by column position rather than by splitting the line on
whitespace: empty cells collapse, so "Rate 15 10 18 0 0" only says which duties
were charged once each figure is placed under its own heading.
"""

import re

from .model import BillOfEntry, Invoice, LineItem
from .pdf_text import page_text, upright_page, word_lines

# How far right of a heading a figure may sit and still belong to that column.
COLUMN_SLACK = 12

DUTY_COLUMNS = {
    "1.BCD": "bcd",
    "3.SWS": "sws",
    "5.IGST": "igst",
    "6.G.CESS": "cess",
    "5.CAIDC": "aidc",
}
ASSESSABLE_VALUE = "29.ASSESS VALUE"
TOTAL_DUTY = "30.TOTAL DUTY"
UNIT_PRICE = "11.UPI"
QUANTITY = "13.C.QTY"

ITEM_ROW = re.compile(r"^(\d+)\s+(\d+)\s+(\d+)\s+(\S+)\s+(.*?)((?:\s+[YN]){5})$")
PART2_ITEM = re.compile(
    r"^(\d+)\s+(\d{4,8})\s+(.+?)\s+([\d.]+)\s+([\d.]+)\s+([A-Z]{2,4})\s+([\d.,]+)$")
NUMBER = re.compile(r"^-?[\d.,]+$")
# The form prints a one-letter status stamp after some names.
TRAILING_STAMP = re.compile(r"\s+[A-Z]$")


def matches(first_page_text):
    flat = " ".join(first_page_text.split()).upper()
    if "BILL OF ENTRY" in flat and "PART - I" in flat:
        return "ICEGATE BOE"
    return None


def _num(value):
    if value in (None, ""):
        return None
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return None


def _field_name(text):
    """Normalise a column heading: the form prints "1. BCD" and "1.BCD"."""
    return re.sub(r"^(\d+)\.\s*", r"\1.", " ".join(text.split())).replace("G. CESS", "G.CESS")


def _headings(line, labels):
    """Where each wanted column heading starts on this line.

    Headings are printed as separate words ("1." then "BCD"), so they are
    rebuilt by walking the line and starting a new field at every "n." token.
    """
    fields, current = [], None
    for word in line:
        if re.match(r"^\d+\.", word["text"]):
            current = {"text": word["text"], "x0": word["x0"]}
            fields.append(current)
        elif current is not None:
            current["text"] += " " + word["text"]
    found = {}
    for index, field in enumerate(fields):
        name = _field_name(field["text"])
        if name in labels:
            end = fields[index + 1]["x0"] if index + 1 < len(fields) else None
            found[name] = (field["x0"], end)
    return found, fields


def _cells(line, fields):
    """Numeric words on a line, grouped under the heading they sit beneath."""
    grouped = {}
    if not fields:
        return grouped
    for word in line:
        if not NUMBER.match(word["text"]):
            continue
        owner = None
        for field in fields:
            if word["x0"] >= field["x0"] - COLUMN_SLACK:
                owner = field
            else:
                break
        if owner is None:
            continue
        name = _field_name(owner["text"])
        # Figures are split by the stamp overlay ("30" "1" ".1"); rejoin them.
        grouped[name] = grouped.get(name, "") + word["text"]
    return grouped


def _duty_block(lines, start):
    """Read one "DUTY ... Rate ... Amount" grid starting at ``lines[start]``."""
    fields = _headings(lines[start], set(DUTY_COLUMNS))[1]
    rates, amounts = {}, {}
    for line in lines[start + 1:start + 8]:
        head = line[0]["text"]
        if head == "Rate":
            rates = _cells(line, fields)
        elif head == "Amount":
            amounts = _cells(line, fields)
        elif head == "DUTY":
            break
    return rates, amounts


def _parse_part3(pdf):
    """Duty and valuation details for every item, keyed by (invoice, item)."""
    details = {}
    for page in pdf.pages:
        lines = word_lines(upright_page(page))
        texts = [" ".join(w["text"] for w in line) for line in lines]
        for index, text in enumerate(texts):
            match = ITEM_ROW.match(text)
            if not match or not text.startswith(tuple("123456789")):
                continue
            invoice_sn, item_sn = int(match.group(1)), int(match.group(2))
            record = {
                "cth": match.group(3),
                "description": match.group(5).strip(),
                "duties": {},
            }
            # The block runs to the next item row on the page.
            end = len(texts)
            for ahead in range(index + 1, len(texts)):
                if ITEM_ROW.match(texts[ahead]) and texts[ahead][0].isdigit():
                    end = ahead
                    break
            for offset in range(index, end):
                head = texts[offset].split(" ")[0]
                if head == "DUTY":
                    rates, amounts = _duty_block(lines, offset)
                    for label, key in DUTY_COLUMNS.items():
                        if label in rates or label in amounts:
                            record["duties"].setdefault(key, [None, None])
                            if label in rates:
                                record["duties"][key][0] = _num(rates[label])
                            if label in amounts:
                                record["duties"][key][1] = _num(amounts[label])
                elif "29." in texts[offset] and "ASSESS" in texts[offset]:
                    values = _cells(lines[offset + 1], _headings(
                        lines[offset], {ASSESSABLE_VALUE, TOTAL_DUTY})[1])
                    record["assessable_value"] = _num(values.get(ASSESSABLE_VALUE))
                    record["duty_amount"] = _num(values.get(TOTAL_DUTY))
                elif texts[offset].startswith("11.UPI") and offset + 1 < len(texts):
                    values = _cells(lines[offset + 1], _headings(
                        lines[offset], {UNIT_PRICE, QUANTITY})[1])
                    record["unit_price"] = _num(values.get(UNIT_PRICE))
                    record["quantity"] = _num(values.get(QUANTITY))
            details[(invoice_sn, item_sn)] = record
    return details


def _parse_part2(pdf):
    """Invoice headers and their item lists, keyed by invoice serial number."""
    invoices = {}
    for page in pdf.pages:
        text = page_text(page)
        header = re.search(r"PART - II - INVOICE & VALUATION DETAILS \(Invoice (\d+)", text)
        if not header:
            continue
        serial = int(header.group(1))
        invoice = invoices.setdefault(serial, {"items": {}})

        # The invoice number sits under the "2.INVOICE NO. & DT." heading, with
        # its date wrapped onto the following line.
        detail = re.search(
            r"1\.S\.NO 2\.INVOICE NO\. & DT\..*?\n\s*%d\s+(\S+)\s*\n\s*([\w-]+)?" % serial,
            text, re.S)
        if detail:
            invoice["number"] = detail.group(1)
            date = detail.group(2) or ""
            if re.match(r"^\d{1,2}-[A-Z]{3}-\d{2,4}$", date):
                invoice["date"] = date

        value = re.search(r"1\.INV VALUE.*?\n\s*([\d.,]+)", text, re.S)
        if value:
            invoice["value"] = _num(value.group(1))
        currency = re.search(r"14\.Cur\s+(\S+)", text)
        if currency:
            invoice["currency"] = currency.group(1)
        supplier = re.search(r"4\.THIRD PARTY NAME & ADDRESS\s*\n\s*(.+)$", text, re.M)
        if supplier:
            invoice["supplier"] = TRAILING_STAMP.sub("", supplier.group(1).strip())

        # Item rows wrap, so the description continues on the following lines.
        rows = text[text.find("1.S NO. 2.CTH"):]
        current = None
        for line in rows.split("\n")[1:]:
            if line.startswith("GLOSSARY"):
                break
            match = PART2_ITEM.match(line.strip())
            if match:
                current = {
                    "cth": match.group(2),
                    "description": match.group(3).strip(),
                    "unit_price": _num(match.group(4)),
                    "quantity": _num(match.group(5)),
                    "uqc": match.group(6),
                    "amount": _num(match.group(7)),
                }
                invoice["items"][int(match.group(1))] = current
            elif current is not None and line.strip():
                current["description"] += " " + line.strip()
    return invoices


def parse(pdf, form_type, source_file=""):
    first = page_text(pdf.pages[0])
    boe = BillOfEntry(form_type=form_type, source_file=source_file)

    summary = re.search(
        r"Port Code BE No BE Date BE Type\s*\n\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+)", first)
    if summary:
        boe.port_code, boe.be_number = summary.group(1), summary.group(2)
        boe.be_date, boe.be_type = summary.group(3), summary.group(4)
    for pattern, field in [(r"IEC/Br\s+(\S+)", "iec"),
                           (r"GSTIN/TYPE\s+(\S+)", "gstin"),
                           (r"AD CODE\s+(\S+)", "ad_code")]:
        match = re.search(pattern, first)
        if match:
            setattr(boe, field, match.group(1))
    origin = re.search(r"13\.COUNTRY OF ORIGIN\s+(.+?)\s+14\.COUNTRY OF CONSIGNMENT\s+(\S+)",
                       first)
    if origin:
        boe.country_of_origin, boe.country_of_consignment = origin.group(1), origin.group(2)
    importer = re.search(r"1\.IMPORTER NAME & ADDRESS\s*\n\s*(.+)$", first, re.M)
    if importer:
        boe.importer_name = TRAILING_STAMP.sub("", importer.group(1).strip())
    rate = re.search(r"1\s*(\w{3})\s*=\s*([\d.]+)\s*INR", first)
    if rate:
        boe.currency, boe.exchange_rate = rate.group(1), _num(rate.group(2))

    part2 = _parse_part2(pdf)
    part3 = _parse_part3(pdf)

    for serial in sorted(part2):
        source = part2[serial]
        invoice = Invoice(number=source.get("number", ""),
                          date=source.get("date", ""),
                          supplier=source.get("supplier", ""),
                          currency=source.get("currency", boe.currency),
                          invoice_value=source.get("value"),
                          exchange_rate=boe.exchange_rate)
        for item_sn in sorted(source["items"]):
            row = source["items"][item_sn]
            duties = part3.get((serial, item_sn), {})
            item = LineItem(item_number=item_sn)
            item.hs_code = row["cth"] or duties.get("cth", "")
            item.description = row["description"] or duties.get("description", "")
            item.quantity = row["quantity"] if row["quantity"] is not None else duties.get("quantity")
            item.unit_of_measure = row["uqc"]
            item.unit_price = row["unit_price"] if row["unit_price"] is not None else duties.get("unit_price")
            item.assessable_value = duties.get("assessable_value")
            for key, (rate_value, amount) in duties.get("duties", {}).items():
                setattr(item, f"{key}_rate", rate_value)
                setattr(item, f"{key}_amount", amount)
            item.duty_amount = duties.get("duty_amount")
            item.backfill_bcd()
            if item.duty_amount is None:
                item.duty_amount = item.total_duty()
            item.exchange_rate = boe.exchange_rate
            invoice.items.append(item)
        boe.invoices.append(invoice)
    return boe
