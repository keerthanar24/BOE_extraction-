"""Checks an extraction against what the document says about itself.

A bill of entry states its own totals, and an ICEGATE one states how many
invoices and items it carries. Those declarations are independent of how the
line items were read, so agreeing with them is real evidence the extraction is
right -- unlike the test suite, which only proves the three sample documents
still parse as they did.

How much can be checked differs by form:

- ICEGATE declares item and invoice counts *and* duty totals.
- CBE-XIII declares its assessable value and total duty.
- CBE-XIV declares neither, so only internal consistency can be checked.
"""

import re

from .cbe_grid import Cell, Marker, read_entries
from .icegate_tables import read_tables
from .pdf_text import page_text

# Per-item figures are rounded to whole rupees, so a total accumulates a little
# rounding. Allow a rupee per ten items, and never less than one.
def _tolerance(items):
    return max(1.0, len(items) / 10)


class Check:
    """One comparison between the document's own figure and the extraction."""

    __slots__ = ("name", "declared", "extracted", "ok", "note")

    def __init__(self, name, declared, extracted, ok, note=""):
        self.name = name
        self.declared = declared
        self.extracted = extracted
        self.ok = ok
        self.note = note

    def __repr__(self):
        mark = "PASS" if self.ok else "FAIL"
        return f"[{mark}] {self.name}: document {self.declared}, extracted {self.extracted}"


def _num(value):
    if value in (None, ""):
        return None
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return None


def _close(declared, extracted, tolerance):
    if declared is None or extracted is None:
        return False
    return abs(declared - extracted) <= tolerance


def _totals(items, field):
    return round(sum(getattr(i, field) or 0 for i in items), 2)


def _icegate_checks(pdf, boe, items):
    checks = []
    first = page_text(pdf.pages[0])

    # "TYPE INV ITEM CONT / Nos 2 9 1" -- the form's own counts.
    counts = re.search(r"Nos\s+(\d+)\s+(\d+)\s+(\d+)", first)
    if counts:
        checks.append(Check("Invoice count", int(counts.group(1)), len(boe.invoices),
                            int(counts.group(1)) == len(boe.invoices)))
        checks.append(Check("Item count", int(counts.group(2)), len(items),
                            int(counts.group(2)) == len(items)))

    summary = {p.label: p.value for p in read_tables(pdf) if p.page == 0}
    tolerance = _tolerance(items)
    for label, field, name in [("18.TOT.ASS VAL", "assessable_value", "Assessable value"),
                               ("14.TOTAL DUTY", "duty_amount", "Total duty"),
                               ("1.BCD", "bcd_amount", "BCD"),
                               ("3.SWS", "sws_amount", "SWS"),
                               ("7.IGST", "igst_amount", "IGST")]:
        declared = _num(summary.get(label))
        if declared is None:
            continue
        extracted = _totals(items, field)
        checks.append(Check(name, declared, extracted, _close(declared, extracted, tolerance)))
    return checks


def _courier_document_cells(pdf):
    """Label/value cells above the first item block."""
    values = {}
    for entry in read_entries(pdf):
        if isinstance(entry, Marker) and entry.text.upper().startswith("ITEM"):
            break
        if isinstance(entry, Cell) and entry.label and entry.label not in values:
            values[entry.label] = entry.value
    return values


def _cbe_xiii_checks(pdf, boe, items):
    cells = _courier_document_cells(pdf)
    tolerance = _tolerance(items)
    checks = []
    for label, field, name in [("Assessable Value", "assessable_value", "Assessable value"),
                               ("Duty(Rs.)", "duty_amount", "Total duty")]:
        declared = _num(cells.get(label))
        if declared is None:
            continue
        extracted = _totals(items, field)
        checks.append(Check(name, declared, extracted, _close(declared, extracted, tolerance)))
    return checks


def _consistency_checks(items):
    """Checks every form supports, because they need no declared figure."""
    checks = []

    priced = [i for i in items
              if None not in (i.unit_price, i.quantity, i.exchange_rate, i.assessable_value)]
    if priced:
        wrong = [i for i in priced
                 if abs(i.unit_price * i.quantity * i.exchange_rate - i.assessable_value) > 1.0]
        checks.append(Check("Assessable value = price x quantity x rate",
                            f"{len(priced)} items", f"{len(priced) - len(wrong)} agree",
                            not wrong,
                            "" if not wrong else
                            "items " + ", ".join(str(i.item_number) for i in wrong[:5])))

    dutied = [i for i in items if i.duty_amount is not None]
    if dutied:
        wrong = [i for i in dutied if abs(i.total_duty() - i.duty_amount) > 1.0]
        checks.append(Check("Duty amount = sum of duty heads",
                            f"{len(dutied)} items", f"{len(dutied) - len(wrong)} agree",
                            not wrong,
                            "" if not wrong else
                            "items " + ", ".join(str(i.item_number) for i in wrong[:5])))

    blank = [i for i in items if not i.description]
    checks.append(Check("Every item has a description", f"{len(items)} items",
                        f"{len(items) - len(blank)} filled", not blank))
    return checks


def verify(pdf, boe):
    """Every check the document supports. An empty list means none were possible."""
    items = boe.all_items()
    if not items:
        return [Check("Line items found", "at least 1", 0, False,
                      "the form was recognised but nothing was extracted")]

    if boe.form_type.startswith("CBE-XIII"):
        checks = _cbe_xiii_checks(pdf, boe, items)
    elif boe.form_type.startswith("CBE"):
        checks = []  # CBE-XIV declares no totals of its own.
    else:
        checks = _icegate_checks(pdf, boe, items)
    return checks + _consistency_checks(items)


def report(checks, form_type):
    """The checks as lines of text, and whether they all passed."""
    lines = []
    declared = [c for c in checks if not c.name.startswith(("Assessable value =",
                                                            "Duty amount =",
                                                            "Every item"))]
    if not declared:
        lines.append(f"  {form_type} declares no totals of its own — "
                     "only internal consistency could be checked.")
    for check in checks:
        mark = "PASS" if check.ok else "FAIL"
        line = f"  [{mark}] {check.name}: document says {check.declared}, extracted {check.extracted}"
        if check.note:
            line += f" ({check.note})"
        lines.append(line)
    return lines, all(c.ok for c in checks)
