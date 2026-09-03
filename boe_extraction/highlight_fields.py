"""Turns the highlights drawn on a Bill of Entry into field/value pairs.

A reviewer marks the fields to extract by highlighting them, but what a
highlight covers differs by form, so each is resolved against the structure
that form already has:

- the courier (CBE) forms, against the label/value grid in ``cbe_grid``;
- the ICEGATE form, against the numbered tables in ``icegate_tables``.

A highlight may cover a label, its value, or both; matching on position rather
than on the highlighted words themselves means all three resolve the same way.
"""

from .cbe_grid import Cell, read_entries
from .highlights import read as read_highlights
from .icegate_tables import read_tables

# How far outside a highlight a row may sit and still count as covered.
ROW_SLACK = 4
# Stands in for a highlight over a heading or a table, which has no one label.
AS_HIGHLIGHTED = "(as highlighted)"

# Blocks whose label is indented over a wider cell than the value beneath it,
# so the column reader clips the value. The parser already has these.
FROM_PARSER = {
    "1.IMPORTER NAME & ADDRESS": lambda boe, _: boe.importer_name,
    "1.BUYER'S NAME & ADDRESS": lambda boe, _: boe.importer_name,
    "3.SUPPLIER NAME & ADDRESS": lambda boe, invoice: invoice and invoice.supplier,
    "4.THIRD PARTY NAME & ADDRESS": lambda boe, invoice: invoice and invoice.supplier,
}


class Field:
    """A highlighted field: what it is called, and what it holds."""

    __slots__ = ("page", "top", "label", "value", "highlighted")

    def __init__(self, page, top, label, value, highlighted):
        self.page = page
        self.top = top
        self.label = label
        self.value = value
        self.highlighted = highlighted

    def __repr__(self):
        return f"Field(p{self.page + 1} {self.label!r}={self.value!r})"


def _covers_row(highlight, page, top, bottom=None):
    """Whether a highlight overlaps a row's vertical band."""
    if page != highlight.page:
        return False
    row_bottom = top if bottom is None else bottom
    for _, quad_top, _, quad_bottom in highlight.quads:
        if quad_top - ROW_SLACK <= row_bottom and top - ROW_SLACK <= quad_bottom:
            return True
    return False


def _covers_column(highlight, x0, until):
    for left, _, right, _ in highlight.quads:
        if left < until and x0 - 6 < right:
            return True
    return False


def _courier_fields(pdf, highlights):
    # A cell whose label never reached a colon is not a label/value pair at all
    # -- it is a row of one of the form's tables (DUTY DETAILS, PAYMENT
    # DETAILS), which the grid cannot represent. Those are reported verbatim.
    cells = [e for e in read_entries(pdf)
             if isinstance(e, Cell) and e.label and e.label_done]
    fields = []
    for highlight in highlights:
        matched = [c for c in cells if _covers_row(highlight, c.page, c.top)]
        if not matched:
            fields.append(Field(highlight.page, highlight.top, "", "", highlight.text))
        for cell in matched:
            fields.append(Field(highlight.page, cell.top, cell.label, cell.value,
                                highlight.text))
    return fields


def _standard_fields(pdf, highlights, boe):
    pairs = read_tables(pdf)
    invoices = {}
    for invoice in getattr(boe, "invoices", []):
        invoices.setdefault("first", invoice)

    fields = []
    for highlight in highlights:
        matched = [p for p in pairs
                   if _covers_row(highlight, p.page, p.top, p.bottom)
                   and _covers_column(highlight, p.x0, p.until)]
        if not matched:
            fields.append(Field(highlight.page, highlight.top, "", "", highlight.text))
        for pair in matched:
            value = pair.value
            override = FROM_PARSER.get(pair.label)
            if override is not None:
                value = override(boe, invoices.get("first")) or value
            fields.append(Field(pair.page, pair.top, pair.label, value,
                                highlight.text))
    return fields


def collect(pdf, form_type, boe=None):
    """Every highlighted field in the document, as label/value pairs."""
    highlights = read_highlights(pdf)
    if not highlights:
        return []
    if form_type.startswith("CBE"):
        fields = _courier_fields(pdf, highlights)
    else:
        fields = _standard_fields(pdf, highlights, boe)

    seen, kept = set(), []
    for field in fields:
        if not field.label and not field.value:
            # Keep a highlight nothing resolved against, so none is lost.
            if field.highlighted:
                field.label = AS_HIGHLIGHTED
                field.value = field.highlighted
            else:
                continue
        key = (field.page, field.label, field.value)
        if key in seen:
            continue
        seen.add(key)
        kept.append(field)
    kept.sort(key=lambda f: (f.page, round(f.top), f.label))
    return kept
