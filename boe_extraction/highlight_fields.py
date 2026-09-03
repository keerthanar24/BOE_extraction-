"""Turns the highlights drawn on a Bill of Entry into field/value pairs.

A reviewer marks the fields to extract by highlighting them, but what a
highlight covers differs by form, so each is resolved against the structure
that form already has:

- the courier (CBE) forms, against the label/value grid in ``cbe_grid``;
- the ICEGATE form, against the numbered tables in ``icegate_tables``.

A highlight may cover a label, its value, or both; matching on position rather
than on the highlighted words themselves means all three resolve the same way.
"""

import re

from .cbe_grid import Cell, Marker, read_entries
from .highlights import read as read_highlights
from .icegate_tables import read_tables
from .pdf_text import is_bold, upright_page, word_lines

PART_HEADING = re.compile(r"^PART - [IVX]+ - .+")
# The narrowest gap between two columns of one of the forms' small tables.
TABLE_COLUMN_GAP = 20

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
    # The rate sits in a two-row table; the parser reads it off page 1.
    "EXCHANGE RATE": lambda boe, _: (f"1 {boe.currency}={boe.exchange_rate}INR"
                                     if boe.exchange_rate else ""),
}


class Field:
    """A highlighted field: what it is called, and what it holds.

    ``section`` is the heading the field sits under. Several sections of a
    courier form carry a "Name" and an "Address", so the heading is what tells
    the importer's name from the supplier's.
    """

    __slots__ = ("page", "top", "label", "value", "highlighted", "section")

    def __init__(self, page, top, label, value, highlighted, section=""):
        self.page = page
        self.top = top
        self.label = label
        self.value = value
        self.highlighted = highlighted
        self.section = section

    @property
    def key(self):
        """The field's name, qualified by its section."""
        return f"{self.section} · {self.label}" if self.section else self.label

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


def _table_columns(header):
    """Group a table's heading words into columns."""
    groups = [[header[0]]]
    for previous, word in zip(header, header[1:]):
        if word["x0"] - previous["x1"] >= TABLE_COLUMN_GAP:
            groups.append([word])
        else:
            groups[-1].append(word)
    columns = []
    for index, group in enumerate(groups):
        until = (groups[index + 1][0]["x0"] - TABLE_COLUMN_GAP
                 if index + 1 < len(groups) else float("inf"))
        columns.append((" ".join(w["text"] for w in group),
                        group[0]["x0"] - TABLE_COLUMN_GAP, until))
    return columns


def _table_fields(page, highlight, section):
    """Read a highlighted table of one row as a field per column.

    The courier form carries a couple of small tables the label/value grid
    cannot represent -- PAYMENT DETAILS is a bold heading row over a single
    row of figures. A table of several rows (DUTY DETAILS) is left verbatim;
    its numbers are already reported against each line item.
    """
    covered = []
    for line in word_lines(upright_page(page)):
        words = [w for w in line if highlight.covers(w)]
        if words:
            covered.append(words)
    if len(covered) != 2:
        return []
    header, row = covered
    if not all(is_bold(w) for w in header) or any(is_bold(w) for w in row):
        return []

    fields = []
    for label, x0, until in _table_columns(header):
        value = " ".join(w["text"] for w in row
                         if x0 <= (w["x0"] + w["x1"]) / 2 < until)
        fields.append(Field(highlight.page, highlight.top, label, value,
                            highlight.text, section))
    return fields


def _courier_fields(pdf, highlights):
    # A cell whose label never reached a colon is not a label/value pair at all
    # -- it is a row of one of the form's tables (DUTY DETAILS, PAYMENT
    # DETAILS), which the grid cannot represent. Those are reported verbatim.
    cells, sections, heading = [], {}, ""
    for entry in read_entries(pdf):
        if isinstance(entry, Marker):
            heading = entry.text.rstrip(" :,")
        elif isinstance(entry, Cell) and entry.label and entry.label_done:
            cells.append(entry)
            sections[id(entry)] = heading

    fields = []
    for highlight in highlights:
        matched = [c for c in cells if _covers_row(highlight, c.page, c.top)]
        if not matched:
            section = _section_before(cells, sections, highlight)
            table = _table_fields(pdf.pages[highlight.page], highlight, section)
            fields.extend(table or
                          [Field(highlight.page, highlight.top, "", "",
                                 highlight.text)])
        for cell in matched:
            fields.append(Field(highlight.page, cell.top, cell.label, cell.value,
                                highlight.text, sections.get(id(cell), "")))
    return fields


def _section_before(cells, sections, highlight):
    """The heading in force where a highlight sits."""
    earlier = [c for c in cells
               if (c.page, c.top) <= (highlight.page, highlight.top)]
    return sections.get(id(earlier[-1]), "") if earlier else ""


def _part_headings(pdf):
    """Where each "PART - n" heading starts, per page."""
    headings = {}
    for page_index, page in enumerate(pdf.pages):
        for line in word_lines(upright_page(page)):
            text = " ".join(w["text"] for w in line)
            if PART_HEADING.match(text):
                headings.setdefault(page_index, []).append((line[0]["top"], text))
    return headings


def _section_at(headings, page, top):
    above = [text for heading_top, text in headings.get(page, [])
             if heading_top <= top]
    return above[-1] if above else ""


def _standard_fields(pdf, highlights, boe):
    pairs = read_tables(pdf)
    headings = _part_headings(pdf)
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
                                highlight.text,
                                _section_at(headings, pair.page, pair.top)))
    return fields


def _qualify_repeats(fields):
    """Keep a field's section in its name only where the label repeats.

    "BOE Number" needs no qualifying; "Name" does, because the importer, the
    supplier and the broker each have one.
    """
    counts = {}
    for field in fields:
        counts[field.label] = counts.get(field.label, 0) + 1
    for field in fields:
        if counts[field.label] < 2:
            field.section = ""
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
        key = (field.page, field.key, field.value)
        if key in seen:
            continue
        seen.add(key)
        kept.append(field)
    kept.sort(key=lambda f: (f.page, round(f.top), f.label))
    _qualify_repeats(kept)
    return kept
