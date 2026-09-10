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
BANNER = re.compile(r"^(BILL OF ENTRY FOR [A-Z ]+?)(?: PKG| G\.WT|$)")
# What opens an item on a courier form: CBE-XIV numbers them, CBE-XIII does not.
ITEM_BLOCK = re.compile(r"^(?:Details\s+Of\s+Item\s*-\s*\d+|ITEM\s*:)$", re.IGNORECASE)
# The narrowest gap between two columns of one of the forms' small tables.
TABLE_COLUMN_GAP = 20

# How far outside a highlight a row may sit and still count as covered.
ROW_SLACK = 4
# Stands in for a highlight over a heading or a table, which has no one label.
AS_HIGHLIGHTED = "(as highlighted)"

# A field the form prints as one cell but which holds two values.
SPLIT_FIELDS = {
    "IEC/Br": ("IEC", "Br"),
    "GSTIN/TYPE": ("GSTIN", "TYPE"),
}
# Stamps and barcodes the form prints beside a value.
STAMP = re.compile(r"\s*(FIRST COPY|SECOND COPY|BE\d{10,})\s*")
# ICEGATE processing flags: not extracted data.
DROP_LABELS = {
    "1.S.NO", "2.INVOICE NO", "3.INV. AMT", "3.DEF BE", "4.KACHA", "5.SEC 48",
    "6.REIMP", "8.ASSESS", "9.EXAM", "10.HSS", "3.AEO", "5.AEO", "10.SAED",
    "11.GSIA", "12.TTA", "1.INVSNO", "23.PRODN24.CNTRL", "7.ADV BE",
    "11.FIRST", "12. PROV/",
    # The Part II item table is reported per item on the Line Items sheet.
    "1.S NO.", "2.CTH", "3.DESCRIPTION", "4.UNIT PRICE", "5.QUANTITY",
    "6.UQC", "7.AMOUNT",
    # Part II repeats per invoice, so a document-level sheet could only ever
    # show the first invoice's figures. The Invoices and Line Items sheets
    # carry them per invoice, which is where they belong.
    "2.INVOICE NO. & DT.", "3.PURCHASE ORDER NO & DT", "1.INV VALUE",
    "2.FREIGHT", "14.Cur", "15.Term", "14.ASS. VALUE",
}

# Blocks whose label is indented over a wider cell than the value beneath it,
# so the column reader clips the value. The parser already has these.
# The name-and-address blocks are read in full from the page itself, so they
# are no longer taken from the parser, which held only the name.
FROM_PARSER = {
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

    __slots__ = ("page", "top", "label", "value", "highlighted", "section",
                 "qualified")

    def __init__(self, page, top, label, value, highlighted, section=""):
        self.page = page
        self.top = top
        self.label = label
        self.value = value
        self.highlighted = highlighted
        self.section = section
        self.qualified = False

    @property
    def key(self):
        """The field's name, qualified by its section where the name repeats."""
        return f"{self.section} · {self.label}" if self.qualified else self.label

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
    """Where each "PART - n" heading starts, per page, and the form's banner.

    The header block above PART - I sits under no heading of its own, so the
    banner the form prints across the top stands as its category.
    """
    headings, banner = {}, ""
    for page_index, page in enumerate(pdf.pages):
        for line in word_lines(upright_page(page)):
            text = " ".join(w["text"] for w in line)
            if PART_HEADING.match(text):
                headings.setdefault(page_index, []).append((line[0]["top"], text))
            elif not banner and BANNER.match(text):
                banner = BANNER.match(text).group(1).strip()
    headings["banner"] = banner
    return headings


def _section_at(headings, page, top):
    above = [text for heading_top, text in headings.get(page, [])
             if heading_top <= top]
    return above[-1] if above else headings.get("banner", "")


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
    """Put a field's section into its name only where the label repeats.

    Every field keeps its section in its own column -- that is the category it
    belongs under -- but "BOE Number" needs no qualifying in the name itself,
    while "Name" does, because the importer, the supplier and the broker each
    have one.
    """
    counts = {}
    for field in fields:
        counts[field.label] = counts.get(field.label, 0) + 1
    for field in fields:
        field.qualified = counts[field.label] > 1 and bool(field.section)
    return fields


def _split_and_clean(fields):
    """Separate a two-in-one field, and drop stamps printed beside a value."""
    result = []
    for field in fields:
        # "FIRST COPY" is a stamp printed beside other values, but it is also
        # the value of 1.BE STATUS, so never strip a field down to nothing.
        stripped = STAMP.sub(" ", field.value or "").strip()
        if stripped:
            field.value = stripped
        parts = SPLIT_FIELDS.get(field.label)
        if parts and "/" in field.value:
            head, _, tail = field.value.partition("/")
            for name, value in zip(parts, (head.strip(), tail.strip())):
                result.append(Field(field.page, field.top, name, value,
                                    field.highlighted, field.section))
            continue
        result.append(field)
    return result


def _item_blocks(pdf, boe):
    """Where each item's block starts, paired with the item itself.

    A courier form opens every item with a marker, so a highlight can be placed
    in the item it actually sits in.
    """
    starts = [(e.page, e.top) for e in read_entries(pdf)
              if isinstance(e, Marker) and ITEM_BLOCK.match(e.text)]
    return list(zip(starts, boe.all_items()))


def _item_at(blocks, field):
    """The item whose block a field falls inside, if any."""
    found = None
    for start, item in blocks:
        if (field.page, field.top) >= start:
            found = item
        else:
            break
    return found


def _shared_item_labels(blocks):
    """Labels every item carries, which are therefore per-item fields.

    Position alone is not enough: everything after the last item heading sits
    "inside" it, including the payment block at the end of a courier form. A
    genuine per-item field appears on more than one item; that block's fields
    appear on only the last.
    """
    if len(blocks) < 2:
        return {label for _, item in blocks for label in item.details}
    counts = {}
    for _, item in blocks:
        for label in item.details:
            counts[label] = counts.get(label, 0) + 1
    return {label for label, seen in counts.items() if seen > 1}


def _belongs_to_an_item(blocks, shared, field):
    item = _item_at(blocks, field)
    return item is not None and field.label in shared


def _use_parsed_item_values(fields, boe, blocks):
    """Repair a per-item field's value from the item the parser built.

    Reading the item table by column clips a value that starts left of its own
    heading -- "3.DESCRIPTION" begins under "2.CTH". The parser has already
    read the row properly, so its value stands.

    Only a field sitting inside an item's own block is repaired, and only from
    that item. A courier form carries an Assessable Value and a Country of
    Origin at document level as well as on every item, and the consignment's
    total must not be replaced by the first item's share of it.
    """
    items = boe.all_items() if boe is not None else []
    if not items:
        return fields
    for field in fields:
        item = _item_at(blocks, field) if blocks else items[0]
        if item is None:
            continue
        parsed = item.details.get(field.label)
        if parsed and field.value:
            field.value = parsed
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

    # An ICEGATE form has no item markers; its item-table labels are distinct
    # from its document-level ones, so the first item is the right source.
    blocks = _item_blocks(pdf, boe) if boe is not None and not form_type.startswith("ICEGATE") else []
    _use_parsed_item_values(fields, boe, blocks)

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
    kept = _split_and_clean(kept)
    shared = _shared_item_labels(blocks)
    kept = [f for f in kept
            if f.label not in DROP_LABELS and f.label != AS_HIGHLIGHTED
            and not _belongs_to_an_item(blocks, shared, f)]
    kept.sort(key=lambda f: (f.page, round(f.top), f.label))
    _qualify_repeats(kept)
    return kept
