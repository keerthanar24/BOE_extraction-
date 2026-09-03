"""Reads an ECCS Courier Bill of Entry as label/value pairs.

The CBE-XIII and CBE-XIV forms are a rigid two-column grid: a right-aligned
label ending in a colon, then its value in a fixed column to the right, twice
across the page.  Labels and values both wrap onto continuation lines, breaking
words with a hyphen ("Courier Registration Num-" / "ber:").

Reading it by line regex goes wrong as soon as a value itself starts with a
capitalised word -- "Unit of Measure : PCS Quantity : 15" then looks like an
empty value followed by a label "PCS Quantity".  So the grid is read by
geometry instead: the two colon columns are located on the page, and every word
is placed relative to them.
"""

import re
from collections import Counter

from .pdf_text import upright_page, word_lines

# A right-aligned label ends within this band around its colon column.
LABEL_END_BAND = (-6, 8)
# The value column starts just past the label column.
VALUE_OFFSET = 5
# The narrowest gap that can separate a left-hand value from a right-hand label.
MIN_LABEL_GAP = 20
# A left-hand value starts within this distance of its own label column.
VALUE_COLUMN_WIDTH = 60
# How near the page centre a line must sit to read as a section heading.
HEADING_CENTRE_SLACK = 45

SECTION_MARKER = re.compile(
    r"^(Details\s+Of\s+(?:Item|Invoice)\s*-\s*\d+)$", re.IGNORECASE)
HEADING_TEXT = re.compile(r"^[A-Z0-9][A-Z0-9 ()\[\]/,.&:;'\-]{7,}$")
PAGE_FOOTER = re.compile(r"^Page\s+\d+\s+of\s+\d+$", re.IGNORECASE)


class Cell:
    """One label and its value, with the position it was read from."""

    __slots__ = ("page", "top", "column", "_label", "_value", "label_done")

    def __init__(self, page, top, column):
        self.page = page
        self.top = top
        self.column = column
        self._label = ""
        self._value = ""
        self.label_done = False

    def add_label(self, text):
        self._label = _join(self._label, text)
        self.label_done = text.endswith(":")

    def add_value(self, text):
        self._value = _join(self._value, text)

    @property
    def label(self):
        return re.sub(r"\s*:\s*$", "", self._label).strip()

    @property
    def value(self):
        return self._value.strip()

    def __repr__(self):
        return f"Cell({self.label!r}={self.value!r} p{self.page} y{int(self.top)})"


class Marker:
    """A section heading or an item/invoice divider."""

    __slots__ = ("page", "top", "text")

    def __init__(self, page, top, text):
        self.page = page
        self.top = top
        self.text = text

    label = ""
    value = ""

    def __repr__(self):
        return f"Marker({self.text!r} p{self.page} y{int(self.top)})"


def _colon_columns(pages):
    """The x of the two label columns, taken from where the colons line up."""
    ends = Counter()
    for page in pages:
        for line in word_lines(upright_page(page)):
            for word in line:
                if word["text"].endswith(":"):
                    ends[round(word["x1"])] += 1
    ranked = [x for x, _ in ends.most_common(6)]
    if not ranked:
        return None
    left = min(ranked[:2]) if len(ranked) > 1 else ranked[0]
    wide = [x for x in ranked if x > left + 100]
    return (left, max(wide)) if wide else None


def _join(head, tail):
    """Append a wrapped fragment to what came before it."""
    if not head:
        return tail
    if not tail:
        return head
    if head.endswith("-"):
        return head[:-1] + tail
    return head + " " + tail


def _text(words):
    return " ".join(w["text"] for w in words).strip()


def _split_off_label(middle, left_column, right_column):
    """Separate a left-hand value from a right-hand label sharing its line.

    A right-hand label is right-aligned on the colon column, so when the middle
    band ends there the label is the run of words after the widest gap.
    """
    if not middle:
        return [], []
    low, high = LABEL_END_BAND
    if not (right_column + low <= middle[-1]["x1"] <= right_column + high):
        return middle, []
    # Nothing sitting in the left value column means the whole band is a label.
    if middle[0]["x0"] > left_column + VALUE_COLUMN_WIDTH:
        return [], middle
    widest, split_at = 0, None
    for index in range(1, len(middle)):
        gap = middle[index]["x0"] - middle[index - 1]["x1"]
        if gap > widest:
            widest, split_at = gap, index
    if split_at is None:
        return [], middle
    if widest < MIN_LABEL_GAP:
        return middle, []
    return middle[:split_at], middle[split_at:]


def _is_heading(line, page_centre, columns):
    text = _text(line)
    if SECTION_MARKER.match(text):
        return True
    if not HEADING_TEXT.match(text):
        return False
    # An all-caps line carrying a colon on a label column is a label, not a
    # heading -- "CTSH : 91021100 CETSH : 91021100".
    low, high = LABEL_END_BAND
    for word in line:
        if not word["text"].endswith(":"):
            continue
        if any(column + low <= word["x1"] <= column + high for column in columns):
            return False
    centre = (line[0]["x0"] + line[-1]["x1"]) / 2
    return abs(centre - page_centre) <= HEADING_CENTRE_SLACK


class _Column:
    """Accumulates the cells of one of the form's two label columns."""

    def __init__(self, name, entries):
        self.name = name
        self.entries = entries
        self.current = None

    def label(self, page, top, text):
        # A finished label -- or a value that arrived without one -- means the
        # next label fragment opens a new cell.
        if (self.current is None or self.current.label_done
                or (self.current.value and not self.current.label)):
            self.current = Cell(page, top, self.name)
            self.entries.append(self.current)
        self.current.add_label(text)

    def value(self, page, top, text):
        if self.current is None:
            self.current = Cell(page, top, self.name)
            self.entries.append(self.current)
        self.current.add_value(text)

    def close(self):
        self.current = None


def read_entries(pdf):
    """Cells and section markers across the document, in reading order."""
    columns = _colon_columns(pdf.pages)
    if columns is None:
        return []
    left_column, right_column = columns
    label_edge = left_column + LABEL_END_BAND[1]
    value_edge = right_column + VALUE_OFFSET

    entries = []
    for page_index, page in enumerate(pdf.pages):
        page_centre = page.width / 2
        # Cells wrap across lines but never across a page break.
        left = _Column("left", entries)
        right = _Column("right", entries)
        for line in word_lines(upright_page(page)):
            top = line[0]["top"]
            if PAGE_FOOTER.match(_text(line)):
                continue
            if _is_heading(line, page_centre, columns):
                left.close()
                right.close()
                entries.append(Marker(page_index, top, _text(line)))
                continue

            label_words = [w for w in line if w["x1"] <= label_edge]
            middle = [w for w in line if w["x1"] > label_edge and w["x0"] < value_edge]
            tail = [w for w in line if w["x0"] >= value_edge]

            # A middle band overflowing the value column is one long left-hand
            # value (an address, say), not a value followed by a label.
            if middle and tail and middle[-1]["x1"] > value_edge:
                middle, tail = middle + tail, []

            value_words, right_label_words = _split_off_label(
                middle, left_column, right_column)

            if label_words:
                left.label(page_index, top, _text(label_words))
            if value_words:
                left.value(page_index, top, _text(value_words))
            if right_label_words:
                right.label(page_index, top, _text(right_label_words))
            if tail:
                right.value(page_index, top, _text(tail))
    return entries


def document_lines(pdf):
    """Every text line of the document as (page index, top, text)."""
    lines = []
    for page_index, page in enumerate(pdf.pages):
        for line in word_lines(upright_page(page)):
            text = _text(line)
            if text and not PAGE_FOOTER.match(text):
                lines.append((page_index, line[0]["top"], text))
    return lines
