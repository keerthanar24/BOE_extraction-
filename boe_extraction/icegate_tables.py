"""Reads the label/value tables of an ICEGATE Bill of Entry.

Nearly every block of the form is a heading line of numbered columns --
"1.BCD 2.ACD 3.SWS ..." -- with its values on the line beneath. Splitting
either line on whitespace loses the alignment, because empty cells collapse.
So a heading line is cut into columns at each "n." marker and every value is
placed in the column it sits under.

Some blocks print the value beside its heading instead ("13.COUNTRY OF ORIGIN
CHINA"). Those are recognised by the line below being another heading, and the
column is then split at the gap between the heading and its value.
"""

import re

from .pdf_text import upright_page, word_lines

# A column number: one or two digits and a dot, not the decimal point of a
# figure -- "1.BCD" and "12. PROV/" are headings, "84001.2" is a value.
NUMBERED = re.compile(r"^\d{1,2}\.(?!\d)")
# The value line sits within this many points below its heading line.
VALUE_LINE_GAP = 14
# The narrowest gap that separates a heading from a value beside it. The form
# sets both in the same run, so the break is only a couple of points wider than
# the spacing between words.
MIN_INLINE_GAP = 4.5
# Headings the form prints without a number.
BARE_LABELS = ("Port Code", "BE No", "BE Date", "BE Type", "IEC/Br",
               "GSTIN/TYPE", "CB CODE", "PKG", "G.WT (KGS)", "AD CODE",
               "EXCHANGE RATE")


class Column:
    """A heading, the words on its own line, and the band of page it owns."""

    __slots__ = ("words", "rest", "x0", "until", "fixed_label")

    def __init__(self, word, fixed_label=None):
        self.words = [word]
        self.rest = []
        self.x0 = word["x0"]
        self.until = float("inf")
        self.fixed_label = fixed_label

    def holds(self, word):
        centre = (word["x0"] + word["x1"]) / 2
        return self.x0 - 6 <= centre < self.until

    def split(self):
        """Separate the heading from a value printed beside it."""
        widest, at = 0, None
        for index in range(1, len(self.words)):
            gap = self.words[index]["x0"] - self.words[index - 1]["x1"]
            if gap > widest:
                widest, at = gap, index
        if at is None or widest < MIN_INLINE_GAP:
            return _text(self.words), ""
        return _text(self.words[:at]), _text(self.words[at:])

    @property
    def label(self):
        return self.fixed_label or self.split()[0]


def _text(words):
    return " ".join(w["text"] for w in words)


def _set_bounds(columns):
    for index, column in enumerate(columns):
        column.until = (columns[index + 1].x0 - 6 if index + 1 < len(columns)
                        else float("inf"))
    return columns


def _numbered_columns(line):
    columns = []
    for word in line:
        if NUMBERED.match(word["text"]):
            columns.append(Column(word))
        elif columns:
            columns[-1].words.append(word)
    return _set_bounds(columns)


def _bare_columns(line):
    """The form's unnumbered headings on a line, in order."""
    words = [w["text"] for w in line]
    columns = []
    for label in BARE_LABELS:
        parts = label.split()
        for start in range(len(words) - len(parts) + 1):
            if words[start:start + len(parts)] == parts:
                column = Column(line[start], fixed_label=label)
                column.words = line[start:start + len(parts)]
                column.rest = line[start + len(parts):]
                columns.append(column)
                break
    columns.sort(key=lambda c: c.x0)
    return _set_bounds(columns)


def _value_line(lines, index, heading):
    """The line of values under a heading line, if there is one."""
    bottom = max(w["bottom"] for w in heading)
    for line in lines[index + 1:]:
        if line[0]["top"] > bottom + VALUE_LINE_GAP:
            return []
        if line[0]["top"] >= bottom - 2:
            return line
    return []


def _is_heading(line):
    return bool(_numbered_columns(line) or _bare_columns(line))


class Pair:
    """A label, its value, and where on the page the two were read from."""

    __slots__ = ("page", "label", "value", "top", "bottom", "x0", "until")

    def __init__(self, page, label, value, top, bottom, x0, until):
        self.page = page
        self.label = label
        self.value = value
        self.top = top
        self.bottom = bottom
        self.x0 = x0
        self.until = until

    def __repr__(self):
        return f"Pair(p{self.page + 1} {self.label!r}={self.value!r})"


def read_tables(pdf):
    """Every label/value pair in the form's tables, in reading order."""
    pairs = []
    for page_index, page in enumerate(pdf.pages):
        lines = word_lines(upright_page(page))
        for index, line in enumerate(lines):
            columns = _numbered_columns(line)
            bare = [] if columns else _bare_columns(line)
            if not columns and not bare:
                continue

            below = _value_line(lines, index, line)
            # A heading line below means the values are printed inline.
            if below and _is_heading(below):
                below = []

            top = line[0]["top"]
            bottom = max([w["bottom"] for w in below] or
                         [w["bottom"] for w in line])

            for column in columns:
                # Only read a value from beside the heading when there is no
                # value line under it; otherwise the heading's own trailing
                # words would be mistaken for the value.
                label, inline = (_text(column.words), "") if below else column.split()
                value = inline or _text([w for w in below if column.holds(w)])
                pairs.append(Pair(page_index, label, value, top, bottom,
                                  column.x0, column.until))

            for column in bare:
                beside = [w for w in column.rest if column.holds(w)]
                value = _text(beside) or _text([w for w in below if column.holds(w)])
                pairs.append(Pair(page_index, column.label, value, top, bottom,
                                  column.x0, column.until))
    return pairs
