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

from .pdf_text import is_bold, upright_page, word_lines

# A column number: one or two digits and a dot, not the decimal point of a
# figure -- "1.BCD" and "12. PROV/" are headings, "84001.2" is a value.
NUMBERED = re.compile(r"^\d{1,2}\.(?!\d)")
# The value line sits within this many points below its heading line.
VALUE_LINE_GAP = 14
# A heading whose value is a name and address printed over several lines.
ADDRESS_BLOCK = re.compile(r"NAME & ADDRESS$")
# The form prints a one-letter status stamp at the end of some address lines.
TRAILING_STAMP = re.compile(r"\s+[A-Z]$")
# A wrapped continuation follows within this many points of the line above it.
CONTINUATION_GAP = 4
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
        """Separate the heading from a value printed beside it.

        The form sets headings in bold and values in regular, so the value
        starts at the first regular word. A few cells are bold throughout
        ("6. AD CODE 6480001"); there the break is the first gap wider than the
        form's word spacing -- the first, not the widest, because in
        "15.Term CIF No" the widest gap falls after the value.
        """
        for index in range(1, len(self.words)):
            if not is_bold(self.words[index]):
                return _text(self.words[:index]), _text(self.words[index:])
        for index in range(1, len(self.words)):
            gap = self.words[index]["x0"] - self.words[index - 1]["x1"]
            if gap >= MIN_INLINE_GAP:
                return _text(self.words[:index]), _text(self.words[index:])
        return _text(self.words), ""

    @property
    def label(self):
        return self.fixed_label or self.split()[0]


def _text(words):
    return " ".join(w["text"] for w in words)


def _join_rows(rows):
    """Join a value's lines, concatenating one that was split mid-token.

    "ONEYSZPGM" over "8034800" is one MAWB number, so it joins up closed. A
    wrapped phrase -- an address over two lines -- keeps its space.
    """
    joined = ""
    for row in rows:
        if not row:
            continue
        if not joined:
            joined = row
        elif " " in joined.strip() or " " in row.strip():
            joined += " " + row
        else:
            joined += row
    return joined.strip()


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


def _separate(numbered, bare):
    """Let an unnumbered heading claim its own space on a numbered line.

    "1.EVENT 2.DATE 3.TIME EXCHANGE RATE" carries both kinds, and without this
    the last numbered column swallows the unnumbered heading.
    """
    kept = []
    for column in bare:
        owner = next((c for c in numbered
                      if c.x0 < column.x0 < c.until), None)
        if owner is None:
            continue
        labels = set(id(w) for w in column.words)
        owner.words = [w for w in owner.words if id(w) not in labels]
        owner.until = min(owner.until, column.x0 - 6)
        column.rest = [w for w in owner.words if w["x0"] > column.x0]
        kept.append(column)
    return [c for c in numbered if c.words], kept


def _value_lines(lines, index, heading):
    """The value lines under a heading, including wrapped continuations.

    A long value runs onto a second line within its own column -- an MAWB
    number is printed as "ONEYSZPGM" over "8034800" and means
    "ONEYSZPGM8034800". Taking only the first line truncates it.
    """
    bottom = max(w["bottom"] for w in heading)
    collected = []
    for line in lines[index + 1:]:
        top = line[0]["top"]
        if not collected:
            if top < bottom - 2:
                continue           # still level with the heading itself
            if top > bottom + VALUE_LINE_GAP:
                break
            if _is_heading(line):
                break              # the values are printed beside the heading
            # The form sets headings in bold, so a wholly bold line that heads
            # nothing is the heading wrapping -- "7.ADV BE 11.FIRST 12. PROV/"
            # over "(Y/N/P) CHECK FINAL" -- and the values are further down.
            if all(is_bold(w) for w in line):
                continue
            collected.append(line)
            continue
        if all(is_bold(w) for w in line):
            break
        # A continuation sits just below the line before it and heads nothing.
        previous = max(w["bottom"] for w in collected[-1])
        if top > previous + CONTINUATION_GAP or _is_heading(line):
            break
        collected.append(line)
    return collected


def _value_line(lines, index, heading):
    """The first line of values under a heading, if there is one."""
    found = _value_lines(lines, index, heading)
    return found[0] if found else []


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
            bare = _bare_columns(line)
            if columns and bare:
                columns, bare = _separate(columns, bare)
            if not columns and not bare:
                continue

            rows = _value_lines(lines, index, line)
            # A heading line below means the values are printed inline.
            if rows and _is_heading(rows[0]):
                rows = []
            below = [w for row in rows for w in row]

            top = line[0]["top"]
            bottom = max([w["bottom"] for w in below] or
                         [w["bottom"] for w in line])

            def block(column, start):
                """A name and address printed as a block under its heading.

                The heading sits indented over a block that starts further
                left, so the block runs from where the previous column ended.
                Each line is cut short at the next label on it, since a
                neighbouring column's value can share the line.
                """
                collected = []
                for below in lines[index + 1:]:
                    # A wholly bold line is the next heading row. A line that
                    # merely carries a label alongside the address is cut, not
                    # stopped at.
                    if all(is_bold(w) for w in below):
                        break
                    labels = [w["x0"] for w in below
                              if is_bold(w) and w["x0"] > start]
                    limit = min([column.until] + labels)
                    held = [w["text"] for w in below if not is_bold(w)
                            and start <= (w["x0"] + w["x1"]) / 2 < limit]
                    if not held:
                        if collected:
                            break
                        continue
                    collected.append(TRAILING_STAMP.sub("", " ".join(held)))
                return ", ".join(c for c in collected if c)

            def under(column):
                """The column's value, with any wrapped line joined back on."""
                return _join_rows([_text([w for w in row if column.holds(w)])
                                   for row in rows])

            for column in columns:
                # Only read a value from beside the heading when there is no
                # value line under it; otherwise the heading's own trailing
                # words would be mistaken for the value.
                label, inline = (_text(column.words), "") if below else column.split()
                value = inline or under(column)
                if ADDRESS_BLOCK.search(label):
                    position = columns.index(column)
                    start = columns[position - 1].until if position else 0
                    value = block(column, start) or value
                pairs.append(Pair(page_index, label, value,
                                  top, bottom, column.x0, column.until))

            for column in bare:
                beside = [w for w in column.rest if column.holds(w)]
                pairs.append(Pair(page_index, column.label,
                                  _text(beside) or under(column),
                                  top, bottom, column.x0, column.until))
    return pairs
