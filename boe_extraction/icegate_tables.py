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
# A wrapped continuation follows within this many points of the line above it.
CONTINUATION_GAP = 4
# The narrowest gap that separates one address block from the next across the
# page. Words inside a block sit a few points apart; the blocks themselves are
# a whole column apart.
BLOCK_GAP = 20
# How far right of a block's left edge a label may start and still count as
# beginning the form's next row: "AD CODE" is set a couple of points in from
# the address above it, while a label belonging to the column beside the block
# sits a whole column away.
BLOCK_ROW_MARGIN = 12
# The narrowest gap that separates a heading from a value beside it. The form
# sets both in the same run, so the break is only a couple of points wider than
# the spacing between words.
MIN_INLINE_GAP = 4.5
# Headings the form prints without a number.
BARE_LABELS = ("Port Code", "BE No", "BE Date", "BE Type", "IEC/Br",
               "GSTIN/TYPE", "CB CODE", "PKG", "G.WT (KGS)", "AD CODE",
               "EXCHANGE RATE", "OOC NO.", "OOC DATE")


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

    def split(self, limit=float("inf")):
        """Separate the heading from a value printed beside it.

        The form sets headings in bold and values in regular, so the value
        starts at the first regular word. A few cells are bold throughout
        ("6. AD CODE 6480001"); there the break is the first gap wider than the
        form's word spacing -- the first, not the widest, because in
        "15.Term CIF No" the widest gap falls after the value.

        A line can also carry the value of a heading printed on the line above
        it -- the "No" in "15.Term CIF No" answers 9.RELTD, one row up -- so
        the value stops at `limit`, where that heading's own column begins.
        """
        value = [w for w in self.words if w["x0"] < limit]
        for index in range(1, len(value)):
            if not is_bold(value[index]):
                return _text(value[:index]), _text(value[index:])
        for index in range(1, len(value)):
            gap = value[index]["x0"] - value[index - 1]["x1"]
            if gap >= MIN_INLINE_GAP:
                return _text(value[:index]), _text(value[index:])
        return _text(value), ""

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


def _next_column(above, column):
    """Where the row above starts its next column, right of this one.

    A heading whose value is printed on the line below leaves that line
    carrying two cells: its own label and value, and the value belonging to
    the heading above. This is the boundary between them.
    """
    beyond = [c.x0 for c in above if c.x0 > column.x0 + MIN_INLINE_GAP]
    return min(beyond) if beyond else float("inf")


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


def _block_edges(body):
    """Where each address block on these lines starts, left to right.

    A block starts wherever a line's words resume after a gap wider than the
    spacing inside one. Taking the leftmost start of each run across the whole
    block gives the column edges, which is what a block must be cut on: the
    headings sit indented over their blocks, so cutting on a heading takes the
    first words of the block beside it.
    """
    starts = set()
    for words in body:
        ordered = sorted(words, key=lambda w: w["x0"])
        starts.add(ordered[0]["x0"])
        for previous, word in zip(ordered, ordered[1:]):
            if word["x0"] - previous["x1"] >= BLOCK_GAP:
                starts.add(word["x0"])
    edges = sorted(starts)
    # Two lines of one block rarely start at the same x to the point, so fold
    # starts that sit within a word's width of each other.
    folded = []
    for edge in edges:
        if not folded or edge - folded[-1] >= BLOCK_GAP:
            folded.append(edge)
    return folded


def _block_columns(address, body):
    """Which stretch of the page each address heading's block occupies.

    Headings are claimed left to right, each taking the rightmost edge that
    still starts at or before it and has not been claimed. A heading with no
    edge left to claim heads an empty block -- the form prints the heading
    whether or not the filer answered it.
    """
    edges = _block_edges(body)
    claimed, bounds = set(), {}
    for column in sorted(address, key=lambda c: c.x0):
        available = [e for e in edges
                     if e <= column.x0 + MIN_INLINE_GAP and e not in claimed]
        if not available:
            continue
        left = available[-1]
        claimed.add(left)
        beyond = [e for e in edges if e > left]
        bounds[id(column)] = ((left - 2, beyond[0] - 2 if beyond else float("inf")),
                              body)
    return bounds


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

            def block_lines():
                """The lines of the address blocks printed under this heading.

                A neighbouring column's label can sit on a block's line -- the
                customs broker's name is printed beside the importer's address
                -- so a label alone does not end the block. What does end it is
                a label starting at the block's own left edge, because that is
                the form beginning its next row.
                """
                found, left = [], None
                for below in lines[index + 1:]:
                    words = [w for w in below if not is_bold(w)]
                    if not words:
                        if found:
                            break
                        continue
                    labels = [w["x0"] for w in below if is_bold(w)]
                    edge = min(w["x0"] for w in words)
                    if left is None:
                        left = edge
                    elif labels and min(labels) <= left + BLOCK_ROW_MARGIN:
                        break
                    found.append(words)
                return found

            def block(edges, body):
                """One address block, read between the edges its column owns."""
                left, right = edges
                collected = []
                for words in body:
                    held = [w["text"] for w in words
                            if left <= (w["x0"] + w["x1"]) / 2 < right]
                    if held:
                        collected.append(" ".join(held))
                return ", ".join(c for c in collected if c)

            def under(column):
                """The column's value, with any wrapped line joined back on."""
                return _join_rows([_text([w for w in row if column.holds(w)])
                                   for row in rows])

            # A heading's own words are its label only where values are
            # printed below it; beside them, the label stops at the value. So
            # the label is settled first, and the address blocks picked from it.
            above = _numbered_columns(lines[index - 1]) if index else []
            read = {id(c): ((_text(c.words), "") if below
                            else c.split(_next_column(above, c)))
                    for c in columns}
            address = [c for c in columns
                       if ADDRESS_BLOCK.search(read[id(c)][0])]
            blocks = _block_columns(address, block_lines()) if address else {}

            for column in columns:
                label, inline = read[id(column)]
                value = inline or under(column)
                if id(column) in blocks:
                    value = block(*blocks[id(column)]) or value
                pairs.append(Pair(page_index, label, value,
                                  top, bottom, column.x0, column.until))

            for column in bare:
                beside = [w for w in column.rest if column.holds(w)]
                pairs.append(Pair(page_index, column.label,
                                  _text(beside) or under(column),
                                  top, bottom, column.x0, column.until))
    return pairs
