"""Reads the highlight annotations a reviewer drew on a Bill of Entry.

The marked-up documents carry PDF highlight annotations over the fields that
must be extracted. Each highlight is one or more quadrilaterals -- one per line
of text it spans.

Text under a highlight is collected by testing each word's centre against those
quads rather than by cropping the page: a crop takes in the neighbouring rows
too, and pdfplumber then merges them into one line, so "Port Code" over
"INNSA1" comes back as "IPNoNrtS CAo1d".
"""

from .pdf_text import upright_page, word_lines

HIGHLIGHT = "/'Highlight'"
# A highlight is drawn a little taller than the text it covers.
QUAD_PADDING = 1.0


class Highlight:
    """One highlight annotation and the text it covers."""

    __slots__ = ("page", "top", "x0", "colour", "quads", "text")

    def __init__(self, page, top, x0, colour, quads, text):
        self.page = page
        self.top = top
        self.x0 = x0
        self.colour = colour
        self.quads = quads
        self.text = text

    def covers(self, word):
        return _inside(word, self.quads)

    def __repr__(self):
        return f"Highlight(p{self.page} y{int(self.top)} {self.text!r})"


def _quads(annot, page_height):
    """The highlight's quads as (x0, top, x1, bottom) in page coordinates."""
    data = annot.get("data", {})
    points = data.get("QuadPoints")
    if not points:
        left, bottom, right, top = (float(v) for v in data["Rect"])
        return [(left, page_height - top, right, page_height - bottom)]
    points = [float(value) for value in points]
    boxes = []
    for start in range(0, len(points) - 7, 8):
        xs = points[start:start + 8:2]
        ys = points[start + 1:start + 8:2]
        boxes.append((min(xs), page_height - max(ys), max(xs), page_height - min(ys)))
    return boxes


def _inside(word, quads):
    """Whether a word's centre falls inside any of the quads."""
    x = (word["x0"] + word["x1"]) / 2
    y = (word["top"] + word["bottom"]) / 2
    for left, top, right, bottom in quads:
        if (left - QUAD_PADDING <= x <= right + QUAD_PADDING
                and top - QUAD_PADDING <= y <= bottom + QUAD_PADDING):
            return True
    return False


def read(pdf):
    """Every highlight in the document, in reading order."""
    found = []
    for page_index, page in enumerate(pdf.pages):
        annots = [a for a in page.annots
                  if str(a.get("data", {}).get("Subtype")) == HIGHLIGHT]
        if not annots:
            continue
        lines = word_lines(upright_page(page))
        for annot in annots:
            quads = _quads(annot, page.height)
            covered = []
            for line in lines:
                words = [w["text"] for w in line if _inside(w, quads)]
                if words:
                    covered.append(" ".join(words))
            colour = tuple(round(float(c), 3) for c in annot.get("data", {}).get("C", ()))
            found.append(Highlight(page_index, annot["top"], annot["x0"],
                                   colour, quads, " ".join(covered).strip()))
    found.sort(key=lambda h: (h.page, round(h.top), h.x0))
    return found
