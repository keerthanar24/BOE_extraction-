"""Shared PDF text helpers for Bill of Entry extraction.

ICEGATE and ECCS bills of entry are stamped with rotated overlay text (the
section labels running up the left margin, the diagonal "PROCESSED" watermark).
pdfplumber returns those glyphs interleaved with the real cell text, which
corrupts words -- "LAPTOP COVER" comes back as "LAPTOSP COVER".  Every rotated
glyph carries a non-zero ``matrix[1]``, so dropping those before laying out the
text gives clean, deterministic output.
"""

X_TOLERANCE = 1.2


def _is_upright(obj):
    if obj.get("object_type") != "char":
        return True
    return round(obj.get("matrix", (1, 0, 0, 1, 0, 0))[1], 2) == 0


def upright_page(page):
    """The page with all rotated glyphs (margin labels, watermark) removed."""
    return page.filter(_is_upright)


def page_text(page):
    return upright_page(page).extract_text(x_tolerance=X_TOLERANCE) or ""


def page_words(page):
    return upright_page(page).extract_words(x_tolerance=X_TOLERANCE)


def document_text(pdf):
    return "\n".join(page_text(p) for p in pdf.pages)


def word_lines(page, y_tolerance=3.0):
    """Words grouped into visual lines, each sorted left to right."""
    lines = []
    for word in sorted(page_words(page), key=lambda w: (w["top"], w["x0"])):
        if lines and abs(word["top"] - lines[-1][0]["top"]) <= y_tolerance:
            lines[-1].append(word)
        else:
            lines.append([word])
    return [sorted(line, key=lambda w: w["x0"]) for line in lines]
