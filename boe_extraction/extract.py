"""Entry point: identify a Bill of Entry PDF and extract its line items."""

from pathlib import Path

import pdfplumber

from . import courier_parser, standard_parser
from .pdf_text import page_text

# Deterministic parsers first; the ICEGATE reader is the general fallback.
PARSERS = (courier_parser, standard_parser)


def identify(pdf):
    """The form type of the document, read from page 1. None if unrecognised."""
    first = page_text(pdf.pages[0])
    for parser in PARSERS:
        form_type = parser.matches(first)
        if form_type:
            return form_type, parser
    return None, None


def _parse(pdf, path):
    form_type, parser = identify(pdf)
    if parser is None:
        raise ValueError(
            f"{path}: page 1 does not identify a supported Bill of Entry form. "
            "A scanned or image-only first page cannot be recognised.")
    return form_type, parser.parse(pdf, form_type, Path(path).name)


def extract(path):
    """Parse a Bill of Entry PDF into a BillOfEntry."""
    with pdfplumber.open(str(path)) as pdf:
        return _parse(pdf, path)[1]


def extract_with_highlights(path):
    """Parse the document and resolve the highlights drawn on it.

    Returns (BillOfEntry, resolved fields, raw highlights).
    """
    from .highlight_fields import collect
    from .highlights import read as read_highlights

    with pdfplumber.open(str(path)) as pdf:
        form_type, boe = _parse(pdf, path)
        return boe, collect(pdf, form_type, boe), read_highlights(pdf)
