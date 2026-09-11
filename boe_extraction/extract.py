"""Entry point: identify a Bill of Entry PDF and extract its line items."""

from dataclasses import dataclass, field
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


@dataclass
class Document:
    """One parsed bill of entry, with whatever was highlighted on it."""

    name: str
    boe: object
    fields: list = field(default_factory=list)
    highlights: list = field(default_factory=list)


def extract_document(path, with_highlights=True):
    """Parse a document, and resolve its highlights unless asked not to."""
    boe, fields, highlights = (extract_with_highlights(path) if with_highlights
                               else (extract(path), [], []))
    return Document(Path(path).name, boe, fields, highlights)


def document_fields(path):
    """Every label/value the document carries above its item blocks."""
    from .cbe_grid import Cell, Marker, read_entries

    with pdfplumber.open(str(path)) as pdf:
        _, boe = _parse(pdf, path)
        rows, section = [], ""
        for entry in read_entries(pdf):
            if isinstance(entry, Marker):
                if entry.text.upper().startswith("ITEM"):
                    break
                section = section_name(entry.text.rstrip(" :,"))
            elif isinstance(entry, Cell) and entry.label and entry.label_done:
                rows.append((section, entry.label, entry.value))
    return boe, rows


def schema_extract(path):
    """A document read into the named schema: its fields and its line items.

    Unlike document_fields, which reports whatever the form prints, this
    returns exactly the fields the schema names -- no more, and each one once.
    """
    from .cbe_grid import Cell, Marker, read_entries
    from .highlight_fields import document_cells
    from .schema import document_rows, item_rows, section_name

    with pdfplumber.open(str(path)) as pdf:
        _, boe = _parse(pdf, path)
        if boe.form_type == "ICEGATE BOE":
            cells = document_cells(pdf, boe)
        else:
            cells, section = [], ""
            for entry in read_entries(pdf):
                if isinstance(entry, Marker):
                    section = section_name(entry.text.rstrip(" :,"))
                elif isinstance(entry, Cell) and entry.label:
                    cells.append((section, entry.label, entry.value))

    fields = document_rows(boe.form_type, cells)
    if fields is None:
        raise ValueError(f"{Path(path).name}: no schema for {boe.form_type}")
    from .model import ITEM_COLUMNS
    items = item_rows(boe, {title: name for name, title in ITEM_COLUMNS})
    return boe, fields, items


def verify_document(path):
    """Parse a document and check it against the totals it declares."""
    from .verify import verify

    with pdfplumber.open(str(path)) as pdf:
        _, boe = _parse(pdf, path)
        return boe, verify(pdf, boe)


def extract_with_highlights(path):
    """Parse the document and resolve the highlights drawn on it.

    Returns (BillOfEntry, resolved fields, raw highlights).
    """
    from .highlight_fields import collect
    from .highlights import read as read_highlights

    with pdfplumber.open(str(path)) as pdf:
        form_type, boe = _parse(pdf, path)
        return boe, collect(pdf, form_type, boe), read_highlights(pdf)
