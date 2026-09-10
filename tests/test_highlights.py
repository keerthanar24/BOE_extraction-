"""Checks against the highlight annotations drawn on the two sample documents.

Every expected value here was read off the highlighted PDF by hand.
"""

import sys
from pathlib import Path

import pdfplumber
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from boe_extraction.extract import extract_with_highlights
from boe_extraction.highlights import read as read_highlights

SAMPLES = Path(__file__).resolve().parents[1] / "samples"
COURIER = SAMPLES / "FBA15M6L9KGF01_courier_cbe_xiv.pdf"
STANDARD = SAMPLES / "BOE_3141398_icegate.pdf"


@pytest.fixture(scope="module")
def courier():
    return extract_with_highlights(COURIER)


@pytest.fixture(scope="module")
def standard():
    return extract_with_highlights(STANDARD)


def _values(fields):
    return {f.label: f.value for f in fields}


def test_every_highlight_is_read(courier, standard):
    assert len(courier[2]) == 55
    assert len(standard[2]) == 84


def test_highlight_text_is_not_merged_across_rows():
    """A crop would return "Port Code" over "INNSA1" as "IPNoNrtS CAo1d"."""
    with pdfplumber.open(STANDARD) as pdf:
        texts = [h.text for h in read_highlights(pdf) if h.page == 0]
    assert "Port Code" in texts
    assert "INNSA1" in texts


def test_courier_highlighted_fields(courier):
    values = _values(courier[1])
    assert values["CBEXIV Number"] == "CBEXIV_DEL_2026-2027_2808_10570"
    assert values["Import export Code"] == "AAFCV5265N"
    assert values["KYC ID"] == "29AAFCV5265N1ZK"
    assert values["Import General Manifest (IGM) Number"] == "3129351"
    assert values["Master Airway Bill (MAWB) Number"] == "31296805785"
    assert values["House Airway Bill (HAWB) Number"] == "FBA15M6L9KGF01"
    assert values["Invoice Number"] == "FBA15M6L9KGF01"
    assert values["Invoice Value"] == "1080"
    assert values["Terms of Invoice"] == "CIF"
    # Per-item fields live on the Line Items sheet, not among the document's.
    assert "CTSH" not in values
    assert "Assessable Value" not in values


def test_section_banners_are_not_reported_as_fields(courier):
    """A highlight over a heading or a multi-row table names no field.

    Its figures are reported against each line item instead, so it is left out
    of the document's fields and kept only in Highlights (raw).
    """
    assert not [f for f in courier[1] if f.label == "(as highlighted)"]
    raw = " ".join(h.text for h in courier[2])
    assert "1 BCD 20 0 0 3908" in raw


def test_single_row_tables_become_fields(courier):
    """A bold heading row over one row of figures is a record, so read it."""
    values = _values(courier[1])
    assert values["Current Status of the CBE"] == "OOC ISSUED on 29-08-2026 07:11"
    assert values["TR-6 Challan Number"] == "2908387251"
    assert values["Total Amount"] == "45833"
    assert values["Challan Date"] == "29/08/2026"
    assert values["Airlines"] == "Indigo Airlines"
    assert values["Flight No."] == "6E 1074"
    assert values["Airport Of Arrival"] == "DEL"
    assert values["Date Of Arrival"] == "27/08/2026"


def test_sections_tell_repeated_labels_apart(courier):
    """The importer, the supplier and the broker each have a Name."""
    names = {f.key: f.value for f in courier[1] if f.label == "Name"}
    assert names["PARTICULARS OF THE IMPORTER · Name"] == "VALUECART PRIVATE LIMITED"
    assert names["SUPPLIER DETAILS · Name"] == "GATI HONG KONG LIMITED"
    assert names["BROKER/ AGENT DETAILS · Name"] == "KBR INTERNATIONAL LOGISTICS"


def test_standard_highlighted_summary_totals(standard):
    """The Part I summary is a heading row over a value row."""
    values = _values(standard[1])
    assert values["1.BCD"] == "84001.2"
    assert values["3.SWS"] == "8400.3"
    assert values["7.IGST"] == "117434"
    assert values["18.TOT.ASS VAL"] == "560008"
    assert values["14.TOTAL DUTY"] == "209835"
    assert values["19.TOT. AMOUNT"] == "209835"


def test_standard_highlighted_header_fields(standard):
    values = _values(standard[1])
    assert values["Port Code"] == "INNSA1"
    assert values["BE No"] == "3141398"
    assert values["BE Date"] == "14/08/2026"
    assert values["G.WT (KGS)"] == "691"      # the barcode beside it is dropped
    assert values["GSTIN"] == "27AAFCV5265N1ZO"
    assert values["TYPE"] == "G"
    assert values["IEC"] == "AAFCV5265N"
    assert values["Br"] == "1"
    assert values["CB CODE"] == "AAHFS9149KCH001"
    assert values["AD CODE"] == "6480001"
    assert values["13.COUNTRY OF ORIGIN"] == "CHINA"
    assert values["15.PORT OF LOADING"] == "Shekou"
    assert values["2.CB NAME"] == "INTERLINK SHIPPING & CLEARING"
    assert values["14.ASS. VALUE"] == "292981.32"


def test_a_name_and_address_is_read_in_full(standard):
    """The heading sits indented over a block that starts further left.

    Reading only the line beneath it returned the name without the address.
    """
    values = _values(standard[1])
    assert values["1.IMPORTER NAME & ADDRESS"] == (
        "VALUECART PRIVATE LIMITED, FLAT No-4, G. S. TOWERS, OPPOSITE, "
        "BIBWEWADI, PUNE, PUNE, 411037")
    assert values["1.BUYER'S NAME & ADDRESS"].endswith("411037")
    assert values["3.SUPPLIER NAME & ADDRESS"].startswith(
        "GATI HONG KONG LIMITED, FLAT/RM C1303")
    # This document names no third party; the block under that heading is
    # empty, and must not borrow the supplier's.
    assert values["4.THIRD PARTY NAME & ADDRESS"] == ""


def test_every_item_is_reported_on_line_items(tmp_path, courier):
    """A field highlighted on item 1 is wanted for all of them."""
    from openpyxl import load_workbook

    from boe_extraction.excel_writer import write_highlighted

    boe, fields, highlights = courier
    path = write_highlighted(boe, fields, highlights, tmp_path / "h.xlsx")
    workbook = load_workbook(path)
    assert workbook.sheetnames == ["Highlighted Fields", "Line Items",
                                   "Highlights (raw)"]
    sheet = workbook["Line Items"]
    header = [c.value for c in sheet[1]]
    assert sheet.max_row == 5  # header plus four items

    column = header.index("Assessable Value") + 1
    assert [sheet.cell(row=r, column=column).value for r in range(2, 6)] == [
        19541.25, 26055, 6513.75, 52110]


def test_workbook_records_every_highlight_verbatim(tmp_path, standard):
    from openpyxl import load_workbook

    from boe_extraction.excel_writer import write_highlighted

    boe, fields, highlights = standard
    path = write_highlighted(boe, fields, highlights, tmp_path / "h.xlsx")
    sheet = load_workbook(path)["Highlights (raw)"]
    assert sheet.max_row == len(highlights) + 1


def test_combined_workbook_holds_every_document(tmp_path):
    """One workbook per run, with every document side by side."""
    from openpyxl import load_workbook

    from boe_extraction.excel_writer import write_combined
    from boe_extraction.extract import extract_document

    documents = [extract_document(COURIER), extract_document(STANDARD)]
    path = write_combined(documents, tmp_path / "combined.xlsx")
    workbook = load_workbook(path)
    assert workbook.sheetnames == ["Invoices", "Line Items",
                                   "Highlighted Fields", "Highlights (raw)"]

    # Every sheet names the document each row came from.
    for name in workbook.sheetnames:
        assert workbook[name]["A1"].value == "Document"

    assert workbook["Invoices"].max_row == 4           # one courier, two ICEGATE
    assert workbook["Line Items"].max_row == 14        # 4 + 9 items
    assert workbook["Highlights (raw)"].max_row == 55 + 84 + 1


def test_combined_totals_match_the_documents(tmp_path):
    from openpyxl import load_workbook

    from boe_extraction.excel_writer import write_combined
    from boe_extraction.extract import extract_document

    documents = [extract_document(COURIER), extract_document(STANDARD)]
    path = write_combined(documents, tmp_path / "combined.xlsx")
    sheet = load_workbook(path)["Invoices"]
    rows = [r for r in sheet.iter_rows(min_row=2, values_only=True)]

    courier = next(r for r in rows if r[0] == COURIER.name)
    assert (courier[1], courier[7], courier[8], courier[9]) == (
        "FBA15M6L9KGF01", 4, 104220, 45814)

    standard = [r for r in rows if r[0] == STANDARD.name]
    assert [r[1] for r in standard] == ["FBA15M13GSD3", "FBA15M16XHDH"]
    assert sum(r[8] for r in standard) == 560008


def test_mandatory_workbook_uses_the_fields_as_columns(tmp_path):
    """The highlights are the required-field list, so they are the columns."""
    from openpyxl import load_workbook

    from boe_extraction.excel_writer import write_mandatory
    from boe_extraction.extract import extract_document

    documents = [extract_document(COURIER), extract_document(STANDARD)]
    path = write_mandatory(documents, tmp_path / "mandatory.xlsx")
    workbook = load_workbook(path)
    assert workbook.sheetnames == ["Mandatory Fields", "Mandatory Item Fields",
                                   "Line Items", "Field Checklist",
                                   "Highlights (raw)"]

    sheet = workbook["Mandatory Fields"]
    assert sheet.max_row == 3                      # header plus two documents
    header = [c.value for c in sheet[1]]
    assert header[:3] == ["Document", "Form Type", "BE Number"]
    for column in ["CBEXIV Number", "TR-6 Challan Number", "1.BCD",
                   "PARTICULARS OF THE IMPORTER · Name"]:
        assert column in header

    rows = {r[0]: r for r in sheet.iter_rows(min_row=2, values_only=True)}
    courier = rows[COURIER.name]
    assert courier[header.index("CBEXIV Number")] == "CBEXIV_DEL_2026-2027_2808_10570"
    assert courier[header.index("TR-6 Challan Number")] == "2908387251"
    standard = rows[STANDARD.name]
    assert standard[header.index("1.BCD")] == "84001.2"
    assert standard[header.index("EXCHANGE RATE")] == "1 USD=96.05INR"


def test_item_rows_cover_every_item_of_every_document(tmp_path):
    from openpyxl import load_workbook

    from boe_extraction.excel_writer import write_mandatory
    from boe_extraction.extract import extract_document

    documents = [extract_document(COURIER), extract_document(STANDARD)]
    path = write_mandatory(documents, tmp_path / "mandatory.xlsx")
    sheet = load_workbook(path)["Mandatory Item Fields"]
    assert sheet.max_row == 4 + 9 + 1


def test_the_item_table_is_reported_on_line_items(standard):
    """Part II's item columns are per item, so they belong on Line Items."""
    values = _values(standard[1])
    for label in ("1.S NO.", "2.CTH", "3.DESCRIPTION", "4.UNIT PRICE",
                  "5.QUANTITY", "6.UQC", "7.AMOUNT"):
        assert label not in values
    first = standard[0].all_items()[0]
    assert first.hs_code == "39269099"
    assert first.description == "X002HTKLZJ MACBOOK PRO 16 INCH CASE - PC"
    assert (first.unit_price, first.quantity, first.unit_of_measure) == (4.62, 40, "NOS")


def test_a_blank_field_does_not_borrow_an_item_value(courier):
    """The courier form carries a Country of Origin at both levels.

    Only the document's own is reported here -- the item's is on Line Items --
    and a blank one must stay blank rather than take the item's value.
    """
    origins = [f for f in courier[1] if f.label == "Country of Origin"]
    assert [f.value for f in origins] == [""]
    assert origins[0].section == "SPECIAL REQUESTS"
    assert courier[0].all_items()[0].details["Country of Origin"] == "CHINA"


CBE_XIII_HIGHLIGHTED = SAMPLES / "FBA15M1ZPS2Y01_cbe_xiii_highlighted.pdf"


def test_highlights_resolve_on_a_cbe_xiii():
    """The highlight path had only ever run on the other two forms."""
    boe, fields, highlights = extract_with_highlights(CBE_XIII_HIGHLIGHTED)
    assert boe.form_type == "CBE-XIII"
    assert len(highlights) == 9
    values = _values(fields)
    assert values["Import Export Code"] == "AAFCV5265N"
    assert values["KYC ID"] == "29AAFCV5265N1ZK"
    # Its per-item fields are reported on Line Items instead.
    assert "CTSH" not in values
    assert boe.all_items()[0].hs_code == "42023290"


def test_a_document_total_is_not_replaced_by_an_items_share():
    """A courier form carries an Assessable Value at both levels.

    The consignment's total must survive; only a field inside an item's own
    block may be repaired from that item.
    """
    _, fields, _ = extract_with_highlights(CBE_XIII_HIGHLIGHTED)
    values = _values(fields)
    assert values["Assessable Value"] == "82062.07"   # not item 1's 2844.07
    assert values["Duty(Rs.)"] == "29648"             # not item 1's 1066
    boe, _, _ = extract_with_highlights(CBE_XIII_HIGHLIGHTED)
    assert boe.be_number == "CBEXIII_DEL_2026-2027_2707_14754"


def test_a_value_wrapped_onto_a_second_line_is_joined(standard):
    """The MAWB number is printed as "ONEYSZPGM" over "8034800"."""
    values = _values(standard[1])
    assert values["6.MAWB NO"] == "ONEYSZPGM8034800"
    assert values["8.HAWB NO"] == "GATI26SE0725"
    # A wrapped phrase keeps its space rather than joining up closed.
    assert values["2.CB NAME"] == "INTERLINK SHIPPING & CLEARING"


def test_a_wrapped_heading_is_not_read_as_values(standard):
    """"7.ADV BE 11.FIRST 12. PROV/" wraps onto "(Y/N/P) CHECK FINAL".

    Both lines are heading; the values are further down. Reading the wrapped
    line as values turned the flag "N" into "CHECKN". Those processing flags
    are not reported, but the fields around them must survive intact.
    """
    values = _values(standard[1])
    assert values["1.BE STATUS"] == "FIRST COPY"
    assert values["2.MODE"] == "Sea"
    # A heading whose value sits beside it still resolves.
    assert values["15.PORT OF LOADING"] == "Shekou"
    assert values["16.PORT OF SHIPMENT"] == "Shekou"
