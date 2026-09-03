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
    assert values["CTSH"] == "91021100"
    assert values["Assessable Value"] == "19541.25"
    assert values["Name of Manufacturer"] == "GATI HONG KONG LIMITED"


def test_multi_row_tables_are_kept_verbatim(courier):
    """DUTY DETAILS has five rows, so it has no one value per column.

    Its figures are already reported against each line item, so the highlight
    is left as it was drawn.
    """
    tables = [f.value for f in courier[1] if f.label == "(as highlighted)"]
    duty = next(t for t in tables if t.startswith("Sr.No. Duty Head"))
    assert "1 BCD 20 0 0 3908" in duty
    assert "5 CMPNSTRY 0 0 0 0" in duty


def test_single_row_tables_become_fields(courier):
    """A bold heading row over one row of figures is a record, so read it."""
    values = _values(courier[1])
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
    assert values["GSTIN/TYPE"] == "27AAFCV5265N1ZO/G"
    assert values["CB CODE"] == "AAHFS9149KCH001"
    assert values["AD CODE"] == "6480001"
    assert values["13.COUNTRY OF ORIGIN"] == "CHINA"
    assert values["15.PORT OF LOADING"] == "Shekou"
    assert values["2.CB NAME"] == "INTERLINK SHIPPING & CLEARING"
    assert values["14.ASS. VALUE"] == "292981.32"


def test_standard_name_blocks_come_from_the_parser(standard):
    """These labels are indented over their value, so the column clips it."""
    values = _values(standard[1])
    assert values["1.IMPORTER NAME & ADDRESS"] == "VALUECART PRIVATE LIMITED"
    assert values["3.SUPPLIER NAME & ADDRESS"] == "GATI HONG KONG LIMITED"


def test_item_fields_are_reported_for_every_item(tmp_path, courier):
    """A field highlighted on item 1 is wanted for all of them."""
    from openpyxl import load_workbook

    from boe_extraction.excel_writer import write_highlighted

    boe, fields, highlights = courier
    path = write_highlighted(boe, fields, highlights, tmp_path / "h.xlsx")
    sheet = load_workbook(path)["Item Details"]
    header = [c.value for c in sheet[1]]
    assert "Name of Manufacturer" in header
    assert "Assessable Value" in header
    assert sheet.max_row == 5  # header plus four items

    column = header.index("Assessable Value") + 1
    assert [sheet.cell(row=r, column=column).value for r in range(2, 6)] == [
        "19541.25", "26055", "6513.75", "52110"]


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
    assert workbook.sheetnames == ["Documents", "Invoices", "Line Items",
                                   "Highlighted Fields", "Item Details",
                                   "Highlights (raw)"]

    # Every sheet names the document each row came from.
    for name in workbook.sheetnames:
        assert workbook[name]["A1"].value == "Document"

    assert workbook["Documents"].max_row == 3          # header plus two documents
    assert workbook["Invoices"].max_row == 4           # one courier, two ICEGATE
    assert workbook["Line Items"].max_row == 14        # 4 + 9 items
    assert workbook["Highlights (raw)"].max_row == 55 + 84 + 1


def test_combined_totals_match_the_documents(tmp_path):
    from openpyxl import load_workbook

    from boe_extraction.excel_writer import write_combined
    from boe_extraction.extract import extract_document

    documents = [extract_document(COURIER), extract_document(STANDARD)]
    path = write_combined(documents, tmp_path / "combined.xlsx")
    sheet = load_workbook(path)["Documents"]
    rows = {r[0]: r for r in sheet.iter_rows(min_row=2, values_only=True)}

    courier = rows[COURIER.name]
    assert courier[1] == "CBE-XIV"
    assert (courier[11], courier[12], courier[13]) == (4, 104220, 45814)

    standard = rows[STANDARD.name]
    assert standard[1] == "ICEGATE BOE"
    assert (standard[11], standard[12]) == (9, 560008)


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


def test_item_table_values_come_from_the_parsed_row(standard):
    """"3.DESCRIPTION" begins under "2.CTH", so reading it by column clips it."""
    values = _values(standard[1])
    assert values["2.CTH"] == "39269099"
    assert values["3.DESCRIPTION"] == "X002HTKLZJ MACBOOK PRO 16 INCH CASE - PC"
    assert values["4.UNIT PRICE"] == "4.62"
    assert values["5.QUANTITY"] == "40"
    assert values["7.AMOUNT"] == "184.8"


def test_a_blank_field_does_not_borrow_an_item_value(courier):
    """The courier form carries a Country of Origin at both levels."""
    origins = {f.key: f.value for f in courier[1] if f.label == "Country of Origin"}
    assert origins["SPECIAL REQUESTS · Country of Origin"] == ""
    assert origins["DETAILED DESCRIPTION OF ITEM · Country of Origin"] == "CHINA"
