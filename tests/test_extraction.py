"""End-to-end checks against the two sample Bills of Entry.

Each figure asserted here was read off the source PDF by hand, and the totals
are checked against the summary the document prints for itself.
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from boe_extraction.extract import extract

SAMPLES = Path(__file__).resolve().parents[1] / "samples"
COURIER = SAMPLES / "FBA15M6L9KGF01_courier_cbe_xiv.pdf"
STANDARD = SAMPLES / "BOE_3141398_icegate.pdf"


@pytest.fixture(scope="module")
def courier():
    return extract(COURIER)


@pytest.fixture(scope="module")
def standard():
    return extract(STANDARD)


def test_courier_is_read_without_a_model(courier):
    assert courier.form_type == "CBE-XIV"
    assert courier.be_number == "CBEXIV_DEL_2026-2027_2808_10570"
    assert courier.importer_name == "VALUECART PRIVATE LIMITED"
    assert courier.gstin == "29AAFCV5265N1ZK"
    assert courier.exchange_rate == 96.5


def test_courier_line_items(courier):
    (invoice,) = courier.invoices
    assert invoice.number == "FBA15M6L9KGF01"
    assert invoice.supplier == "GATI HONG KONG LIMITED"
    assert len(invoice.items) == 4

    first = invoice.items[0]
    assert first.hs_code == "91021100"
    assert first.description == "X0021B4KHX Wrist Watch - Mechanical Display"
    assert (first.quantity, first.unit_of_measure, first.unit_price) == (15, "PCS", 13.5)
    assert first.assessable_value == 19541.25
    assert (first.bcd_rate, first.bcd_amount) == (20, 3908)
    assert (first.sws_rate, first.sws_amount) == (10, 391)
    assert (first.igst_rate, first.igst_amount) == (18, 4291)
    assert first.duty_amount == 8590


def test_courier_assessable_value_is_price_times_quantity(courier):
    for item in courier.all_items():
        expected = item.unit_price * item.quantity * item.exchange_rate
        assert item.assessable_value == pytest.approx(expected, abs=0.01)


def test_standard_form_is_identified_from_page_one(standard):
    assert standard.form_type == "ICEGATE BOE"
    assert standard.be_number == "3141398"
    assert standard.port_code == "INNSA1"
    assert standard.gstin == "27AAFCV5265N1ZO/G"
    assert standard.exchange_rate == 96.05


def test_standard_splits_invoices(standard):
    assert [i.number for i in standard.invoices] == ["FBA15M13GSD3", "FBA15M16XHDH"]
    assert [len(i.items) for i in standard.invoices] == [5, 4]
    assert len(standard.all_items()) == 9


def test_standard_line_items(standard):
    first = standard.invoices[0].items[0]
    assert first.hs_code == "39269099"
    assert first.description == "X002HTKLZJ MACBOOK PRO 16 INCH CASE - PC"
    assert (first.quantity, first.unit_of_measure, first.unit_price) == (40, "NOS", 4.62)
    assert first.assessable_value == 17750.04
    assert (first.bcd_rate, first.bcd_amount) == (15, 2662.5)
    assert (first.sws_rate, first.sws_amount) == (10, 266.3)
    assert (first.igst_rate, first.igst_amount) == (18, 3722.2)
    assert first.duty_amount == 6651

    # The stamp overlay splits this row's SWS figure into "30", "1" and ".1".
    second = standard.invoices[0].items[1]
    assert second.sws_amount == 301.1


def test_standard_reconciles_with_its_own_summary(standard):
    """Part I prints the totals; the extracted rows must add up to them."""
    items = standard.all_items()
    assert sum(i.assessable_value for i in items) == pytest.approx(560008, abs=0.5)
    assert sum(i.bcd_amount for i in items) == pytest.approx(84001.2, abs=0.5)
    assert sum(i.sws_amount for i in items) == pytest.approx(8400.3, abs=0.5)
    assert sum(i.igst_amount for i in items) == pytest.approx(117434, abs=0.5)
    assert sum(i.duty_amount for i in items) == pytest.approx(209835, abs=0.5)


def test_invoice_assessable_values_match_part_two(standard):
    totals = [round(sum(i.assessable_value for i in invoice.items), 2)
              for invoice in standard.invoices]
    assert totals == [292981.32, 267026.68]


@pytest.mark.parametrize("sample", [COURIER, STANDARD])
def test_duty_amount_is_the_sum_of_the_duty_heads(sample):
    for item in extract(sample).all_items():
        assert item.duty_amount == pytest.approx(item.total_duty(), abs=0.5)


def test_workbook_matches_the_established_column_layout(tmp_path, courier):
    from openpyxl import load_workbook

    from boe_extraction.excel_writer import write_invoice

    path = write_invoice(courier, courier.invoices[0], tmp_path / "out.xlsx")
    sheet = load_workbook(path)["Sheet1"]
    assert [c.value for c in sheet[1]] == [
        "TRUE", "Item Number", "HS Code", "Description", "Quantity",
        "Unit of Measure", "Unit Price", "Assessable Value", "BCD Rate",
        "BCD Amount", "SWS Rate", "SWS Amount", "IGST Rate", "IGST Amount",
        "AIDC Rate", "AIDC Amount", "CMPNSTRY Rate", "CMPNSTRY Amount",
        "Duty Amount", "Exchange Rate"]
    assert sheet.max_row == 5
    assert sheet["A2"].value is True


CBE_XIII = SAMPLES / "FBA15M1ZPS2Y01_courier_cbe_xiii.pdf"


@pytest.fixture(scope="module")
def cbe_xiii():
    return extract(CBE_XIII)


def test_cbe_xiii_is_identified_and_read(cbe_xiii):
    """This form lists items flat under "ITEM :", with no invoice sections."""
    assert cbe_xiii.form_type == "CBE-XIII"
    assert cbe_xiii.importer_name == "VALUECART PRIVATE LIMITED"
    assert cbe_xiii.iec == "AAFCV5265N"
    assert cbe_xiii.gstin == "29AAFCV5265N1ZK"
    assert cbe_xiii.exchange_rate == 97.2
    assert len(cbe_xiii.all_items()) == 44


def test_cbe_xiii_groups_items_into_their_invoice(cbe_xiii):
    (invoice,) = cbe_xiii.invoices
    assert invoice.number == "FBA15M1ZPS2Y01"
    assert invoice.supplier == "GATI HONG KONG LIMITED"
    assert len(invoice.items) == 44


def test_cbe_xiii_line_items(cbe_xiii):
    first = cbe_xiii.all_items()[0]
    assert first.hs_code == "42023290"
    assert first.description == "X002447XTT Meta Ray-Ban Glasses Carrying Case"
    assert (first.quantity, first.unit_of_measure, first.unit_price) == (7, "PCS", 4.18)
    assert first.assessable_value == 2844.07
    assert (first.bcd_rate, first.bcd_amount) == (15, 427)
    assert (first.sws_rate, first.sws_amount) == (10, 43)
    assert (first.igst_rate, first.igst_amount) == (18, 596)
    assert first.duty_amount == 1066


def test_cbe_xiii_reconciles_with_its_own_header(cbe_xiii):
    """The form prints the consignment's assessable value and duty on page 1."""
    items = cbe_xiii.all_items()
    assert sum(i.assessable_value for i in items) == pytest.approx(82062.07, abs=0.01)
    assert sum(i.duty_amount for i in items) == pytest.approx(29648, abs=0.5)


def test_a_quantity_is_not_lost_to_a_colonless_label(cbe_xiii):
    """"Marks on Packages 0" used to swallow the "Quantity : 7" beside it."""
    assert all(item.quantity is not None for item in cbe_xiii.all_items())
