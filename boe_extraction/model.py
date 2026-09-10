"""Row and header shapes shared by every Bill of Entry parser."""

from dataclasses import dataclass, field, asdict
from typing import Optional

# The 20 line-item columns, in the order they are written to Excel.
# Every duty head the forms carry, in the order they are written to Excel.
# A form that has no such head leaves the pair blank.
DUTY_HEADS = [
    ("bcd", "BCD"),
    ("sws", "SWS"),
    ("igst", "IGST"),
    ("aidc", "AIDC"),
    ("add", "ADD"),
    ("chcess", "CHCESS"),
    ("cess", "CESS"),
    ("cmpnstry", "CMPNSTRY"),
]

ITEM_COLUMNS = [
    ("item_number", "Item Number"),
    ("hs_code", "HS Code"),
    ("description", "Description"),
    ("quantity", "Quantity"),
    ("unit_of_measure", "Unit of Measure"),
    ("unit_price", "Unit Price"),
    ("assessable_value", "Assessable Value"),
    ("notification_number", "Notification number"),
    ("notification_serial", "serial number of notification"),
    ("bcd_rate", "BCD Rate"),
    ("bcd_specific_rate", "BCD Specific rate"),
    ("bcd_amount", "BCD Amount"),
    ("sws_rate", "SWS Rate"),
    ("sws_amount", "SWS Amount"),
    ("igst_rate", "IGST Rate"),
    ("igst_amount", "IGST Amount"),
    ("aidc_rate", "AIDC Rate"),
    ("aidc_amount", "AIDC Amount"),
    ("add_rate", "ADD Rate"),
    ("add_amount", "ADD Amount"),
    ("chcess_rate", "CHCESS rate"),
    ("chcess_amount", "CHCESS Amount"),
    ("cess_rate", "CESS rate"),
    ("cess_amount", "CESS Amount"),
    ("cmpnstry_rate", "CMPNSTRY Rate"),
    ("cmpnstry_amount", "CMPNSTRY Amount"),
    ("duty_amount", "Duty Amount"),
    ("exchange_rate", "Exchange Rate"),
]


@dataclass
class LineItem:
    """One line of the bill.

    ``details`` holds every other field the form carries for this item, keyed
    by the form's own label, so a document can be asked for a field the 20
    standard columns do not cover.
    """

    item_number: int
    hs_code: str = ""
    description: str = ""
    quantity: Optional[float] = None
    unit_of_measure: str = ""
    unit_price: Optional[float] = None
    assessable_value: Optional[float] = None
    bcd_rate: Optional[float] = None
    bcd_amount: Optional[float] = None
    sws_rate: Optional[float] = None
    sws_amount: Optional[float] = None
    igst_rate: Optional[float] = None
    igst_amount: Optional[float] = None
    aidc_rate: Optional[float] = None
    aidc_amount: Optional[float] = None
    add_rate: Optional[float] = None
    add_amount: Optional[float] = None
    chcess_rate: Optional[float] = None
    chcess_amount: Optional[float] = None
    cess_rate: Optional[float] = None
    cess_amount: Optional[float] = None
    cmpnstry_rate: Optional[float] = None
    cmpnstry_amount: Optional[float] = None
    bcd_specific_rate: Optional[float] = None
    notification_number: str = ""
    notification_serial: str = ""
    duty_amount: Optional[float] = None
    exchange_rate: Optional[float] = None
    details: dict = field(default_factory=dict)

    def duty_amounts(self):
        return [getattr(self, f"{head}_amount") for head, _ in DUTY_HEADS]

    def backfill_bcd(self):
        """Recover a BCD amount the table reader dropped.

        A blank BCD cell in the source must not become a zero in the output, so
        derive it from the other duty heads and the item's total duty.
        """
        if self.bcd_amount is not None or self.duty_amount is None:
            return
        others = [getattr(self, f"{head}_amount") for head, _ in DUTY_HEADS
                  if head != "bcd"]
        if any(o is None for o in others):
            return
        self.bcd_amount = round(self.duty_amount - sum(others), 2)

    def total_duty(self):
        return round(sum(a for a in self.duty_amounts() if a is not None), 2)


@dataclass
class Invoice:
    number: str = ""
    date: str = ""
    supplier: str = ""
    currency: str = ""
    invoice_value: Optional[float] = None
    exchange_rate: Optional[float] = None
    items: list = field(default_factory=list)


@dataclass
class BillOfEntry:
    form_type: str = ""
    source_file: str = ""
    be_number: str = ""
    be_date: str = ""
    be_type: str = ""
    port_code: str = ""
    importer_name: str = ""
    iec: str = ""
    gstin: str = ""
    ad_code: str = ""
    country_of_origin: str = ""
    country_of_consignment: str = ""
    exchange_rate: Optional[float] = None
    currency: str = ""
    invoices: list = field(default_factory=list)

    def all_items(self):
        return [item for inv in self.invoices for item in inv.items]

    def header_fields(self):
        skip = {"invoices"}
        return [(k, v) for k, v in asdict(self).items() if k not in skip]
