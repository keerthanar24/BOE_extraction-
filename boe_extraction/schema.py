"""The fields the extract carries, and nothing else.

The ECCS courier forms are fixed government layouts, so what is wanted from
them is a named list rather than whatever a given document happens to print.
Naming the fields here is what keeps the extract free of fields nobody asked
for, and free of the same value appearing under two headings.

Each document field is a (section, label) pair, because the forms reuse a
label -- Name, Address -- under several sections.
"""

# Dropped as duplicates, and where the value is kept instead:
#   BOE Number          -> ORIGINAL COPY - CBEXIV Number (same number twice)
#   CTSH, CETSH         -> HS Code
#   Description of Goods-> Description
#   Duty(Rs.)           -> Duty Amount
#   Rate of Exchange    -> Exchange Rate
#   Invoice Number      -> Invoice
#   Currency of Invoice -> Currency of Unit Price
# Part II repeats per invoice, so the fields it holds -- 1.INV VALUE and
# 15.Term among them -- report the first invoice's. The Line Items sheet
# names the invoice each item belongs to.
# Dropped as not data: Charge Type is the constant heading "DUTY DETAILS", and
# Charge Amount(in rs.) is the empty cell beside it.

CBE_XIV_FIELDS = [
    ('ORIGINAL COPY', 'Current Status of the CBE'),
    ('ORIGINAL COPY', 'CBEXIV Number'),
    ('DETAILS OF AUTHORIZED COURIER', 'Courier Registration Number'),
    ('DETAILS OF AUTHORIZED COURIER', 'Name of the Authorized Courier'),
    ('DETAILS OF AUTHORIZED COURIER', 'Address of Authorized Courier'),
    ('PARTICULARS OF THE IMPORTER', 'Import Export Branch Code'),
    ('PARTICULARS OF THE IMPORTER', 'Import export Code'),
    ('PARTICULARS OF THE IMPORTER', 'Address'),
    ('PARTICULARS OF THE IMPORTER', 'Name'),
    ('PARTICULARS OF THE IMPORTER', 'Category Of Importer'),
    ('PARTICULARS OF THE IMPORTER', 'Type Of Importer'),
    ('PARTICULARS OF THE IMPORTER', 'Authorised Dealer Code Of Bank'),
    ('PARTICULARS OF THE IMPORTER', 'Class Code'),
    ('PARTICULARS OF THE IMPORTER', 'BOE Date'),
    ('PARTICULARS OF THE IMPORTER', 'Category Of BOE'),
    ('PARTICULARS OF THE IMPORTER', 'Type Of BOE'),
    ('PARTICULARS OF THE IMPORTER', 'Whether Import Using eCommerce'),
    ('PARTICULARS OF THE IMPORTER', 'KYC Document'),
    ('PARTICULARS OF THE IMPORTER', 'KYC ID'),
    ('PARTICULARS OF THE IMPORTER', 'State Code'),
    ('SPECIAL REQUESTS', 'Country of Consignment'),
    ('SPECIAL REQUESTS', 'Country of Origin'),
    ('IGM DETAILS', 'Airlines'),
    ('IGM DETAILS', 'Airport Of Arrival'),
    ('IGM DETAILS', 'Date Of Arrival'),
    ('IGM DETAILS', 'Flight No.'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Date of Entry Inward'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Import General Manifest (IGM) Number'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Date Of MAWB'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Master Airway Bill (MAWB) Number'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Date of HAWB'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'House Airway Bill (HAWB) Number'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Marks and Numbers'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Number of Packages'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Interest Amount'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Type of Packages'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Gross Weight'),
    ('IMPORT GENERAL MANIFEST DETAILS', 'Unit of Measure for Gross Weight'),
    ('Details Of Invoice - 1', 'Date of Invoice'),
    ('Details Of Invoice - 1', 'Invoice Number'),
    ('Details Of Invoice - 1', 'Date of Purchase Order'),
    ('Details Of Invoice - 1', 'Purchase Order Number'),
    ('SUPPLIER DETAILS', 'Address'),
    ('SUPPLIER DETAILS', 'Name'),
    ('IF SUPPLIER IS NOT THE SELLER', 'Address'),
    ('IF SUPPLIER IS NOT THE SELLER', 'Name'),
    ('BROKER/ AGENT DETAILS', 'Address'),
    ('BROKER/ AGENT DETAILS', 'Name'),
    ('BROKER/ AGENT DETAILS', 'Method of Valuation'),
    ('BROKER/ AGENT DETAILS', 'Terms of Invoice'),
    ('BROKER/ AGENT DETAILS', 'Currency'),
    ('BROKER/ AGENT DETAILS', 'Invoice Value'),
    ('PAYMENT DETAILS', 'Challan Date'),
    ('PAYMENT DETAILS', 'TR-6 Challan Number'),
    ('PAYMENT DETAILS', 'Total Amount'),
]

CBE_XIII_FIELDS = [
    ('ORIGINAL COPY', 'Current Status of the CBE'),
    # The form prints this label without its colon, so it reaches the schema
    # only because the schema names it. It is the form's own BE number, the
    # counterpart of CBE-XIV's CBEXIV Number.
    ('ORIGINAL COPY', 'CBE-XIII Number'),
    ('ORIGINAL COPY', 'Courier Registration Number'),
    ('ORIGINAL COPY', 'Name of the Authorized Courier'),
    ('ORIGINAL COPY', 'Address of Authorized Courier'),
    ('IGM DETAILS', 'Airlines'),
    ('IGM DETAILS', 'Flight No.'),
    ('IGM DETAILS', 'Airport Of Arrival'),
    ('IGM DETAILS', 'First Port Of Arrival'),
    ('IGM DETAILS', 'Date Of Arrival'),
    ('IGM DETAILS', 'Time Of Arrival'),
    ('IGM DETAILS', 'Airport of Shipment'),
    ('IGM DETAILS', 'Country of Exportation'),
    ('IGM DETAILS', 'HAWB Number'),
    ('IGM DETAILS', 'Name of Consignor'),
    ('IGM DETAILS', 'Address of Consignor'),
    ('IGM DETAILS', 'Name of Consignee'),
    ('IGM DETAILS', 'Address of Consignee'),
    ('IGM DETAILS', 'Import Export Code'),
    ('IGM DETAILS', 'IEC Branch Code'),
    ('IGM DETAILS', 'Special Request'),
    ('IGM DETAILS', 'No of Packages'),
    ('IGM DETAILS', 'Gross Weight'),
    ('IGM DETAILS', 'Net Weight'),
    ('IGM DETAILS', 'Assessable Value'),
    ('IGM DETAILS', 'Duty(Rs.)'),
    ('IGM DETAILS', 'Invoice Value'),
    ('IGM DETAILS', 'Case of CRN'),
    ('IGM DETAILS', 'KYC Document'),
    ('IGM DETAILS', 'KYC ID'),
    ('IGM DETAILS', 'State Code'),
    ('IGM DETAILS', 'Interest Amount'),
    ('IGM DETAILS', 'Government / NonGovernment'),
    ('IGM DETAILS', 'AD Code'),
    ('IGM DETAILS', 'Import Using e-Commerce'),
]

ICEGATE_FIELDS = [
    ('BILL OF ENTRY FOR HOME CONSUMPTION', 'BE Date'),
    ('BILL OF ENTRY FOR HOME CONSUMPTION', 'BE No'),
    ('BILL OF ENTRY FOR HOME CONSUMPTION', 'Port Code'),
    ('BILL OF ENTRY FOR HOME CONSUMPTION', 'Br'),
    ('BILL OF ENTRY FOR HOME CONSUMPTION', 'IEC'),
    ('BILL OF ENTRY FOR HOME CONSUMPTION', 'GSTIN'),
    ('BILL OF ENTRY FOR HOME CONSUMPTION', 'TYPE'),
    ('BILL OF ENTRY FOR HOME CONSUMPTION', 'CB CODE'),
    ('BILL OF ENTRY FOR HOME CONSUMPTION', 'G.WT (KGS)'),
    ('BILL OF ENTRY FOR HOME CONSUMPTION', 'PKG'),
    ('PART - I - BILL OF ENTRY SUMMARY', '1.BE STATUS'),
    ('PART - I - BILL OF ENTRY SUMMARY', '2.MODE'),
    ('PART - I - BILL OF ENTRY SUMMARY', '13.COUNTRY OF ORIGIN'),
    ('PART - I - BILL OF ENTRY SUMMARY', '14.COUNTRY OF CONSIGNMENT'),
    ('PART - I - BILL OF ENTRY SUMMARY', '15.PORT OF LOADING'),
    ('PART - I - BILL OF ENTRY SUMMARY', '16.PORT OF SHIPMENT'),
    ('PART - I - BILL OF ENTRY SUMMARY', '1.IMPORTER NAME & ADDRESS'),
    ('PART - I - BILL OF ENTRY SUMMARY', '2.CB NAME'),
    ('PART - I - BILL OF ENTRY SUMMARY', 'AD CODE'),
    ('PART - I - BILL OF ENTRY SUMMARY', '1.BCD'),
    ('PART - I - BILL OF ENTRY SUMMARY', '18.TOT.ASS VAL'),
    ('PART - I - BILL OF ENTRY SUMMARY', '2.ACD'),
    ('PART - I - BILL OF ENTRY SUMMARY', '3.SWS'),
    ('PART - I - BILL OF ENTRY SUMMARY', '4.NCCD'),
    ('PART - I - BILL OF ENTRY SUMMARY', '5.ADD'),
    ('PART - I - BILL OF ENTRY SUMMARY', '6.CVD'),
    ('PART - I - BILL OF ENTRY SUMMARY', '7.IGST'),
    ('PART - I - BILL OF ENTRY SUMMARY', '8.G.CESS'),
    ('PART - I - BILL OF ENTRY SUMMARY', '13.HEALTH'),
    ('PART - I - BILL OF ENTRY SUMMARY', '14.TOTAL DUTY'),
    ('PART - I - BILL OF ENTRY SUMMARY', '15.INT'),
    ('PART - I - BILL OF ENTRY SUMMARY', '16.PNLTY'),
    ('PART - I - BILL OF ENTRY SUMMARY', '17.FINE'),
    ('PART - I - BILL OF ENTRY SUMMARY', '19.TOT. AMOUNT'),
    ('PART - I - BILL OF ENTRY SUMMARY', '9.SG'),
    ('PART - I - BILL OF ENTRY SUMMARY', '10.PKG'),
    ('PART - I - BILL OF ENTRY SUMMARY', '11.GW'),
    ('PART - I - BILL OF ENTRY SUMMARY', '6.MAWB NO'),
    ('PART - I - BILL OF ENTRY SUMMARY', '7.DATE'),
    ('PART - I - BILL OF ENTRY SUMMARY', '8.HAWB NO'),
    ('PART - I - BILL OF ENTRY SUMMARY', '9.DATE'),
    ('PART - I - BILL OF ENTRY SUMMARY', '3.WBE SITE'),
    ('PART - I - BILL OF ENTRY SUMMARY', '4.WH CODE'),
    ('PART - I - BILL OF ENTRY SUMMARY', 'EXCHANGE RATE'),
    # Empty until the bill is given out of charge, so a first copy reports
    # both blank; the form names them, and they are read when filled.
    ('PART - I - BILL OF ENTRY SUMMARY', 'OOC NO.'),
    ('PART - I - BILL OF ENTRY SUMMARY', 'OOC DATE'),
    ('PART - II - INVOICE & VALUATION DETAILS', '1.INV VALUE'),
    ('PART - II - INVOICE & VALUATION DETAILS', '15.Term'),
    ('PART - II - INVOICE & VALUATION DETAILS', '4.LC NO & DATE'),
    ('PART - II - INVOICE & VALUATION DETAILS', '5.CONTRACT NO & DATE'),
    ('PART - II - INVOICE & VALUATION DETAILS', "1.BUYER'S NAME & ADDRESS"),
    ('PART - II - INVOICE & VALUATION DETAILS', "2.SELLER'S NAME & ADDRESS"),
    ('PART - II - INVOICE & VALUATION DETAILS', '3.SUPPLIER NAME & ADDRESS'),
    ('PART - II - INVOICE & VALUATION DETAILS', '4.THIRD PARTY NAME & ADDRESS'),
    ('PART - II - INVOICE & VALUATION DETAILS', 'AD CODE'),
]

DOCUMENT_FIELDS = {"ICEGATE BOE": ICEGATE_FIELDS,
                   "CBE-XIV": CBE_XIV_FIELDS,
                   "CBE-XIII": CBE_XIII_FIELDS}

# The two families the extract is delivered in: the ICEGATE cargo bill, and
# the two ECCS courier bills, which are one family in everything but which
# regulation they are filed under.
FAMILIES = ("Cargo", "Courier")
FAMILY = {"ICEGATE BOE": "Cargo", "CBE-XIV": "Courier", "CBE-XIII": "Courier"}

# One row per line item. The columns the model exposes as attributes come
# first; the rest are read from the item's own label/value cells.
ITEM_FIELDS = [
    'Invoice',
    'Item Number',
    'HS Code',
    'Description',
    'Quantity',
    'Unit of Measure',
    'Unit Price',
    'Assessable Value',
    'Notification number',
    'serial number of notification',
    'BCD Rate',
    'BCD Specific rate',
    'BCD Amount',
    'SWS Rate',
    'SWS Amount',
    'IGST Rate',
    'IGST Amount',
    'AIDC Rate',
    'AIDC Amount',
    'ADD Rate',
    'ADD Amount',
    'CHCESS rate',
    'CHCESS Amount',
    'CESS rate',
    'CESS Amount',
    'CMPNSTRY Rate',
    'CMPNSTRY Amount',
    'Duty Amount',
    'Exchange Rate',
    'License Type',
    'License Number',
    'Country of Origin',
    'Name of Manufacturer',
    'Address of Manufacturer',
    'Number of Packages',
    'Marks on Packages',
    'Invoice Value',
    'Currency of Unit Price',
    'Invoice Term',
    'Landing Charges',
    'Insurance',
    'Freight',
    'Discount Amount',
    'Currency of Discount',
]


def document_fields_for(form_type):
    """The schema's document fields for a form, or None if it has no schema."""
    return DOCUMENT_FIELDS.get(form_type)


def document_rows(form_type, cells):
    """Fill the schema's document fields from the cells a document carries.

    The schema is the allow-list: only a named field is read, so a table
    fragment or a declaration paragraph can never reach the extract. A field
    the document leaves blank still gets its row, because an empty cell is an
    answer -- the form asked and the filer left it empty.

    Where a bill carries several invoices the per-invoice sections repeat, and
    what is reported here is the first invoice's. The Line Items sheet names
    the invoice each item belongs to.
    """
    wanted = DOCUMENT_FIELDS.get(form_type)
    if wanted is None:
        return None
    found = {}
    for section, label, value in cells:
        key = (section, label)
        # A section repeats where a bill carries several invoices, so the
        # first answer is kept: the fields then all describe invoice 1,
        # rather than the number coming from the first and the value from
        # the last. A blank first answer still yields to a filled later one.
        if found.get(key):
            continue
        found[key] = value
    return [(section, label, found.get((section, label), ""))
            for section, label in wanted]


# The two courier forms word some item labels differently. A column is one
# field, so the other form's wording is read into the same column rather than
# added as a column of its own.
ITEM_LABEL_ALIASES = {
    "Currency of Unit Price": ("Currency for Unit Price",),
}


def item_rows(boe, columns_by_title):
    """One row per line item, in the schema's column order.

    A column is either something the model parsed onto the item, or one of the
    item's own label/value cells read straight off the form.
    """
    rows = []
    for invoice in boe.invoices:
        for item in invoice.items:
            row = []
            for title in ITEM_FIELDS:
                if title == "Invoice":
                    row.append(invoice.number)
                elif title in columns_by_title:
                    row.append(getattr(item, columns_by_title[title]))
                else:
                    row.append(_detail(item, title))
            rows.append(row)
    return rows


def _detail(item, title):
    """A column's value from the item's own cells, under either form's wording."""
    for label in (title,) + ITEM_LABEL_ALIASES.get(title, ()):
        value = item.details.get(label)
        if value:
            return value
    return None
