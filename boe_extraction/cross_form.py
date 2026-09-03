"""Applies the field list highlighted on one bill of entry to another form.

The highlights on an ICEGATE bill are a required-field list, but a courier
CBE-XIII names almost nothing the same way: the ICEGATE "18.TOT.ASS VAL" is a
CBE-XIII "Assessable Value", "2.CTH" is "CTSH", and a good many ICEGATE fields
(the sea-freight flags, the CVD and SAD duty heads) have no counterpart at all.

So the spec is applied through an explicit map. A field with no counterpart is
reported as absent rather than silently left blank, because a blank cell reads
as "the document says nothing here" when the truth is "this form has no such
field".
"""

from .cbe_grid import Cell, Marker, read_entries

ABSENT = "— not on this form —"

# Spec fields that name a per-item column rather than a document-level one.
ITEM_LEVEL = {
    "1.S NO.": lambda item: item.item_number,
    "1.INVSNO": lambda item: item.item_number,
    "2.CTH": lambda item: item.hs_code,
    "3.DESCRIPTION": lambda item: item.description,
    "4.UNIT PRICE": lambda item: item.unit_price,
    "5.QUANTITY": lambda item: item.quantity,
    "6.UQC": lambda item: item.unit_of_measure,
    "7.AMOUNT": lambda item: item.details.get("Invoice Value", ""),
    "14.ASS. VALUE": lambda item: item.assessable_value,
}


def _totals(items, field):
    values = [getattr(item, field) for item in items]
    return round(sum(v for v in values if v is not None), 2) if values else None


def _first_invoice(boe):
    return boe.invoices[0] if boe.invoices else None


# Each entry is (what the CBE-XIII calls it, how to read it). A None reader
# means the form carries no such field.
DOCUMENT_LEVEL = {
    "BE No": ("CBE-XIII Number", lambda b, c: b.be_number),
    "BE Date": ("Current Status of the CBE", lambda b, c: c.get("Current Status of the CBE", "")),
    "IEC/Br": ("Import Export Code / IEC Branch Code",
               lambda b, c: "/".join(x for x in (b.iec, c.get("IEC Branch Code", "")) if x)),
    "GSTIN/TYPE": ("KYC ID / KYC Document",
                   lambda b, c: "/".join(x for x in (b.gstin, c.get("KYC Document", "")) if x)),
    "CB CODE": ("Courier Registration Number",
                lambda b, c: c.get("Courier Registration Number", "")),
    "2.CB NAME": ("Name of the Authorized Courier",
                  lambda b, c: c.get("Name of the Authorized Courier", "")),
    "PKG": ("No of Packages", lambda b, c: c.get("No of Packages", "")),
    "10.PKG": ("No of Packages", lambda b, c: c.get("No of Packages", "")),
    "G.WT (KGS)": ("Gross Weight", lambda b, c: c.get("Gross Weight", "")),
    "11.GW": ("Gross Weight", lambda b, c: c.get("Gross Weight", "")),
    "13.COUNTRY OF ORIGIN": ("Country of Origin", lambda b, c: b.country_of_origin),
    "14.COUNTRY OF CONSIGNMENT": ("Country of Exportation",
                                  lambda b, c: b.country_of_consignment),
    "15.PORT OF LOADING": ("Airport of Shipment", lambda b, c: c.get("Airport of Shipment", "")),
    "16.PORT OF SHIPMENT": ("Airport of Shipment", lambda b, c: c.get("Airport of Shipment", "")),
    "1.IMPORTER NAME & ADDRESS": ("Name of Consignee", lambda b, c: b.importer_name),
    "1.BUYER'S NAME & ADDRESS": ("Name of Consignee", lambda b, c: b.importer_name),
    "3.SUPPLIER NAME & ADDRESS": ("Name of Consignor",
                                  lambda b, c: c.get("Name of Consignor", "")),
    "4.THIRD PARTY NAME & ADDRESS": ("Name of Consignor",
                                     lambda b, c: c.get("Name of Consignor", "")),
    "AD CODE": ("AD Code", lambda b, c: b.ad_code),
    "8.HAWB NO": ("HAWB Number", lambda b, c: c.get("HAWB Number", "")),
    "EXCHANGE RATE": ("Rate of Exchange",
                      lambda b, c: (f"1 {b.currency}={b.exchange_rate}INR"
                                    if b.exchange_rate else "")),
    "1.S.NO": ("Invoice serial", lambda b, c: 1 if b.invoices else ""),
    "2.INVOICE NO": ("Invoice Number",
                     lambda b, c: getattr(_first_invoice(b), "number", "")),
    "2.INVOICE NO. & DT.": ("Invoice Number",
                            lambda b, c: getattr(_first_invoice(b), "number", "")),
    "3.INV. AMT": ("Invoice Value (sum of items)",
                   lambda b, c: getattr(_first_invoice(b), "invoice_value", "")),
    "1.INV VALUE": ("Invoice Value (sum of items)",
                    lambda b, c: getattr(_first_invoice(b), "invoice_value", "")),
    "14.Cur": ("Currency of Invoice", lambda b, c: b.currency),
    "15.Term": ("Invoice Term", lambda b, c: c.get("Invoice Term", "")),
    "2.FREIGHT": ("Freight", lambda b, c: c.get("Freight", "")),
    "15.INT": ("Interest Amount", lambda b, c: c.get("Interest Amount", "")),
    # The CBE-XIII prints no duty summary, so the heads are totalled from the
    # items -- which reconcile exactly with the Assessable Value and Duty(Rs.)
    # the form does print.
    "1.BCD": ("Sum of item BCD", lambda b, c: _totals(b.all_items(), "bcd_amount")),
    "3.SWS": ("Sum of item SWS", lambda b, c: _totals(b.all_items(), "sws_amount")),
    "7.IGST": ("Sum of item IGST", lambda b, c: _totals(b.all_items(), "igst_amount")),
    "8.G.CESS": ("Sum of item compensation cess",
                 lambda b, c: _totals(b.all_items(), "cess_amount")),
    "18.TOT.ASS VAL": ("Assessable Value", lambda b, c: c.get("Assessable Value", "")),
    "14.TOTAL DUTY": ("Duty(Rs.)", lambda b, c: c.get("Duty(Rs.)", "")),
    "19.TOT. AMOUNT": ("Duty(Rs.)", lambda b, c: c.get("Duty(Rs.)", "")),
    # Present on the ICEGATE form only.
    "Port Code": (None, None),
    "1.BE STATUS": (None, None),
    "2.MODE": (None, None),
    "3.DEF BE": (None, None),
    "4.KACHA": (None, None),
    "5.SEC 48": (None, None),
    "6.REIMP": (None, None),
    "8.ASSESS": (None, None),
    "9.EXAM": (None, None),
    "10.HSS": (None, None),
    "3.AEO": (None, None),
    "5.AEO": (None, None),
    "2.SELLER'S NAME & ADDRESS": (None, None),
    "2.ACD": (None, None),
    "4.NCCD": (None, None),
    "5.ADD": (None, None),
    "6.CVD": (None, None),
    "9.SG": (None, None),
    "10.SAED": (None, None),
    "11.GSIA": (None, None),
    "12.TTA": (None, None),
    "13.HEALTH": (None, None),
    "16.PNLTY": (None, None),
    "17.FINE": (None, None),
    "6.MAWB NO": (None, None),
    "7.DATE": (None, None),
    "9.DATE": (None, None),
    "23.PRODN24.CNTRL": (None, None),
}


def document_cells(pdf):
    """The label/value cells above the first item block, as a dict."""
    values = {}
    for entry in read_entries(pdf):
        if isinstance(entry, Marker) and entry.text.upper().startswith("ITEM"):
            break
        if isinstance(entry, Cell) and entry.label and entry.label not in values:
            values[entry.label] = entry.value
    return values


def spec_labels(fields):
    """The field names a highlight spec asks for, in order, without banners."""
    labels, seen = [], set()
    for field in fields:
        label = field.label
        if not label or label.startswith("("):
            continue
        if label not in seen:
            seen.add(label)
            labels.append(label)
    return labels


def apply_spec(labels, boe, cells):
    """Resolve each spec field against the target document.

    Returns (document rows, item column labels). A document row is
    (spec field, what this form calls it, value).
    """
    rows, item_labels = [], []
    for label in labels:
        if label in ITEM_LEVEL:
            item_labels.append(label)
            continue
        known, reader = DOCUMENT_LEVEL.get(label, (None, None))
        if reader is None:
            rows.append((label, ABSENT, ""))
        else:
            rows.append((label, known, reader(boe, cells)))
    return rows, item_labels
