# The extract schema

`boe_extraction/schema.py` names every field the extract carries. It is an
allow-list, and that is what enforces the three rules this extract is held to:

- **No extra fields.** Only a field the schema names is written. A table
  fragment, a heading, or a declaration paragraph can never reach the output,
  because the schema does not name it.
- **No duplicated fields.** A name appears once. The columns that repeated
  another column's value on every single item were dropped.
- **No duplicated information.** Where two headings carried the same value,
  one is kept.

## What is written

One workbook per document, two sheets, nothing else:

| Sheet | Shape |
|---|---|
| `Document Fields` | Section, Field, Value — one row per schema field |
| `Line Items` | 44 columns — one row per line item |

A field the form left blank still gets its row: the form asked, and an empty
answer is an answer. Blank is written as blank, never as `0` — a duty head the
form does not carry reads empty, not zero.

## Field counts

| Form | Document fields | Line item columns |
|---|---|---|
| CBE-XIV | 55 | 44 |
| CBE-XIII | 35 | 44 |

The two forms name their header fields differently, so their document lists
differ. The line item columns are one list across both.

## What was dropped, and what holds the value instead

Item columns, each of which repeated another column on all 44 items of the
CBE-XIII sample:

| Dropped | Kept |
|---|---|
| `CTSH`, `CETSH` | `HS Code` |
| `Description of Goods` | `Description` |
| `Duty(Rs.)` | `Duty Amount` |
| `Rate of Exchange` | `Exchange Rate` |
| `Invoice Number` | `Invoice` |
| `Currency of Invoice` | `Currency of Unit Price` |

Document fields:

| Dropped | Kept |
|---|---|
| `PARTICULARS OF THE IMPORTER · BOE Number` | `ORIGINAL COPY · CBEXIV Number` |

Dropped as not being data at all: `Charge Type`, which held the constant
heading `DUTY DETAILS` on every item, and `Charge Amount(in rs.)`, the empty
cell printed beside it.

## Added

`ORIGINAL COPY · CBE-XIII Number` — the CBE-XIII form's own bill of entry
number, the counterpart of CBE-XIV's `CBEXIV Number`. The form prints this
label without its colon, so it was being dropped before it reached the
extract.

## Completeness

Two fields were being cut short, and both are fixed:

- `Address of Authorized Courier` lost its first line. The form right-aligns a
  label to its colon column, so a long label wraps with the colon left on the
  last line — `Address of Authorized` over `Courier :`. The second line was
  read as a new field called `Courier`, which started the address at its own
  second line. Now the two columns are treated as one grid row, so a label
  fragment only starts a new row when every column taking a label on that line
  is starting one.
- `Import Using e-Commerce` is empty on the form, and was picking up the
  centred heading that follows it, `DETAILS OF CRN (if present)`. The heading
  test rejected it for the lower-case words in its brackets.

## Checks the extract is held to

Run `python -m pytest tests -q`.

- Every field of both specs is reproduced exactly.
- Duty heads add up to `Duty Amount` on every item of both forms, within 1.0.
- Every column the form answers is non-empty for at least one item.
- No field name appears twice, in either list.
