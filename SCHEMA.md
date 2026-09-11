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

    python -m boe_extraction.cli <pdf> [<pdf> ...] --schema --combined out.xlsx

writes every document given into **one workbook**, each on its own two sheets:

| Sheet | Shape |
|---|---|
| `<form> Fields` | Section, Field, Value — one row per schema field |
| `<form> Line Items` | 44 columns — one row per line item |

Combined into one file, never into one sheet: a sheet holds a single bill of
entry, and its name says which — `Cargo BOE`, `Courier CBE-XIV`,
`Courier CBE-XIII`. Drop `--combined` to get one workbook per document
instead, where the sheets are named `Document Fields` and `Line Items`.

A field the form left blank still gets its row: the form asked, and an empty
answer is an answer. Blank is written as blank, never as `0` — a duty head the
form does not carry reads empty, not zero.

## Field counts

| Form | Sheet prefix | Document fields | Line item columns |
|---|---|---|---|
| ICEGATE BOE | `Cargo BOE` | 51 | 44 |
| CBE-XIV | `Courier CBE-XIV` | 55 | 44 |
| CBE-XIII | `Courier CBE-XIII` | 35 | 44 |

Each form names its header fields differently, so each has its own document
list. The line item columns are one list across all three.

The cargo list is the corrected one: the addresses read in full, and the seven
Part II fields that repeat per invoice are not in it. It is read straight from
the form, so the schema path does not need a reviewer to have highlighted the
PDF first -- a test asserts the two paths return the same 51 fields.

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

## Why a line item column can be blank

The column list is one list across all three forms, and the forms do not print
the same things per item. A blank column is never a field that was looked for
and missed -- each one is accounted for:

| Column | CBE-XIV | CBE-XIII | Cargo BOE |
|---|---|---|---|
| `ADD`, `CHCESS`, `CESS` rate and amount | not on the form | not on the form | `ADD Amount` filled |
| `License Type`, `License Number` | not filled | printed, left blank | not on the form |
| `Address of Manufacturer` | not on the form | printed, left blank | not on the form |
| `Discount Amount`, `Currency of Discount` | printed, left blank | not filled | not on the form |
| `Number of Packages`, `Marks on Packages` | **document level** | per item | not per item |
| `Invoice Value`, `Invoice Term` | **document level** | per item | document level |
| `Landing Charges`, `Insurance`, `Freight` | charges table, left blank | per item | not per item |

Where a form prints something once per invoice rather than once per item, it
is extracted -- on that document's Fields sheet -- and is not repeated down
every item row, because that would be the same value under two headings.
CBE-XIV's `Number of Packages` is `1`, its `Invoice Value` `1080`, its
`Terms of Invoice` `CIF`; all three are on `Courier CBE-XIV Fields`.

A test asserts this list: if a column goes blank for a reason not named above,
it fails.

## One field, either form's wording

The two courier forms word an item label differently, so a column is read
under either: CBE-XIV prints `Currency for Unit Price` where CBE-XIII prints
`Currency of Unit Price`. The alias fills the one column rather than adding a
second. `ITEM_LABEL_ALIASES` in `schema.py` holds these.

## Checks the extract is held to

Run `python -m pytest tests -q`.

- Every field of both specs is reproduced exactly.
- Duty heads add up to `Duty Amount` on every item of both forms, within 1.0.
- Every column the form answers is non-empty for at least one item.
- No field name appears twice, in either list.
