# Packing List Checker

A small Python tool that automates the packing-list checks I used to do by hand during my international trade internship.

It reads supplier packing lists (Excel), checks each line, compares them with the commercial invoice, and writes one Excel report listing every discrepancy, plus totals per HS code for customs declarations.

## Why

Checking a packing list against the invoice by hand is slow and easy to get wrong. Each supplier also uses different column names ("CTNS" vs "No. of Cartons", "G.W. (kg)" vs "Gross Weight (kg)"). A wrong quantity, a missing HS code or an item left off the invoice can hold up customs clearance. This tool runs the same checks on every file in seconds.

## What it checks

| Check | Example |
|---|---|
| Missing value | HS code or weight left empty |
| Invalid HS code | `94049`: 5 digits, expected 6 (HS), 8 (EU CN) or 10 (TARIC) |
| Quantity arithmetic | 30 cartons × 8 pcs = 240, but the total says 250 |
| Weight logic | Gross weight lower than net weight |
| Possible duplicate line | Same item listed twice in one shipment (warning) |
| Quantity mismatch | Packing list 25 pcs vs invoice 30 pcs |
| HS code mismatch | Packing list `940429` vs invoice `940421` |
| Not on invoice | Item packed (e.g. a free sample) but missing from the invoice |
| Not in packing list | Item invoiced but never packed |

## Output

`output/packing_list_report.xlsx` has four sheets:

- **Summary**: number of files, lines, errors and warnings, and a count for each check
- **Issues**: one row per problem, colour-coded (red = error, yellow = warning), filterable
- **HS_Summary**: cartons, quantity, net and gross weight and value per shipment, HS code and country of origin
- **Packing_Lists**: every packing list combined into one table, with its source file

## How to run

```bash
pip install -r requirements.txt
python make_sample_data.py        # creates sample files in data/
python packing_list_checker.py    # writes output/packing_list_report.xlsx
```

To run it on your own files:

```bash
python packing_list_checker.py --packing-lists path/to/packing_lists \
    --invoice path/to/invoice.xlsx --out path/to/report.xlsx
```

If a supplier uses a column header the tool doesn't recognise, add it to `COLUMN_ALIASES` at the top of `packing_list_checker.py`.

## Project structure

```
packing_list_checker.py   main script: load -> check -> report
make_sample_data.py       creates sample packing lists and an invoice
data/                     input files (sample data is invented)
output/                   generated report
```

## Tech

Python, pandas (loading, cleaning, merging and grouping data), openpyxl (Excel formatting).

## Ideas for next steps

- Look up duty rates for each HS code from the EU TARIC database
- Read packing lists that arrive as PDF
- Check that weights per carton stay within a normal range for each item
