# excel-report-generator

A command-line tool that turns a CSV file of sales into a formatted Excel report.

## Features
- Checks that the required columns are present in the CSV
- Writes an Excel file with two sheets:
  - `Data`: the original rows from the CSV
  - `Recap`: total sales per product and per category, plus a grand total
- Prices and totals are shown with two decimals, and column widths adjust to the content
- Accepts a decimal comma in prices (for example `"6,50"`, between quotes)
- Clear error messages for a missing file, missing columns, invalid numbers, and when the output file cannot be saved

## Requirements
- Python 3.8 or newer
- [openpyxl](https://pypi.org/project/openpyxl/) (installed with the command below)

## Installation
```bash
git clone https://github.com/ricky-gg-cyber/excel-report-generator.git
cd excel-report-generator
python -m venv .venv
```
Activate the virtual environment (`source .venv/bin/activate` on Linux/macOS, `.venv\Scripts\activate` on Windows), then:
```bash
pip install -r requirements.txt
```

## Usage
```bash
python report.py examples/sales.csv report.xlsx
```
The first argument is the CSV file to read, the second is the Excel file to create.

## Input format
The CSV must have a header row with these columns (lowercase):

| Column | Meaning |
|---|---|
| `product` | Product name |
| `category` | Product category |
| `quantity` | Number of items sold (integer) |
| `unit_price` | Price of one item |

Other columns (for example `date`) are copied to the `Data` sheet as they are.

```
date,product,category,quantity,unit_price
2026-10-01,USB cable,Electronics,10,6.50
2026-10-01,Notebook,Office,25,3.20
```

## Example output
Running the command on `examples/sales.csv` gives a `Recap` sheet with these totals per category:

| Category | Total Sales |
|---|---|
| Electronics | 1102.80 |
| Office | 554.30 |
| Home | 509.90 |
| **Total** | **2167.00** |

## Known limitations
- Column names are fixed and must match exactly
- Only the comma is supported as CSV separator (not `;`)
- No automated tests yet
