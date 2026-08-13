# CSV Report Automator

Turns a messy CSV export into a clean, formatted Excel report — automatically.

## The problem this solves

Businesses often have raw data exports (sales logs, inventory, form
responses) full of inconsistent formatting: mismatched capitalization,
stray whitespace, duplicate rows, and missing values. Cleaning this by
hand in Excel every week is slow and error-prone. This script does it
in seconds.

## What it does

- **Cleans the data:** strips whitespace, normalizes text casing,
  removes exact duplicate rows, detects and converts numeric columns
- **Summarizes it:** totals and averages for every numeric column,
  category breakdowns (e.g. counts by region, by product)
- **Reports missing data:** flags which columns have gaps and how many
- **Outputs a formatted `.xlsx`** with a Summary sheet (including a
  bar chart) and a Cleaned Data sheet

## Usage

```bash
pip install pandas openpyxl
python csv_report_automator.py input.csv output_report.xlsx
```

## Example

`sample_messy_sales.csv` is included as a demo — inconsistent casing
("south" vs "South"), stray whitespace in headers and values, a
duplicated row, and missing values in the Units Sold and Revenue
columns.

Run:
```bash
python csv_report_automator.py sample_messy_sales.csv sample_report.xlsx
```

Output: `sample_report.xlsx` — cleaned data, deduplicated, with a
summary sheet and chart, ready to hand to a client.

## Why this is useful for freelance/automation work

This is a template for a common paid request: *"I have messy data
exports every week and want a clean summary report automatically."*
The column detection is generic — it works on most simple tabular
CSVs, not just this sample — so it can be adapted quickly per client
without a rewrite.
