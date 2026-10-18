# Excel Sales Report Automation

A Python automation tool that converts raw sales data into a structured Excel business report.

## Features

- Automatically processes sales data from CSV files
- Calculates total revenue and units sold
- Generates product-wise sales summaries
- Generates monthly revenue summaries
- Creates an Excel workbook with multiple worksheets
- Includes a monthly revenue chart
- Exports a JSON summary report
- Creates sample sales data if no input file exists

## Technologies Used

- Python
- Pandas-style data processing with Python's built-in CSV tools
- OpenPyXL
- Excel reporting and data visualization

## How to Run

1. Install Python 3.12 or newer.
2. Install the required library:

   `py -3.12 -m pip install openpyxl`

3. Run the script:

   `py -3.12 sales_report.py`

4. Find the generated reports in the `output` folder.

## Output Files

- `output/sales_report.xlsx`
- `output/summary_report.json`

## Example Results

- Records processed: 8
- Total revenue: 1875.00
- Total units sold: 45
- Top product: Monitor

These results are based on the sample dataset included with the project.

## Purpose

This project demonstrates Python automation for sales reporting, spreadsheet generation, and basic business analytics.