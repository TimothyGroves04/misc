# Transurban Group Financial Model Generator

A Python script that programmatically generates a comprehensive 3-way financial model for Transurban Group (ASX: TCL) in Excel format.

## Overview

This tool creates a fully-linked Excel workbook containing:
- **Assumptions & Drivers**: Key assumptions driving the forecast
- **Income Statement**: Revenue, expenses, and profitability metrics
- **Balance Sheet**: Assets, liabilities, and equity
- **Cash Flow Statement**: Operating, investing, and financing cash flows

All three financial statements are interconnected following best-practice financial modeling principles.

## Features

- **Historical Data**: Covers FY21-FY25 (years ended 30 June)
- **Forecast Period**: FY26-FY30 with formula-driven calculations
- **Fully Integrated**: All statements link to each other with proper accounting flows
- **Professional Formatting**: Color-coded historical vs. forecast periods, consistent styling
- **Australian Standards**: Figures in Australian Dollars, millions (A$m)

## Requirements

- Python 3.7+
- openpyxl library

## Installation

1. Clone this repository:
```bash
git clone https://github.com/TimothyGroves04/misc.git
cd misc
```

2. Install the required dependency:
```bash
pip install openpyxl
```

## Usage

Run the script from the command line:

```bash
python3 generate_model.py
```

The script will generate `Transurban_Group_3Way_Financial_Model.xlsx` in the same directory.

## Files in This Repository

- `generate_model.py`: Main Python script that generates the financial model
- `Transurban_Group_3Way_Financial_Model.xlsx`: Sample output Excel file
- `Model template.xlsb`: Template file for reference
- `.gitignore`: Git ignore configuration

## Output

The generated Excel workbook includes:
- Freeze panes for easy navigation
- Color-coded columns (blue for historical, yellow for forecast)
- Accounting number formats
- Professional styling and layout
- Grid lines disabled for cleaner presentation

## Technical Details

The script uses the `openpyxl` library to:
- Create and format Excel worksheets
- Apply custom styles (fonts, colors, borders)
- Insert formulas for financial calculations
- Link data between sheets
- Format numbers in accounting style

## Author

Timothy Groves (t.groves@uqconnect.edu.au)

## License

This project is available for educational and reference purposes.