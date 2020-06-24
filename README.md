# Scraping Stock Prices
This repository contains a small script that gets **stocks prices**, **dividend** **frequency**, and **dividend** from the TMX website using the ticker. The prices are updated in the original Excel file from which the ticker is retrieved.

## Setup

### Python Version
Python Version: 3.7.6

### Libraries
1. time
2. openpyxl
3. bs4
4. requests
5. pandas
6. re
7. sys

### Requirements
1. The Excel file and script must be in the same folder
2. The Excel file must be named 'Stock_Analysis.xlsx'
3. The Excel file must contain the following sheets: 'Potential_Investments', 'Current_Holdings'
4. The sheets must contain the following columns and they must be populated:
- 'Ticker': Stock symbol (CTC.A, BMO, etc.)
- 'Currency': 'CAD' or 'USD'
5. The sheets must contain the following columns but do not need to be populated:'Current_Price', 'Div_Frequency', 'Dividend', 'Link'

**Note**: The Excel file attached in this repo is already formatted to work with this script.

## How-to Use
The script can be run from the command line based on your Python configurations or manually run using tools like VSCode, Spyder, etc.
- Just ensure the script and Excel file are in the same folder
- A sample Excel file has been upload to test

## Helpful Links
1. https://automatetheboringstuff.com/2e/chapter13/
2. https://realpython.com/openpyxl-excel-spreadsheets-python/
