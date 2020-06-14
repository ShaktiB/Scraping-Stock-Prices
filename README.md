# Scraping Stock Prices
This repository contains a small script that gets stocks prices from the TMX website using the ticker. The prices are updated in the original Excel file from which the ticker is retrieved.

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

### Requirements
1. The Excel file and script must be in the same folder
2. The Excel file must be named 'Stock_Analysis.xlsx'
3. The Excel file must contain the following columns:
- 'Stock Name': name of the stocks
- 'Ticker': the stock symbol (ex: Canadian Tire -> CTC.A)
- 'Currency': 'CAD' or 'US'
- 'Current_Price': this can be empty or filed in, the values will be overwritten with new scraped values

## How-to Use
The script can be run from the command line based on your Python configurations or manually run using tools like VSCode, Spyder, etc.
- Just ensure the script and Excel file are in the same folder
- A sample Excel file has been upload to test

## Helpful Links
1. https://automatetheboringstuff.com/2e/chapter13/
2. https://realpython.com/openpyxl-excel-spreadsheets-python/

git branch test
