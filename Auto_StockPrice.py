import time
start_time = time.time()

from openpyxl import load_workbook
import time
import bs4
import requests as rqs
from bs4 import BeautifulSoup
import pandas as pd
import re

wb = load_workbook(filename = 'Stock_Analysis.xlsx')
#wb = load_workbook(filename = 'Stock_Analysis.xlsx', data_only=True)
df = pd.read_excel('Stock_Analysis.xlsx', keep_default_na=False)

stocks = df.loc[:,['Stock Name','Ticker','Currency']]
stocks['Currency'] = stocks['Currency'].str.lower()
tckrs = list(df['Ticker'])

sheet = wb.active # takes the first sheet in the workbook

def get_price(ticker, c):
    u = 'https://web.tmxmoney.com/quote.php?qm_symbol='
    
    if c == 'us': # For US stocks
       url = u + ticker + ':us' 
    else:
        url = u + ticker
        
    print(url)
    response = rqs.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    
    for tr in soup.find_all('span', {"class" : 'price'}):
        items = " ".join(tr.text.split())
        p = re.findall("\d+\.\d+", items)
        price = float(p[0])
        return price

stocks['Updated_Price'] = stocks.apply(lambda x: get_price(x['Ticker'], x['Currency']), axis=1)

stocks.set_index('Stock Name', inplace=True)

print(sheet['A1'].value) # .value is needed to actually return the value in the cell
print(sheet.cell(row=2,column=1).value)

cp_indx = list(df.columns).index('Current_Price') + 1 # Get column index of for 'Current Price'

for rn in range(2,sheet.max_row+1):
    rowName = sheet.cell(row=rn, column = 1).value
    if rowName in df['Stock Name'].values:
        sheet.cell(row=rn,column=cp_indx).value = stocks.at[rowName,'Updated_Price']

wb.save('Stock_Analysis.xlsx')

print("--- %s seconds ---" % (time.time() - start_time))
#test



