import time
start_time = time.time()

from openpyxl import load_workbook
import time
#import bs4
import requests as rqs
from bs4 import BeautifulSoup
import pandas as pd
import re

def get_data(ticker, c, info):
    u = 'https://web.tmxmoney.com/quote.php?qm_symbol='
    
    if c == 'us': # For US stocks
       url = u + ticker + ':us' 
    else:
        url = u + ticker
        
    #print(url)
    response = rqs.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    
    #if info == 'price':
    for pr in soup.find_all('span', {"class" : 'price'}):
        items = " ".join(pr.text.split())
        p = re.findall("\d+\.\d+", items)
        price = float(p[0])
    #return price
    
    #if info == 'div_freq':
    for pr in soup.find_all('div', {'class':'dq-card'}):
        if pr(text=re.compile('Div. Frequency:')):
            freq = pr.strong.text
    #return freq
    
    #if info == 'div':
    for pr in soup.find_all('div', {'class':'dq-card'}):
        if pr(text=re.compile('Dividend')):
            div_ = pr.strong.text.split()[0]
            try:
                div_ = float(div_)
            except:
                pass
    
    return [price, div_, freq, url]  

if __name__ == "__main__":

    wb = load_workbook(filename = 'Stock_Analysis.xlsx')
    #wb = load_workbook(filename = 'Stock_Analysis.xlsx', data_only=True)
    df = pd.read_excel('Stock_Analysis.xlsx', keep_default_na=False)
    
    # Remove whitespaces from column titles and data values
    df.columns = df.columns.str.strip()
    df['Ticker'] = df['Ticker'].str.strip()
    df['Currency'] =  df['Currency'].str.strip()
    
    stocks = df.loc[:,['Stock Name','Ticker','Currency']]
    stocks['Currency'] = stocks['Currency'].str.lower()
    #tckrs = list(df['Ticker'])
    
    sheet = wb.active # takes the first sheet in the workbook
    
    # Create new columns for data that will be scraped
    stocks['Updated_Price'] = ''
    stocks['Div_Freq'] = ''
    stocks['Dividend'] = ''
    stocks['Link'] = ''
    
    stocks[['Updated_Price','Dividend','Div_Freq','Link']] = stocks.apply(lambda x: get_data(x['Ticker'], x['Currency'], 'price'), axis=1, result_type='expand')
    
    stocks.set_index('Stock Name', inplace=True) # Need to assign names as index for use later when adding in scraped data
    
    # Get Indexes for the columns being updated 
    cp_indx = list(df.columns).index('Current_Price') + 1 # Get column index of for 'Current Price'
    divFreq_indx = list(df.columns).index('Div_Frequency') + 1 # Get column index of for 'Current Price'
    div_indx = list(df.columns).index('Dividend') + 1 # Get column index of for 'Current Price'
    link_indx = list(df.columns).index('Link') + 1 # Get column index of for 'Current Price'
    
    for rn in range(2,sheet.max_row+1):
        rowName = sheet.cell(row=rn, column = 1).value
        if rowName in df['Stock Name'].values:
            sheet.cell(row=rn,column=cp_indx).value = stocks.at[rowName,'Updated_Price'] # Requires stock names to be the row index
            sheet.cell(row=rn,column=divFreq_indx).value = stocks.at[rowName,'Div_Freq']
            sheet.cell(row=rn,column=div_indx).value = stocks.at[rowName,'Dividend']
            sheet.cell(row=rn,column=link_indx).value = stocks.at[rowName,'Link']
    
    wb.save('Stock_Analysis.xlsx')
    
    print("--- %s seconds ---" % (time.time() - start_time))
    #test



