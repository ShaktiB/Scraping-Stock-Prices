import time
start_time = time.time()

from openpyxl import load_workbook
import time
#import bs4
import requests as rqs
from bs4 import BeautifulSoup
import pandas as pd
import re
import sys

def get_data(ticker, c):
    
    u = 'https://web.tmxmoney.com/quote.php?qm_symbol='
    

    if c == 'usd': # For US stocks
       url = u + ticker + ':us' 
    elif c == 'cad':
        url = u + ticker
    else:
        print("Make sure the currency is in 'CAD' or 'USD'")
        sys.exit(0)

    #print(url)
    response = rqs.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    
    # Get the Stock/ETF Price
    for pr in soup.find_all('span', {"class" : 'price'}):
        items = " ".join(pr.text.split())
        p = re.findall("\d+\.\d+", items)
        price = float(p[0])

    # Get the Dividend Frequency
    for pr in soup.find_all('div', {'class':'dq-card'}):
        if pr(text=re.compile('Div. Frequency:')):
            freq = pr.strong.text
    #return freq
    
    # Get the Dividend
    for pr in soup.find_all('div', {'class':'dq-card'}):
        if pr(text=re.compile('Dividend')):
            div_ = pr.strong.text.split()[0]
            #div_ = float(div_)
            try:
                div_ = float(div_)
            except:
                print("No dividend was available for: {}.".format(ticker))
    
    return [price, div_, freq, url]  
    

if __name__ == "__main__":

    file_name = 'Stock_Analysis_Personal.xlsx'
    sheet_names = ['Potential_Investments', 'Current_Holdings']
    
    try:
        df = pd.read_excel(file_name, sheet_names[0], keep_default_na=False) # Store Potential_Investments sheet in a dataframe
        curr_holdings = pd.read_excel(file_name, sheet_names[1], keep_default_na=False) # Store Current_Holdings sheet in a dataframe
        
        wb = load_workbook(filename = file_name)
        
    except FileNotFoundError:
        print(f"No file named '{file_name}' in the directory")
        sys.exit(0)
        
    except:
        sys.exit(1)
    
    # Remove whitespaces from column titles and data values
    df.columns = df.columns.str.strip()
    df['Ticker'] = df['Ticker'].str.strip()
    df['Currency'] =  df['Currency'].str.strip()
    
    curr_holdings.columns = curr_holdings.columns.str.strip()
    curr_holdings['Ticker'] = curr_holdings['Ticker'].str.strip()
    curr_holdings['Currency'] =  curr_holdings['Currency'].str.strip()
    
    # Take required columns from the Potential_Investments dataframe
    stocks = df.loc[:,['Stock Name','Ticker','Currency']]
    stocks['Currency'] = stocks['Currency'].str.lower()
    
    #tckrs = list(df['Ticker'])
    
    sheets = (wb.sheetnames)
    
    pot_inv = wb['Potential_Investments'] # Retrieve the Potential_Investments sheet
    
    # Create new columns for data that will be scraped
    stocks['Updated_Price'] = ''
    stocks['Div_Freq'] = ''
    stocks['Dividend'] = ''
    stocks['Link'] = ''
    
    stocks[['Updated_Price','Dividend','Div_Freq','Link']] = stocks.apply(lambda x: get_data(x['Ticker'], x['Currency']), axis=1, result_type='expand')
    
    stocks.set_index('Stock Name', inplace=True) # Need to assign names as index for use later when adding in scraped data
    
    # Get Indexes for the columns being updated 
    cp_indx = list(df.columns).index('Current_Price') + 1 # Get column index of for 'Current Price'
    divFreq_indx = list(df.columns).index('Div_Frequency') + 1 # Get column index of for 'Current Price'
    div_indx = list(df.columns).index('Dividend') + 1 # Get column index of for 'Current Price'
    link_indx = list(df.columns).index('Link') + 1 # Get column index of for 'Current Price'
    
    for rn in range(2,pot_inv.max_row+1):
        rowName = pot_inv.cell(row=rn, column = 1).value
        if rowName in df['Stock Name'].values:
            pot_inv.cell(row=rn,column=cp_indx).value = stocks.at[rowName,'Updated_Price'] # Requires stock names to be the row index
            pot_inv.cell(row=rn,column=divFreq_indx).value = stocks.at[rowName,'Div_Freq']
            pot_inv.cell(row=rn,column=div_indx).value = stocks.at[rowName,'Dividend']
            pot_inv.cell(row=rn,column=link_indx).value = stocks.at[rowName,'Link']
    
    wb.save(file_name)
    
    print("--- %s seconds ---" % (time.time() - start_time))
    #test



