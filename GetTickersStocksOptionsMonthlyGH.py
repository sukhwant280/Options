import pandas as pd
import requests
import numpy as np

# Get Nifty 50 Tickers
niftydf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickers.xlsx', sheet_name= 'IndexMembers')  

# -----------------------------------Get the list of Options tickers as per expiry date
urlk = "https://margincalculator.angelbroking.com/OpenAPI_File/files/OpenAPIScripMaster.json"
resp = requests.get(urlk)
data1 = resp.json()
tickerlist = pd.json_normalize(data1,errors='ignore')

# select Option Tickers from the list
optiontickerlist = tickerlist[(tickerlist['exch_seg'] == 'NFO') & (tickerlist['instrumenttype'] == 'OPTSTK')]

# select Stock Tickers from the list
stocktickerlist = tickerlist[(tickerlist['exch_seg'] == 'NSE') & (tickerlist['symbol'].str.contains('-EQ', regex = True) == True)]
stocktickerlist['expiry'] = np.where(stocktickerlist.index%2==0, 'stock', 'stock') 

#append the above tables
alltickers = pd.concat([optiontickerlist,stocktickerlist], sort =False).reset_index()

# Assign columns to Dataframe 
alltickers = alltickers[['symbol', 'token','name','expiry','strike','lotsize','instrumenttype','exch_seg','tick_size']]

# Join the tables
tickerAll = pd.merge(alltickers, niftydf, left_on=['name'], right_on=['Symbol'], how='inner')

# Drop the columns
tickerAll = tickerAll.drop(['Symbol'], axis = 1)

# Export the data
tickerAll.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionTickers.xlsx', index = None, header=True)


