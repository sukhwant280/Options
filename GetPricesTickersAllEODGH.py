import pandas as pd
import datetime as dt
from smartapi import SmartConnect
from datetime import date, timedelta
import time

st=time.time()
# declare variables
vdate = (dt.datetime.today() - dt.timedelta(days=0)).strftime("%d%m%Y")
expirydate= "26AUG2021|30SEP2021|stock" # Need to make this dynamic and change this at expiry of every month

# Get all Tickers
tickerdf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionTickers.xlsx', sheet_name= 'Sheet1')  

# select Stock/Options Tickers for 2 months
tickerdf = tickerdf[(tickerdf['expiry'].str.contains(expirydate, regex = True) == True)]

# add name in the table
stocknames = tickerdf[(tickerdf['exch_seg'] == 'NSE') | (tickerdf['exch_seg'] == 'NFO')]
stocknames.drop(stocknames.columns[[1,3,5,6,7,8,9,10,11,12]], axis=1, inplace=True)


# Setting the access to get Bid/ask of options

smartApi =SmartConnect(api_key="YOURAPIKEY")
login = smartApi.generateSession('YOURCLIENTID', 'YOURPASSWORD')
refreshToken = login['data']['refreshToken']
feedToken = smartApi.getfeedToken()
smartApi.getProfile(refreshToken)
smartApi.generateToken(refreshToken)

prices_df = pd.DataFrame()

# ------------------------------------ get Open, high, Low, Close of the Option Tickers

for index,row in tickerdf.iterrows():
        
    exchange = row['exch_seg']
    tradingsymbol = row['symbol']
    symboltoken = int(row['token'])
     
    data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
    prices = pd.json_normalize(data['data'],errors='ignore')
    
    prices['BussDate'] = vdate
    
    prices_df = prices_df.append(prices).reset_index(drop=True)
    
    # time.sleep(2)

# Join tables
prices_df1 = pd.merge(prices_df, stocknames, left_on=['tradingsymbol'], right_on=['symbol'], how='left')


# Export the data
prices_df1.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\alltickers\OptionPricesAll{}.xlsx'.format(vdate), index = None, header=True)

et=time.time()
print("run time: %g Minutes" % ((et-st)/60))

