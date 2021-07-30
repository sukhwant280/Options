import pandas as pd
import datetime as dt
from smartapi import SmartConnect
from datetime import date, timedelta

def getprices():
    
    vdate = (dt.datetime.today() - dt.timedelta(days=0)).strftime("%d%m%Y")
    
    
    # Get Tickers
    tickerdf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickers.xlsx', sheet_name= 'Sheet1')  
    
    
    # Setting the access to get Bid/ask of options
    
    smartApi =SmartConnect(api_key="YourKey")
    login = smartApi.generateSession('YourClientID', 'YourPassword') # Angel broking
    refreshToken = login['data']['refreshToken']
    feedToken = smartApi.getfeedToken()
    smartApi.getProfile(refreshToken)
    smartApi.generateToken(refreshToken)
    
    
    
    prices_df = pd.DataFrame()
    
    # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in tickerdf.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['Ticker']
        symboltoken = int(row['Token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
    
    # Export the data
    prices_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices.xlsx', index = None, header=True)
    prices_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\OptionPrices{}.xlsx'.format(vdate), index = None, header=True)
    
    


