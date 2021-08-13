import pandas as pd
import datetime as dt
from smartapi import SmartConnect
from datetime import date, timedelta

def getprices():
    
    vdate = (dt.datetime.today() - dt.timedelta(days=0)).strftime("%d%m%Y")
    
    
    # Get Tickers
    optiontickerdf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickers.xlsx', sheet_name= 'OptionTickers')  
    tickerdf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionTickers.xlsx', sheet_name= 'Sheet1')  
    
    # Drop Null Rows from the tables
    optiontickerdf = optiontickerdf.dropna()
    
    # New Dataframe
    opdf = optiontickerdf[['Name']]
    opdf = opdf.drop_duplicates()
    
    # Join the tables
    stocktickerdf = pd.merge(tickerdf, opdf, left_on=['name'], right_on=['Name'], how='inner')
    
    # select Stock Tickers from the list
    stocktickerdf = stocktickerdf[(stocktickerdf['exch_seg'] == 'NSE')]
    
    
    # Setting the access to get Bid/ask of options
    
    smartApi =SmartConnect(api_key="YOURAPIKEY")
    login = smartApi.generateSession('YOURCLIENTID', 'YOURPASSWORD')
    refreshToken = login['data']['refreshToken']
    feedToken = smartApi.getfeedToken()
    smartApi.getProfile(refreshToken)
    smartApi.generateToken(refreshToken)
    
    
    
    prices_df = pd.DataFrame()
    
    # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in optiontickerdf.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['Ticker']
        symboltoken = int(row['Token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
        
    # ------------------------------------ get Open, high, Low, Close of Stock Tikcers
    
    for index,row in stocktickerdf.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['symbol']
        symboltoken = int(row['token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
       
        
    # Join the tables
    pricesAll_df = pd.merge(prices_df, tickerdf, left_on=['tradingsymbol'], right_on=['symbol'], how='left')
    # Drop columns
    pricesAll_df = pricesAll_df.drop(['symbol','token','expiry','strike','lotsize','instrumenttype','exch_seg','tick_size','Company Name','Industry','Series','ISIN Code'], axis = 1)
    #Align Columns
    pricesAll_df = pricesAll_df[['name', 'exchange','tradingsymbol','symboltoken','open','high','low','close','ltp','BussDate']]
    # add change % column
    pricesAll_df['Diffpercent'] = (pricesAll_df['ltp'] - pricesAll_df['close']) / pricesAll_df['close']
    
    # Export the data
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices.xlsx', index = None, header=True)
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\specifictickers\OptionPrices{}.xlsx'.format(vdate), index = None, header=True)
    
        
