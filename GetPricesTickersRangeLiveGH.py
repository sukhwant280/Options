import pandas as pd
import datetime as dt
from smartapi import SmartConnect
from datetime import date, timedelta
import time

def getpricesrange():
    
    vdate = (dt.datetime.today() - dt.timedelta(days=0)).strftime("%d%m%Y")
    
    
    # Get Tickers
    optiontickerdf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickers.xlsx', sheet_name= 'OptionTickers')  
    optiontickerdf1 = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickersRange.xlsx', sheet_name= 'a')  
    optiontickerdf2 = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickersRange.xlsx', sheet_name= 'b')  
    optiontickerdf3 = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickersRange.xlsx', sheet_name= 'c')  
    optiontickerdf4 = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickersRange.xlsx', sheet_name= 'd')  
    optiontickerdf5 = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickersRange.xlsx', sheet_name= 'e')  
    optiontickerdf6 = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickersRange.xlsx', sheet_name= 'f')  
    optiontickerdf7 = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickersRange.xlsx', sheet_name= 'g')  
    optiontickerdf8 = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickersRange.xlsx', sheet_name= 'h')  
    tickerdf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionTickers.xlsx', sheet_name= 'Sheet1')  
    
    # Drop Null Rows from the tables
    optiontickerdf = optiontickerdf.dropna()
    optiontickerdf1 = optiontickerdf1.dropna()
    optiontickerdf2 = optiontickerdf2.dropna()
    optiontickerdf3 = optiontickerdf3.dropna()
    optiontickerdf4 = optiontickerdf4.dropna()
    optiontickerdf5 = optiontickerdf5.dropna()
    optiontickerdf6 = optiontickerdf6.dropna()
    optiontickerdf7 = optiontickerdf7.dropna()
    optiontickerdf8 = optiontickerdf8.dropna()
    
    # New Dataframe
    opdf = optiontickerdf[['Name']]
    opdf = opdf.drop_duplicates()
    opdf1 = optiontickerdf1[['name']]
    opdf1 = opdf1.drop_duplicates()
    opdf2 = optiontickerdf2[['name']]
    opdf2 = opdf2.drop_duplicates()
    opdf3 = optiontickerdf3[['name']]
    opdf3 = opdf3.drop_duplicates()
    opdf4 = optiontickerdf4[['name']]
    opdf4 = opdf4.drop_duplicates()
    opdf5 = optiontickerdf5[['name']]
    opdf5 = opdf5.drop_duplicates()
    opdf6 = optiontickerdf6[['name']]
    opdf6 = opdf6.drop_duplicates()
    opdf7 = optiontickerdf7[['name']]
    opdf7 = opdf7.drop_duplicates()
    opdf8 = optiontickerdf8[['name']]
    opdf8 = opdf8.drop_duplicates()
    
    # Join the tables
    stocktickerdf = pd.merge(tickerdf, opdf, left_on=['name'], right_on=['Name'], how='inner')
    stocktickerdf1 = pd.merge(tickerdf, opdf1, left_on=['name'], right_on=['name'], how='inner')
    stocktickerdf2 = pd.merge(tickerdf, opdf2, left_on=['name'], right_on=['name'], how='inner')
    stocktickerdf3 = pd.merge(tickerdf, opdf3, left_on=['name'], right_on=['name'], how='inner')
    stocktickerdf4 = pd.merge(tickerdf, opdf4, left_on=['name'], right_on=['name'], how='inner')
    stocktickerdf5 = pd.merge(tickerdf, opdf5, left_on=['name'], right_on=['name'], how='inner')
    stocktickerdf6 = pd.merge(tickerdf, opdf6, left_on=['name'], right_on=['name'], how='inner')
    stocktickerdf7 = pd.merge(tickerdf, opdf7, left_on=['name'], right_on=['name'], how='inner')
    stocktickerdf8 = pd.merge(tickerdf, opdf8, left_on=['name'], right_on=['name'], how='inner')
    
    # select Stock Tickers from the list
    stocktickerdf = stocktickerdf[(stocktickerdf['exch_seg'] == 'NSE')]
    stocktickerdf1 = stocktickerdf1[(stocktickerdf1['exch_seg'] == 'NSE')]
    stocktickerdf2 = stocktickerdf2[(stocktickerdf2['exch_seg'] == 'NSE')]
    stocktickerdf3 = stocktickerdf3[(stocktickerdf3['exch_seg'] == 'NSE')]
    stocktickerdf4 = stocktickerdf4[(stocktickerdf4['exch_seg'] == 'NSE')]
    stocktickerdf5 = stocktickerdf5[(stocktickerdf5['exch_seg'] == 'NSE')]
    stocktickerdf6 = stocktickerdf6[(stocktickerdf6['exch_seg'] == 'NSE')]
    stocktickerdf7 = stocktickerdf7[(stocktickerdf7['exch_seg'] == 'NSE')]
    stocktickerdf8 = stocktickerdf8[(stocktickerdf8['exch_seg'] == 'NSE')]
    
    
    # Setting the access to get Bid/ask of options
    smartApi =SmartConnect(api_key="YOURAPIKEY")
    login = smartApi.generateSession('YOURCLIENTID', 'YOURPASSWORD')
    refreshToken = login['data']['refreshToken']
    feedToken = smartApi.getfeedToken()
    smartApi.getProfile(refreshToken)
    smartApi.generateToken(refreshToken)
    
    prices_df = pd.DataFrame()
    prices_df1 = pd.DataFrame()
    
    st=time.time()
    
      # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in optiontickerdf.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['Ticker']
        symboltoken = int(row['Token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df1 = prices_df1.append(prices).reset_index(drop=True)
        
    # ------------------------------------ get Open, high, Low, Close of Stock Tikcers
    
    for index,row in stocktickerdf.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['symbol']
        symboltoken = int(row['token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df1 = prices_df1.append(prices).reset_index(drop=True)
       
        
    # Join the tables
    pricesAll_df1 = pd.merge(prices_df1, tickerdf, left_on=['tradingsymbol'], right_on=['symbol'], how='left')
    # Drop columns
    pricesAll_df1 = pricesAll_df1.drop(['symbol','token','expiry','strike','lotsize','instrumenttype','exch_seg','tick_size','Company Name','Industry','Series','ISIN Code'], axis = 1)
    #Align Columns
    pricesAll_df1 = pricesAll_df1[['name', 'exchange','tradingsymbol','symboltoken','open','high','low','close','ltp','BussDate']]
    # add change % column
    pricesAll_df1['Diffpercent'] = (pricesAll_df1['ltp'] - pricesAll_df1['close']) / pricesAll_df1['close']
    
    # Export the data
    pricesAll_df1.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices.xlsx', index = None, header=True)
    pricesAll_df1.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\specifictickers\OptionPrices{}.xlsx'.format(vdate), index = None, header=True)
    
    et=time.time()
    print("run time: %g Minutes" % ((et-st)/60))
    time.sleep(2)
    
       
    
    st1=time.time()
    # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in optiontickerdf1.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['symbol']
        symboltoken = int(row['token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
        
    # ------------------------------------ get Open, high, Low, Close of Stock Tikcers
    
    for index,row in stocktickerdf1.iterrows():
            
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
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPricesRange.xlsx', index = None, header=True)
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\alltickersrange\OptionPricesRange{}.xlsx'.format(vdate), index = None, header=True)
    
    et1=time.time()
    print("run time: %g Minutes" % ((et1-st1)/60))
    time.sleep(2)
    
    st2=time.time()
    # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in optiontickerdf2.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['symbol']
        symboltoken = int(row['token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
        
    # ------------------------------------ get Open, high, Low, Close of Stock Tikcers
    
    for index,row in stocktickerdf2.iterrows():
            
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
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPricesRange.xlsx', index = None, header=True)
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\alltickersrange\OptionPricesRange{}.xlsx'.format(vdate), index = None, header=True)
    
    et2=time.time()
    print("run time: %g Minutes" % ((et2-st2)/60))
    time.sleep(2)
    
    
    st3=time.time()
    # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in optiontickerdf3.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['symbol']
        symboltoken = int(row['token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
        
    # ------------------------------------ get Open, high, Low, Close of Stock Tikcers
    
    for index,row in stocktickerdf3.iterrows():
            
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
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPricesRange.xlsx', index = None, header=True)
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\alltickersrange\OptionPricesRange{}.xlsx'.format(vdate), index = None, header=True)
    
    et3=time.time()
    print("run time: %g Minutes" % ((et3-st3)/60))
    time.sleep(2)
    
    
    st4=time.time()
    # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in optiontickerdf4.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['symbol']
        symboltoken = int(row['token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
        
    # ------------------------------------ get Open, high, Low, Close of Stock Tikcers
    
    for index,row in stocktickerdf4.iterrows():
            
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
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPricesRange.xlsx', index = None, header=True)
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\alltickersrange\OptionPricesRange{}.xlsx'.format(vdate), index = None, header=True)
    
    et4=time.time()
    print("run time: %g Minutes" % ((et4-st4)/60))
    time.sleep(2)
      
    
    st5=time.time()
    # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in optiontickerdf5.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['symbol']
        symboltoken = int(row['token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
        
    # ------------------------------------ get Open, high, Low, Close of Stock Tikcers
    
    for index,row in stocktickerdf5.iterrows():
            
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
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPricesRange.xlsx', index = None, header=True)
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\alltickersrange\OptionPricesRange{}.xlsx'.format(vdate), index = None, header=True)
    
    et5=time.time()
    print("run time: %g Minutes" % ((et5-st5)/60))
    time.sleep(2)
    
    
    st6=time.time()
    # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in optiontickerdf6.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['symbol']
        symboltoken = int(row['token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
        
    # ------------------------------------ get Open, high, Low, Close of Stock Tikcers
    
    for index,row in stocktickerdf6.iterrows():
            
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
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPricesRange.xlsx', index = None, header=True)
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\alltickersrange\OptionPricesRange{}.xlsx'.format(vdate), index = None, header=True)
    
    et6=time.time()
    print("run time: %g Minutes" % ((et6-st6)/60))
    time.sleep(2)
    
    
    st7=time.time()
    # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in optiontickerdf7.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['symbol']
        symboltoken = int(row['token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
        
    # ------------------------------------ get Open, high, Low, Close of Stock Tikcers
    
    for index,row in stocktickerdf7.iterrows():
            
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
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPricesRange.xlsx', index = None, header=True)
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\alltickersrange\OptionPricesRange{}.xlsx'.format(vdate), index = None, header=True)
    
    et7=time.time()
    print("run time: %g Minutes" % ((et7-st7)/60))
    time.sleep(2)
    
    
    st8=time.time()
    # ------------------------------------ get Open, high, Low, Close of the Option Tikcers
    
    for index,row in optiontickerdf8.iterrows():
            
        exchange = row['exch_seg']
        tradingsymbol = row['symbol']
        symboltoken = int(row['token'])
         
        data = smartApi.ltpData(exchange, tradingsymbol, symboltoken)
        prices = pd.json_normalize(data['data'],errors='ignore')
        
        prices['BussDate'] = vdate
        
        prices_df = prices_df.append(prices).reset_index(drop=True)
        
    # ------------------------------------ get Open, high, Low, Close of Stock Tikcers
    
    for index,row in stocktickerdf8.iterrows():
            
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
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPricesRange.xlsx', index = None, header=True)
    pricesAll_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\alltickersrange\OptionPricesRange{}.xlsx'.format(vdate), index = None, header=True)
    
    et8=time.time()
    print("run time: %g Minutes" % ((et8-st8)/60))
    time.sleep(2)
