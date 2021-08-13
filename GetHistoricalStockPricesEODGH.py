import pandas as pd
import datetime as dt
import numpy as np
import os
from openpyxl import load_workbook
import time
from smartapi import SmartConnect

### Get Tickers
historydata = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\historyprices\historyprices.xlsx', sheet_name= 'Sheet1') 
tickersdf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionTickers.xlsx', sheet_name= 'Sheet1') 

## Select stocks and ignore option tickers
stocktickerdf = tickersdf[(tickersdf['exch_seg'] == 'NSE')]

####generate session #####

obj =SmartConnect(api_key="YOURAPIKEY")
login = obj.generateSession('YOURCLIENTID', 'YOURPASSWORD')
refreshToken = login['data']['refreshToken']
feedToken = obj.getfeedToken()

#################################################

histanalysis_df = pd.DataFrame()
histfinal_df = pd.DataFrame()
historydata1 = pd.DataFrame()

   
# ------------------------------------ get Open, high, Low, Close of the Option Tikcers

for index,row in stocktickerdf.iterrows():
    
    exchange = row['exch_seg']
    token = row['token']
    interval = 'ONE_DAY'
    start = "2021-07-26 09:00" # Change this every Month
    end = "2021-08-26 16:00" # Change this every Month
        
    historicParam={"exchange": exchange,"symboltoken": token,"interval": interval,"fromdate":start ,"todate": end}
    df = obj.getCandleData(historicParam)
    histprices = pd.json_normalize(df, "data", meta=["status", "errorcode","message"])
    histprices = histprices.drop(['status','errorcode','message'], axis = 1)
    histprices['token'] = token
    histprices.columns = ['date','open','high','low','close','volume','token']
       
    # format date
    histprices["date"] = pd.to_datetime(histprices["date"]).dt.strftime('%Y%m%d')
    
    date={'months':['Aug 21'], # Change this every Month
          'datestart':[start],
          'dateend':[end]}
    datedf = pd.DataFrame(date)
    
    # format date
    datedf["datestart"] = pd.to_datetime(datedf["datestart"]).dt.strftime('%Y%m%d')
    datedf["dateend"] = pd.to_datetime(datedf["dateend"]).dt.strftime('%Y%m%d')
    
    for index,row in datedf.iterrows():
    
        months = row['months']
        datestart = row['datestart']
        dateend = row['dateend']
        
        history_df = histprices[(histprices['date'] >= datestart) & (histprices['date'] <= dateend)]
        maxhigh = history_df['high'].max()
        minlow = history_df['low'].min()
        maxclose = history_df['close'].max()
        minclose = history_df['close'].min()
        closeAvg = (maxclose - minclose)/ minclose
        vol = (maxhigh - minlow)/ minlow
        
        # Make a final table
        historyavg={'token':[token],
          'months':[months],
          'close':[closeAvg],
          'vol':[vol]}
        data1 = pd.DataFrame(historyavg)
    
        histanalysis_df = histanalysis_df.append(data1).reset_index(drop=True)
    
        
    # Join tables
    histprices_df = pd.merge(histanalysis_df, tickersdf, left_on=['token'], right_on=['token'], how='left')
    # drop columns
    histprices_df.drop(histprices_df.columns[[0,4,6,7,8,9,10,11,12,13,14,15]], axis=1, inplace=True)
    #Align Columns
    histprices_df = histprices_df[['name','months','close','vol']]
    
    # histfinal_df = histfinal_df.append(histprices_df).reset_index(drop=True)
    time.sleep(1)
    
# show percentage of numbers    
histprices_df['close'] = (histprices_df['close']*100).round(decimals=1)
histprices_df['vol'] = (histprices_df['vol']*100).round(decimals=1)

# remove current month
historydata.drop(historydata[historydata['months'] == months ].index, inplace = True) 

# Add current month in master table
historydata = historydata.append(histprices_df).reset_index(drop=True)

# Remove all Nulls in all columns
historydata = historydata.dropna()

# Export the data
historydata.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\historyprices\historyprices.xlsx', index = None, header=True)

