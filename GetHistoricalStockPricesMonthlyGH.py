import pandas as pd
import datetime as dt
import numpy as np
import os
from openpyxl import load_workbook
import time
from smartapi import SmartConnect

### Get Tickers
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
# histfinal_df = pd.DataFrame()
   
# ------------------------------------ get Open, high, Low, Close of the Option Tikcers

for index,row in stocktickerdf.iterrows():
    
    exchange = row['exch_seg']
    token = row['token']
    interval = 'ONE_DAY'
    start = "2020-01-01 09:00"
    end = "2021-08-04 16:00"
        
    historicParam={"exchange": exchange,"symboltoken": token,"interval": interval,"fromdate":start ,"todate": end}
    df = obj.getCandleData(historicParam)
    histprices = pd.json_normalize(df, "data", meta=["status", "errorcode","message"])
    histprices = histprices.drop(['status','errorcode','message'], axis = 1)
    histprices['token'] = token
    histprices.columns = ['date','open','high','low','close','volume','token']
       
    # format date
    histprices["date"] = pd.to_datetime(histprices["date"]).dt.strftime('%Y%m%d')
    
    date={'months':['Feb 20','March 20','April 20','May 20','June 20','July 20','Aug 20', 'Sep 20','Oct 20','Nov 20','Dec 20','Jan 21','Feb 21','March 21','April 21','May 21', 'June 21','July 21', 'Aug 21'],
          'datestart':['01/26/2020','02/26/2020','03/26/2020','04/26/2020','05/26/2020','06/26/2020','07/26/2020','08/26/2020','09/26/2020','10/26/2020','11/26/2020','12/26/2020','01/26/2021','02/26/2021','03/26/2021','04/26/2021','05/26/2021','06/26/2021','07/26/2021'],
          'dateend':['02/25/2020','03/25/2020','04/25/2020','05/25/2020','06/25/2020','07/25/2020','08/25/2020','09/25/2020','10/25/2020','11/25/2020','12/25/2020','01/25/2021','02/25/2021','03/25/2021','04/25/2021','05/25/2021','06/25/2021','07/25/2021','08/26/2021']}
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
    time.sleep(2)

# Export the data
histprices_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\historyprices\historyprices.xlsx', index = None, header=True)
