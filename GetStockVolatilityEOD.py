import pandas as pd
import datetime as dt
import numpy as np
import math
import os
from openpyxl import load_workbook
import time
from smartapi import SmartConnect

st=time.time()

m = 0

# while m != -1:
 
# Dates
startdate = (dt.datetime.today() - dt.timedelta(days=365+m)).strftime("%Y-%m-%d %H:%M")
enddate = (dt.datetime.today() - dt.timedelta(days=m)).strftime("%Y-%m-%d %H:%M")
vdate = (dt.datetime.today() - dt.timedelta(days=m)).strftime("%Y%m%d")

### Get Tickers
tickersdf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionTickers.xlsx', sheet_name= 'Sheet1') 

## Select stocks and ignore option tickers
stocktickerdf = tickersdf[(tickersdf['exch_seg'] == 'NSE')]

####generate session #####

obj =SmartConnect(api_key="l7Y6WJSy")
login = obj.generateSession('S855039', 'TahL0r123')
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
    start = startdate
    end = enddate
        
    historicParam={"exchange": exchange,"symboltoken": token,"interval": interval,"fromdate":start ,"todate": end}
    df = obj.getCandleData(historicParam)
    histprices = pd.json_normalize(df, "data", meta=["status", "errorcode","message"])
    histprices = histprices.drop(['status','errorcode','message'], axis = 1)
    histprices['token'] = token
    histprices.columns = ['date','open','high','low','close','volume','token']
       
    # format date
    histprices["date"] = pd.to_datetime(histprices["date"]).dt.strftime('%Y%m%d')
    # sort values
    histprices = histprices.sort_values(by='date', ascending=False)
    # add difference column
    histprices['diff'] = (np.log(histprices.close /histprices.close.shift(-1)))
    daily_std = np.std(histprices['diff']) * 100
    monthly_std = (daily_std * 21 ** 0.5) 
    annual_std = (daily_std * 252 ** 0.5)
   
    # Make a final table
    historyavg={'token':[token],
      'dailyvolatility':[daily_std],
      'monthlyvolatility':[monthly_std],
      'annualvolatility':[annual_std]}
    data1 = pd.DataFrame(historyavg)

    histanalysis_df = histanalysis_df.append(data1).reset_index(drop=True)

    time.sleep(2)
   
# Join tables
histprices_df = pd.merge(histanalysis_df, tickersdf, left_on=['token'], right_on=['token'], how='left')
# drop columns
histprices_df.drop(histprices_df.columns[[4,6,7,8,9,10,11,12,13,14,15]], axis=1, inplace=True)
# Align Columns
histprices_df = histprices_df[['name','token', 'dailyvolatility','monthlyvolatility','annualvolatility']]
# Sort Table
histprices_df = histprices_df.sort_values(by = ['name'], ascending = [True]).reset_index(drop=True)


# # Export the data
histprices_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\StockVolatility.xlsx', index = None, header=True)
histprices_df.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\volatility\StockVolatility{}.xlsx'.format(vdate), index = None, header=True)
    
    # m -= 1

et=time.time()
print("run time: %g Minutes" % ((et-st)/60))
