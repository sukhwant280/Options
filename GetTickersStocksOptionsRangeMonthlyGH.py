import pandas as pd
import numpy as np
import datetime as dt

vdate = (dt.datetime.today() - dt.timedelta(days=0)).strftime("%d%m%Y")
expirydate= "26AUG2021" # Need to make this dynamic and need to chnage every month on expiry date


# Get Tickers
tickerdf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionTickers.xlsx', sheet_name= 'Sheet1')  
tickerSOMdf = pd.read_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\alltickers\OptionPricesAll30072021.xlsx', sheet_name= 'Sheet1')  

# select Stock Tickers from the list
stocktickerdf = tickerSOMdf[(tickerSOMdf['exchange'] == 'NSE')]
# stocktickerdf = stocktickerdf.iloc[:5,]

# select Option Tickers for current month
optiontickerdf = tickerdf[(tickerdf['exch_seg'] == 'NFO') & (tickerdf['expiry'] == expirydate)]
optiontickerdf['strike'] = optiontickerdf['strike']/100

# Drop columns
optiontickerdf = optiontickerdf.drop(['lotsize','instrumenttype','tick_size','Company Name','Industry','Series','ISIN Code'], axis = 1)

# add call put column
optiontickerdf['category'] = optiontickerdf['symbol'].str[-2:]

   
# get 15% / 25% up and down strike range
stocktickerdf['strikeup'] = stocktickerdf['close'] + (stocktickerdf['close'] * 0.15)
stocktickerdf['strikeup1'] = stocktickerdf['close'] + (stocktickerdf['close'] * 0.25)
stocktickerdf['strikedown'] = stocktickerdf['close'] - (stocktickerdf['close'] * 0.15)
stocktickerdf['strikedown1'] = stocktickerdf['close'] - (stocktickerdf['close'] * 0.25)
   
# Join the tables
pricesAll_df = pd.merge(stocktickerdf, tickerdf, left_on=['tradingsymbol'], right_on=['symbol'], how='left')
# Drop columns
pricesAll_df = pricesAll_df.drop(['symbol','token','expiry','strike','lotsize','instrumenttype','exch_seg','tick_size','Company Name','Industry','Series','ISIN Code','exchange','tradingsymbol','symboltoken','open','high','low','ltp','BussDate'], axis = 1)
#Align Columns
pricesAll_df = pricesAll_df[['name', 'close','strikeup','strikeup1','strikedown','strikedown1']]

# Join the tables
tickerlistdf = pd.merge(optiontickerdf, pricesAll_df, how='left')
# Drop Null Rows from the tables
tickerlistdf = tickerlistdf.dropna()

tickerlistdf['check'] = np.where((tickerlistdf['category'] == 'CE') & ((tickerlistdf['strike'] >= tickerlistdf['strikeup']) & (tickerlistdf['strike'] <= tickerlistdf['strikeup1'])),'y',
                        np.where((tickerlistdf['category'] == 'PE') & ((tickerlistdf['strike'] <= tickerlistdf['strikedown']) & (tickerlistdf['strike'] >= tickerlistdf['strikedown1'])),'y','n'))

tickerlistdf = tickerlistdf[tickerlistdf['check'] == 'y']

# # Export the data
tickerlistdf.to_excel (r'C:\Users\sukhw\OneDrive\Documents\Options\SpecificTickersRange.xlsx', index = None, header=True)


