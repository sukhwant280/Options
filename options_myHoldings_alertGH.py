import glob
import os.path
import pandas as pd
import requests
import matplotlib.pyplot as plt
from dataframe_to_image import dataframe_to_image
# from pandas.table.plotting import table

def options_myHoldings_alert():
    
    # to send the message to telegram
    def telegram_bot_sendtext(bot_message):
            
            bot_token = 'Your Token' # cryto_arb_alert_bot - token
            bot_chatID = 'Your Chat ID' # sending to sukhi - '1523871308'
            send_text = 'https://api.telegram.org/bot' + bot_token + '/sendMessage?chat_id=' + bot_chatID + '&parse_mode=Markdown&text=' + bot_message
            response = requests.get(send_text)
            return response.json()
    
    # find the latest excel file from the shared drive
    folder_path = r'C:\Users\sukhw\OneDrive\Documents\Options\OptionPrices\specifictickers'
    file_type = '\*xlsx'
    files = glob.glob(folder_path + file_type)
    max_file = max(files, key=os.path.getctime)
    
    # save in dataframe
    df = pd.read_excel (r'{}'.format(max_file))
    
    # drop columns
    df.drop(df.columns[[0,3,4,5,6,7,9]], axis=1, inplace=True)
    
    # select option Tickers from the list
    df = df[(df['exchange'] == 'NFO')]
    
    # Sort the table ascending
    df = df.sort_values(by = ['Diffpercent'], ascending = [True]).reset_index(drop=True)
    df1 = df[(df['Diffpercent'] < -0.4 )] # decrease is more than this %
    df1 = df1.head(3) # select top 3 rows
    r1 = len(df1.index) # store rows of the table in a variable
    
    # Sort the table descending
    df = df.sort_values(by = ['Diffpercent'], ascending = [False]).reset_index(drop=True)
    df2 = df[(df['Diffpercent'] > 0.4 )] # increase is more than this %
    df2 = df2.head(3) # select top 3 rows
    r2 = len(df2.index) # store rows of the table in a variable
    
    # get price and % data
    if r1 >= 1:
        lt1 = df1.iat[0,1]
        lp1 = df1.iat[0,2]
        ld1 = df1.iat[0,3]*100
    
    if r1 >= 2:
        lt2 = df1.iat[1,1]
        lp2 = df1.iat[1,2]
        ld2 = df1.iat[1,3]*100
    
    if r1 >= 3:
        lt3 = df1.iat[2,1]
        lp3 = df1.iat[2,2]
        ld3 = df1.iat[2,3]*100
    
    if r2 >= 1:
        gt1 = df2.iat[0,1]
        gp1 = df2.iat[0,2]
        gd1 = df2.iat[0,3]*100
    
    if r2 >= 2:
        gt2 = df2.iat[1,1]
        gp2 = df2.iat[1,2]
        gd2 = df2.iat[1,3]*100
    
    if r2 >= 3:
        gt3 = df2.iat[2,1]
        gp3 = df2.iat[2,2]
        gd3 = df2.iat[2,3]*100
    
    # # ############################     decision making
    if r1 >= 1 :
        lt1 = str(lt1)
        lp1 = str(int(lp1))
        ld1 = str(int(ld1))
    else:
        lt1 = str("")
        lp1 = str("")
        ld1 = str("")
        
    if r1 >= 2 :   
        lt2 = str(lt2)
        lp2 = str(int(lp2))
        ld2 = str(int(ld2))
    else:
        lt2 = str("")
        lp2 = str("")
        ld2 = str("")
        
    if r1 >= 3 :
        lt3 = str(lt3)
        lp3 = str(int(lp3))
        ld3 = str(int(ld3))
    else:
        lt3 = str("")
        lp3 = str("")
        ld3 = str("")
    
    if r2 >= 1 :
        gt1 = str(gt1)
        gp1 = str(int(gp1))
        gd1 = str(int(gd1))
    else:
        gt1 = str("")
        gp1 = str("")
        gd1 = str("")
        
    if r2 >= 2 :   
        gt2 = str(gt2)
        gp2 = str(int(gp2))
        gd2 = str(int(gd2))
    else:
        gt2 = str("")
        gp2 = str("")
        gd2 = str("")
        
    if r2 >= 3 :
        gt3 = str(gt3)
        gp3 = str(int(gp3))
        gd3 = str(int(gd3))
    else:
        gt3 = str("")
        gp3 = str("")
        gd3 = str("")
    
    if r2 == 1 :
        bot_message = "\nGainers:\n\nSymbol: "+gt1+"\nPrice: "+gp1+"\n% Change: "+gd1
        test = telegram_bot_sendtext(bot_message)
    elif r2 == 2 :
        bot_message = "\nGainers:\n\nSymbol: "+gt1+"\nPrice: "+gp1+"\n% Change: "+gd1+"\n\nSymbol: "+gt2+"\nPrice: "+gp2+"\n% Change: "+gd2
        test = telegram_bot_sendtext(bot_message)
    elif r2 == 3 :
        bot_message = "\nGainers:\n\nSymbol: "+gt1+"\nPrice: "+gp1+"\n% Change: "+gd1+"\n\nSymbol: "+gt2+"\nPrice: "+gp2+"\n% Change: "+gd2+"\n\nSymbol: "+gt3+"\nPrice: "+gp3+"\n% Change: "+gd3
        test = telegram_bot_sendtext(bot_message)

    
    if r1 == 1 :
        bot_message = "\nLosers:\n\nSymbol: "+lt1+"\nPrice: "+lp1+"\n% Change: "+ld1
        test = telegram_bot_sendtext(bot_message)
    elif r1 == 2 :
        bot_message = "\nLosers:\n\nSymbol: "+lt1+"\nPrice: "+lp1+"\n% Change: "+ld1+"\n\nSymbol: "+lt2+"\nPrice: "+lp2+"\n% Change: "+ld2
        test = telegram_bot_sendtext(bot_message)
    elif r1 == 3 :
        bot_message = "\nLosers:\n\nSymbol: "+lt1+"\nPrice: "+lp1+"\n% Change: "+ld1+"\n\nSymbol: "+lt2+"\nPrice: "+lp2+"\n% Change: "+ld2+"\n\nSymbol: "+lt3+"\nPrice: "+lp3+"\n% Change: "+ld3
        test = telegram_bot_sendtext(bot_message)
