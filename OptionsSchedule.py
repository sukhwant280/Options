from GetPricesTickersSpecific import getprices 
import schedule
import time

schedule.every(45).seconds.do(getprices)



while True:
  
    # Checks whether a scheduled task 
    # is pending to run or not
    schedule.run_pending()
    time.sleep(1)
    
    


