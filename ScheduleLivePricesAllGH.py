from GetPricesTickersRangeLive import getpricesrange
from options_alert import options_alert
from options_myHoldings_alert import options_myHoldings_alert 
import schedule
import time
from subprocess import run

# Path and name to the script you are trying to start

def start_script():
    try:
        # Make sure 'python' command is available
        # run("python "+file_path, check=True)
        schedule.every(1).minutes.do(getpricesrange)
        schedule.every(2).minutes.do(options_myHoldings_alert)
        schedule.every(1.2).minutes.do(options_alert)

        while True:
          
            # Checks whether a scheduled task 
            # is pending to run or not
            schedule.run_pending()
            time.sleep(3)
    except:
        # Script crashed, lets restart it!
        print("Crashed - Restarting")
        handle_crash()

def handle_crash():
    time.sleep(3)  # Restarts the script after 2 seconds
    start_script()

start_script()



    
    


