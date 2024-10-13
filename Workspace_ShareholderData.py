# Script created to test downloading Shareholder data through Python directly
# instead of directly using the Excel addin or Jupyter NoteBook CodeBook.
# Use conditions:
# - you require a working/active LSEG Workspace account
# - before running the script you need to login using either the Excel addin or
#   the Workspace environment through the software or the website
# - the file paths in the script will need to be changed to be in line with what is
#   needed on your computer
# - the Refinitiv API has the same limits as the original Eikon API:
#   https://researchfinancial.wordpress.com/2024/09/04/workspace-excel-and-code-book-limitations/
#
# Created by: Mark Bruyneel
# Date: 2024-09-26
#
# Used: Python 3.10

import csv
import time
from datetime import datetime
import pandas as pd # version 2.1.2
import warnings
from loguru import logger # version 0.7.0
import refinitiv.data as rd # version 1.6.2

# RuntimeWarning: invalid value encountered in cast > Pandas updaten? Doing this will present different errors
# related to the same problems with dtypes. To avoid the error from popping up I chose to include the ignore option.
warnings.filterwarnings("ignore")

# Show all data in screen
pd.set_option("display.max.columns", None)

# Create year and date variable for filenames etc.
today = datetime.now()
year = today.strftime("%Y")
runday = str(datetime.today().date())
runyear = str(datetime.today().year)
timestr = time.strftime("%Y-%m-%d_%H-%M-%S")

# Create a date + time for file logging
now = str(datetime.now())
nowt = time.time()

logger.add(r'U:\Werk\financiele bestanden\Workspace\Python_CodeBook_tests\Output\Workspace_'+runday+'.log', backtrace=True, diagnose=True, rotation="100 MB", retention="12 months")
@logger.catch()

def main():
    my_file = open('U:\Werk\\financiele bestanden\Workspace\Python_CodeBook_tests\ISIN_list.txt', 'r')
    # reading the file
    data = my_file.read()
    # replacing end splitting the text when newline ('\n') is seen.
    isins = data.split('\n')
    my_file.close()

    # Create empty DataFrame/table to send the data to
    All_ShareHolder_data = pd.DataFrame()
    nrofisins = len(isins)

    for i, item in enumerate(isins, 1):
        rd.open_session()
        print("Retrieving data for ISIN nr.", i, ' of ', nrofisins, ': ', item, ' ...')
        try:
            df = rd.get_data(
                universe = [item],
                fields = ['TR.SharesHeld.investorname','TR.SharesHeld','TR.SharesHeldValue','TR.PctOfSharesOutHeld','TR.HoldingsDate','TR.FilingType','TR.InvestorType','TR.InvAddrCountry'],
                parameters = {
                    'SDate': '2022-12-31',
                }
            )
            All_ShareHolder_data = pd.concat([df, All_ShareHolder_data], ignore_index=True)
            # To prevent making too many requests in too short of a time for the API
            time.sleep(3)
        except rd.errors.RDError as e:
            # This code logs errors in the main log to make sure they can be checked later
            logger.debug('Error occurred: '+ str(e))
            # To prevent making too many requests in too short of a time for the API
            time.sleep(3)
            continue
    # Export data as tab-delimited file which can be opened
    All_ShareHolder_data.to_csv(f'U:\Werk\\financiele bestanden\Workspace\Python_CodeBook_tests\Output\Workspace_result_' + timestr + '.tsv', sep='\t', encoding='utf-8')

    # Logging of script run:
    end = str(datetime.now())
    logger.debug('Processing started at: ' + now)
    logger.debug('Processing completed at: ' + end)
    duration_s = (round((time.time() - nowt), 2))
    if duration_s > 3600:
        duration = str(duration_s / 3600)
        logger.debug('Search took: ' + duration + ' hours.')
    elif duration_s > 60:
        duration = str(duration_s / 60)
        logger.debug('Search took: ' + duration + ' minutes.')
    else:
        duration = str(duration_s)
        logger.debug('Search took: ' + duration + ' seconds.')

if __name__ == "__main__":
    main()