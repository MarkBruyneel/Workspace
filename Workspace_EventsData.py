# Script created to test downloading event type data through Python directly
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
# Date: 2024-09-29
#
# Used: Python 3.10

import csv
import time
from datetime import datetime
import pandas as pd # version 2.1.2
import warnings
from loguru import logger # version 0.7.0
import refinitiv.data as rd # version 1.6.2

# RuntimeWarning: invalid value encountered in cast > Pandas update? Doing this will present different errors
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
    # The script works with Excel files with the following structured fields:
    # 'ISIN', 'Issuer/Borrower Name Full', 'TradeDate', 'Price'
    excelfile = input('Please provide the location and name of the Excel file.\nExample: C:\\temp\keyword_list.xlsx \n')
    Events = pd.read_excel(f'{excelfile}', sheet_name='Sheet1') # This assumes that the Excel file has the default name
    eventcount = len(Events)

    # Create empty DataFrame/table to send the data to.
    # tochangelater to: Event_data = pd.DataFrame(columns = ['ISIN', 'Issuer/Borrower Name Full', 'TradeDate', 'Price'])
    Event_data = pd.DataFrame()

    # Create a loop to gather event data for each event
    eventnr = 0
    while eventnr < eventcount:
        # To print event number being processed
        eventcnt = eventnr + 1
        # Establishing which columns to use for request
        colnr = 0
        item = Events.iloc[eventnr, colnr] # ISIN
        event_start = Events.iloc[eventnr, (colnr + 2)] # FirstTradeDate
        event_end = Events.iloc[eventnr, (colnr + 3)] # EndDate

        # To show progress this line is printed:
        print("Retrieving data for nr.", eventcnt, ' of ', eventcount, ': ', item, ' ...')

        rd.open_session()
        try:
            df = rd.get_data(
                universe=[item],
                fields=['TR.PriceClose', 'TR.PriceCloseDate'],
                parameters={
                    'SDate': event_start,
                    'EDate': event_end,
                    'Frq': 'D',
                    'Curn': 'EUR',
                }
            )
            # Add data to table for output
            Event_data = pd.concat([df, Event_data], ignore_index=True)
            eventnr += 1
            # To prevent making too many requests in too short of a time for the API
            time.sleep(3)
        except rd.errors.RDError as e:
            # This code logs errors in the main log to make sure they can be checked later
            logger.debug('Error occurred: '+ str(e))
            eventnr += 1
            # To prevent making too many requests in too short of a time for the API
            time.sleep(3)
            continue
    # Export data as tab-delimited file which can be opened
    Event_data.to_csv(f'U:\Werk\\financiele bestanden\Workspace\Python_CodeBook_tests\Output\Workspace_events_result_' + timestr + '.tsv', sep='\t', encoding='utf-8')

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