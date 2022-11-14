import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

def importcsv(fname,fworksheet_name):
    #READ CSV
    data = pd.read_csv(fname)

    #SORT DATES
    data.sort_values(by = 'Date', ascending = True, inplace = True)

    #DROP DATAFRAME COLUMNS
    df=data.drop ('Original Description', axis=1)
    df = df[['Status','Date','Description', 'Category', 'Amount']]
    #DELTE ROWS WITH CONTAINED WORDS
    df = df[df["Status"].str.contains("Recurring|Scheduled|Pending") == False]
    #CHANGE DATAFRAME Status to Lower
    df['Status']=df['Status'].str.lower()

    # BEGINNING WITH CSV IMPORT
    wb1_destination = openpyxl.load_workbook('<insert_filepath_here>')
    ws1_destination = wb1_destination[fworksheet_name]
    #DETERMINE LAST ROW (NOT NEEDED FOR CODE TO WORK)
    start_row=ws1_destination.max_row+1
    print(start_row)

    #IMPORT DATAFRAME TO EXCEL
    for r in dataframe_to_rows(df, index=False, header=False):
        ws1_destination.append (r)
    wb1_destination.save("<insert_filepath_here>")
    print ('import complete')



importcsv('<insert_filepath1>', '<insert_fworksheet_name>')

