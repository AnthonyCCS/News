import pandas as pd
from pandas.io import excel
import numpy as np
import os.path
'''''
Convert Excel xlsx file to csv file.
If csv file already there. it will append to that file
else it will create new csv file.
'''''

def excel_to_csv(from_excel_file, to_csv_file):

    excel_file = from_excel_file
    csv_file = to_csv_file
   

    try:

        if os.path.isfile(to_csv_file):
            print('write to existing csv file') 
            from_excel = pd.read_excel (excel_file)
            from_excel.to_csv(csv_file, mode='a', index=False, header=False)
        else:
            print("csv file created...")
            from_excel = pd.read_excel (excel_file)
            #remove dulicates
            from_excel.drop_duplicates(['date','header','timeline'],keep= 'last')
            from_excel.to_csv (csv_file, index = None, header=True)

    except Exception as ex:
        print('========================================================')
        print('Excel to csv conversion failed...')
        print('Is Excel file and csv file available?')
        print('========================================================')



if __name__ == "__main__":
    excel_to_csv( r'C:\Users\deadw\Documents\Algo\May2021\News\all_news.xlsx', r'C:\Users\deadw\Documents\Algo\May2021\News\to_sql_news.csv')