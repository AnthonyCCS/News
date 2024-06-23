from Data_Transform.excel_to_csv import excel_to_csv

'''
this is just a script to demostrate 
importing a self written module.
the excel_to_csv module is in the same folder as this script.
So can just use 'from excel_to_csv import excel_to_csv'.
The module name and function name happens to be the same.
'''


#function call, takes 2 arguments, Excel file address and csv file address
excel_to_csv( r'C:\Users\deadw\Documents\Algo\May2021\News\all_news.xlsx', r'C:\Users\deadw\Documents\Algo\May2021\News\to_sql_news.csv')