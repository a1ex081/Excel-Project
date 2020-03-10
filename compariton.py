import os
import time
import pandas as pd 
import win32com.client 

start = time.time()

# import sheets
excel_path = r'G:/Users2 (Temp)/AlexB/Private/Macro'
os.chdir(excel_path)

my_spreadsheet1 = pd.read_excel('Pricing for Week 09 - 2020-2.xlsm', sheet_name='buying worksheet')
my_spreadsheet2 = pd.read_excel('temp.xlsx', sheet_name='Sheet1')

column1 = my_spreadsheet1['Unnamed: 2'].tolist()
column2 = my_spreadsheet2['test'].tolist()

excelApp = win32com.client.GetActiveObject('Excel.Application')
excelBook = excelApp.workBooks(r'Pricing for Week 09 - 2020-2.xlsm')
excelWorkSheet = excelBook.worksheets(r'buying worksheet')

b = column1
a = column2

for i in b:
    try: 
        if i in a:
            print('{} is in both sets'.format(i))
        else:
            print('{} is not in either set'.format(i))
    except:
        pass        

stop = time.time()
elapsed = stop - start
total_time = time.strftime('%H:%M:%S', time.gmtime(elapsed))
print('\nTotal time elapsed: {}'.format(total_time))
