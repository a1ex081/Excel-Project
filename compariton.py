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

listA = my_spreadsheet1['Unnamed: 2'].tolist()
tinyList = my_spreadsheet2['test'].tolist()

listB = my_spreadsheet1['Unnamed: 2']

index = pd.Index(listB)
#x = index.get_loc('02952')+2


excelApp = win32com.client.GetActiveObject('Excel.Application')
excelBook = excelApp.workBooks(r'Pricing for Week 09 - 2020-2.xlsm')
excelWorkSheet = excelBook.worksheets(r'buying worksheet')

for item in listA:
    
    try:
        for item1 in tinyList:
            if not str(item) in str(item1):
                x = index.get_loc(item)+2
                excelWorkSheet.Range('b{}:t{}'.format(x, x)).Interior.ColorIndex = 0
                excelWorkSheet.Range('w{}:af{}'.format(x, x)).Interior.ColorIndex = 0   
            elif not str(item) in str(item1).zfill(5):
                x = index.get_loc(item)+2
                excelWorkSheet.Range('b{}:t{}'.format(x, x)).Interior.ColorIndex = 0
                excelWorkSheet.Range('w{}:af{}'.format(x, x)).Interior.ColorIndex = 0   
            else:
                pass
    except:
        pass
    
    try:
        for item1 in tinyList:
            if str(item) in str(item1):
                #print(item, item1, tinyList)
                x = index.get_loc(item)+2
                excelWorkSheet.Range('b{}:t{}'.format(x, x)).Interior.ColorIndex = 3
                excelWorkSheet.Range('w{}:af{}'.format(x, x)).Interior.ColorIndex = 3
                
            elif str(item) in str(item1).zfill(5):
                #print(item, item1, tinyList)
                x = index.get_loc(item)+2
                excelWorkSheet.Range('b{}:t{}'.format(x, x)).Interior.ColorIndex = 3
                excelWorkSheet.Range('w{}:af{}'.format(x, x)).Interior.ColorIndex = 3
            else: 
                pass
                
    except:
        pass

stop = time.time()
elapsed = stop - start
total_time = time.strftime('%H:%M:%S', time.gmtime(elapsed))
print('\nTotal time elapsed: {}'.format(total_time))
