import os
import time
import pandas as pd 
import win32com.client 

def import_spreadsheets(test):

    # Test override
    test = False

    if test == False:

        excel_path = r'G:/Users2 (Temp)/AlexB/Private/Macro'
        os.chdir(excel_path)

        my_spreadsheet1, my_spreadsheet2 = pd.read_excel('Pricing for Week 09 - 2020-2.xlsm', sheet_name='buying worksheet'), pd.read_excel('temp.xlsx', sheet_name='Sheet1')
        #print(my_spreadsheet1)
        #print(my_spreadsheet2)

        colA, colB = my_spreadsheet1['Unnamed: 2'].tolist(), my_spreadsheet2['test'].tolist()
        # Test colA
        #print('\npm_id type: {}'.format(type(colA)))
        #print('\nLength of list A: {}\n'.format(len(colA)))
        #for x in range(len(colA)): print('{}'.format(colA[x]))
        # test colB
        #print('\ntest_val type: {}'.format(type(colB)))
        #print('\nLenght of list B: {}\n'.format(len(colB)))
        #for y in range(len(colB)): print('{}'.format(colB[y]))

        return colA, colB

    else:
        pass

def check(list1, test_val):
    
    #print('list1 = {}'.format(list1))
    #print('test_val = {}'.format(test_val))

    # traverse in the list
    result1 = test_val.count(list1)
    print(result1)

    if result1 > 0:
        print('\nCheck 1\nList1: {} - test_val: {}'.format(list1, test_val))
        print('\nYes, element exist within list')
        return True
    else: 
        print('List1: {} - test_val: {}'.format(list1, test_val))
        print('\nNo, element does not exist within list')
        return False

def process_data(test, pm_id, test_val):
    
    # Test Master override
    test = False

    if test == False:
        
        # Setting up to read/write engine for excel
        excelApp = win32com.client.GetActiveObject('Excel.Application')

        # create a reference to the actual Excel Workbook - Allows script to modify in real time
        # Path is mutable dependent on where the file resides. Network paths need UNC path /
        
        #path = r'G:Users2 (temp)/AlexB/Private/Macro/'
        #os.chdir(path)
        #print('\nCurrent Directory: {}'.format(os.listdir(path)))
        
        excelBook = excelApp.workBooks(r'Pricing for Week 09 - 2020-2.xlsm')
        excelWorkSheet = excelBook.worksheets(r'buying worksheet')

        #print('\nWorksheet Name: {}'.format(excelWorkSheet.name))
        #print('\npm_id type: {}'.format(type(pm_id)))
        #print('\ntest_val type: {}'.format(type(test_val)))

        true, false = 0, 0 
        #check(pm_id[10], test_val)
        for x in range(4, len(pm_id)):    
            
            result = check(pm_id[x], test_val)

            # sync check
            #excelWorkSheet.Range('c{}'.format(x)).Value
            try: 
                if result:
                    true += 1
                    excelWorkSheet.Range('b{}:t{}'.format(x, x)).Interior.ColorIndex = 3
                    #borderA = excelWorkSheet.Range('b{}:t{}'.format(x, x))
                    excelWorkSheet.Range('w{}:af{}'.format(x, x)).Interior.ColorIndex = 3

                    #excelWorkSheet.Range('w{}:af{}'.format(x, x)).BorderAround.ColorIndex = 1
                    
                else:
                    false += 1
                    excelWorkSheet.Range('b{}:t{}'.format(x, x)).Interior.ColorIndex = 2
                    excelWorkSheet.Range('w{}:af{}'.format(x, x)).Interior.ColorIndex = 2

                    #excelWorkSheet.Range('b{}:t{}'.format(x, x)).BorderAround.ColorIndex = 1
                    #excelWorkSheet.Range('w{}:af{}'.format(x, x)).BorderAround.ColorIndex = 1
            except:
                pass    
    
    print('\nTrue: {}\nFalse: {}\n'.format(true, false))

def main():
    
    # Start tracking runtime 
    start = time.time()

    # Test Master
    test = True 

    # import spreadsheet values
    pm_id, test_val = import_spreadsheets(test)

    process_data(test, pm_id, test_val)

    # Stop tracking runtime & calculate elapsed time
    stop = time.time()
    elapsed = stop - start
    total_time = time.strftime('%H:%M:%S', time.gmtime(elapsed))
    print('\nTotal time elapsed: {}'.format(total_time))

if __name__=="__main__":
    main()