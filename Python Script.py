import os
import time
import pandas as pd 

def import_spreadsheets(test):

    # Test override
    test = False

    if test == False:

        excel_path = r'G:/Users2 (Temp)/AlexB/Private/Macro'
        os.chdir(excel_path)

        my_spreadsheet1, my_spreadsheet2 = pd.read_excel('Pricing for Week 09 - 2020-2.xlsm', sheet_name='buying worksheet'), pd.read_excel('temp.xlsx', sheet_name='Sheet1')
        #print(my_spreadsheet1)
        #print(my_spreadsheet2)

        colA, colB = my_spreadsheet1['Unnamed: 2'],my_spreadsheet2['test']
        #print('\nLength of list A: {}\n'.format(len(colA)))
        #for x in range(len(colA)): print('{}'.format(colA[x]))
        #print('\nLenght of list B: {}\n'.format(len(colB)))
        #for y in range(len(colB)): print('{}'.format(colB[y]))

    else:
        pass

def main():
    
    # Start tracking runtime 
    start = time.time()

    # Test Master
    test = True 

    # import spreadsheet values
    import_spreadsheets(test)

    # Stop tracking runtime & calculate elapsed time
    stop = time.time()
    elapsed = stop - start
    total_time = time.strftime('%H:%M:%S', time.gmtime(elapsed))
    print('\nTotal time elapsed: {}'.format(total_time))

if __name__=="__main__":
    main()