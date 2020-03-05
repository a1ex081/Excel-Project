import os
import time
import pandas as pd 

def import_spreadsheets(test):

    # Test override
    test = False

    if test == False:

        excel_path = r'G:/Users2 (Temp)/AlexB/Private/Macro'
        os.chdir(excel_path)

        my_spreadsheet2 = pd.read_excel('temp.xlsx', sheet_name='Sheet1')
        #print(my_spreadsheet2)
        
        col1 = my_spreadsheet2['test']
        #for x in range(len(col1)): print('{}'.format(col1[x]))

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