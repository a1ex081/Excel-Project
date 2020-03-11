import os
import time
import pandas as pd 
import win32com.client 

def input(test):
    pass

def process(test):
    pass

def output(test):
    pass

def main():
    
    start = time.time()

    test = True

    stop = time.time()
    elapsed = stop - start
    total_time = time.strftime('%H:%M:%S', time.gmtime(elapsed))
    print('\nTotal time elapsed: {}'.format(total_time))


if __name__=="__main__":
    main()