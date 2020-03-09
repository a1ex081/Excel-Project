def check(list1, test_val):
    
    #print('list1 = {}'.format(list1))
    #print('test_val = {}'.format(test_val))

    # traverse in the list
    #for i in test_val:
    
    result1 = test_val.count(list1)
    print(result1)

    if result1 > 0:
        print('\nCheck 1\nList1: {} - test_val: {}'.format(list1, test_val))
        print('\nYes, element exist within list')
        return True
    else: 
        #print('List1: {} - test_val: {}'.format(list1, test_val))
        print('\nNo, element does not exist within list')
        return False
    
    """
    if str(i) in str(list1):
        print('\nCheck 1\nList1: {} - test_val: {}'.format(list1, test_val))
        print('\nYes, element exist within list')
        return True

    elif str(i).zfill(5) in str(list1):
        print('List1: {} - test_val: {}'.format(list1, test_val))
        return True
    else:
        #print('List1: {} - test_val: {}'.format(list1, test_val))
        print('\nNo, element does not exist within list')
        return False
    """
def main():

    list1 = '02952'
    test_val = '02952', '03238', '04980', '03593', '3845', '62606', '08942'

    result = check(list1, test_val)

    print('\nResult: {}'.format(result))

if __name__ == '__main__':
    main()