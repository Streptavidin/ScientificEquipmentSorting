# ScientificEquipmentSorting
# Python code for converting and sorting separated Excel billing forms of a scientific instructment called EM based on the columns in Excel files.

################################################################################
# Part 1: Data cleaning and reading files
################################################################################


dirname='/Users/Equip1/bill'
old_ext = ".xlsx"
new_ext = ".csv"

# fileList = []
# print "Data need to clean:"
# print "......................................................."
# for fn in os.listdir(dirname):
#     if os.path.splitext(fn)[1] == old_ext:
#         fn1 = os.path.join(dirname, fn)
#         if os.path.isfile(fn1):
#             fileList.append(fn1)
# 
# for fn in fileList:
#     dn, fn1 = os.path.split(fn)
# #   os.system('xlsx2csv.py '+ fn1 + ' ' + os.path.splitext(fn1)[0]+new_ext)  
#     Xlsx2csv(fn1).convert(os.path.splitext(fn1)[0]+new_ext)                             
#     print "Converting",fn1, "to", os.path.splitext(fn1)[0]+new_ext
# print "......................................................."

def read_csv(filename):
    """Reads the CSV file at path, and returns a list of rows from the file.

    Parameters:
        path: path to a CSV file. 

    Returns:
        list of dictionaries: Each dictionary maps the columns of the CSV file
        to the values found in one row of the CSV file. Although this function 
        will work for any csv file, for our purposes, depending on the contents
        of the CSV file, example: [{'': '','EM Training': '0',
        'EM Usage': '1.50', 'End': '0.49305555555555558','Group': 'Wang','Start': \
        '0.41666666666666669','Totals': '1.5','User Name': 'LW','date': '01/10/2013'}]
    """
    output = []
    for row in csv.DictReader(open(filename)):
        output.append(row)
    return output



def read_billings(dirname, startdate, enddate):
    """Reads the CSV files, and returns a list of rows from the file btw start and end.

    Parameters:
        dirname: path to a CSV file. 
        startdate: string format(i.e. 201301), the starting billing date(inclusive)
        enddate: string format, the end billing date(inclusive)
        note startdate should be earlier or equal to enddate. Note you can only run
        it under dirname defined at this script.

    Returns:
        list of dictionaries: Each dictionary maps the columns of the CSV file
        to the values found in one row of the CSV file. Although this function 
        will work for any csv file, for our purposes, depending on the contents
        of the CSV file.
    """
    outputlist = list()
    for fn in os.listdir(dirname):
        if os.path.splitext(fn)[1] == new_ext:
            fnint = int(os.path.splitext(fn)[0])
            if fnint >= int(startdate) and fnint <= int(enddate):
                output1 = read_csv(fn) 
                outputlist.extend(output1)
    return outputlist


################################################################################
# Part2: Entry dictionary
################################################################################

def billing_collection_all(outputlist, case, option):
    """Given a list of billing dictionary and return a dictionary of dictionary
    with interested fields.

    Parameters:
    case is a string from ['User Name', 'Weekday','Month', 'Group', 'Year']. The first
    character of case can be lower case.
    option is either 0 or 1. When option is 1, the program will sum all EM usage time.
    If option is 0, the program will sum all EM training time.
    
    Returns:
    dictionary from case(string) to a dictionary with keys(entries,total hours)

    """
    collectiondict = dict()
    entry = 0
    hour = 0
    if option == 1:
        total = 'EM Usage'
    else:
        total = 'EM Training' 
    for ii in outputlist:
        hour = ii[total]
        if hour != '':
            if case == 'User Name' or case == 'user name':
                field = ii[case]
            elif case == 'Group' or case == 'group':
                field = ii[case]
            elif case == 'month' or case == 'Month':
                field = ii['date'].split('/')[0]
            elif case == 'year' or case == 'Year':
                field = ii['date'].split('/')[2]
            elif case == 'weekday' or case == 'Weekday':
                 field = datetime.datetime.strptime(ii['date'],"%m/%d/%Y").strftime('%A')
            if field not in collectiondict:
                collectiondict[field] = dict()
                collectiondict[field]['entries'] = 1
                collectiondict[field]['total hours'] = float("{0:.2f}".format(float(hour)))
            else:
                collectiondict[field]['entries'] += 1
                collectiondict[field]['total hours'] += float("{0:.2f}".format(float(hour)))             
    return collectiondict


################################################################################
# Part3: Calculate averages and percentages
################################################################################

def collection_average(collectiondict):
    """Given a dictionary of dictionary with interested fields, calculate the average 
    time.

    Parameters:
    collectiondict is a dict of dict generated from  billing_collection_all. In this
    program, we calculate the average time by dividing total hours by entries of the 
    nested dict.
    
    Returns:
    dictionary from case(string) to a float number.

    """
    aver_dict = dict()
    for key in collectiondict:
        value = collectiondict[key]['total hours']/collectiondict[key]['entries']
        value = float("{0:.2f}".format(value))
        aver_dict[key] = value
    return aver_dict

def collection_sum(collectiondict, option):
    """Given a dictionary of dictionary with interested fields, extract only the hour
    or entries dictionary from it.

    Parameters:
    collectiondict is a dict of dict generated from  billing_collection_all. In this
    program, we generate a dict using either entries(0) or total hours(1).
    
    Returns:
    dictionary from case(string) to a float number.

    """
    col_dict = dict()
    if option is 1:
        field = 'total hours'
    else:
        field = 'entries'
    for key in collectiondict:
        value = collectiondict[key][field]
        value = float("{0:.2f}".format(value))
        col_dict[key] = value
    return col_dict

def collection_percentage(aver_dict):
    """Given a dictionary of dictionary mapping from case to float number, calculate the 
    percentage.

    Parameters:
    aver_dict is a dict from  collection_average. In this
    program, we calculate the average time by dividing individual time by total time of 
    all the users.
    
    Returns:
    dictionary from case(string) to a float number(percentage).

    """
    per_dict = dict()
    totaltimes = sum(aver_dict.values())
    for key in aver_dict:
        value = aver_dict[key]/totaltimes
        per_dict[key] = value
    return per_dict


################################################################################
# Part4: Calculate averages and percentages
################################################################################

def pie_chart(dict):
    """Given a dictionary of dictionary mapping from case to float number, draw pie chart.

    Parameters:
    aver_dict is a dict from  collection_average. In this
    program, we display the each user running time in piechart.
    
    SideEffects:
    A piechart
    """
    plt.clf()
    #creates the figure and sets its size
    figure(1, figsize=(7,7))
    #centers the figure
    ax = axes([.2, .2, .6, .6])
    label = dict.keys()
    fracs = dict.values()  

    #autopct places the percentages inside their corresponding section
    plt.pie(fracs, labels=label, autopct='%1.1f%%')
    plt.title('Distribution of using time')
    plt.show()
    
#pie_chart(aver_dict)


def main():
    """Main function, executed when hw7code.py is run as a Python script.
    """
    dirname='/Users/Equip1/bill'
    old_ext = ".xlsx"
    new_ext = ".csv"
    outputlist=read_billings(dirname, '201403','201404')
    
    colldict_user = billing_collection_all(outputlist,'User Name', 1)
    col_user_entry = collection_sum(colldict_user, 0)
    max_entry = max(col_user_entry.values())
    print "Question1: Which user used the equipment most often in \
times?" , col_user_entry.keys()[col_user_entry.values().index(max_entry)]  
    col_user_time = collection_sum(colldict_user, 1) 
    max_time = max(col_user_time.values())
    print "Question1: Which user used the equipment most often in \
hours?" , col_user_time.keys()[col_user_time.values().index(max_time)]  
    aver_dict_user = collection_average(colldict_user)
    pie_chart(aver_dict_user)
    
    print "......................................................."
    
    colldict_month = billing_collection_all(outputlist,'Month', 1)
    col_month_time = collection_sum(colldict_month, 1) 
    max_time = max(col_month_time.values())
    min_time = min(col_month_time.values())
    print "Question2: Which month used the equipment most often in \
hours?" , col_month_time.keys()[col_month_time.values().index(max_time)] 
    print "Question2: Which month used the equipment least often in \
hours?" , col_month_time.keys()[col_month_time.values().index(min_time)]   
    aver_dict_month = collection_average(colldict_month)
    pie_chart(aver_dict_month)
    print "......................................................."
  
    colldict_weekday = billing_collection_all(outputlist,'weekday', 1)
    col_weekday_time = collection_sum(colldict_weekday, 1) 
    max_time = max(col_weekday_time.values())
    min_time = min(col_weekday_time.values())
    print "Question3: Which weekday used the equipment most often in \
hours?" , col_weekday_time.keys()[col_weekday_time.values().index(max_time)] 
    print "Question3: Which weekday used the equipment least often in \
hours?" , col_weekday_time.keys()[col_weekday_time.values().index(min_time)]   
    aver_dict_weekday = collection_average(colldict_weekday)
    pie_chart(aver_dict_weekday)
   
  

# If this file, election.py, is run as a Python script (such as by typing
# "python election.py" at the command shell), then run the main() function.
if __name__ == "__main__":
    main()


###
### Collaboration
###

# no one.
