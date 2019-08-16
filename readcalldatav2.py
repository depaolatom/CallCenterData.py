import openpyxl, pprint
from collections import defaultdict
print("Opening Workbook...")
wb = openpyxl.load_workbook('Fakecalldata.xlsx')
sheet = wb.get_sheet_by_name('tdepaola')
callDict = {}

value = 0
callnumber = 0;
#TODO : FIll in countyData with each county's population and tracts
print("Reading rows...")

for row in range(2, sheet.max_row +1):
    #name = sheet['A' + str(row)].value
    time = sheet['B' + str(row)].value
    date = sheet['C' + str(row)].value
    if time and date != None:
        dateStr = date.strftime("%m/%d/%Y")
        print(dateStr)
        #timeStr = time.strftime("%H:%M:%S")
        #print(timeStr)
# TODO: open a new text file and write the contents of countyData to it.
#DATA STRUCTURE = Dictionary
#make sure key for state exists
#convert string to Datetime datetime.strptime()
    callDict.setdefault(dateStr, []).append(time)

    #callDict[dateStr] = time
    #callDict.setdefault(dateStr, {})
    #callDict[dateStr].setdefault(time)
    #each row represents a census tract so increment by One
    #increase county pop by pop in census tract
#creates a list of all keys in dictionary so can parse through
#open new text file and write contents of countydata to it
callList = list(callDict)

print("Writing Results...")
resultFile = open('calldata.py', 'w')
resultFile.write('import datetime\n\n')
resultFile.write('callDict = ' + pprint.pformat(callDict))
resultFile.write('\n\n\ncallList = ' + pprint.pformat(callList))
resultFile.close()
print("Done. ")
