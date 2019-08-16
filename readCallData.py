import openpyxl, pprint
print("Opening Workbook...")
wb = openpyxl.load_workbook('Fakecalldata.xlsx')
sheet = wb.get_sheet_by_name('tdepaola')
countyData = {}

callnumber = 0;
#TODO : FIll in countyData with each county's population and tracts
print("Reading rows...")

for row in range(2, sheet.max_row + 1):
    name = sheet['A' + str(row)].value
    time = sheet['B' + str(row)].value
    date = sheet['C' + str(row)].value
# TODO: open a new text file and write the contents of countyData to it.
#DATA STRUCTURE = Dictionary
#make sure key for state exists


    countyData.setdefault(name, {})
    countyData[name].setdefault(date, {})
    #make sure key for county exists
    countyData[name][date].setdefault(time, {'calls': 1})
    #each row represents a census tract so increment by One
    countyData[name][date][time]['calls'] += int(callnumber)
    #increase county pop by pop in census tract

#open new text file and write contents of countydata to it

print("Writing Results...")
resultFile = open('calldata.py', 'w')
resultFile.write('allData = ' + pprint.pformat(countyData))
resultFile.close()
print("Done. ")
