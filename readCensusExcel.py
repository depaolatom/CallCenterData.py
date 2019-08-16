import openpyxl, pprint
print("Opening Workbook...")
wb = openpyxl.load_workbook('censuspopdata.xlsx')
sheet = wb.get_sheet_by_name('Population by Census Tract')
countyData = {}

#TODO : FIll in countyData with each county's population and tracts
print("Reading rows...")
for row in range(2, sheet.max_row + 1):
    state = sheet['B' + str(row)].value
    county = sheet['C' + str(row)].value
    pop = sheet['D' + str(row)].value
# TODO: open a new text file and write the contents of countyData to it.
#DATA STRUCTURE = Dictionary
#make sure key for state exists
    countyData.setdefault(state, {})
    #make sure key for county exists
    countyData[state].setdefault(county, {'tracts': 0, 'pop': 0})
    #each row represents a census tract so increment by One
    countyData[state][county]['tracts'] += 1
    #increase county pop by pop in census tract
    countyData[state][county]['pop'] += int(pop)
#open new text file and write contents of countydata to it

print("Writing Results...")
resultFile = open('census2010.py', 'w')
resultFile.write('allData = ' + pprint.pformat(countyData))
resultFile.close()
print("Done. ")
