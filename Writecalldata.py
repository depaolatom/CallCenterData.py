import openpyxl, calldata, datetime, xlsxwriter, readcalldatav2, collections

#call list to find keys, the values of the dictionary are the time
#somehow check existing file for date, then find time interval, increment

wbw = xlsxwriter.Workbook('test1.xlsx')
sheetw = wbw.add_worksheet()

sheetw.set_column(0,0,15)
#time intervals
times = ["8 - 8:30", "8:30 - 9", "9 - 9:30", "9:30 - 10", "10 - 10:30",
        "10:30 - 11", "11 - 11:30", "11:30 - 12","12 - 12:30", "12:30 - 1",
        "1 - 1:30", "1:30 - 2", "2 - 2:30", "2:30 - 3", "3 - 3:30", "3:30 - 4",
        "4 - 4:30", "4:30 - 5","5 - 5:30", "5:30 - 6", ]
#dates list from read xlsx file
key = calldata.callList

row = 1
column = 0

for item in key:
    sheetw.write(row, column, item)
    row += 1
#reset variables to write time intervals
row = 0
column = 1
for item in times:
    sheetw.write(row, column, item)
    column += 1
print("Write Finished")

d = calldata.callDict

#valueList = list(d.values())
row = 1
cl = 1


def writeData(rw,col,i):
    sheetw.write(rw,col,1) #allows for increment
    count[i] = count[i] + 1
    print(count[i])
    if count[i] > 1:
        print("conditional")
        sheetw.write(rw,col, count[i])

valList = list(d.values())
print(valList)
for value in d.values():
    count = [0] * 20
    for i in range(len(value)):
        if value[i] != None:
            if datetime.time(8,0) < value[i] < datetime.time(8,30):
                writeData(row, 1, 0)

            if datetime.time(8,30) < value[i] < datetime.time(9,00):
                writeData(row, 2, 1)

            if datetime.time(9,0) < value[i] < datetime.time(9,30):
                writeData(row, 3, 2)

            if datetime.time(9,30) < value[i] < datetime.time(10,00):
                writeData(row, 4, 3)

            if datetime.time(10,0) < value[i] < datetime.time(10,30):
                writeData(row, 5, 4)

            if datetime.time(10,30) < value[i] < datetime.time(11,00):
                writeData(row, 6, 5)

            if datetime.time(11,0) < value[i] < datetime.time(11,30):
                writeData(row, 7, 6)

            if datetime.time(11,30) < value[i] < datetime.time(12,00):
                writeData(row, 8, 7)

            if datetime.time(12,0) < value[i] < datetime.time(12,30):
                writeData(row, 9, 8)

            if datetime.time(12,30) < value[i] < datetime.time(13,00):
                writeData(row, 10, 9)

            if datetime.time(13,0) < value[i] < datetime.time(13,30):
                writeData(row, 11, 10)

            if datetime.time(13,30) < value[i] < datetime.time(14,00):
                writeData(row, 12, 11)

            if datetime.time(14,0) < value[i] < datetime.time(14,30):
                writeData(row, 13, 12)

            if datetime.time(14,30) < value[i] < datetime.time(15,00):
                writeData(row, 14, 13)

            if datetime.time(15,0) < value[i] < datetime.time(15,30):
                writeData(row, 15, 14)

            if datetime.time(15,30) < value[i] < datetime.time(16,00):
                writeData(row, 16, 15)

            if datetime.time(16,00) < value[i] < datetime.time(16,30):
                writeData(row, 17, 16)

            if datetime.time(16,30) < value[i] < datetime.time(17,00):
                writeData(row, 18, 17)

            if datetime.time(17,00) < value[i] < datetime.time(17,30):
                writeData(row, 19, 18)

            if datetime.time(17,30) < value[i] < datetime.time(18,00):
                writeData(row, 20, 19)

            #print(sheet.cell(row = row, column = 1))
    row += 1
wbw.close()
