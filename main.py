import openpyxl
from datetime import datetime

formatDate = lambda d: datetime.strptime(d, '%d/%m/%Y %H:%M:%S')

DAY = 19
MONTH = 8
YEAR = 2017

doc = openpyxl.load_workbook('test.xlsx')
fillSheet = doc.get_sheet_by_name("Bowl_fill_history")
fillSheet = [r for r in fillSheet.rows if r[3].value[0] == "A"]
formatedRow = {}
begEnd = {}
for row in fillSheet:
    if row[1].value in formatedRow.keys():
        formatedRow[row[1].value].append(row[4].value)
    else:
        formatedRow[row[1].value] = [row[4].value]
for key in formatedRow.keys():
    #print(key)
    for item in formatedRow[key]:
        item = formatDate(item)
        if item.day == DAY and item.hour == 8 and item.minute > 45 and item.minute < 59:
            if key in begEnd.keys():
                begEnd[key]['begin'] = item
            else:
                begEnd[key] = {'begin': item}
    if "begin" not in begEnd[key].keys():
        begEnd[key]['begin'] = formatDate(str(DAY) + "/" + str(MONTH) + "/" + str(YEAR) + " 08:55:00")
    for item in formatedRow[key]:
        item = formatDate(item)
        if item.day == DAY and item.hour == 10 and item.minute < 30:
            begEnd[key]['end'] = item
    if "end" not in begEnd[key].keys():
        begEnd[key]['end'] = formatDate(str(DAY)+"/"+str(MONTH)+"/"+str(YEAR)+" 10:05:00")
print(begEnd)
