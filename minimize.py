from xlrd import open_workbook
import xlwt
import operator

wb = open_workbook('ALL.xls', encoding_override="cp1251")

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('sheet1')

first_sheet = wb.sheet_by_index(0)

for row_num in range(0, first_sheet.nrows):
    for col_num in range(0, 3):
        value = first_sheet.cell(row_num,col_num).value
        sheet.write(row_num, col_num, value)

countMap = {}

for row_num in range(0, first_sheet.nrows):
    count = 3

    valCount = 0
    for col_num in range(0, first_sheet.ncols):
        if col_num - 3 >= 0 and (col_num - 3) % 3 == 1:
            value = first_sheet.cell(row_num,col_num).value
            sheet.write(row_num, count, value)
            count += 1

            if isinstance(value, float) and row_num > 7:
                valCount += 1

    countMap[row_num] = valCount

sortedMap = sorted(countMap.items(), key=operator.itemgetter(1))

workbook.save('result.xls')


result_wb = open_workbook('result.xls', encoding_override="cp1251")
first_sheet = result_wb.sheet_by_index(0)

final_workbook = xlwt.Workbook()
sheet = final_workbook.add_sheet('sheet1')

for row_num in range(0, first_sheet.nrows):
    count = 0

    for col_num in range(0, first_sheet.ncols):
            value = first_sheet.cell(sortedMap[row_num][0],col_num).value
            if row_num > 7:
                sheet.write(count, 7 + first_sheet.nrows - row_num, value)
            else:
                sheet.write(count, row_num, value)
            count += 1

final_workbook.save('final_result.xls')