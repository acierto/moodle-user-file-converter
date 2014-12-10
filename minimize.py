from xlrd import open_workbook
import xlwt

wb = open_workbook('ALL.xls', encoding_override="cp1251")

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('sheet1')

first_sheet = wb.sheet_by_index(0)

for row_num in range(0, first_sheet.nrows):
    for col_num in range(0, 3):
        value = first_sheet.cell(row_num, col_num).value
        sheet.write(col_num, row_num, value)

for row_num in range(0, first_sheet.nrows):
    count = 3
    for col_num in range(0, first_sheet.ncols):
        if col_num - 3 >= 0 and (col_num - 3) % 3 == 1:
            value = first_sheet.cell(row_num,col_num).value
            sheet.write(count, row_num, value)
            count += 1

workbook.save('result.xls')