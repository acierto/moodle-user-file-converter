from xlrd import open_workbook
from xlrd import xldate
import xlwt

wb = open_workbook('logs_20150416-0118.xlsx', encoding_override="cp1251")
workbook = xlwt.Workbook()

final_workbook = xlwt.Workbook()
sheet = final_workbook.add_sheet('sheet1')

first_sheet = wb.sheet_by_index(0)

for row_num in range(0, first_sheet.nrows):
    for col_num in range(0, 6):
        value = first_sheet.cell(row_num,col_num).value

        if row_num > 2:
            if col_num == 1:
                time_value = first_sheet.cell(row_num, col_num).value
                date_and_time = str(xldate.xldate_as_datetime(float(time_value), 0))
                date=date_and_time.split()[0]
                time=date_and_time.split()[1]

                sheet.write(row_num, col_num, date)
                sheet.write(row_num, col_num + 1, time)
            elif col_num > 1:
                sheet.write(row_num, col_num + 1, value)
            else:
                sheet.write(row_num, col_num, value)
        elif row_num == 2:
            if col_num == 1:
                sheet.write(row_num, col_num, "Date")
                sheet.write(row_num, col_num + 1, "Time")
            elif col_num > 1:
                sheet.write(row_num, col_num + 1, value)
            else:
                sheet.write(row_num, col_num, value)
        else:
            sheet.write(row_num, col_num, value)

final_workbook.save('final_result.xls')