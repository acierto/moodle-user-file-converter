from xlrd import open_workbook, xldate_as_tuple
from unidecode import unidecode
import datetime

wb = open_workbook('SGTiRmoodle.xls')
students = []
first_sheet = wb.sheet_by_index(0)

counter = 0
mail_cnt = 1
# read a row

students.append([])
students[counter].append("username")
students[counter].append("password")
students[counter].append("firstname")
students[counter].append("lastname")
students[counter].append("email")
students[counter].append("cohort1")
counter += 1

for row_num in range(0, first_sheet.nrows):
    date_of_birth_float = first_sheet.cell(row_num,6).value
    cohort = str(int(first_sheet.cell(row_num,5).value))

    date_of_birth = datetime.datetime(*xldate_as_tuple(date_of_birth_float, wb.datemode))

    formatted_date_of_birth = str(date_of_birth.strftime("%d%m%y"))

    first_name = unidecode(first_sheet.cell(row_num,2).value.strip().replace(' ', ''))
    last_name = unidecode(first_sheet.cell(row_num,3).value.strip().replace(' ', ''))

    email = first_sheet.cell(row_num,4).value.strip().replace(' ', '')
    if not email:
        email = "test%s@mytest.com" % (1 + mail_cnt)
        mail_cnt += 1

    login = first_name[0:1].lower() + last_name.lower()

    password = last_name + formatted_date_of_birth + '@'

    students.append([])
    students[counter].append(login) # login
    students[counter].append(password) # password
    students[counter].append(first_name) # name
    students[counter].append(last_name) # surname
    students[counter].append(email) # email
    students[counter].append(cohort) # cohort

    counter += 1

import csv
with open('students.csv', 'wb') as csvfile:
    spamwriter = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)

    for row_num in range(0, len(students)):
        spamwriter.writerow(students[row_num])