from xlrd import open_workbook
from unidecode import unidecode

wb = open_workbook('SGTiRmoodle.xls')
students = []
first_sheet = wb.sheet_by_index(0)

counter = 0
# read a row
for row_num in range(0, first_sheet.nrows):
    date_of_birth = first_sheet.cell(row_num,6).value
    formatted_date_of_birth = date_of_birth[0:2] + date_of_birth[3:5] + date_of_birth[-2:]

    first_name = unidecode(first_sheet.cell(row_num,2).value.strip().replace(' ', ''))
    last_name = unidecode(first_sheet.cell(row_num,3).value.strip().replace(' ', ''))

    login = first_name[0:1].lower() + last_name.lower()

    password = unicode(last_name + formatted_date_of_birth + "@")

    students.append([])
    students[counter].append(login) # login
    students[counter].append(password) # password
    students[counter].append(first_name) # name
    students[counter].append(last_name) # surname
    students[counter].append("test%s@test.com" % (1 + counter)) # email
    students[counter].append('SGTiR_Students_Z_2014/2015') # cohort

    counter += 1

import csv
with open('students.csv', 'wb') as csvfile:
    spamwriter = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)

    for row_num in range(0, len(students)):
        print students[row_num]
        spamwriter.writerow(students[row_num])