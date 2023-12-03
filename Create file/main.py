import openpyxl
from openpyxl.utils.cell import column_index_from_string
import re

''' Creating an excel file based on another excel file '''
wb = openpyxl.load_workbook("IPA 248 Admitad MTD 15th November'23 (1).xlsx")
sheet = wb['DATA']
wb.create_sheet(title='filtered rows')
filtered_rows = wb['filtered rows']
correct_value = re.compile(r'(\d{13})')
for row in range(1, sheet.max_row):
    if sheet.cell(row=row, column=column_index_from_string('S')).value:
        for groups in correct_value.findall(sheet.cell(row=row, column=column_index_from_string('S')).value):
            filtered_rows['S' + str(row)] = groups
            filtered_rows['O' + str(row)] = sheet.cell(row=row, column=column_index_from_string('O')).value
            filtered_rows['N' + str(row)] = sheet.cell(row=row, column=column_index_from_string('N')).value
            filtered_rows['B' + str(row)] = sheet.cell(row=row, column=column_index_from_string('B')).value
            filtered_rows['A' + str(row)] = sheet.cell(row=row, column=column_index_from_string('A')).value
            filtered_rows['E' + str(row)] = sheet.cell(row=row, column=column_index_from_string('E')).value
            filtered_rows['C' + str(row)] = sheet.cell(row=row, column=column_index_from_string('C')).value
            filtered_rows['Y' + str(row)] = sheet.cell(row=row, column=column_index_from_string('Y')).value
wb.save('test.xlsx')