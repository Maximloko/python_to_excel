import openpyxl
import re
phone_regex = re.compile(r'''(
    (\d)?
    (\s|-|\.)?           # delimiter
    (\d{3}|\(\d{3}\))  # region code
    (\s|-|\.)?           # delimiter
    (\d{3}) 
    (\s|-|\.)?           # delimiter   
    (\d{2})
     (\s|-|\.)?           # delimiter 
     (\d{2})   
    )''', re.VERBOSE)
phone_6_regex = re.compile(r'''(
    (\d{2})  
    (\s|-|\.)?           # delimiter
    (\d{2}) 
    (\s|-|\.)?           # delimiter   
    (\d{2})
    )''', re.VERBOSE)
wb = openpyxl.load_workbook('10.xlsx')
sheet = wb['Лист1']
# print(tuple(sheet['A1':'A10']))
phone_nunbers = []
for row in sheet['A1':'A10']:
    for v in row:
        # print(v.value)
        for groups in phone_regex.findall(v.value):
            phone_num = '-'.join([groups[1], groups[3], groups[5], groups[7], groups[9]])
            phone_nunbers.append(phone_num)
        for groups in phone_6_regex.findall(v.value):
            phone_6_num = '-'.join([groups[1], groups[3], groups[5]])
            phone_nunbers.append(phone_6_num)
# print(phone_nunbers)
with open('phone numbers.txt', 'w') as f:
    f.write('\n'.join(phone_nunbers))
    # for v in row:
    #     print(v.value)
# sheet = wb.active
# # sheet.columns[1]
# for cellObj in sheet.columns[1]:
#     print(cellObj.value)


