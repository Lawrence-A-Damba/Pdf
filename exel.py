import openpyxl
import os
wb = openpyxl.load_workbook('Sheet.xlsx')
#print(wb.sheetnames)
#sheet = wb.get_sheet_by_name('Permits')
sheet = wb['Permits']
#print(sheet['D1'].value)
for i in range(1,200):
    print(sheet['A'+str(i)].value, '\t', sheet['D'+str(i)].value, '\t', sheet['E'+str(i)].value)
