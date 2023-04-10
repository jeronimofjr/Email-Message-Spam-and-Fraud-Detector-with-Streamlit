from openpyxl import Workbook

# create a new workbook
wb = Workbook()

# select the active sheet
ws = wb.active

# modify the active sheet
ws['A1'] = 'Hello, World!'

# save the workbook
wb.save('example.xlsx')