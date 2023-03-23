import sys
from openpyxl import load_workbook
import xlrd
import os

wb = load_workbook("Enter File path/name")
sheet = wb.worksheets[0]
row_count = sheet.max_row
column_count = sheet.max_column
workbook = xlrd.open_workbook("Enter File path/name")
worksheet = workbook.sheet_by_index(0)
for i in range(0, row_count):
        name=worksheet.cell_value(i, 3)
        value=worksheet.cell_value(i, 1)
        type=worksheet.cell_value(i, 2)
        action=worksheet.cell_value(i, 4)
        if(name!="SSM Location" and name!="" and value!="" and value!="Value" and type!="" and type!="String"):
            print('on parameter ',name,'performing operation ',action)            
            if(action=='ADD' or action=='UPDATE'):
                os.system(f"aws ssm put-parameter --name {name} --value '{value}' --type '{type}' --overwrite")
            elif(action=='REMOVE'):
                os.system(f"aws ssm delete-parameter --name {name}")
            else:
                sys.exit('invalid action')
