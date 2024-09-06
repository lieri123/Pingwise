import openpyxl, smtplib, sys

wb = openpyxl.load_workbook('due_Records.xlsx') #placeholder name for file 
sheet = wb.get_sheet_by_name("Sheet1")
lastCol = sheet.max_column
latestMonth = sheet.cell(row=1, column=lastCol).value

