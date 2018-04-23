import openpyxl
wb = openpyxl.load_workbook('PivotCIM.xlsx', data_only=True)
print(wb('raw-data'))