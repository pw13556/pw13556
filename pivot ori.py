import win32com.client
Excel = win32com.client.gencache.EnsureDispatch('Excel.Application') # Excel = win32com.client.Dispatch('Excel.Application')

win32c = win32com.client.constants

wb = Excel.Workbooks.Add()
Sheet1 = wb.Worksheets("Sheet1")

TestData = [['Country','Name','Gender','Sign','Amount'],
             ['CH','Max' ,'M','Plus',123.4567],
             ['CH','Max' ,'M','Minus',-23.4567],
             ['CH','Max' ,'M','Plus',12.2314],
             ['CH','Max' ,'M','Minus',-2.2314],
             ['CH','Sam' ,'M','Plus',453.7685],
             ['CH','Sam' ,'M','Minus',-53.7685],
             ['CH','Sara','F','Plus',777.666],
             ['CH','Sara','F','Minus',-77.666],
             ['DE','Hans','M','Plus',345.088],
             ['DE','Hans','M','Minus',-45.088],
             ['DE','Paul','M','Plus',222.455],
             ['DE','Paul','M','Minus',-22.455]]

for i, TestDataRow in enumerate(TestData):
    for j, TestDataItem in enumerate(TestDataRow):
        Sheet1.Cells(i+2,j+4).Value = TestDataItem

cl1 = Sheet1.Cells(2,4)
cl2 = Sheet1.Cells(2+len(TestData)-1,4+len(TestData[0])-1)
PivotSourceRange = Sheet1.Range(cl1,cl2)

PivotSourceRange.Select()

Sheet2 = wb.Worksheets(2)
cl3=Sheet2.Cells(4,1)
PivotTargetRange=  Sheet2.Range(cl3,cl3)
PivotTableName = 'ReportPivotTable'

PivotCache = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=PivotSourceRange, Version=win32c.xlPivotTableVersion14)

PivotTable = PivotCache.CreatePivotTable(TableDestination=PivotTargetRange, TableName=PivotTableName, DefaultVersion=win32c.xlPivotTableVersion14)

PivotTable.PivotFields('Name').Orientation = win32c.xlRowField
PivotTable.PivotFields('Name').Position = 1
PivotTable.PivotFields('Gender').Orientation = win32c.xlPageField
PivotTable.PivotFields('Gender').Position = 1
PivotTable.PivotFields('Gender').CurrentPage = 'M'
PivotTable.PivotFields('Country').Orientation = win32c.xlColumnField
PivotTable.PivotFields('Country').Position = 1
PivotTable.PivotFields('Country').Subtotals = [False, False, False, False, False, False, False, False, False, False, False, False]
PivotTable.PivotFields('Sign').Orientation = win32c.xlColumnField
PivotTable.PivotFields('Sign').Position = 2

DataField = PivotTable.AddDataField(PivotTable.PivotFields('Amount'))
DataField.NumberFormat = '#\'##0.00'

Excel.Visible = 1

wb.SaveAs('ranges_and_offsets.xlsx')
Excel.Application.Quit()