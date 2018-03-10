def addpivot(wb,sourcedata,title,filters=(),columns=(),
             rows=(),sumvalue=(),sortfield=""):
    """Build a pivot table using the provided source location data
    and specified fields
    """
    ...
    for fieldlist,fieldc in ((filters,win32c.xlPageField),
                            (columns,win32c.xlColumnField),
                            (rows,win32c.xlRowField)):
        for i,val in enumerate(fieldlist):
            wb.ActiveSheet.PivotTables(tname).PivotFields(val).Orientation = fieldc
        wb.ActiveSheet.PivotTables(tname).PivotFields(val).Position = i+1
    ...