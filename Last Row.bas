Sub LastRow()

Dim LR As Long, BookName As String, SheetName As String
BookName = ActiveWorkbook.Name
SheetName = ActiveSheet.Name

LR = Workbooks(BookName).Worksheets(SheetName).Cells(Rows.Count, 1).End(xlUp).Row

End Sub
