Sub TableSorter()

Dim BookName As String, Sheet As String, Colum As String
BookName = ActiveWorkbook.Name
Sheet = InputBox("What day are you looking to filter?")
Column = InputBox("What Column are you looking to filter on?")

Workbooks(BookName).Worksheets(Sheet).ListObjects(Sheet).Sort.SortFields.Clear
Workbooks(BookName).Worksheets(Sheet).ListObjects(Sheet).Sort.SortFields.Add Key:=Range(Sheet & "[[#All],[" & Column & "]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With Workbooks(BookName).Worksheets(Sheet).ListObjects(Sheet).Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

End Sub
