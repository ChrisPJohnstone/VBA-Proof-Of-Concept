Attribute VB_Name = "Validation"
Function GetFileNames(Campaign)

GetFileNames = DLookup("[Linked File Name]", "[Campaign Key]", "[Campaign] = " & Chr(34) & Campaign & Chr(34))

End Function
Function CheckInput(Campaign As String, CampaignCode As String, xlName As String, WDate As String)

'Declaring Variables
Dim FilePath As String, DataFile As Integer, LR As String, UndelimitedFile As Variant, DataRow As Long, DelimitedFile As Variant
FilePath = CurrentProject.Path & "\"

'Open Excel File
On Error GoTo ErrHandler:
Excel.Workbooks.Open (FilePath & "Imports\" & xlName & ".csv")
On Error GoTo 0

'Commit Undelimited Data to Memory
UndelimitedFile = Excel.Workbooks(xlName & ".csv").Sheets(Left(xlName, 31)).Range("A1").CurrentRegion.Value
    
'Delimit and Complete Checks row by row
For DataRow = LBound(UndelimitedFile) To UBound(UndelimitedFile)
    DelimitedFile = Split(UndelimitedFile(DataRow, 1), "|")
    If DelimitedFile(2) <> Format(DelimitedFile(2), "YYYY-MM-DD") Then Call ErrMsg("Extract Date", xlName, DataRow)
    If DelimitedFile(6) <> Format(WDate, "YYYY-MM-DD") Then Call ErrMsg("Dial date", xlName, DataRow)
    If DelimitedFile(8) <> Format(DelimitedFile(8), "YYYY-MM-DD") Then Call ErrMsg("Return Date", xlName, DataRow)
Next DataRow

'Close Excel File
Excel.Workbooks(xlName & ".csv").Close savechanges = False

'Close Excel
Excel.Quit

Exit Function
ErrHandler:
Excel.Quit
MsgBox (xlName & " not found, please save report from Genius")
End Function
Function CreateOutput(Campaign As String, CampaignCode As String)
'Turn off notifications
DoCmd.SetWarnings False

'Run relevant query
DoCmd.OpenQuery ("02_" & CampaignCode & "_" & Campaign & "_Backup")

'Export to text file
DoCmd.RunSavedImportExport ("Export-Dialler_Outcomes_Webhelp_" & CampaignCode)

'Delete Table from Access
DoCmd.DeleteObject acTable, CampaignCode
End Function
Function CheckOutput(CampaignCode As String, WDate As String)
'Find file path of current wokbook and name of text file
Dim FilePath As String, txtFile As String
FilePath = CurrentProject.Path & "\Exports\"
txtFile = "Dialler_Outcomes_Webhelp_" & CampaignCode

'Open Text File
Excel.Workbooks.Open (FilePath & txtFile & ".txt")

'CommitUndelimited file to memory in VBA array
Dim UndelimitedFile As Variant, DelimitedFile As Variant, DataRow As Long
UndelimitedFile = Excel.Workbooks(txtFile & ".txt").Sheets(Left(txtFile, 31)).Range("A1").CurrentRegion.Value

'Delimit and Complete Checks row by row
For DataRow = (LBound(UndelimitedFile) + 1) To UBound(UndelimitedFile)
    DelimitedFile = Split(UndelimitedFile(DataRow, 1), "|")
    If DelimitedFile(0) <> CampaignCode Then Call ErrMsg("Campaign Code", txtFile, DataRow)
    If DelimitedFile(1) = "" Then Call ErrMsg("Cell Code", txtFile, DataRow)
    If DelimitedFile(2) <> Format(DelimitedFile(2), "YYYY-MM-DD") Then Call ErrMsg("Extract Date", txtFile, DataRow)
    If DelimitedFile(4) = "" Then Call ErrMsg("MSISDN", txtFile, DataRow)
    If DelimitedFile(5) <> "Webhelp" Then Call ErrMsg("Call Centre", txtFile, DataRow)
    If DelimitedFile(6) = "" Then Call ErrMsg("Call Direction", txtFile, DataRow)
    If DelimitedFile(7) = "" Then Call ErrMsg("Caller Type", txtFile, DataRow)
    If DelimitedFile(9) <> Format(WDate, "YYYY-MM-DD") Then Call ErrMsg("Extract Date", txtFile, DataRow)
    If DelimitedFile(10) <> Format(DelimitedFile(10), "hh:mm:ss") Then Call ErrMsg("Extract Date", txtFile, DataRow)
    If DelimitedFile(11) <> Format(DelimitedFile(11), "YYYY-MM-DD") Then Call ErrMsg("Return Date", txtFile, DataRow)
    If DelimitedFile(12) = "" Then Call ErrMsg("Outcome", txtFile, DataRow)
    If DelimitedFile(13) = "" Then Call ErrMsg("Call Centre Outcome", txtFile, DataRow)
    If DelimitedFile(14) = "" Then Call ErrMsg("Dialling Complete", txtFile, DataRow)
Next DataRow

Excel.Workbooks(txtFile & ".txt").Close savechanges = False
Excel.Quit

Name (FilePath & txtFile & ".txt") As (FilePath & txtFile & "_" & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & ".txt")

End Function
Function ErrMsg(Err As String, xlName As String, DataRow As Long)
With Excel
    .DisplayAlerts = False
    .Quit
End With
MsgBox (Err & " does not match for " & xlName & " in row " & DataRow)
End
End Function
