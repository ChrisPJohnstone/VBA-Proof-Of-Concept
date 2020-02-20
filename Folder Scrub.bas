Attribute VB_Name = "Module1"
Sub CheckFolders()
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .DisplayAlerts = False
End With

Dim FileExplorer As Object, Folder As Object, SubFolder As Object, File As Object

'Create an instance of the FileSystemObject
Set FileExplorer = CreateObject("Scripting.FileSystemObject")

'Get the folder object
Set Folder = FileExplorer.GetFolder("\\dvnas01\Share\DV-DiallerTeam-Only\Vodafone\SMS Campaign\Sent")

For Each SubFolder In Folder.subfolders

    For Each File In SubFolder.Files
    
        Call GetNumbers(File.Path, File.Name)
        
    Next File

Next SubFolder

ThisWorkbook.Sheets("Data").Range("A1").Value = "Telephone Numbers"
ThisWorkbook.Sheets("Data").Range("A1").Value = "FileName"
ThisWorkbook.Sheets("Data").Columns(1).NumberFormat = "0##########"

With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .DisplayAlerts = True
End With

End Sub
Sub GetNumbers(ByVal FilePath As String, ByVal FileName As String)

Dim ResultLR As Long, DataLR As Long, SheetName As String

SheetName = Left(FileName, Len(FileName) - 4)
If Len(SheetName) > 31 Then SheetName = Left(SheetName, 31)

Workbooks.Open FilePath

ResultLR = ThisWorkbook.Sheets("Data").Cells(Rows.Count, 1).End(xlUp).Row
DataLR = Workbooks(FileName).Sheets(SheetName).Cells(Rows.Count, 1).End(xlUp).Row

ThisWorkbook.Sheets("Data").Range("A" & (ResultLR + 1) & ":A" & (ResultLR + DataLR - 1)).Value = Workbooks(FileName).Sheets(SheetName).Range("A2:A" & DataLR).Value
ThisWorkbook.Sheets("Data").Range("B" & (ResultLR + 1) & ":B" & (ResultLR + DataLR - 1)).Value = FileName

Workbooks(FileName).Close

End Sub
