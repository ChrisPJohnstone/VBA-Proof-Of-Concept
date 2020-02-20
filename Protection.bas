Attribute VB_Name = "Protection"
Private Sub ProtectSheets()

Application.ScreenUpdating = "False"

'Protect all sheets in workbook
Dim WS As Worksheet
For Each WS In ThisWorkbook.Worksheets
      WS.Protect Password:="braw", UserInterfaceOnly:=True
Next WS

'Protect workbook
ActiveWorkbook.Protect "braw", True, False

End Sub
Private Sub UnProtectSheets()

Application.ScreenUpdating = "False"

'Remove protection from all sheets in workbook
Dim WS As Worksheet
For Each WS In ThisWorkbook.Worksheets
      WS.Unprotect Password:="braw"
Next WS

'Remove workbook protection
ActiveWorkbook.Unprotect "braw"

End Sub
Private Sub RequestUnlock()

'Get name of user
Dim User As String
User = Application.UserName

'Check user has access
If User = "Christopher Johnstone" Or User = "Mark McGrath" Or User = "Drew Naylor" Then
    Application.Run "UnProtectSheets"
    MsgBox "Access Granted"
    Dim Response As String
    Response = MsgBox("Would you like to know the password?", vbYesNo)
    If Response = vbYes Then MsgBox "Password is " & Chr(34) & "braw" & Chr(34)
Else
    MsgBox "Access Denied"
End If

End Sub

