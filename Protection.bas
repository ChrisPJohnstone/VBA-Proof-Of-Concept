Private Sub ProtectSheets()

Application.ScreenUpdating = "False"

'Protect all sheets in workbook
Dim WS As Worksheet
For Each WS In ThisWorkbook.Worksheets
      WS.Protect Password:="Redacted", UserInterfaceOnly:=True
Next WS

'Protect workbook
ActiveWorkbook.Protect "Redacted", True, False

End Sub



Private Sub UnProtectSheets()

Application.ScreenUpdating = "False"

'Remove protection from all sheets in workbook
Dim WS As Worksheet
For Each WS In ThisWorkbook.Worksheets
WS.Unprotect Password:="Redacted"
Next WS

'Remove workbook protection
ActiveWorkbook.Unprotect "Redacted"

End Sub



Private Sub RequestUnlock()

'Get name of user
Dim User As String
User = Application.UserName

'Check user has access
If User = "Xxxx" Or User = "Xxxx" Then
    Application.Run "UnProtectSheets"
    MsgBox "Access Granted"
    Dim Response As String
    Response = MsgBox("Would you like to know the password?", vbYesNo)
If Response = vbYes Then MsgBox "Password is " & Chr(34) & "Redacted" & Chr(34)
Else
    MsgBox "Access Denied"
End If

End Sub
