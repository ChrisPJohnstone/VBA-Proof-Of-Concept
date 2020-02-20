Attribute VB_Name = "Module1"
Function DB()
Dim Access As Object, Cnt As Long
On Error GoTo Err:
Set Access = CreateObject("Access.Application")

Access.Visible = True
Access.OpenCurrentDatabase (FilePath & "\Automation\Agent Level Performance DB.accdb")
Cnt = 0
Access.DoCmd.OpenQuery "00 Clear Holding"
Cnt = 1
Access.DoCmd.OpenQuery "01 Linked to Holding"
Cnt = 2
Access.DoCmd.OpenQuery "02 Holding to Archive"

Exit Function
Err:
If Cnt = 0 Then
    MsgBox "Error has occured with query " & Chr(34) & "00 Clear Holding" & Chr(34) & ". Please ensure you have 1. Registered ODBC Connection & 2. Enabled all macros in database trust settings. If you are unsure of how to complete either of these steps please refer to process guide"
    Else
    If Cnt = 1 Then
        MsgBox "Error has occured with query " & Chr(34) & "01 Linked to Holding" & Chr(34)
        Else
        If Cnt = 2 Then
            MsgBox "Error has occured with query " & Chr(34) & "02 Holding to Archive" & Chr(34) & ". Please ensure this file has not already been added to archive"
        End If
    End If
End If
End
End Function

