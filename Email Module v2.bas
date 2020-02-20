Attribute VB_Name = "Module1"
Function CreateEmailFile()

On Error GoTo Abort

Dim Outlook As Object: Set Outlook = CreateObject("Outlook.Application")

Dim Email As Object: Set Email = Outlook.CreateItem(olMailItem)
Dim Doc As Object: Set Doc = Email.GetInspector.WordEditor

Email.Display
        Email.To = "xxx@xxx.com"
        Email.SentOnBehalfOfName = "xxx@xxx.com"
Email.Subject = "Subject"

Doc.Range(0, 0).InsertAfter (vbNewLine & vbNewLine & "For any queries around this, please contact our Dialler Support Team via dialler.support@uk.webhelp.com and one of our team will be in touch as soon as possible.")

Doc.Range(0, 0).InsertAfter ("Hello" & vbNewLine & vbNewLine & "Please find the latest update")

Exit Function
Abort:
On Error Resume Next
Email.Close olDiscard
EntEmail.Close olDiscard
On Error GoTo 0

End Function
