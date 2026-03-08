Attribute VB_Name = "Module1"
Function GetHTMLSignature() As String
    Dim fso As Object
    Dim ts As Object
    Dim sigPath As String

    sigPath = Environ("APPDATA") & "\Microsoft\Signatures\"

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(sigPath) Then
        Dim file As Object
        For Each file In fso.GetFolder(sigPath).Files
            If LCase(fso.GetExtensionName(file.Name)) = "htm" Then
                Set ts = file.OpenAsTextStream(1, -2)
                GetHTMLSignature = ts.ReadAll
                ts.Close
                Exit Function
            End If
        Next
    End If

    GetHTMLSignature = ""
End Function

Sub Bulk_Email_Automation()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim nameVal As String, emailVal As String, subjectVal As String
    Dim bodyVal As String, attachPath As String
    Dim signature As String

    Set ws = Sheet1
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    signature = GetHTMLSignature()

    Set OutApp = CreateObject("Outlook.Application")

    For i = 2 To lastRow

        nameVal = ws.Cells(i, 2).Value
        emailVal = ws.Cells(i, 3).Value
        subjectVal = ws.Cells(i, 4).Value
        bodyVal = ws.Cells(i, 5).Value
        attachPath = ws.Cells(i, 6).Value

        If emailVal = "" Then GoTo NextRow

        Set OutMail = OutApp.CreateItem(0)

        OutMail.HTMLBody = "Dear " & nameVal & ",<br><br>" & bodyVal & "<br><br>" & signature
        OutMail.To = emailVal
        OutMail.Subject = subjectVal

        If attachPath <> "" Then
            If Dir(attachPath) <> "" Then
                OutMail.Attachments.Add attachPath
            End If
        End If

        OutMail.Send
        ws.Cells(i, 7).Value = "Sent"

NextRow:
    Next i

    MsgBox "All emails sent successfully!", vbInformation

End Sub
