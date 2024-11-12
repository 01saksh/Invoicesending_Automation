# Invoicesending_Automation

## VBA Script (Code)
The following VBA script should be placed in a module in the Excel workbook:
Sub SendEmailsFromSheet()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ws As Worksheet
    Dim i As Long
    Dim LastRow As Long
    Dim ToEmail As String
    Dim CCEmail As String
    Dim EmailSubject As String
    Dim EmailBody As String
    Dim ClientName As String
    Dim AttachmentFileName As String
    Dim AttachmentFolderPath As String
    Dim FullAttachmentPath As String
    Dim FileExtension As String
    Dim FileFound As String
    
    ' Define the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your actual sheet name

    ' Define the attachment folder path
    AttachmentFolderPath = "C:\Users\ANAROCK\your file name\"

    ' Define the file extension
    FileExtension = ".pdf" ' Change this to your actual file extension

    ' Find the last row with data
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Create Outlook application object
    Set OutApp = CreateObject("Outlook.Application")

    ' Loop through each row in the sheet
    For i = 2 To LastRow ' Assuming the first row is header
        ' Read values from the sheet
        ClientName = ws.Cells(i, 1).Value ' "Client Name" column
        ToEmail = ws.Cells(i, 3).Value ' "To" column
        CCEmail = ws.Cells(i, 4).Value ' "CC" column
        EmailSubject = ws.Cells(i, 5).Value ' "Subject" column
        EmailBody = ws.Cells(i, 6).Value ' "Body" column

        ' Construct the search pattern for the file (assuming the format includes ClientName)
        AttachmentFileName = "*" & ClientName & "*" & FileExtension

        ' Find the file using Dir with the pattern
        FileFound = Dir(AttachmentFolderPath & AttachmentFileName)

        ' Check if the attachment file exists
        If FileFound <> "" Then
            FullAttachmentPath = AttachmentFolderPath & FileFound

            ' Create a new email item
            Set OutMail = OutApp.CreateItem(0)
            With OutMail
                .To = ToEmail
                .CC = CCEmail
                .Subject = EmailSubject
                .Body = EmailBody
                .Attachments.Add FullAttachmentPath
                .Send ' or use .Display to show the email before sending
            End With
            Set OutMail = Nothing

            ' Update status column to "Done"
            ws.Cells(i, 8).Value = "Done" ' Assuming the "Status" column is the 8th column (Column H)
        Else
            MsgBox "Attachment file not found for client: " & ClientName, vbExclamation
        End If
    Next i

    ' Clean up
    Set OutApp = Nothing

    MsgBox "Emails sent successfully!"
End Sub




