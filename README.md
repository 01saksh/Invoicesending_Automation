# Invoicesending_Automation

## Objective
This SOP outlines the process for sending automated emails with attachments using a VBA script in Excel. The script reads email details from an Excel sheet, attaches relevant files, and sends the emails through Outlook.

## Prerequisites
•	Microsoft Excel installed
•	Microsoft Outlook installed
•	Verify that the ActiveX Installer service startup type is set to Automatic in the Services of the laptop
•	Access to the folder containing the attachment files
•	An Excel workbook with a sheet containing the necessary email details
•	An Excel workbook should be saved as an Excel Macro-Enabled Workbook (*.xlsm)
•	In the VBA editor, you must first attach the correct Microsoft Outlook reference in the Tools dialog box

## Sheet Setup
Ensure the Excel sheet ("Sheet1") contains the following columns:
1.	Client Name (Column A)
2.	From (Column B)
3.	To (Column C)
4.	CC (Column D)
5.	Subject (Column E)
6.	Body (Column F)
7.	Status (Column H) - This will be updated by the script
8.	Dialog Box (Anywhere) – Type "Send Email" and then attach the VBA code to it.
   
## Folder Structure
Ensure all attachment files are stored in the specified folder:
Note : Please provide the correct attachment link of your PC (Change)
•	C:\Users\ANAROCK\OneDrive - Anarock Property Consultants Pvt Ltd\Documents\Automatic email\Checking invoice\
##File Naming Convention
The attachment files should contain the Client Name and have a .pdf extension.

## Procedure
1.	Open the Excel Workbook
o	Open the workbook containing the email details.
2.	Access the VBA Editor
o	Press Alt + F11 to open the VBA editor.
3.	  Open the Services Manager:
o	Press Windows + R to open the Run dialog box.
o	Type services.msc and press Enter. This will open the Services Manager.
4.	 Locate the ActiveX Installer Service:
o	In the Services Manager, scroll through the list of services to find "ActiveX Installer (AxInstSV)."
5.	 Open the Properties Dialog Box:
o	Right-click on "ActiveX Installer (AxInstSV)" and select "Properties" from the context menu. This will open the Properties dialog box for the service.
6.	 Set the Startup Type:
o	In the Properties dialog box, find the "Startup type" dropdown menu.
o	Select "Automatic" from the dropdown options.
7.	 Apply the Changes:
o	This needs the admin excess in your latop.
o	Click the "Apply" button to save the changes.
o	Click "OK" to close the Properties dialog box.
8.	 Start the Service (if not already running):
o	If the service is not already running, click the "Start" button in the Services Manager toolbar or right-click on the service and select "Start."
9.	 Verify the Changes:
o	Ensure that the "Status" column for the "ActiveX Installer (AxInstSV)" service shows "Running" and that the "Startup Type" column shows "Automatic."
10.	 Access the References Dialog Box:
o	In the VBA editor, click on the "Tools" menu.
o	Select "References" from the drop-down menu. This will open the References dialog box.
11.	  Select the Microsoft Outlook Library:
o	In the References dialog box, scroll through the list of available references.
o	Look for "Microsoft Outlook XX.X Object Library," where "XX.X" represents the version number of Outlook installed on your computer.
o	Check the box next to "Microsoft Outlook XX.X Object Library."
12.	 Confirm Your Selection:
o	Click the "OK" button to close the References dialog box and save your selection.
13.	Verify the Reference:
o	Ensure that the reference to the Microsoft Outlook Object Library is now listed under "Available References" with a check mark next to it.
14.	Insert a New Module
o	In the VBA editor, insert a new module by right-clicking on any existing module or workbook name, selecting Insert, and then Module.
15.	Paste the VBA Script
o	Copy and paste the provided VBA script into the new module.
16.	Run the Script
o	Close the VBA editor and return to Excel.
o	Press Alt + F8, select SendEmailsFromSheet, and click Run.
17.	Monitor the Status
o	The script will send emails and update the "Status" column to "Done" for each successfully sent email.
o	If an attachment is not found, a message box will notify you of the missing file for the specific client.

## Error Handling
•	If an attachment file is not found, the script will display a message box with the client's name.
•	Ensure all attachment files are correctly named and placed in the specified folder.
Completion
•	Once the script has run, a message box will confirm that all emails have been sent successfully.

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
    AttachmentFolderPath = "C:\Users\ANAROCK\OneDrive - Anarock Property Consultants Pvt Ltd\Documents\Automatic email\Checking invoice\"

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




