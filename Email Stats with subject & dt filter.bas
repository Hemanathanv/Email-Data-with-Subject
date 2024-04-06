Attribute VB_Name = "Module1"
Sub ExportEmailDataByDateRange()
    Dim olApp As Object ' Outlook.Application
    Dim olNamespace As Object ' Outlook.Namespace
    Dim olInbox As Object ' Outlook.Folder
    Dim olItems As Object ' Outlook.Items
    Dim olMail As Object
    Dim xlApp As Object ' Excel.Application
    Dim xlWorkbook As Object ' Excel.Workbook
    Dim xlSheet As Object ' Excel.Worksheet
    Dim rowNum As Long
    Dim mailData As Object ' Dictionary to store mail data

    ' Initialize date variables
    Dim startDate As Date
    Dim endDate As Date

    ' Log start date and time
    Dim logStartDate As Date
    logStartDate = Now

    ' Loop until valid start date is entered or canceled
    Do
        startDateInput = InputBox("Enter the start date (MM/DD/YYYY format):", "Start Date")
        If startDateInput = "" Then
            MsgBox "Operation canceled by user.", vbExclamation
            Exit Sub ' Exit the subroutine if user cancels
        ElseIf IsDate(startDateInput) Then
            startDate = DateValue(startDateInput)
            Exit Do
        Else
            MsgBox "Invalid date format. Please enter a valid date in MM/DD/YYYY format.", vbExclamation
        End If
    Loop

    ' Loop until valid end date is entered or canceled
    Do
        endDateInput = InputBox("Enter the end date (MM/DD/YYYY format):", "End Date")
        If endDateInput = "" Then
            MsgBox "Operation canceled by user.", vbExclamation
            Exit Sub ' Exit the subroutine if user cancels
        ElseIf IsDate(endDateInput) Then
            endDate = DateValue(endDateInput)
            Exit Do
        Else
            MsgBox "Invalid date format. Please enter a valid date in MM/DD/YYYY format.", vbExclamation
        End If
    Loop

    ' Check if the user clicked Cancel or provided empty inputs
    If startDate = 0 Or endDate = 0 Then
        MsgBox "Operation canceled by user.", vbExclamation
        Exit Sub ' Exit the subroutine if user cancels or leaves inputs empty
    End If

    ' Create Outlook application and namespace
    Set olApp = CreateObject("Outlook.Application")
    Set olNamespace = olApp.GetNamespace("MAPI")

    ' Get the Inbox folder of the specified email account
    Set olInbox = olNamespace.Folders("hemanathan.v@sprandco.com").Folders("Inbox")

    ' Create Excel application and workbook
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True ' Optional: Set to False if you don't want Excel to be visible
    Set xlWorkbook = xlApp.Workbooks.Add
    Set xlSheet = xlWorkbook.Sheets(1)

    ' Clear previous data in columns A to F
    xlSheet.Range("A:F").ClearContents

    ' Initialize dictionary to store mail data
    Set mailData = CreateObject("Scripting.Dictionary")

    ' Process each email in the Inbox folder within the date range
    Set olItems = olInbox.Items
    For Each olMail In olItems
        If TypeOf olMail Is Object  Then ' Outlook.MailItem
            Dim receivedDateTime As Date
            Dim senderAddress As String
            Dim isReplied As Boolean
            Dim subject As String

            receivedDateTime = olMail.ReceivedTime
            If receivedDateTime >= startDate And receivedDateTime <= endDate Then
                senderAddress = olMail.senderEmailAddress
                isReplied = CheckIfReplied(olMail)
                subject = olMail.subject

                ' Check if the mail received date & time exists in mailData dictionary
                If mailData.Exists(receivedDateTime) Then
                    ' Increment the Total Mails count for this date & time
                    mailData(receivedDateTime)("Total Mails") = mailData(receivedDateTime)("Total Mails") + 1

                    ' Increment Replied Mails count if the email is a reply
                    If isReplied Then
                        mailData(receivedDateTime)("Replied Mails") = mailData(receivedDateTime)("Replied Mails") + 1
                    Else
                        ' Otherwise, increment Unreplied Mails count
                        mailData(receivedDateTime)("Unreplied Mails") = mailData(receivedDateTime)("Unreplied Mails") + 1
                    End If
                Else
                    ' Create a new dictionary for this date & time and initialize counts
                    Dim counts As Object
                    Set counts = CreateObject("Scripting.Dictionary")
                    counts("Sender Email Address") = senderAddress
                    counts("Replied Mails") = IIf(isReplied, 1, 0)
                    counts("Unreplied Mails") = IIf(isReplied, 0, 1)
                    counts("Total Mails") = 1
                    counts("Subject") = subject
                    mailData.Add receivedDateTime, counts
                End If
            End If
        End If
    Next olMail

    ' Write the data to the worksheet
    rowNum = 1
    xlSheet.Cells(rowNum, 1).Value = "Mail Received Date & Time"
    xlSheet.Cells(rowNum, 2).Value = "Sender Email Address"
    xlSheet.Cells(rowNum, 3).Value = "Replied Mails"
    xlSheet.Cells(rowNum, 4).Value = "Unreplied Mails"
    xlSheet.Cells(rowNum, 5).Value = "Total Mails"
    xlSheet.Cells(rowNum, 6).Value = "Subject"

    Dim key As Variant
    For Each key In mailData.keys
        rowNum = rowNum + 1
        xlSheet.Cells(rowNum, 1).Value = key ' Mail Received Date & Time in Column A
        xlSheet.Cells(rowNum, 2).Value = mailData(key)("Sender Email Address") ' Sender Email Address in Column B
        xlSheet.Cells(rowNum, 3).Value = mailData(key)("Replied Mails") ' Replied Mails in Column C
        xlSheet.Cells(rowNum, 4).Value = mailData(key)("Unreplied Mails") ' Unreplied Mails in Column D
        xlSheet.Cells(rowNum, 5).Value = mailData(key)("Total Mails") ' Total Mails in Column E
        xlSheet.Cells(rowNum, 6).Value = mailData(key)("Subject") ' Subject in Column F
    Next key

    ' Save the Excel workbook to the desktop
    Dim desktopPath As String
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    xlWorkbook.SaveAs desktopPath & "\EmailStats_" & Format(startDate, "MMDDYYYY") & "-" & Format(endDate, "MMDDYYYY") & ".xlsx"

    ' Log end date, end time, process time, process name, and user name
    Dim logEndDate As Date
    logEndDate = Now
    Dim processTime As Double
    processTime = DateDiff("s", logStartDate, logEndDate)
    Dim processTimeFormatted As String
    processTimeFormatted = Format(Int(processTime / 60), "00") & ":" & Format(processTime Mod 60, "00")

    Dim processName As String
    processName = "ExportEmailDataByDateRange"
    Dim userName As String
    userName = Environ("USERNAME")

    ' Create a new workbook for logging or open an existing one
    Dim logWorkbook As Object
    Dim logSheet As Object
    Dim logRow As Long
    Dim logFilePath As String
    logFilePath = desktopPath & "\BotLog.xlsx"

    On Error Resume Next
    Set logWorkbook = xlApp.Workbooks.Open(logFilePath)
    On Error GoTo 0

    If logWorkbook Is Nothing Then
        Set logWorkbook = xlApp.Workbooks.Add
        Set logSheet = logWorkbook.Sheets(1)
        logSheet.Cells(1, 1).Value = "Date"
        logSheet.Cells(1, 2).Value = "Start Time"
        logSheet.Cells(1, 3).Value = "End Time"
        logSheet.Cells(1, 4).Value = "Process Time (mm:ss)"
        logSheet.Cells(1, 5).Value = "Process Name"
        logSheet.Cells(1, 6).Value = "User Name"
        logRow = 2
    Else
        Set logSheet = logWorkbook.Sheets(1)
        logRow = logSheet.Cells(logSheet.Rows.count, 1).End(-4162).Row + 1
    End If

    logSheet.Cells(logRow, 1).Value = Date
    logSheet.Cells(logRow, 2).Value = Format(logStartDate, "hh:mm:ss AM/PM")
    logSheet.Cells(logRow, 3).Value = Format(logEndDate, "hh:mm:ss AM/PM")
    logSheet.Cells(logRow, 4).Value = processTimeFormatted ' Display process time in mm:ss format
    logSheet.Cells(logRow, 5).Value = processName
    logSheet.Cells(logRow, 6).Value = userName

    ' Save and close the log workbook
    logWorkbook.SaveAs logFilePath
    logWorkbook.Close

    ' Clean up objects and close Outlook
    xlWorkbook.Close

    Set xlSheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    Set olItems = Nothing
    Set olInbox = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing

    MsgBox "Email data exported successfully to Desktop\EmailStats_" & Format(startDate, "DDMMYYYY") & "-" & Format(endDate, "DDMMYYYY") & ".xlsx.", vbInformation
End Sub

Function CheckIfReplied(olMail As Object) As Boolean
    Dim olPropertyAccessor As Object
    Dim PR_LAST_VERB_EXECUTED As String
    Dim lastVerb As Integer

    ' Define the property tag for the last executed verb
    PR_LAST_VERB_EXECUTED = "http://schemas.microsoft.com/mapi/proptag/0x10810003"

    ' Get the PropertyAccessor object for the email
    Set olPropertyAccessor = olMail.PropertyAccessor

    ' Get the last executed verb property value
    lastVerb = olPropertyAccessor.GetProperty(PR_LAST_VERB_EXECUTED)

    ' Check if the last executed verb indicates a reply (102 for reply)
    CheckIfReplied = (lastVerb = 102)
End Function








