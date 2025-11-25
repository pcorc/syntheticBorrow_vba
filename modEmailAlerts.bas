Attribute VB_Name = "modEmailAlerts"

' ===================================================================
' EMAIL ALERTS MODULE - Version 2
' Synthetic Borrow Trading System
' ===================================================================

Option Explicit

' ===================================================================
' MARGIN VERIFICATION EMAIL ALERTS
' ===================================================================

Public Sub SendMarginVerificationSummary()
    On Error GoTo ErrorHandler
    
    ' Get summary data from VerificationSummary worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("VerificationSummary")
    
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then
        MsgBox "No verification data to summarize", vbInformation
        Exit Sub
    End If
    
    ' Count verification statuses
    Dim totalTrades As Long, marginPassed As Long, bbgPassed As Long, bothPassed As Long
    Dim approvedTrades As String, deniedTrades As String
    
    totalTrades = lastRow - 1
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Range("C" & i).Value = "PASS" Then marginPassed = marginPassed + 1
        If ws.Range("D" & i).Value = "PASS" Then bbgPassed = bbgPassed + 1
        
        If ws.Range("G" & i).Value = "APPROVED" Then
            bothPassed = bothPassed + 1
            approvedTrades = approvedTrades & BuildVerificationRow(ws, i, True)
        Else
            deniedTrades = deniedTrades & BuildVerificationRow(ws, i, False)
        End If
    Next i
    
    ' Build and send email
    Dim emailBody As String
    emailBody = BuildVerificationEmailBody(totalTrades, marginPassed, bbgPassed, bothPassed, _
                                          approvedTrades, deniedTrades)
    
    ' Send email
    Call SendEmail( _
        Range("email_margin_to").Value, _
        Range("email_margin_cc").Value, _
        "Trade Verification Summary - " & Format(Date, "mm/dd/yyyy"), _
        emailBody, _
        "" _
    )
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error sending verification summary: " & Err.description, vbCritical
End Sub

Public Function BuildVerificationRow(ws As Worksheet, row As Long, isApproved As Boolean) As String
    Dim rowColor As String
    rowColor = IIf(isApproved, "#d4edda", "#f8d7da") ' Green for approved, red for denied
    
    BuildVerificationRow = "<tr style='background-color: " & rowColor & ";'>" & _
                          "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("A" & row).Value & "</td>" & _
                          "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("B" & row).Value & "</td>" & _
                          "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("C" & row).Value & "</td>" & _
                          "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("D" & row).Value & "</td>" & _
                          "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("E" & row).Value, "0.00%") & "</td>" & _
                          "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("F" & row).Value, "0.00%") & "</td>" & _
                          "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("G" & row).Value & "</td>" & _
                          "</tr>"
End Function

Public Function BuildVerificationEmailBody(totalTrades As Long, marginPassed As Long, _
                                           bbgPassed As Long, bothPassed As Long, _
                                           approvedTrades As String, deniedTrades As String) As String
    Dim html As String
    html = "<html><body style='font-family: Arial, sans-serif;'>" & _
           "<h2 style='color: #003242;'>Trade Verification Summary</h2>" & _
           "<p>Date: " & Format(Date, "mm/dd/yyyy") & "</p>" & _
           "<p>Time: " & Format(Now(), "hh:mm AM/PM") & "</p>" & _
           "<h3>Summary Statistics:</h3>" & _
           "<ul>" & _
           "<li>Total Trades: " & totalTrades & "</li>" & _
           "<li>Margin Verification Passed: " & marginPassed & "</li>" & _
           "<li>Bloomberg Validation Passed: " & bbgPassed & "</li>" & _
           "<li>Both Checks Passed: " & bothPassed & "</li>" & _
           "</ul>"
    
    If approvedTrades <> "" Or deniedTrades <> "" Then
        html = html & "<h3>Detailed Results:</h3>" & _
               "<table style='border-collapse: collapse; width: 100%;'>" & _
               "<thead style='background-color: #003242; color: white;'>" & _
               "<tr>" & _
               "<th style='border: 1px solid #ddd; padding: 8px;'>Trade ID</th>" & _
               "<th style='border: 1px solid #ddd; padding: 8px;'>Client</th>" & _
               "<th style='border: 1px solid #ddd; padding: 8px;'>Margin</th>" & _
               "<th style='border: 1px solid #ddd; padding: 8px;'>Bloomberg</th>" & _
               "<th style='border: 1px solid #ddd; padding: 8px;'>Vest vs Tsy</th>" & _
               "<th style='border: 1px solid #ddd; padding: 8px;'>BBG vs Tsy</th>" & _
               "<th style='border: 1px solid #ddd; padding: 8px;'>Status</th>" & _
               "</tr></thead><tbody>" & _
               approvedTrades & deniedTrades & _
               "</tbody></table>"
    End If
    
    html = html & "<p style='margin-top: 20px; font-size: 10pt; color: #666;'>" & _
           "Generated by Synthetic Borrow Trading System</p>" & _
           "</body></html>"
    
    BuildVerificationEmailBody = html
End Function


Public Function CreateExpirationCSV(ws As Worksheet) As String
    On Error GoTo ErrorHandler
    
    Dim fileName As String
    fileName = Range("file_directory").Value & "Expirations_" & Format(Date, "YYYYMMDD") & ".csv"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open fileName For Output As #fileNum
    
    ' Write headers from specification
    Print #fileNum, "id,original_trade_id,client_name,client_account,client_email," & _
                    "execution_date,expiry_date,box_structure,executed_premium," & _
                    "payback_amount,annualized_rate,trading_system_id,order_id"
    
    ' Write expiring positions
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Range("O" & i).Value = "ALERT" Then
            Print #fileNum, _
                ws.Range("A" & i).Value & "," & _
                ws.Range("B" & i).Value & "," & _
                """" & ws.Range("C" & i).Value & """," & _
                ws.Range("D" & i).Value & "," & _
                ws.Range("E" & i).Value & "," & _
                ws.Range("F" & i).Value & "," & _
                ws.Range("G" & i).Value & "," & _
                """" & ws.Range("H" & i).Value & """," & _
                ws.Range("I" & i).Value & "," & _
                ws.Range("J" & i).Value & "," & _
                ws.Range("K" & i).Value & "," & _
                ws.Range("L" & i).Value & "," & _
                ws.Range("M" & i).Value
        End If
    Next i
    
    Close #fileNum
    CreateExpirationCSV = fileName
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    CreateExpirationCSV = ""
End Function

Public Function BuildExpirationEmailBody(alertCount As Long) As String
    ' Build HTML table of expiring positions
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ClientPortfolio")
    
    Dim html As String
    html = "<html><body style='font-family: Arial, sans-serif;'>" & _
           "<h2>Synthetic Borrow Expiration Alert</h2>" & _
           "<p>The following " & alertCount & " positions are expiring within " & _
           Range("expiration_alert_days").Value & " days:</p>" & _
           "<table style='border-collapse: collapse; width: 100%;'>" & _
           "<thead style='background-color: #003242; color: white;'>" & _
           "<tr>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Client</th>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Account</th>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Expiry Date</th>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Days Left</th>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Premium</th>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Payback</th>" & _
           "</tr></thead><tbody>"
    
    ' Add rows
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Range("O" & i).Value = "ALERT" Then
            Dim rowColor As String
            If ws.Range("N" & i).Value <= 1 Then
                rowColor = "#ffcccc" ' Light red
            ElseIf ws.Range("N" & i).Value <= 2 Then
                rowColor = "#ffe6cc" ' Light orange
            Else
                rowColor = "#ffffcc" ' Light yellow
            End If
            
            html = html & "<tr style='background-color: " & rowColor & ";'>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("C" & i).Value & "</td>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("D" & i).Value & "</td>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("G" & i).Value, "mm/dd/yyyy") & "</td>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("N" & i).Value & "</td>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("I" & i).Value, "$#,##0") & "</td>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("J" & i).Value, "$#,##0") & "</td>" & _
                   "</tr>"
        End If
    Next i
    
    html = html & "</tbody></table>" & _
           "<p style='margin-top: 20px;'>Please review and take appropriate action.</p>" & _
           "<p style='font-size: 10pt; color: #666;'>Generated by Synthetic Borrow Trading System</p>" & _
           "</body></html>"
    
    BuildExpirationEmailBody = html
End Function



' ===================================================================
' TRADE EXECUTION SUMMARY EMAIL
' ===================================================================

Public Sub SendExecutionSummary()
    On Error GoTo ErrorHandler
    
    ' Get execution data from ExecutionResults worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ExecutionResults")
    
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then
        MsgBox "No execution data to summarize", vbInformation
        Exit Sub
    End If
    
    ' Build execution summary
    Dim executionTable As String
    Dim totalExecutions As Long
    Dim totalPremium As Double
    
    totalExecutions = lastRow - 1
    
    Dim i As Long
    For i = 2 To lastRow
        totalPremium = totalPremium + CDbl(ws.Range("D" & i).Value) ' Executed Premium column
        executionTable = executionTable & BuildExecutionRow(ws, i)
    Next i
    
    ' Build email body
    Dim emailBody As String
    emailBody = "<html><body style='font-family: Arial, sans-serif;'>" & _
               "<h2 style='color: #003242;'>Daily Execution Summary</h2>" & _
               "<p>Date: " & Format(Date, "mm/dd/yyyy") & "</p>" & _
               "<p>Execution Time: " & Format(Now(), "hh:mm AM/PM") & "</p>" & _
               "<h3>Summary:</h3>" & _
               "<ul>" & _
               "<li>Total Executions: " & totalExecutions & "</li>" & _
               "<li>Total Premium Generated: " & Format(totalPremium, "$#,##0.00") & "</li>" & _
               "</ul>" & _
               "<h3>Execution Details:</h3>" & _
               BuildExecutionTable(executionTable) & _
               "<p style='margin-top: 20px; font-size: 10pt; color: #666;'>" & _
               "This summary was generated automatically by the Synthetic Borrow Trading System." & _
               "</p></body></html>"
    
    ' Send email
    Call SendEmail( _
        Range("email_execution_to").Value, _
        Range("email_execution_cc").Value, _
        "Daily Execution Summary - " & Format(Date, "mm/dd/yyyy"), _
        emailBody, _
        "" _
    )
    
    MsgBox "Execution summary sent for " & totalExecutions & " trades", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error sending execution summary: " & Err.description, vbCritical
End Sub

Public Function BuildExecutionRow(ws As Worksheet, row As Long) As String
    BuildExecutionRow = "<tr>" & _
                       "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("A" & row).Value & "</td>" & _
                       "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("B" & row).Value & "</td>" & _
                       "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("C" & row).Value & "</td>" & _
                       "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("D" & row).Value, "$#,##0") & "</td>" & _
                       "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("E" & row).Value, "$#,##0") & "</td>" & _
                       "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("F" & row).Value, "0.00%") & "</td>" & _
                       "</tr>"
End Function

Public Function BuildExecutionTable(tableRows As String) As String
    BuildExecutionTable = "<table style='border-collapse: collapse; width: 100%; margin: 10px 0;'>" & _
                         "<thead style='background-color: #007bff; color: white;'>" & _
                         "<tr>" & _
                         "<th style='border: 1px solid #ddd; padding: 8px;'>Order ID</th>" & _
                         "<th style='border: 1px solid #ddd; padding: 8px;'>Trading System ID</th>" & _
                         "<th style='border: 1px solid #ddd; padding: 8px;'>Client Account</th>" & _
                         "<th style='border: 1px solid #ddd; padding: 8px;'>Premium</th>" & _
                         "<th style='border: 1px solid #ddd; padding: 8px;'>Payback</th>" & _
                         "<th style='border: 1px solid #ddd; padding: 8px;'>Rate</th>" & _
                         "</tr>" & _
                         "</thead>" & _
                         "<tbody>" & _
                         tableRows & _
                         "</tbody>" & _
                         "</table>"
End Function

' ===================================================================
' CORE EMAIL SENDING FUNCTION
' ===================================================================

Public Sub SendEmail(toRecipients As String, ccRecipients As String, subject As String, htmlBody As String, attachmentPath As String)
    On Error GoTo ErrorHandler
    
    Dim objOutlook As Object
    Dim objMail As Object
    
    ' Create Outlook application object
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0) ' 0 = olMailItem
    
    With objMail
        .To = toRecipients
        .CC = ccRecipients
        .subject = subject
        .htmlBody = htmlBody
        
        ' Add attachment if provided
        If attachmentPath <> "" And Len(Dir(attachmentPath)) > 0 Then
            .Attachments.Add attachmentPath
        End If
        
        ' Display email (don't send automatically)
        .Display
    End With
    
    Set objMail = Nothing
    Set objOutlook = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error sending email: " & Err.description, vbCritical
    Set objMail = Nothing
    Set objOutlook = Nothing
End Sub




Public Function CheckExpirationAlerts(ws As Worksheet) As Long
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then
        CheckExpirationAlerts = 0
        Exit Function
    End If
    
    Dim alertCount As Long
    alertCount = 0
    
    Dim alertDays As Long
    alertDays = Val(Range("expiration_alert_days").Value)
    
    ' Check each position for expiration
    Dim i As Long
    For i = 2 To lastRow
        Dim expiryDate As Date
        Dim daysToExpiry As Long
        
        On Error Resume Next
        expiryDate = CDate(ws.Range("D" & i).Value) ' Assuming expiry date in column D
        On Error GoTo ErrorHandler
        
        If expiryDate > 0 Then
            daysToExpiry = WorksheetFunction.NetworkDays(Date, expiryDate)
            
            If daysToExpiry <= alertDays And daysToExpiry >= 0 Then
                ' Mark as alert
                ws.Range("K" & i).Value = "ALERT"
                ws.Range("L" & i).Value = daysToExpiry & " days"
                
                ' Highlight the row
                ws.Range("A" & i & ":L" & i).Interior.Color = RGB(255, 235, 156) ' Light yellow
                
                alertCount = alertCount + 1
            Else
                ws.Range("K" & i).Value = "OK"
                ws.Range("L" & i).Value = daysToExpiry & " days"
            End If
        End If
    Next i
    
    CheckExpirationAlerts = alertCount
    
    If alertCount > 0 Then
        MsgBox "WARNING: " & alertCount & " positions expire within " & _
               alertDays & " business days!", _
               vbExclamation, "Expiration Alert"
    End If
    
    Exit Function
    
ErrorHandler:
    CheckExpirationAlerts = 0
    MsgBox "Error checking expiration alerts: " & Err.description, vbCritical
End Function




Public Sub CreateExpirationEmail(alertCount As Long, csvPath As String)
    On Error GoTo ErrorHandler
    
    Dim objOutlook As Object
    Dim objMail As Object
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    
    ' Build email body with expiring positions table
    Dim emailBody As String
    emailBody = BuildExpirationEmailBody(alertCount)
    
    With objMail
        .To = Range("email_expiration_to").Value
        .CC = Range("email_expiration_cc").Value
        .subject = "Expiration Alert - " & Format(Date, "mm/dd/yyyy") & " - " & alertCount & " Positions"
        .htmlBody = emailBody
        
        If csvPath <> "" And Len(Dir(csvPath)) > 0 Then
            .Attachments.Add csvPath
        End If
        
        .Display ' Don't send automatically
    End With
    
    Set objMail = Nothing
    Set objOutlook = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set objMail = Nothing
    Set objOutlook = Nothing
End Sub




''''''''''''''''''''''''''''''''' dont need
''''''''''''''''''''''''''''''''' dont need
''''''''''''''''''''''''''''''''' dont need
''''''''''''''''''''''''''''''''' dont need


Public Function BuildExpirationEmailBodyBad(ws As Worksheet, alertCount As Long) As String
    Dim html As String
    html = "<html><body style='font-family: Arial, sans-serif;'>" & _
           "<h2 style='color: #003242;'>Synthetic Borrow Expiration Alert</h2>" & _
           "<p>The following " & alertCount & " positions are expiring within " & _
           Range("expiration_alert_days").Value & " business days:</p>" & _
           "<table style='border-collapse: collapse; width: 100%;'>" & _
           "<thead style='background-color: #003242; color: white;'>" & _
           "<tr>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Client</th>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Account</th>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Expiry Date</th>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Days Left</th>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Premium</th>" & _
           "<th style='border: 1px solid #ddd; padding: 8px;'>Payback</th>" & _
           "</tr></thead><tbody>"
    
    ' Add rows
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    ' Create screenshot table in email body
    Dim i As Long
    For i = 2 To lastRow
        If ws.Range("O" & i).Value = "ALERT" Then
            Dim rowColor As String
            If ws.Range("N" & i).Value <= 1 Then
                rowColor = "#ffcccc" ' Light red
            ElseIf ws.Range("N" & i).Value <= 2 Then
                rowColor = "#ffe6cc" ' Light orange
            Else
                rowColor = "#ffffcc" ' Light yellow
            End If
            
            html = html & "<tr style='background-color: " & rowColor & ";'>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("C" & i).Value & "</td>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("D" & i).Value & "</td>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("G" & i).Value, "mm/dd/yyyy") & "</td>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & ws.Range("N" & i).Value & "</td>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("I" & i).Value, "$#,##0") & "</td>" & _
                   "<td style='border: 1px solid #ddd; padding: 8px;'>" & Format(ws.Range("J" & i).Value, "$#,##0") & "</td>" & _
                   "</tr>"
        End If
    Next i
    
    html = html & "</tbody></table>" & _
           "<p style='margin-top: 20px;'>A detailed CSV file is attached with complete information.</p>" & _
           "<p>Please review and take appropriate action.</p>" & _
           "<p style='font-size: 10pt; color: #666;'>Generated by Synthetic Borrow Trading System</p>" & _
           "</body></html>"
    
    BuildExpirationEmailBodyBad = html
End Function


Public Sub SendExpirationAlert()
    On Error GoTo ErrorHandler
    
    ' Get expiring positions from ClientPortfolio worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ClientPortfolio")
    
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then Exit Sub
    
    ' Count alerts
    Dim alertCount As Long
    alertCount = 0
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Range("O" & i).Value = "ALERT" Then
            alertCount = alertCount + 1
        End If
    Next i
    
    If alertCount = 0 Then Exit Sub
    
    ' Create CSV attachment
    Dim csvPath As String
    csvPath = CreateExpirationCSV(ws)
    
    ' Build email body
    Dim emailBody As String
    emailBody = BuildExpirationEmailBody(ws, alertCount)
    
    ' Send email
    Call SendEmail( _
        Range("email_expiration_to").Value, _
        Range("email_expiration_cc").Value, _
        "Expiration Alert - " & Format(Date, "mm/dd/yyyy") & " - " & alertCount & " Positions", _
        emailBody, _
        csvPath _
    )
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error sending expiration alert: " & Err.description, vbCritical
End Sub

