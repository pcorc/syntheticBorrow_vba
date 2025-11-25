Attribute VB_Name = "modCompliance"

' ===================================================================
' COMPLETE BLOOMBERG VALIDATION WITH MARGIN CHECK
' ===================================================================
Public Sub RunBloombergValidationWithMarginCheck()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Ensure database connection
    If Not EnsureConnection() Then
        MsgBox "Cannot connect to database", vbCritical
        Exit Sub
    End If
    
    ' Reference worksheets
    Dim wsBbg As Worksheet
    Dim wsCompliance As Worksheet
    Dim wsMargin As Worksheet
    
    Set wsBbg = ThisWorkbook.Worksheets("BBG_Validation")
    Set wsMargin = ThisWorkbook.Worksheets("MarginVerification")
    Set wsCompliance = ThisWorkbook.Worksheets("Compliance")
    
    ' Setup
    Call CleanupDailyTables
    Call ExportMarginVerification
    Call AddComplianceHeaders(wsCompliance)
    
    Dim lastRowBBG As Long
    lastRowBBG = wsBbg.Range("A" & Rows.Count).End(xlUp).row
    
    If lastRowBBG < 2 Then
        MsgBox "No trades found in BBG_Validation sheet", vbExclamation
        GoTo Cleanup
    End If
    
    ' Process each trade
    Dim passCount As Long, failCount As Long
    passCount = 0
    failCount = 0
    
    Dim i As Long
    For i = 2 To lastRowBBG
        Dim tradeID As String
        tradeID = wsBbg.Range("A" & i).Value
        If tradeID = "" Then Exit For
        
        ' Process individual trade compliance check
        Dim tradeStatus As String
        tradeStatus = ProcessSingleTradeCompliance(i, wsBbg, wsMargin, wsCompliance)
        
        If tradeStatus = "APPROVED" Then
            passCount = passCount + 1
        Else
            failCount = failCount + 1
        End If
    Next i
    
    ' Format Compliance sheet
    wsCompliance.Columns.AutoFit
    
    ' Show summary message
    Dim totalTrades As Long
    totalTrades = lastRowBBG - 1
    
    
    ' ===================================================================
    ' KEEPING THIS CODE HERE FOR NOW, BUT WE MESSAGE BOX THE OVERALL SUMMARY LATER IN THE CODE
    ' ===================================================================
    
'    MsgBox "Bloomberg Validation with Margin Check Complete" & vbCrLf & vbCrLf & _
'           "Total Trades: " & totalTrades & vbCrLf & _
'           "Approved: " & passCount & vbCrLf & _
'           "Rejected: " & failCount & vbCrLf & vbCrLf & _
'           "Results saved to Compliance sheet and database.", vbInformation


    Call UpdateComplianceSummary
    wsCompliance.Activate
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error in Bloomberg validation with margin check: " & Err.description, vbCritical
End Sub

' ===================================================================
' PROCESS SINGLE TRADE COMPLIANCE
' ===================================================================

Private Function ProcessSingleTradeCompliance(rowNum As Long, _
                                              wsBbg As Worksheet, _
                                              wsMargin As Worksheet, _
                                              wsCompliance As Worksheet) As String
    ' Process compliance for a single trade
    ' Returns: "APPROVED" or "REJECTED"
    
    On Error GoTo ErrorHandler
    
    Dim tradeID As String, clientName As String, accountNum As String
    Dim marginStatus As String, bbgStatus As String, overallStatus As String
    Dim notes As String
    
    ' Get trade details from BBG_Validation
    tradeID = wsBbg.Range("A" & rowNum).Value
    clientName = wsBbg.Range("B" & rowNum).Value
    accountNum = wsBbg.Range("C" & rowNum).Value
    
    ' ============================================================
    ' STEP 1: Get Margin Status
    ' ============================================================
    marginStatus = GetMarginStatus(tradeID, wsMargin, notes)
    
    ' ============================================================
    ' STEP 2: Validate Bloomberg Rates
    ' ============================================================
    Dim expiryDate As Date
    Dim vestRate As Double, bbgRate As Double, tsyRate As Double
    Dim vestVsTsy As Double, bbgVsTsy As Double
    Dim maturityTolerance As Double
    
    expiryDate = wsBbg.Range("E" & rowNum).Value
    maturityTolerance = GetMaturityTolerance(expiryDate)
    
    vestRate = wsBbg.Range("K" & rowNum).Value   ' ICE Implied Rate
    bbgRate = wsBbg.Range("Q" & rowNum).Value    ' BBG Implied Rate
    tsyRate = wsBbg.Range("G" & rowNum).Value    ' Treasury Rate
    
    vestVsTsy = Abs(vestRate - tsyRate)
    bbgVsTsy = Abs(bbgRate - tsyRate)
    
    ' Check Bloomberg validation with maturity-based tolerance
    If bbgVsTsy <= maturityTolerance Then
        bbgStatus = "PASS"
    Else
        bbgStatus = "FAIL"
        If notes <> "" Then notes = notes & "; "
        notes = notes & "BBG validation failed (diff: " & Format(bbgVsTsy * 100, "0.00") & _
                "% exceeds " & Format(maturityTolerance * 100, "0.00") & "% tolerance)"
    End If
    
    ' ============================================================
    ' STEP 3: Determine Overall Status
    ' ============================================================
    If marginStatus = "PASS" And bbgStatus = "PASS" Then
        overallStatus = "APPROVED"
    Else
        overallStatus = "REJECTED"
    End If
    
    ' ============================================================
    ' STEP 4: Write to Compliance Sheet
    ' ============================================================
    wsCompliance.Range("A" & rowNum).Value = tradeID
    wsCompliance.Range("B" & rowNum).Value = clientName
    wsCompliance.Range("C" & rowNum).Value = accountNum
    wsCompliance.Range("D" & rowNum).Value = marginStatus
    wsCompliance.Range("E" & rowNum).Value = bbgStatus
    wsCompliance.Range("F" & rowNum).Value = Format(maturityTolerance * 100, "0.00") & "%"
    wsCompliance.Range("G" & rowNum).Value = overallStatus
    wsCompliance.Range("H" & rowNum).Value = notes
    wsCompliance.Range("I" & rowNum).Value = Range("trader").Value
    wsCompliance.Range("J" & rowNum).Value = Date
    
    ' Apply formatting
    If overallStatus = "APPROVED" Then
        wsCompliance.Range("A" & rowNum & ":J" & rowNum).Interior.Color = RGB(144, 238, 144) ' Light green
    Else
        wsCompliance.Range("A" & rowNum & ":J" & rowNum).Interior.Color = RGB(255, 182, 193) ' Light red
    End If
    
    ' ============================================================
    ' STEP 5: Export to Database
    ' ============================================================
    Dim todayStr As String
    todayStr = Format(Range("today").Value, "yyyy-mm-dd")
    
    Call ExportComplianceRecord(tradeID, clientName, accountNum, marginStatus, _
                               bbgStatus, vestVsTsy, bbgVsTsy, overallStatus, _
                               todayStr, Range("trader").Value, notes)
    
    ProcessSingleTradeCompliance = overallStatus
    Exit Function
    
ErrorHandler:
    ProcessSingleTradeCompliance = "ERROR"
    Debug.Print "Error processing trade " & tradeID & ": " & Err.description
End Function


Private Function GetMarginStatus(tradeID As String, _
                                 wsMargin As Worksheet, _
                                 ByRef notes As String) As String
    ' Returns: "PASS", "FAIL", or "PENDING"
    
    On Error GoTo ErrorHandler
    
    Dim marginRow As Long
    Dim j As Long
    Dim lastRowMargin As Long
    
    lastRowMargin = wsMargin.Range("A1").CurrentRegion.Rows.Count
    
    ' Find trade in MarginVerification sheet
    marginRow = 0
    For j = 2 To lastRowMargin
        If wsMargin.Range("A" & j).Value = tradeID Then
            marginRow = j
            Exit For
        End If
    Next j
    
    If marginRow > 0 Then
        ' Get margin status from column J
        Dim rawMarginStatus As String
        rawMarginStatus = Trim(UCase(wsMargin.Range("J" & marginRow).Value))
        
        If rawMarginStatus = "YES" Then
            GetMarginStatus = "PASS"
        ElseIf rawMarginStatus = "NO" Then
            GetMarginStatus = "FAIL"
            notes = "Margin check: FAIL"
        Else
            GetMarginStatus = "PENDING"
            notes = "Margin check: PENDING"
        End If
    Else
        ' Trade not found in MarginVerification sheet
        GetMarginStatus = "PENDING"
        notes = "Trade not found in MarginVerification sheet"
    End If
    
    Exit Function
    
ErrorHandler:
    GetMarginStatus = "ERROR"
    notes = "Error checking margin status"
End Function


Public Sub ExportComplianceRecord(tradeID As String, clientName As String, clientAccount As String, _
                                   marginCheck As String, bloombergCheck As String, _
                                   vestVsTsy As Double, bbgVsTsy As Double, overallStatus As String, _
                                   checkDate As String, checkedBy As String, note As String)
    On Error GoTo ErrorHandler
    
    ' Ensure database connection
    If Not modSQLConnections.EnsureConnection() Then Exit Sub
    
    Dim dbConn As ADODB.Connection
    Set dbConn = modSQLConnections.GetConnection()
    If dbConn Is Nothing Then Exit Sub
    
    ' Escape single quotes to prevent SQL injection
    tradeID = Replace(tradeID, "'", "''")
    clientName = Replace(clientName, "'", "''")
    clientAccount = Replace(clientAccount, "'", "''")
    marginCheck = Replace(marginCheck, "'", "''")
    bloombergCheck = Replace(bloombergCheck, "'", "''")
    overallStatus = Replace(overallStatus, "'", "''")
    checkedBy = Replace(checkedBy, "'", "''")
    note = Replace(note, "'", "''")
    
    ' Build INSERT statement with ON DUPLICATE KEY UPDATE
    Dim sql As String
    
    ' Build INSERT portion
    sql = "INSERT INTO synthetic_borrow_compliance "
    sql = sql & "(synthetic_borrow_app_id, client_name, client_account, margin_check, "
    sql = sql & "bloomberg_check, vest_vs_tsy, bbg_vs_tsy, overall_status, check_date, checked_by, note) "
    sql = sql & "VALUES ("
    sql = sql & "'" & tradeID & "', "
    sql = sql & "'" & clientName & "', "
    sql = sql & "'" & clientAccount & "', "
    sql = sql & "'" & marginCheck & "', "
    sql = sql & "'" & bloombergCheck & "', "
    sql = sql & vestVsTsy & ", "
    sql = sql & bbgVsTsy & ", "
    sql = sql & "'" & overallStatus & "', "
    sql = sql & "'" & checkDate & "', "
    sql = sql & "'" & checkedBy & "', "
    sql = sql & "'" & note & "') "
    
    ' Build ON DUPLICATE KEY UPDATE portion
    sql = sql & "ON DUPLICATE KEY UPDATE "
    sql = sql & "client_name = '" & clientName & "', "
    sql = sql & "client_account = '" & clientAccount & "', "
    sql = sql & "margin_check = '" & marginCheck & "', "
    sql = sql & "bloomberg_check = '" & bloombergCheck & "', "
    sql = sql & "vest_vs_tsy = " & vestVsTsy & ", "
    sql = sql & "bbg_vs_tsy = " & bbgVsTsy & ", "
    sql = sql & "overall_status = '" & overallStatus & "', "
    sql = sql & "check_date = '" & checkDate & "', "
    sql = sql & "checked_by = '" & checkedBy & "', "
    sql = sql & "note = '" & note & "'"
    
    ' Execute insert
    dbConn.Execute sql
    
    Exit Sub
    
ErrorHandler:
    ' Log error but don't stop processing
    Debug.Print "Error exporting compliance record for trade " & tradeID & ": " & Err.description
    MsgBox "Error exporting compliance record: " & Err.description & vbCrLf & _
           "Trade ID: " & tradeID, vbExclamation
End Sub


Public Function GetMaturityTolerance(expiryDate As Date) As Double
    On Error GoTo ErrorHandler
    
    Dim todayDate As Date
    todayDate = Range("today").Value
    
    Dim monthsToMaturity As Long
    monthsToMaturity = DateDiff("m", todayDate, expiryDate)
    
    ' Get starting cell for threshold table
    Dim startCell As Range
    Set startCell = Range("treasury_threshold_start")
    
    ' Loop through rows until blank cell in column 1 (Max_Months)
    Dim i As Long
    i = 2 ' Start at row 2 (skip header)
    
    Do While startCell.Offset(i - 1, 0).Value <> ""
        Dim maxMonths As Long
        Dim threshold As Double
        
        maxMonths = startCell.Offset(i - 1, 0).Value    ' Column M (Max_Months)
        threshold = startCell.Offset(i - 1, 1).Value    ' Column N (Threshold)
        
        If monthsToMaturity <= maxMonths Then
            GetMaturityTolerance = threshold
            Exit Function
        End If
        
        i = i + 1
    Loop
    
    ' Default if not found (shouldn't happen if you have 9999 row)
    GetMaturityTolerance = 0.005
    Exit Function
    
ErrorHandler:
    GetMaturityTolerance = 0.005 ' Default fallback
End Function




Public Sub UpdateComplianceSummary()
    ' ===================================================================
    ' COMPLIANCE SUMMARY UPDATE ROUTINE
    ' Called at the end of RunBloombergValidationWithMarginCheck()
    ' Analyzes compliance results and updates Control tab summary
    ' ===================================================================
    
    On Error GoTo ErrorHandler
    
    Dim wsCompliance As Worksheet
    Dim wsControl As Worksheet
    Dim wsPortfolio As Worksheet
    
    Set wsCompliance = ThisWorkbook.Worksheets("Compliance")
    Set wsControl = ThisWorkbook.Worksheets("Control")
    Set wsPortfolio = ThisWorkbook.Worksheets("ClientPortfolio")
    
    ' Initialize counters
    Dim totalTrades As Long
    Dim approvedTrades As Long
    Dim rejectedTrades As Long
    Dim marginFailures As Long
    Dim bbgFailures As Long
    Dim pendingTrades As Long
    Dim expirationAlerts As Long
    
    ' Find last row in Compliance sheet
    Dim lastRow As Long
    lastRow = wsCompliance.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then
        ' No compliance data - clear summary
        Call ClearComplianceSummary(wsControl)
        Exit Sub
    End If
    
    ' Loop through compliance results and count statuses
    Dim i As Long
    For i = 2 To lastRow
        If wsCompliance.Range("A" & i).Value <> "" Then
            totalTrades = totalTrades + 1
            
            Dim overallStatus As String
            Dim marginStatus As String
            Dim bbgStatus As String
            
            marginStatus = Trim(UCase(wsCompliance.Range("D" & i).Value))
            bbgStatus = Trim(UCase(wsCompliance.Range("E" & i).Value))
            overallStatus = Trim(UCase(wsCompliance.Range("G" & i).Value))
            
            ' Count overall status
            If overallStatus = "APPROVED" Then
                approvedTrades = approvedTrades + 1
            ElseIf overallStatus = "REJECTED" Then
                rejectedTrades = rejectedTrades + 1
            End If
            
            ' Count specific failures
            If marginStatus = "FAIL" Then
                marginFailures = marginFailures + 1
            ElseIf marginStatus = "PENDING" Then
                pendingTrades = pendingTrades + 1
            End If
            
            If bbgStatus = "FAIL" Then
                bbgFailures = bbgFailures + 1
            End If
        End If
    Next i
    
    ' Check for expiration alerts in Client Portfolio
    expirationAlerts = CountExpirationAlerts(wsPortfolio)
    
    ' Update Control tab summary
    Range("compliance_total_trades").Value = totalTrades
    Range("compliance_approved_trades").Value = approvedTrades
    Range("compliance_rejected_trades").Value = rejectedTrades
    Range("compliance_margin_failures").Value = marginFailures
    Range("compliance_bbg_failures").Value = bbgFailures
    Range("compliance_pending_trades").Value = pendingTrades
    
    ' Determine overall status
    Dim overallPass As Boolean
    overallPass = (rejectedTrades = 0 And pendingTrades = 0)
    
    If overallPass Then
        Range("compliance_overall_status").Value = "PASS"
        Range("compliance_overall_status").Interior.Color = RGB(144, 238, 144) ' Light green
    Else
        Range("compliance_overall_status").Value = "FAIL"
        Range("compliance_overall_status").Interior.Color = RGB(255, 99, 71) ' Tomato red
    End If
    
    ' Build summary message
    Dim summaryMsg As String
    summaryMsg = "=== COMPLIANCE SUMMARY ===" & vbCrLf & vbCrLf
    summaryMsg = summaryMsg & "Total Trades Processed: " & totalTrades & vbCrLf
    summaryMsg = summaryMsg & "Approved: " & approvedTrades & vbCrLf
    summaryMsg = summaryMsg & "Rejected: " & rejectedTrades & vbCrLf
    
    If pendingTrades > 0 Then
        summaryMsg = summaryMsg & vbCrLf & "?? WARNING: " & pendingTrades & " trade(s) pending verification" & vbCrLf
    End If
    
    If marginFailures > 0 Then
        summaryMsg = summaryMsg & vbCrLf & "? MARGIN FAILURES: " & marginFailures & " trade(s) failed margin verification" & vbCrLf
    End If
    
    If bbgFailures > 0 Then
        summaryMsg = summaryMsg & vbCrLf & "? BLOOMBERG FAILURES: " & bbgFailures & " trade(s) exceeded pricing tolerance" & vbCrLf
    End If
    
    If expirationAlerts > 0 Then
        summaryMsg = summaryMsg & vbCrLf & "? EXPIRATION ALERT: " & expirationAlerts & " position(s) expiring within " & _
                     Range("expiration_alert_days").Value & " days" & vbCrLf
    End If
    
    summaryMsg = summaryMsg & vbCrLf & "================================" & vbCrLf
    
    If overallPass Then
        summaryMsg = summaryMsg & vbCrLf & "? STATUS: ALL CHECKS PASSED" & vbCrLf & vbCrLf
        summaryMsg = summaryMsg & "All trades approved for execution."
    Else
        summaryMsg = summaryMsg & vbCrLf & "?? STATUS: COMPLIANCE ISSUES DETECTED" & vbCrLf & vbCrLf
        summaryMsg = summaryMsg & "Review Compliance sheet for details before proceeding."
    End If
    
    ' Display summary message
    If overallPass Then
        MsgBox summaryMsg, vbInformation, "Compliance Summary - PASSED"
    Else
        MsgBox summaryMsg, vbCritical, "Compliance Summary - ATTENTION REQUIRED"
    End If
    
    ' Optionally send email alert if there are failures
'    If Not overallPass Then
'        Dim sendAlert As VbMsgBoxResult
'        sendAlert = MsgBox("Would you like to send an alert email to the compliance team?", _
'                          vbYesNo + vbQuestion, "Send Alert")
'        If sendAlert = vbYes Then
'            Call SendComplianceAlert(summaryMsg)
'        End If
'    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error updating compliance summary: " & Err.description, vbCritical
End Sub


Public Function CountExpirationAlerts(ws As Worksheet) As Long
    ' Count positions expiring within alert threshold
    On Error Resume Next
    
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then
        CountExpirationAlerts = 0
        Exit Function
    End If
    
    Dim alertCount As Long
    alertCount = 0
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Range("K" & i).Value = "ALERT" Then
            alertCount = alertCount + 1
        End If
    Next i
    
    CountExpirationAlerts = alertCount
End Function


Public Sub ClearComplianceSummary(wsControl As Worksheet)
    ' Clear all summary fields
    On Error Resume Next
    
    Range("compliance_total_trades").Value = 0
    Range("compliance_approved_trades").Value = 0
    Range("compliance_rejected_trades").Value = 0
    Range("compliance_margin_failures").Value = 0
    Range("compliance_bbg_failures").Value = 0
    Range("compliance_pending_trades").Value = 0
    Range("compliance_overall_status").Value = ""
    Range("compliance_status_indicator").Interior.colorIndex = xlNone
End Sub



Public Function OrderGenerationComplianceChecks() As Boolean
    ' Centralized compliance validation for order generation
    ' Returns True if all checks pass, False otherwise
    
    On Error GoTo ErrorHandler
    
    OrderGenerationComplianceChecks = False ' Default to fail
    
    ' ============================================================
    ' CHECK 1: Compliance worksheet exists
    ' ============================================================
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Compliance")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "Compliance worksheet not found." & vbCrLf & vbCrLf & _
               "Please run Compliance first.", _
               vbExclamation, "Missing Compliance Data"
        Exit Function
    End If
    
    ' ============================================================
    ' CHECK 2: Master account configuration
    ' ============================================================
    Dim masterAccount As String
    On Error Resume Next
    masterAccount = CStr(Range("vest_master_account").Value)
    On Error GoTo ErrorHandler
    
    If masterAccount <> "8471863" Then
        MsgBox "Invalid master account configuration." & vbCrLf & _
               "Expected: 8471863" & vbCrLf & _
               "Found: " & masterAccount & vbCrLf & vbCrLf & _
               "Please verify the 'vest_master_account' named range.", _
               vbCritical, "Configuration Error"
        Exit Function
    End If
    
    ' ============================================================
    ' CHECK 3: Overall compliance status
    ' ============================================================
    Dim complianceStatus As String
    On Error Resume Next
    complianceStatus = UCase(Trim(CStr(Range("compliance_overall_status").Value)))
    On Error GoTo ErrorHandler
    
    If complianceStatus <> "PASS" Then
        MsgBox "Compliance status must be 'PASS' before processing orders." & vbCrLf & _
               "Current status: " & complianceStatus & vbCrLf & vbCrLf & _
               "Please complete compliance checks before proceeding.", _
               vbExclamation, "Compliance Check Failed"
        Exit Function
    End If
    
    ' ============================================================
    ' CHECK 4: Compliance data exists
    ' ============================================================
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then
        MsgBox "No trades found in Compliance sheet." & vbCrLf & vbCrLf & _
               "Please run Compliance first.", _
               vbExclamation, "No Compliance Data"
        Exit Function
    End If
    
    ' ============================================================
    ' CHECK 5: At least one approved trade exists
    ' ============================================================
    Dim approvedCount As Long
    approvedCount = 0
    
    Dim i As Long
    For i = 2 To lastRow
        If UCase(Trim(ws.Range("G" & i).Value)) = "APPROVED" Then
            approvedCount = approvedCount + 1
        End If
    Next i
    
    If approvedCount = 0 Then
        MsgBox "No approved trades found." & vbCrLf & vbCrLf & _
               "All trades failed compliance checks. Please review.", _
               vbExclamation, "No Approved Trades"
        Exit Function
    End If
    
    ' ============================================================
    ' ALL CHECKS PASSED
    ' ============================================================
    OrderGenerationComplianceChecks = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error during compliance checks: " & Err.description, vbCritical, "Compliance Check Error"
    OrderGenerationComplianceChecks = False
End Function







'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''' DONT USE


Public Sub SendComplianceAlert(summaryMsg As String)
    ' Send email alert about compliance failures
    On Error GoTo ErrorHandler
    
    Dim emailBody As String
    emailBody = "<div style='font-family: Arial, sans-serif; font-size: 11pt;'>" & _
               "<h2 style='color: #dc3545;'>?? Compliance Alert - Action Required</h2>" & _
               "<p>Compliance issues detected in today's trading:</p>" & _
               "<pre style='background-color: #f8f9fa; padding: 15px; border-left: 3px solid #dc3545;'>" & _
               summaryMsg & _
               "</pre>" & _
               "<p><strong>Action Required:</strong> Review the Compliance sheet and resolve issues before proceeding with execution.</p>" & _
               "<p style='font-size: 9pt; color: #666;'>Generated: " & Format(Now(), "mm/dd/yyyy hh:mm AM/PM") & "</p>" & _
               "</div>"
    
    Call modEmailAlerts.SendEmail( _
        Range("email_margin_to").Value, _
        Range("email_margin_cc").Value, _
        "?? COMPLIANCE ALERT - " & Format(Range("today").Value, "mm/dd/yyyy"), _
        emailBody, _
        "" _
    )
    
    MsgBox "Compliance alert sent to team", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error sending compliance alert: " & Err.description, vbExclamation
End Sub




Public Sub ClearComplianceProcessingSheets()
    ' ===================================================================
    ' CLEAR COMPLIANCE AND PROCESSING SHEETS
    ' Called at the start of RunBloombergValidationWithMarginCheck()
    ' Prevents stacking results from multiple runs on same day
    ' ===================================================================
    
    On Error Resume Next
    
    Application.ScreenUpdating = False
    
    Dim wsCompliance As Worksheet
    Dim wsOrderGen As Worksheet
    Dim wsExecution As Worksheet
    
    ' Reference worksheets
    Set wsCompliance = ThisWorkbook.Worksheets("Compliance")
    Set wsOrderGen = ThisWorkbook.Worksheets("OrderGen")
    Set wsExecution = ThisWorkbook.Worksheets("Execution Results")
    
    ' ==================================================================
    ' CLEAR COMPLIANCE SHEET
    ' ==================================================================
    If Not wsCompliance Is Nothing Then
        ' Clear all data rows (keep headers in row 3)
        wsCompliance.Range("A4:Z10000").ClearContents
        wsCompliance.Range("A4:Z10000").ClearFormats
        wsCompliance.Range("A4:Z10000").Interior.colorIndex = xlNone
        
        ' Optional: Clear summary area if you have one
        wsCompliance.Range("AA4:AZ10000").ClearContents
    End If
    
    ' ==================================================================
    ' CLEAR ORDER GENERATION SHEET
    ' ==================================================================
    If Not wsOrderGen Is Nothing Then
        ' Clear trade processing area (keep headers in row 1)
        wsOrderGen.Range("A2:AM10000").ClearContents
        wsOrderGen.Range("A2:AM10000").ClearFormats
        wsOrderGen.Range("A2:AM10000").Interior.colorIndex = xlNone
        
        ' Clear blocked orders summary area (columns AO onwards)
        wsOrderGen.Range("AO4:BA10000").ClearContents
        wsOrderGen.Range("AO4:BA10000").ClearFormats
        wsOrderGen.Range("AO4:BA10000").Interior.colorIndex = xlNone
    End If
    
    ' ==================================================================
    ' CLEAR EXECUTION RESULTS SHEET
    ' ==================================================================
    If Not wsExecution Is Nothing Then
        ' Clear all execution data (keep headers in row 3)
        wsExecution.Range("A4:Z10000").ClearContents
        wsExecution.Range("A4:Z10000").ClearFormats
        wsExecution.Range("A4:Z10000").Interior.colorIndex = xlNone
    End If
    
    ' ==================================================================
    ' CLEAR COMPLIANCE SUMMARY ON CONTROL TAB
    ' ==================================================================
    Call ClearComplianceSummary(ThisWorkbook.Worksheets("Control"))
    
    Application.ScreenUpdating = True
    
    On Error GoTo 0
End Sub





