Attribute VB_Name = "modUtilities"

' ===================================================================
' UTILITIES MODULE - Version 2
' Synthetic Borrow Trading System
' ===================================================================

Option Explicit


Public Sub ImportAndAllocateExecutions()
    ' Combined execution import and allocation
    On Error GoTo ErrorHandler
    
    ' First import execution results
    Call ImportExecutionResults
    
    ' Ask if user wants to generate allocations
    Dim response As VbMsgBoxResult
    response = MsgBox("Execution results imported. Generate allocation file now?", _
                     vbYesNo + vbQuestion, "Generate Allocations")
    
    If response = vbYes Then
        Call GenerateAllocationFile
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in execution processing: " & Err.description, vbCritical
End Sub



' ===================================================================
' VALIDATION HELPERS
' ===================================================================

Public Function ValidateDataExists() As Boolean
    ' Check if we have data to process
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("RawTradeImport")
    
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    ValidateDataExists = (lastRow > 1)
    Exit Function
    
ErrorHandler:
    ValidateDataExists = False
End Function




' ===================================================================
' HEADER CREATION FUNCTIONS
' ===================================================================

Public Sub AddTradeSubmissionHeaders(ws As Worksheet)
    ' Clear and recreate headers in row 1
    ws.Range("1:1").ClearContents
    
    Dim headers As Variant
    headers = Array("synthetic_borrow_app_id", "Client Name", "Account", "Email", "Buying Power", _
                   "Requested Amount", "Quoted Borrow", "Payback Amount", "Rate", "Limit %", _
                   "Expiry Date", "Request Time", "Execution Time", _
                   "L1 Action", "L1 Qty", "L1 Type", "L1 Strike", "L1 Price", "L1 Notional", _
                   "L2 Action", "L2 Qty", "L2 Type", "L2 Strike", "L2 Price", "L2 Notional", _
                   "L3 Action", "L3 Qty", "L3 Type", "L3 Strike", "L3 Price", "L3 Notional", _
                   "L4 Action", "L4 Qty", "L4 Type", "L4 Strike", "L4 Price", "L4 Notional", _
                   "User ID", "User Email", "Created At", "Updated At")
    
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    ws.Range("A1").Resize(1, UBound(headers) + 1).Font.Bold = True
    
    ' Format headers
    With ws.Range("A1").Resize(1, UBound(headers) + 1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196) ' Blue background
        .Font.Color = RGB(255, 255, 255) ' White text
        .HorizontalAlignment = xlCenter
    End With
    
End Sub

Public Sub AddMarginVerificationHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("Trade ID", "Client Name", "Account", "Email", "Buying Power", _
                   "Requested Amount", "Quoted Borrow", "Payback Amount", "Rate", _
                   "Margin Status", "Notes")
    
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    ' Format headers
    With ws.Range("A1").Resize(1, UBound(headers) + 1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196) ' Blue background
        .Font.Color = RGB(255, 255, 255) ' White text
        .HorizontalAlignment = xlCenter
    End With
End Sub

Public Sub AddPortfolioHeaders(ws As Worksheet)
    ' Clear and recreate headers
    ws.Range("1:1").ClearContents
    
    Dim headers As Variant
    headers = Array("ID", "Synthetic Borrow Trade ID", "Client Name", "Account", "Email", _
                   "Execution Date", "Expiry Date", "Box Structure", "Premium", _
                   "Payback", "Rate", "System ID", "Days to Expiry", "Alert Status")
    
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    ws.Range("A1").Resize(1, UBound(headers) + 1).Font.Bold = True
    
    ' Format headers
    With ws.Range("A1").Resize(1, UBound(headers) + 1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196) ' Blue background
        .Font.Color = RGB(255, 255, 255) ' White text
        .HorizontalAlignment = xlCenter
    End With
End Sub



Public Sub AddOrderTrackingHeaders(ws As Worksheet)
    ' Clear and recreate headers
    ws.Range("1:1").ClearContents
    
    Dim headers As Variant
    headers = Array("ID", "Synthetic Borrow Trade ID", "Client Name", "Account", "Email", _
                   "Execution Date", "Expiry Date", "Box Structure", "Premium", _
                   "Payback", "Rate", "System ID", "Days to Expiry", "Alert Status")
    
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    ws.Range("A1").Resize(1, UBound(headers) + 1).Font.Bold = True
    
    ' Format headers
    With ws.Range("A1").Resize(1, UBound(headers) + 1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196) ' Blue background
        .Font.Color = RGB(255, 255, 255) ' White text
        .HorizontalAlignment = xlCenter
    End With
End Sub


Public Sub AddComplianceHeaders(wsCompliance As Worksheet)
    ' Setup Compliance sheet headers with formatting
    Dim headers As Variant
    headers = Array("Synthetic Borrow App ID", "Client Name", "Account", _
                   "Margin Status", "Bloomberg Check", "Tolerance Used", _
                   "Overall Status", "Notes", "Checked By", "Check Date")
    
    ' Add headers to row 1
    wsCompliance.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    ' Format headers
    With wsCompliance.Range("A1").Resize(1, UBound(headers) + 1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196) ' Blue background
        .Font.Color = RGB(255, 255, 255) ' White text
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Auto-fit columns
    wsCompliance.Columns("A:J").AutoFit
End Sub


Public Sub SetupOrderGenHeaders(ws As Worksheet)
    ws.Range("A1:O1").Value = Array("#ACCOUNT", "ORDER TYPE", "LIMIT PRICE", _
                                    "LEG 1 SYMBOL", "LEG 1 QUANTITY", "LEG 1 OPEN CLOSE", _
                                    "LEG 2 SYMBOL", "LEG 2 QUANTITY", "LEG 2 OPEN CLOSE", _
                                    "LEG 3 SYMBOL", "LEG 3 QUANTITY", "LEG 3 OPEN CLOSE", _
                                    "LEG 4 SYMBOL", "LEG 4 QUANTITY", "LEG 4 OPEN CLOSE")
    ws.Range("A1:O1").Font.Bold = True
End Sub


Public Sub AddExecutionHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("Allocation ID", "Trading System ID", "Execution Date", _
                   "Quantity", "Leg 1 Price", "Leg 2 Price", "Leg 3 Price", _
                   "Leg 4 Price", "Net Premium", "Status", "Import Time")
    
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    ' Format headers
    With ws.Range("A1").Resize(1, UBound(headers) + 1)
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 80) ' Green background
        .Font.Color = RGB(255, 255, 255) ' White text
        .HorizontalAlignment = xlCenter
    End With
End Sub



' ===================================================================
' FORMATTING FUNCTIONS
' ===================================================================

Public Sub FormatTradeSubmissionData(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then Exit Sub
    
    ' Format currency columns
    ws.Range("E2:H" & lastRow).NumberFormat = "$#,##0.00"
    ws.Range("I2:I" & lastRow).NumberFormat = "0.00%"
    ws.Range("J2:J" & lastRow).NumberFormat = "0.00%"
    
    ' Format date columns
    ws.Range("K2:K" & lastRow).NumberFormat = "mm/dd/yyyy"
    ws.Range("L2:M" & lastRow).NumberFormat = "mm/dd/yyyy hh:mm"
    ws.Range("AI2:AJ" & lastRow).NumberFormat = "mm/dd/yyyy hh:mm"
    
    ' Format option prices
    ws.Range("R2:R" & lastRow & ",S2:S" & lastRow).NumberFormat = "$#,##0.00"
    ws.Range("X2:X" & lastRow & ",Y2:Y" & lastRow).NumberFormat = "$#,##0.00"
    ws.Range("AD2:AD" & lastRow & ",AE2:AE" & lastRow).NumberFormat = "$#,##0.00"
    ws.Range("AJ2:AJ" & lastRow & ",AK2:AK" & lastRow).NumberFormat = "$#,##0.00"
    
End Sub


Public Sub FormatPortfolioData(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then Exit Sub
    
    ' Format currency and percentage columns
    ws.Range("I2:J" & lastRow).NumberFormat = "$#,##0.00"
    ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
    ws.Range("F2:G" & lastRow).NumberFormat = "mm/dd/yyyy"
    
    ' Add alert status and conditional formatting
    Dim alertDays As Integer
    alertDays = Val(Range("expiration_alert_days").Value)
    
    Dim i As Long
    For i = 2 To lastRow
        Dim daysToExpiry As Long
        daysToExpiry = ws.Range("M" & i).Value
        
        If daysToExpiry <= alertDays Then
            ws.Range("O" & i).Value = "ALERT"
            ' Apply urgency colors
            If daysToExpiry <= 1 Then
                ws.Range("A" & i & ":O" & i).Interior.Color = RGB(255, 0, 0) ' Red
            ElseIf daysToExpiry <= 2 Then
                ws.Range("A" & i & ":O" & i).Interior.Color = RGB(255, 165, 0) ' Orange
            Else
                ws.Range("A" & i & ":O" & i).Interior.Color = RGB(255, 255, 0) ' Yellow
            End If
        Else
            ws.Range("O" & i).Value = "OK"
        End If
    Next i
End Sub


Public Sub FormatOrderExport(ws As Worksheet)
    ws.Columns.AutoFit
    ws.Range("D:D").NumberFormat = "0.00%"
    
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    ws.Range("A1:P" & lastRow).Borders.LineStyle = xlContinuous
End Sub


Public Sub FormatBBGValidationHeaders()
    On Error Resume Next
    
    Dim wsBbg As Worksheet
    Set wsBbg = ThisWorkbook.Worksheets("BBG_Validation")
    
    If wsBbg Is Nothing Then Exit Sub
    
    ' Find last column with headers
    Dim lastCol As Long
    lastCol = wsBbg.Cells(1, wsBbg.Columns.Count).End(xlToLeft).Column
    
    If lastCol < 1 Then Exit Sub ' No headers found
    
    ' Define colors for each group
    Dim colorGeneral As Long      ' For non-leg columns
    Dim colorLeg1 As Long         ' For Leg 1 columns
    Dim colorLeg2 As Long         ' For Leg 2 columns
    Dim colorLeg3 As Long         ' For Leg 3 columns
    Dim colorLeg4 As Long         ' For Leg 4 columns
    Dim colorBBG As Long          ' For BBG columns
    Dim colorICE As Long          ' For ICE columns
    
    ' Set distinct colors (RGB values)
    colorGeneral = RGB(68, 114, 196)   ' Blue - General columns
    colorLeg1 = RGB(112, 173, 71)      ' Green - Leg 1
    colorLeg2 = RGB(237, 125, 49)      ' Orange - Leg 2
    colorLeg3 = RGB(165, 105, 189)     ' Purple - Leg 3
    colorLeg4 = RGB(255, 192, 0)       ' Gold - Leg 4
    colorBBG = RGB(192, 0, 0)          ' Dark Red - BBG columns
    colorICE = RGB(0, 176, 240)        ' Cyan - ICE columns
    
    ' Loop through each header column and apply formatting
    Dim col As Long
    Dim headerText As String
    
    For col = 1 To lastCol
        headerText = Trim(wsBbg.Cells(1, col).Value)
        
        ' Apply formatting based on column content
        With wsBbg.Cells(1, col)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255) ' White text
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            
            ' Determine which color to apply based on header text
            ' Check BBG and ICE first (higher priority)
            If InStr(1, headerText, "BBG", vbTextCompare) > 0 Then
                .Interior.Color = colorBBG
            ElseIf InStr(1, headerText, "ICE", vbTextCompare) > 0 Then
                .Interior.Color = colorICE
            ElseIf InStr(1, headerText, "Leg 1", vbTextCompare) > 0 Then
                .Interior.Color = colorLeg1
            ElseIf InStr(1, headerText, "Leg 2", vbTextCompare) > 0 Then
                .Interior.Color = colorLeg2
            ElseIf InStr(1, headerText, "Leg 3", vbTextCompare) > 0 Then
                .Interior.Color = colorLeg3
            ElseIf InStr(1, headerText, "Leg 4", vbTextCompare) > 0 Then
                .Interior.Color = colorLeg4
            Else
                .Interior.Color = colorGeneral
            End If
        End With
    Next col
    
    ' Auto-fit columns
    wsBbg.Columns.AutoFit

    Set wsBbg = Nothing
    
    On Error GoTo 0
End Sub



' ===================================================================
' DATE UTILITIES
' ===================================================================

Public Function GetTenorInMonths(expiryDate As Date) As Long
    ' Calculate tenor in months from today
    GetTenorInMonths = DateDiff("m", Date, expiryDate)
End Function

Public Function GetBusinessDaysToExpiry(expiryDate As Date) As Long
    ' Calculate business days to expiry
    Dim currentDate As Date
    currentDate = Date
    
    Dim businessDays As Long
    businessDays = 0
    
    Do While currentDate < expiryDate
        If Weekday(currentDate) <> vbSaturday And Weekday(currentDate) <> vbSunday Then
            businessDays = businessDays + 1
        End If
        currentDate = currentDate + 1
    Loop
    
    GetBusinessDaysToExpiry = businessDays
End Function


Public Sub SendExecutionSummaryEmail()
    ' Wrapper to send execution summary
    Call modEmailAlerts.SendExecutionSummary
End Sub

' ===================================================================
' ERROR LOGGING
' ===================================================================

Public Sub LogError(moduleName As String, procedureName As String, errorDesc As String)
    On Error Resume Next
    
    ' Log to immediate window
    Debug.Print Now() & " | " & moduleName & "." & procedureName & " | " & errorDesc
    
    ' Optionally log to a sheet
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets("ErrorLog")
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add
        wsLog.Name = "ErrorLog"
        wsLog.Range("A1:D1").Value = Array("Timestamp", "Module", "Procedure", "Error")
        wsLog.Range("A1:D1").Font.Bold = True
    End If
    
    Dim nextRow As Long
    nextRow = wsLog.Range("A1").CurrentRegion.Rows.Count + 1
    
    wsLog.Range("A" & nextRow).Value = Now()
    wsLog.Range("B" & nextRow).Value = moduleName
    wsLog.Range("C" & nextRow).Value = procedureName
    wsLog.Range("D" & nextRow).Value = errorDesc
End Sub






''''''''''''''''''''''''''''''''''''''''''''''''''''''''' DON THINK THIS IS USED

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ===================================================================
' HELPER FUNCTIONS
' ===================================================================

Public Sub AddOrderExportHeaders(ws As Worksheet)
    ws.Range("A1:P1").Value = Array("#ORDER IDS", "#ACCOUNT", "ORDER TYPE", "LIMIT PRICE", _
                                    "LEG 1 SYMBOL", "LEG 1 QUANTITY", "LEG 1 OPEN CLOSE", _
                                    "LEG 2 SYMBOL", "LEG 2 QUANTITY", "LEG 2 OPEN CLOSE", _
                                    "LEG 3 SYMBOL", "LEG 3 QUANTITY", "LEG 3 OPEN CLOSE", _
                                    "LEG 4 SYMBOL", "LEG 4 QUANTITY", "LEG 4 OPEN CLOSE")
    ws.Range("A1:P1").Font.Bold = True
End Sub

Public Sub AddExecutionHeadersOld(ws As Worksheet)
    ' Clear and add headers based on execution file format
    ws.Range("1:1").ClearContents
    ws.Range("A1:Z1").Value = Array("Order ID", "Trading System ID", "Client Account", _
                                    "Executed Premium", "Payback Amount", "Annualized Rate", _
                                    "Leg1 Price", "Leg2 Price", "Leg3 Price", "Leg4 Price", _
                                    "Leg1 Action", "Leg1 Qty", "Leg1 Type", "Leg1 Strike", _
                                    "Leg2 Action", "Leg2 Qty", "Leg2 Type", "Leg2 Strike", _
                                    "Leg3 Action", "Leg3 Qty", "Leg3 Type", "Leg3 Strike", _
                                    "Leg4 Action", "Leg4 Qty", "Leg4 Type", "Leg4 Strike")
    ws.Range("A1:Z1").Font.Bold = True
End Sub

