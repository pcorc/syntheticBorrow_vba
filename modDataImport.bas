Attribute VB_Name = "modDataImport"

' ===================================================================
' DATA IMPORT MODULE - Version 2
' Synthetic Borrow Trading System
' ===================================================================

Option Explicit

' ===================================================================
' MAIN IMPORT ORCHESTRATOR
' ===================================================================

Public Sub ImportAllData()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Ensure database connection ONCE at the start
    If Not EnsureConnection() Then
        MsgBox "Cannot connect to database", vbCritical
        GoTo Cleanup
    End If
    
    ' Clear data sheets before import (preserving headers and formulas)
    Call ClearAndSetupSheets
    
    ' Phase 1: Import all data silently
    Call ImportTradeSubmissions
    Call ImportMarginVerificationStatus
    Call ImportClientPortfolio
    Call CalculateTenorForBBGValidation

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Show import completion with expiration check notice
'    MsgBox "Data import complete" & vbCrLf & vbCrLf & _
'           "The system will now check for upcoming position expirations." & vbCrLf & _
'           "An email alert will be generated if any positions are expiring soon.", _
'           vbInformation, "Import Complete"
    
    ' Check for expiring positions and send alerts if needed
    Call CheckAndAlertExpirations
    
    ' Navigate to BBG_Validation tab and show completion message
    ThisWorkbook.Worksheets("BBG_Validation").Activate
    
    MsgBox "Synthetic Borrow Setup is complete", vbInformation, "Setup Complete"
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error during data import: " & Err.description, vbCritical, "Import Error"
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


' ===================================================================
' CLEAR IMPORT SHEETS (called at start of import)
' ===================================================================

Public Sub ClearImportSheets()
    On Error Resume Next
    
    ' Clear sheets that get populated by import (not formula sheets)
    Dim sheetsToClean As Variant
    sheetsToClean = Array("RawTradeImport", "ClientPortfolio", "Compliance", _
                         "OrderGen", "ExecutionResults")
    
    Dim sheetName As Variant
    For Each sheetName In sheetsToClean
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
        If Not ws Is Nothing Then
            ' Clear all data but preserve row 1 (headers) if they exist
            ws.Range("A2:AZ10000").ClearContents
            ws.Range("A2:AZ10000").Interior.colorIndex = xlNone ' Clear formatting too
        End If
        Set ws = Nothing
    Next sheetName
    
    On Error GoTo 0
End Sub


Public Sub ClearAndSetupSheets()
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Check vest_master_account named range
    Dim masterAccount As String
    masterAccount = CStr(Range("vest_master_account").Value)
    
    If masterAccount <> "8471863" Then
        MsgBox "Invalid master account configuration." & vbCrLf & _
               "Expected: 8471863" & vbCrLf & _
               "Found: " & masterAccount & vbCrLf & vbCrLf & _
               "Please verify the 'vest_master_account' named range before proceeding.", _
               vbCritical, "Configuration Error"
        Exit Sub
    End If
    
    ' Set compliance_overall_status to "MODEL SETUP" with purple color
    On Error Resume Next
    With Range("compliance_overall_status")
        .Value = "MODEL SETUP"
        .Interior.Color = RGB(155, 89, 182)  ' Off purple
        .Font.Color = RGB(255, 255, 255)     ' White text
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    On Error GoTo 0
    
    ' ===================================================================
    ' CLEAR AND SETUP: COMPLIANCE SHEET
    ' ===================================================================
    Set ws = ThisWorkbook.Worksheets("Compliance")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Compliance"
    Else
        ws.Cells.ClearContents
        ws.Cells.Interior.colorIndex = xlNone
    End If
    Call AddComplianceHeaders(ws)
    Set ws = Nothing
    
    ' ===================================================================
    ' CLEAR AND SETUP: MARGIN VERIFICATION SHEET
    ' ===================================================================
    Set ws = ThisWorkbook.Worksheets("MarginVerification")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "MarginVerification"
    Else
        ws.Cells.ClearContents
        ws.Cells.Interior.colorIndex = xlNone
    End If
    Call AddMarginVerificationHeaders(ws)
    Set ws = Nothing
    
    ' ===================================================================
    ' CLEAR AND SETUP: RAW TRADE IMPORT SHEET
    ' ===================================================================
    Set ws = ThisWorkbook.Worksheets("RawTradeImport")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "RawTradeImport"
    Else
        ws.Cells.ClearContents
        ws.Cells.Interior.colorIndex = xlNone
    End If
    Call AddTradeSubmissionHeaders(ws)
    Set ws = Nothing
    
    ' ===================================================================
    ' CLEAR AND SETUP: CLIENT PORTFOLIO SHEET
    ' ===================================================================
    Set ws = ThisWorkbook.Worksheets("ClientPortfolio")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ClientPortfolio"
    Else
        ws.Cells.ClearContents
        ws.Cells.Interior.colorIndex = xlNone
    End If
    Call AddPortfolioHeaders(ws)
    Set ws = Nothing
    
    ' ===================================================================
    ' CLEAR AND SETUP: ORDGEN SHEET
    ' ===================================================================
    Set ws = ThisWorkbook.Worksheets("OrderGen")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "OrderGen"
    Else
        ws.Cells.ClearContents
        ws.Cells.Interior.colorIndex = xlNone
    End If
    ' OrderGen headers are set during BlockSimilarOrders
    Call SetupOrderGenHeaders(ws)
    Set ws = Nothing
    
    ' ===================================================================
    ' CLEAR AND SETUP: EXECUTIONRESULTS SHEET
    ' ===================================================================
    Set ws = ThisWorkbook.Worksheets("ExecutionResults")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ExecutionResults"
    Else
        ws.Cells.ClearContents
        ws.Cells.Interior.colorIndex = xlNone
    End If
    ' ExecutionResults headers are set during execution import
    Set ws = Nothing
        
    ' ===================================================================
    ' CLEAR AND SETUP: ORDERTRACKING SHEET
    ' ===================================================================
    Set ws = ThisWorkbook.Worksheets("OrderTracking")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "OrderTracking"
    Else
        ws.Cells.ClearContents
        ws.Cells.Interior.colorIndex = xlNone
    End If
    Call AddOrderTrackingHeaders(ws)
    Set ws = Nothing
    
    ' ===================================================================
    ' FORMAT: BBG_Validation SHEET
    ' ===================================================================
    Call FormatBBGValidationHeaders
    
    On Error GoTo 0
End Sub


' ===================================================================
' TRADE SUBMISSIONS IMPORT
' ===================================================================
Public Sub ImportTradeSubmissions()
    On Error GoTo ErrorHandler
    
    ' Ensure database connection
    If Not modSQLConnections.EnsureConnection() Then
        MsgBox "Database connection failed", vbCritical
        Exit Sub
    End If
    
    Dim dbConn As ADODB.Connection
    Set dbConn = modSQLConnections.GetConnection()
    If dbConn Is Nothing Then
        MsgBox "Cannot establish database connection", vbCritical
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("RawTradeImport")
    
    ' Clear existing data
    ws.Range("A2:AO1000").ClearContents
    
    ' Build SQL query for today's submissions
    Dim todayStr As String
    todayStr = Format(Range("today").Value, "yyyy-mm-dd")
    
    Dim sql As String
    sql = "SELECT " & _
          "synthetic_borrow_app_id, client_name, client_account_number, user_email, " & _
          "client_buying_power, requested_amount, quoted_borrow_amount, " & _
          "payback_amount, annualized_rate, maximum_limit_percentage, " & _
          "expiry_date, request_time, planned_execution_time, " & _
          "leg1_action, leg1_quantity, leg1_option_type, leg1_strike, leg1_theoretical_price, leg1_notional_value, " & _
          "leg2_action, leg2_quantity, leg2_option_type, leg2_strike, leg2_theoretical_price, leg2_notional_value, " & _
          "leg3_action, leg3_quantity, leg3_option_type, leg3_strike, leg3_theoretical_price, leg3_notional_value, " & _
          "leg4_action, leg4_quantity, leg4_option_type, leg4_strike, leg4_theoretical_price, leg4_notional_value, " & _
          "user_id, created_at, updated_at " & _
          "FROM " & Range("table_synthetic_borrow").Value & " " & _
          "WHERE DATE(created_at) = '" & todayStr & "' " & _
          "ORDER BY request_time ASC"
    
    ' Execute query and populate worksheet
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    ' Use client-side cursor for CopyFromRecordset
    rs.CursorLocation = adUseClient
    rs.Open sql, dbConn, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        ' Check record count
        rs.MoveLast
        rs.MoveFirst
        
        ' Copy recordset data starting from row 2
        ws.Range("A2").CopyFromRecordset rs
        
        ' Format the data
        Call FormatTradeSubmissionData(ws)
    Else
        MsgBox "No trade submissions found for " & todayStr, vbInformation
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    MsgBox "Error importing trade submissions: " & Err.description, vbCritical
End Sub

' ===================================================================
' MARGIN VERIFICATION STATUS IMPORT
' ===================================================================
Public Sub ImportMarginVerificationStatus()
    On Error GoTo ErrorHandler
    
    Dim wsMargin As Worksheet
    Set wsMargin = ThisWorkbook.Worksheets("MarginVerification")
    
    ' Copy today's trades from RawTradeImport
    Dim wsRaw As Worksheet
    Set wsRaw = ThisWorkbook.Worksheets("RawTradeImport")
    
    Dim lastRow As Long
    lastRow = wsRaw.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow > 1 Then
        
        ' Copy specific columns
        Dim i As Long
        For i = 2 To lastRow
            wsMargin.Range("A" & i).Value = wsRaw.Range("A" & i).Value ' synthetic_borrow_app_id
            wsMargin.Range("B" & i).Value = wsRaw.Range("B" & i).Value ' client_name
            wsMargin.Range("C" & i).Value = wsRaw.Range("C" & i).Value ' client_account_number
            wsMargin.Range("D" & i).Value = wsRaw.Range("D" & i).Value ' client_email
            wsMargin.Range("E" & i).Value = wsRaw.Range("E" & i).Value ' client_buying_power
            wsMargin.Range("F" & i).Value = wsRaw.Range("F" & i).Value ' requested_amount
            wsMargin.Range("G" & i).Value = wsRaw.Range("G" & i).Value ' quoted_borrow_amount
            wsMargin.Range("H" & i).Value = wsRaw.Range("H" & i).Value ' payback_amount
            wsMargin.Range("I" & i).Value = wsRaw.Range("I" & i).Value ' annualized_rate
            
            ' Set all margin status to PENDING initially
            wsMargin.Range("J" & i).Value = "PENDING"
        Next i
        
        ' Add data validation to Margin Status column (J)
        With wsMargin.Range("J2:J" & lastRow).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="PENDING,YES,NO"
            .IgnoreBlank = False
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
            .InputTitle = "Margin Status"
            .ErrorTitle = "Invalid Entry"
            .InputMessage = "Select margin verification status"
            .ErrorMessage = "Please select PENDING, YES, or NO"
        End With
        
        ' Format the Margin Status column with conditional formatting
        With wsMargin.Range("J2:J" & lastRow)
            .FormatConditions.Delete
            
            ' PENDING - Yellow
            .FormatConditions.Add Type:=xlTextString, String:="PENDING", TextOperator:=xlContains
            With .FormatConditions(1)
                .Interior.Color = RGB(255, 235, 156) ' Light yellow
                .Font.Color = RGB(156, 101, 0) ' Dark orange
                .Font.Bold = True
            End With
            
            ' YES - Green
            .FormatConditions.Add Type:=xlTextString, String:="YES", TextOperator:=xlContains
            With .FormatConditions(2)
                .Interior.Color = RGB(198, 239, 206) ' Light green
                .Font.Color = RGB(0, 97, 0) ' Dark green
                .Font.Bold = True
            End With
            
            ' NO - Red
            .FormatConditions.Add Type:=xlTextString, String:="NO", TextOperator:=xlContains
            With .FormatConditions(3)
                .Interior.Color = RGB(255, 199, 206) ' Light red
                .Font.Color = RGB(156, 0, 6) ' Dark red
                .Font.Bold = True
            End With
        End With
        
        ' Auto-fit columns
        wsMargin.Columns("A:J").AutoFit
        
        ' Add borders to data range
        With wsMargin.Range("A1:J" & lastRow)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error importing margin verification: " & Err.description, vbCritical
End Sub


' ===================================================================
' CLIENT PORTFOLIO IMPORT
' ===================================================================
Public Sub ImportClientPortfolio()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ClientPortfolio")
    
    ' Get all active positions (not expired)
    Dim todayStr As String
    todayStr = Format(Range("today").Value, "yyyy-mm-dd")
    
    Dim sql As String
    sql = "SELECT id, synthetic_borrow_app_id, client_name, client_account, user_email, " & _
          "execution_date, expiry_date, box_structure, executed_premium, " & _
          "payback_amount, annualized_rate, trading_system_id, " & _
          "DATEDIFF(expiry_date, '" & todayStr & "') as days_to_expiry " & _
          "FROM " & Range("table_executed_trades").Value & " " & _
          "WHERE expiry_date > '" & todayStr & "' " & _
          "AND status = 'EXECUTED' " & _
          "ORDER BY days_to_expiry ASC, client_name ASC"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open sql, conn
    
    If Not rs.EOF Then
        ' Copy data starting from row 2
        ws.Range("A2").CopyFromRecordset rs
        
        ' Format and add alerts
        Call FormatPortfolioData(ws)
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error importing client portfolio: " & Err.description, vbCritical
End Sub



' ===================================================================
' ECECUTION RESULTS
' ===================================================================

Public Sub ProcessExecutionData(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then Exit Sub
    
    ' Calculate net premiums and format data
    Dim i As Long
    For i = 2 To lastRow
        ' Calculate total premium from leg prices and quantities
        Dim leg1Premium As Double, leg2Premium As Double
        Dim leg3Premium As Double, leg4Premium As Double
        
        ' Assuming columns follow the execution format
        ' Adjust column references based on your actual CSV structure
        leg1Premium = ws.Range("E" & i).Value * ws.Range("D" & i).Value * 100
        leg2Premium = ws.Range("F" & i).Value * ws.Range("D" & i).Value * 100
        leg3Premium = ws.Range("G" & i).Value * ws.Range("D" & i).Value * 100
        leg4Premium = ws.Range("H" & i).Value * ws.Range("D" & i).Value * 100
        
        ' Calculate net premium (sells - buys)
        ws.Range("I" & i).Value = leg2Premium + leg4Premium - leg1Premium - leg3Premium
        
        ' Add execution status
        ws.Range("J" & i).Value = "EXECUTED"
        
        ' Add timestamp
        ws.Range("K" & i).Value = Now()
    Next i
    
    ' Format columns
    ws.Range("E:I").NumberFormat = "$#,##0.00"
    ws.Range("K:K").NumberFormat = "mm/dd/yyyy hh:mm"
    
    ' Auto-fit columns
    ws.Columns.AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error processing execution data: " & Err.description, vbCritical
End Sub




' ===================================================================
' EXPIRATION CHECKING
' ===================================================================

Public Sub CheckAndAlertExpirations()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ClientPortfolio")
    
    ' Navigate to ClientPortfolio
    ws.Activate
    
    ' Get today's date from named range
    Dim todayDate As Date
    todayDate = Range("today").Value
    
    ' Get alert threshold from named range
    Dim alertDays As Long
    alertDays = Range("expiration_alert_days").Value
    
    ' Find last row
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then
        Exit Sub ' No data to check
    End If
    
    Dim alertCount As Long
    alertCount = 0
    
    ' Clear any existing alert column values first
    If ws.Range("O1").Value <> "Alert Status" Then
        ws.Range("O1").Value = "Alert Status"
        ws.Range("O1").Font.Bold = True
    End If
    ws.Range("O2:O" & lastRow).ClearContents
    ws.Range("O2:O" & lastRow).Interior.colorIndex = xlNone
    
    Dim i As Long
    For i = 2 To lastRow
        ' Get expiry date from column D (adjust if different column)
        Dim expiryDate As Date
        If IsDate(ws.Range("D" & i).Value) Then
            expiryDate = ws.Range("E" & i).Value
            
            ' Calculate business days until expiry
            Dim daysToExpiry As Long
            daysToExpiry = WorksheetFunction.NetworkDays(todayDate, expiryDate)
            
            ' Check if within alert threshold
            If daysToExpiry <= alertDays And daysToExpiry >= 0 Then
                ws.Range("O" & i).Value = "ALERT"
                ws.Range("A" & i & ":O" & i).Interior.Color = RGB(255, 182, 193) ' Light red
                alertCount = alertCount + 1
            Else
                ws.Range("O" & i).Value = "OK"
            End If
            
            ' Optional: Add days to expiry in another column for reference
            If ws.Range("N1").Value <> "Days to Expiry" Then
                ws.Range("N1").Value = "Days to Expiry"
                ws.Range("N1").Font.Bold = True
            End If
            ws.Range("N" & i).Value = daysToExpiry
        End If
    Next i
    
    ' Format columns
    ws.Columns("N:O").AutoFit
    
    If alertCount > 0 Then
        ' Generate CSV attachment
        Dim csvPath As String
        csvPath = CreateExpirationCSV(ws)
        
        ' Create email draft
        Call CreateExpirationEmail(alertCount, csvPath)
        
        ' Show alert message
        MsgBox alertCount & " position" & IIf(alertCount > 1, "s", "") & _
               " expiring within " & alertDays & " business days!" & vbCrLf & vbCrLf & _
               "Email draft has been created with details.", _
               vbExclamation, "Expiration Alert"
    Else
        ' Optional: Show confirmation that no alerts were found
        MsgBox "No positions expiring within " & alertDays & " business days.", _
               vbInformation, "No Expiration Alerts"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error checking expirations: " & Err.description, vbCritical
End Sub


