Attribute VB_Name = "modDataExport"

' ===================================================================
' DATA EXPORT MODULE - Version 2
' Synthetic Borrow Trading System
' ===================================================================

Option Explicit

' ===================================================================
' ORDER FILE GENERATION
' ===================================================================


Public Sub CleanupDailyTables()
        
    ' Clear sheets that get re-populated during compliance/export workflow
    ' Headers will be recreated by the import/export routines
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' Clear Compliance sheet
    Set ws = ThisWorkbook.Worksheets("Compliance")
    If Not ws Is Nothing Then
        ws.Cells.ClearContents
        ws.Cells.Interior.colorIndex = xlNone
    End If
    
    ' Clear OrderGen sheet (gets populated during order processing)
    Set ws = ThisWorkbook.Worksheets("OrderGen")
    If Not ws Is Nothing Then
        ws.Cells.ClearContents
        ws.Cells.Interior.colorIndex = xlNone
    End If
    
    ' Clear ExecutionResults sheet
    Set ws = ThisWorkbook.Worksheets("ExecutionResults")
    If Not ws Is Nothing Then
        ws.Cells.ClearContents
        ws.Cells.Interior.colorIndex = xlNone
    End If
    
    On Error GoTo ErrorHandler
    
    
    On Error GoTo ErrorHandler
    
    ' Ensure database connection
    If Not modSQLConnections.EnsureConnection() Then
        MsgBox "Cannot connect to database for cleanup", vbCritical
        Exit Sub
    End If
    
    Dim dbConn As ADODB.Connection
    Set dbConn = modSQLConnections.GetConnection()
    If dbConn Is Nothing Then Exit Sub
    
    ' Get today's date as string
    Dim todayStr As String
    todayStr = Format(Range("today").Value, "yyyy-mm-dd")
    
    ' Delete from trade_staging (by execution_date)
    Dim sql As String
    sql = "DELETE FROM synthetic_borrow.trade_staging WHERE DATE(execution_date) = '" & todayStr & "'"
    dbConn.Execute sql
    
    ' Delete from trade_staging (by execution_date)
    sql = "DELETE FROM synthetic_borrow.trade_ticket WHERE DATE(execution_date) = '" & todayStr & "'"
    dbConn.Execute sql
    
    ' Delete from margin_verification (by verification_date)
    sql = "DELETE FROM synthetic_borrow.margin_verification WHERE DATE(verification_date) = '" & todayStr & "'"
    dbConn.Execute sql
    
    ' Delete from order_tracking (by created_at timestamp)
    sql = "DELETE FROM synthetic_borrow.blocked_order_tracking WHERE DATE(created_at) = '" & todayStr & "'"
    dbConn.Execute sql
    
    ' Delete from synthetic_borrow_compliance (by check_date)
    sql = "DELETE FROM synthetic_borrow.synthetic_borrow_compliance WHERE DATE(check_date) = '" & todayStr & "'"
    dbConn.Execute sql
    
    ' Delete from trade_executions (by execution_date)
    sql = "DELETE FROM synthetic_borrow.trade_executions WHERE DATE(execution_date) = '" & todayStr & "'"
    dbConn.Execute sql
    
    Debug.Print "Daily table cleanup completed for date: " & todayStr
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error during daily table/sheet cleanup: " & Err.description, vbCritical
End Sub




Public Sub ExportToTradeStaging()
    On Error GoTo ErrorHandler
    
    If Not modSQLConnections.EnsureConnection() Then
        MsgBox "Database connection failed", vbCritical
        Exit Sub
    End If
    
    Dim dbConn As ADODB.Connection
    Set dbConn = modSQLConnections.GetConnection()
    If dbConn Is Nothing Then Exit Sub
    
    Dim todayDate As Date
    todayDate = Range("today").Value
    Dim formattedDate As String
    formattedDate = Format(todayDate, "yyyy-mm-dd")
    
    Dim wsBbg As Worksheet
    Set wsBbg = ThisWorkbook.Worksheets("BBG_Validation")
    
    Dim lastRow As Long
    lastRow = wsBbg.Range("A" & Rows.Count).End(xlUp).row
    
    If lastRow < 2 Then
        MsgBox "No BBG data found", vbExclamation
        Exit Sub
    End If
    
    Dim insertCount As Long
    insertCount = 0
    
    Dim i As Long
    For i = 2 To lastRow
        If wsBbg.Range("A" & i).Value <> "" Then
            
            ' ============================================================
            ' COLUMN DECLARATIONS - MAPPED TO EXCEL COLUMNS
            ' ============================================================
            
            ' Basic Trade Info (Columns A-F)
            Dim tradeID As String, clientName As String, marginStatus As String
            Dim tenor As String, expiryDate As Date, formattedExpiryDate As String
            Dim paybackAmount As Double
            
            ' Treasury and Comparison (Columns G-H)
            Dim treasuryYield As Double, primeVsSecondPx As Double
            
            ' ICE Data (Columns I-N)
            Dim iceCreditToday As Double, iceMidMarket As Double
            Dim iceLimitMid As Double, iceImpliedRate As Double
            Dim iceLimitImpliedRate As Double, iceVsTsy As Double
            
            ' BBG Data (Columns O-R)
            Dim bbgCreditToday As Double, bbgMidMarket As Double
            Dim bbgImpliedRate As Double, bbgVsTsy As Double
            
            ' Leg Market Values (Columns S, Z, AG, AN)
            Dim leg1MarketVal As Double, leg2MarketVal As Double
            Dim leg3MarketVal As Double, leg4MarketVal As Double
            
            ' Leg 1 Details (Columns T-Y)
            Dim leg1Ticker As String, leg1Price As Double, leg1Type As String
            Dim leg1Qty As Long, leg1Mult As Long, leg1Strike As Double
            
            ' Leg 2 Details (Columns AA-AF)
            Dim leg2Ticker As String, leg2Price As Double, leg2Type As String
            Dim leg2Qty As Long, leg2Mult As Long, leg2Strike As Double
            
            ' Leg 3 Details (Columns AH-AM)
            Dim leg3Ticker As String, leg3Price As Double, leg3Type As String
            Dim leg3Qty As Long, leg3Mult As Long, leg3Strike As Double
            
            ' Leg 4 Details (Columns AO-AT)
            Dim leg4Ticker As String, leg4Price As Double, leg4Type As String
            Dim leg4Qty As Long, leg4Mult As Long, leg4Strike As Double
            
            ' Calculated fields
            Dim boxQuantity As Long, boxWidth As Double
            
            ' ============================================================
            ' READ VALUES FROM EXCEL - UPDATE THESE COLUMN LETTERS
            ' ============================================================
            
            tradeID = wsBbg.Range("A" & i).Value
            clientName = wsBbg.Range("B" & i).Value
            marginStatus = wsBbg.Range("C" & i).Value
            tenor = wsBbg.Range("D" & i).Value
            expiryDate = wsBbg.Range("E" & i).Value
            formattedExpiryDate = Format(expiryDate, "yyyy-mm-dd")
            paybackAmount = wsBbg.Range("F" & i).Value
            treasuryYield = wsBbg.Range("G" & i).Value        ' Tsy Curve
            primeVsSecondPx = wsBbg.Range("H" & i).Value
            
            ' Credit and Payback (J-K)
            iceCreditToday = wsBbg.Range("I" & i).Value       ' ICE Credit Today
            iceMidMarket = wsBbg.Range("J" & i).Value         ' ICE Mid
            iceLimitMid = wsBbg.Range("K" & i).Value          ' ICE Cushion Mid
            iceImpliedRate = wsBbg.Range("L" & i).Value       ' ICE Rate
            iceLimitImpliedRate = wsBbg.Range("M" & i).Value  ' ICE Cushion Limit
            iceVsTsy = wsBbg.Range("N" & i).Value             ' ICE vs Tsy
            
            ' Columns O-R: BBG Data
            bbgCreditToday = wsBbg.Range("O" & i).Value       ' BBG Credit Today
            bbgMidMarket = wsBbg.Range("P" & i).Value         ' BBG Mid
            bbgImpliedRate = wsBbg.Range("Q" & i).Value       ' BBG Rate
            bbgVsTsy = wsBbg.Range("R" & i).Value
            
            
            ' Leg 1 Details - ADJUST STARTING COLUMN
            leg1MarketVal = wsBbg.Range("S" & i).Value
            leg1Ticker = wsBbg.Range("T" & i).Value
            leg1Price = wsBbg.Range("U" & i).Value
            leg1Type = wsBbg.Range("V" & i).Value
            leg1Qty = wsBbg.Range("W" & i).Value
            leg1Mult = wsBbg.Range("X" & i).Value
            leg1Strike = wsBbg.Range("Y" & i).Value
            
            ' Leg 2 Details
            leg2MarketVal = wsBbg.Range("Z" & i).Value
            leg2Ticker = wsBbg.Range("AA" & i).Value
            leg2Price = wsBbg.Range("AB" & i).Value
            leg2Type = wsBbg.Range("AC" & i).Value
            leg2Qty = wsBbg.Range("AD" & i).Value
            leg2Mult = wsBbg.Range("AE" & i).Value
            leg2Strike = wsBbg.Range("AF" & i).Value
            
            ' Leg 3 Details
            leg3MarketVal = wsBbg.Range("AG" & i).Value
            leg3Ticker = wsBbg.Range("AH" & i).Value
            leg3Price = wsBbg.Range("AI" & i).Value
            leg3Type = wsBbg.Range("AJ" & i).Value
            leg3Qty = wsBbg.Range("AK" & i).Value
            leg3Mult = wsBbg.Range("AL" & i).Value
            leg3Strike = wsBbg.Range("AM" & i).Value

            ' Leg 4 Details
            leg4MarketVal = wsBbg.Range("AN" & i).Value
            leg4Ticker = wsBbg.Range("AO" & i).Value
            leg4Price = wsBbg.Range("AP" & i).Value
            leg4Type = wsBbg.Range("AQ" & i).Value
            leg4Qty = wsBbg.Range("AR" & i).Value
            leg4Mult = wsBbg.Range("AS" & i).Value
            leg4Strike = wsBbg.Range("AT" & i).Value
            
            ' Calculate box metrics
            boxQuantity = Abs(leg1Qty)
            boxWidth = Abs(leg2Strike - leg1Strike)
            
            ' Map Excel margin status to SQL verified enum
            Select Case marginStatus
                Case "YES"
                    marginStatus = "Y"
                Case "NO"
                    marginStatus = "N"
                Case "PENDING"
                    marginStatus = "N"  ' Treat PENDING as N
                Case Else
                    marginStatus = "N"  ' Default to N for any other values
            End Select
            
            
            ' ============================================================
            ' ESCAPE QUOTES IN TEXT FIELDS
            ' ============================================================
            tradeID = Replace(tradeID, "'", "''")
            clientName = Replace(clientName, "'", "''")
            marginStatus = Replace(marginStatus, "'", "''")
            tenor = Replace(tenor, "'", "''")
            leg1Ticker = Replace(leg1Ticker, "'", "''")
            leg2Ticker = Replace(leg2Ticker, "'", "''")
            leg3Ticker = Replace(leg3Ticker, "'", "''")
            leg4Ticker = Replace(leg4Ticker, "'", "''")
            leg1Type = Replace(leg1Type, "'", "''")
            leg2Type = Replace(leg2Type, "'", "''")
            leg3Type = Replace(leg3Type, "'", "''")
            leg4Type = Replace(leg4Type, "'", "''")
            
            ' ============================================================
            ' BUILD SQL INSERT STATEMENT IN CHUNKS
            ' ============================================================
            Dim sql As String
            Dim sqlCols As String
            Dim sqlVals As String
            
            ' Build column list
            sqlCols = "INSERT INTO trade_staging ("
            sqlCols = sqlCols & "synthetic_borrow_app_id, execution_date, client_name, margin_verified, tenor, "
            sqlCols = sqlCols & "box_quantity, box_expiry_date, box_width, "
            sqlCols = sqlCols & "leg1_market_value, leg2_market_value, leg3_market_value, leg4_market_value, "
            sqlCols = sqlCols & "ice_credit_today, ice_mid_market, ice_limit_mid, "
            sqlCols = sqlCols & "ice_implied_rate, ice_limit_implied_rate, ice_vs_tsy, "
            sqlCols = sqlCols & "bbg_credit_today, bbg_mid_market, bbg_implied_rate, bbg_vs_tsy, "
            sqlCols = sqlCols & "treasury_yield, payback_amount, "
            sqlCols = sqlCols & "leg1_ticker, leg1_price, leg1_type, leg1_quantity, leg1_multiplier, leg1_strike, "
            sqlCols = sqlCols & "leg2_ticker, leg2_price, leg2_type, leg2_quantity, leg2_multiplier, leg2_strike, "
            sqlCols = sqlCols & "leg3_ticker, leg3_price, leg3_type, leg3_quantity, leg3_multiplier, leg3_strike, "
            sqlCols = sqlCols & "leg4_ticker, leg4_price, leg4_type, leg4_quantity, leg4_multiplier, leg4_strike, "
            sqlCols = sqlCols & "status, import_date, processed) VALUES ("
            
            ' Build values list
            sqlVals = "'" & tradeID & "', '" & formattedDate & "', '" & clientName & "', "
            sqlVals = sqlVals & "'" & marginStatus & "', '" & tenor & "', "
            sqlVals = sqlVals & boxQuantity & ", '" & formattedExpiryDate & "', " & boxWidth & ", "
            sqlVals = sqlVals & leg1MarketVal & ", " & leg2MarketVal & ", " & leg3MarketVal & ", " & leg4MarketVal & ", "
            sqlVals = sqlVals & iceCreditToday & ", " & iceMidMarket & ", " & iceLimitMid & ", "
            sqlVals = sqlVals & iceImpliedRate & ", " & iceLimitImpliedRate & ", " & iceVsTsy & ", "
            sqlVals = sqlVals & bbgCreditToday & ", " & bbgMidMarket & ", " & bbgImpliedRate & ", " & bbgVsTsy & ", "
            sqlVals = sqlVals & treasuryYield & ", " & paybackAmount & ", "
            
            ' Leg 1
            sqlVals = sqlVals & "'" & leg1Ticker & "', " & leg1Price & ", '" & leg1Type & "', "
            sqlVals = sqlVals & leg1Qty & ", " & leg1Mult & ", " & leg1Strike & ", "
            
            ' Leg 2
            sqlVals = sqlVals & "'" & leg2Ticker & "', " & leg2Price & ", '" & leg2Type & "', "
            sqlVals = sqlVals & leg2Qty & ", " & leg2Mult & ", " & leg2Strike & ", "
            
            ' Leg 3
            sqlVals = sqlVals & "'" & leg3Ticker & "', " & leg3Price & ", '" & leg3Type & "', "
            sqlVals = sqlVals & leg3Qty & ", " & leg3Mult & ", " & leg3Strike & ", "
            
            ' Leg 4
            sqlVals = sqlVals & "'" & leg4Ticker & "', " & leg4Price & ", '" & leg4Type & "', "
            sqlVals = sqlVals & leg4Qty & ", " & leg4Mult & ", " & leg4Strike & ", "
            
            ' Status fields
            sqlVals = sqlVals & "'STAGED', NOW(), 'N')"
            
            ' Combine
            sql = sqlCols & sqlVals
            
            dbConn.Execute sql
            insertCount = insertCount + 1
            
            Debug.Print "Staged " & tradeID & " | Client: " & clientName & " | Qty: " & boxQuantity
        End If
    Next i
    
    If insertCount > 0 Then
        Debug.Print vbCrLf & "=== Staged " & insertCount & " BBG records ==="
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error staging BBG data: " & Err.description & " at row " & i, vbCritical
End Sub




Public Sub SendOrderEmail(attachmentPath As String, orderCount As Long)
    On Error GoTo ErrorHandler
    
    Dim objOutlook As Object
    Dim objMail As Object
    
    ' Create Outlook application object
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0) ' 0 = olMailItem
    
    ' Get email parameters from named ranges
    Dim toRecipients As String
    Dim ccRecipients As String
    Dim emailSubject As String
    Dim emailBody As String
    
    toRecipients = Range("email_trade_to").Value
    ccRecipients = Range("email_trade_cc").Value
    emailSubject = Range("email_trade_subject").Value
    emailBody = Range("email_trade_body").Value
    
    ' Replace placeholders in subject and body
    emailSubject = Replace(emailSubject, "#DATE#", Format(Now(), "mm/dd/yyyy"))
    emailSubject = Replace(emailSubject, "#TIME#", Format(Now(), "hh:mm AM/PM"))
    emailSubject = Replace(emailSubject, "#COUNT#", orderCount)
    
    emailBody = Replace(emailBody, "#DATE#", Format(Now(), "mm/dd/yyyy"))
    emailBody = Replace(emailBody, "#TIME#", Format(Now(), "hh:mm AM/PM"))
    emailBody = Replace(emailBody, "#COUNT#", orderCount)
    emailBody = Replace(emailBody, "#FILENAME#", Mid(attachmentPath, InStrRev(attachmentPath, "\") + 1))
    
    ' Configure email
    With objMail
        .To = toRecipients
        .CC = ccRecipients
        .subject = emailSubject
        .body = emailBody
        
        ' Add attachment
        If Len(Dir(attachmentPath)) > 0 Then
            .Attachments.Add attachmentPath
        Else
            MsgBox "Warning: Order file not found for attachment", vbExclamation
        End If
        
        ' Display email (change to .Send to send automatically)
        .Display
    End With
    
    ' Clean up
    Set objMail = Nothing
    Set objOutlook = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating email: " & Err.description, vbCritical
    Set objMail = Nothing
    Set objOutlook = Nothing
End Sub



' ===================================================================
' WORKFLOW WRAPPER FUNCTIONS (For Macro Button Assignment)
' ===================================================================


Public Sub ExportMarginVerification()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Ensure SQL connection is open
    If Not modSQLConnections.EnsureConnection() Then
        MsgBox "Failed to connect to database. Please check connection settings.", vbCritical
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("MarginVerification")
    
    ' Find last row with data
    Dim lastRow As Long
    lastRow = ws.Range("A2").End(xlDown).row
    
    If lastRow < 2 Then
        MsgBox "No data found to export in MarginVerification sheet.", vbInformation
        Exit Sub
    End If
    
    ' Get named range values
    Dim verificationDate As String
    Dim verifiedBy As String
    Dim updatedAt As String
    
    verificationDate = Format(Range("today").Value, "yyyy-mm-dd")
    verifiedBy = Range("trader").Value
    updatedAt = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    ' Counter for successful exports
    Dim exportCount As Long
    exportCount = 0
    
    ' Get database connection
    Dim dbConn As ADODB.Connection
    Set dbConn = modSQLConnections.GetConnection()
    
    ' Process each row
    Dim i As Long
    For i = 2 To lastRow
        ' Skip empty rows
        If ws.Range("A" & i).Value <> "" Then
            
            ' Extract data from Excel columns
            Dim tradeID As String
            Dim clientName As String
            Dim clientAccount As String
            Dim marginStatus As String
            Dim verified As String
            Dim notes As String
            
            tradeID = ws.Range("A" & i).Value          ' synthetic_borrow_app_id
            clientName = ws.Range("B" & i).Value       ' Client Name
            clientAccount = ws.Range("C" & i).Value    ' Account
            marginStatus = UCase(Trim(ws.Range("J" & i).Value))  ' Margin Status
            
            ' Map Excel margin status to SQL verified enum
            Select Case marginStatus
                Case "YES"
                    verified = "Y"
                Case "NO"
                    verified = "N"
                Case "PENDING"
                    verified = "N"  ' Treat PENDING as N
                Case Else
                    verified = "N"  ' Default to N for any other values
            End Select
            
            ' Handle notes - check if column K exists and has value
            notes = ""
            If ws.Range("K" & i).Value <> "" Then
                notes = ws.Range("K" & i).Value
            End If
            
            ' Escape single quotes in text fields to prevent SQL errors
            clientName = Replace(clientName, "'", "''")
            clientAccount = Replace(clientAccount, "'", "''")
            notes = Replace(notes, "'", "''")
            tradeID = Replace(tradeID, "'", "''")
            verifiedBy = Replace(verifiedBy, "'", "''")
            
            ' Build SQL INSERT statement with ON DUPLICATE KEY UPDATE
            Dim sql As String
            sql = "INSERT INTO margin_verification " & _
                  "(synthetic_borrow_app_id, client_name, client_account, verified, verification_date, notes, verified_by, updated_at) " & _
                  "VALUES (" & _
                  "'" & tradeID & "', " & _
                  "'" & clientName & "', " & _
                  "'" & clientAccount & "', " & _
                  "'" & verified & "', " & _
                  "'" & verificationDate & "', " & _
                  "'" & notes & "', " & _
                  "'" & verifiedBy & "', " & _
                  "'" & updatedAt & "') " & _
                  "ON DUPLICATE KEY UPDATE " & _
                  "client_name = '" & clientName & "', " & _
                  "client_account = '" & clientAccount & "', " & _
                  "verified = '" & verified & "', " & _
                  "verification_date = '" & verificationDate & "', " & _
                  "notes = '" & notes & "', " & _
                  "verified_by = '" & verifiedBy & "', " & _
                  "updated_at = '" & updatedAt & "'"
            
            ' Execute SQL statement
            dbConn.Execute sql
            exportCount = exportCount + 1
            
        End If
    Next i
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error exporting margin verification data: " & Err.description & vbCrLf & _
           "Error occurred at row " & i, vbCritical
End Sub





Public Sub PopulateOrderExportRow(ws As Worksheet, wsExport As Worksheet, summaryRow As Long, sourceRow As Long, exportRow As Long)
    ' Get summary data
    Dim orderID As String
    Dim destinationAccount As String
    Dim limitPrice As Double
    Dim totalQuantity As Long
    
    orderID = ws.Range("AO" & summaryRow).Value
    destinationAccount = ws.Range("AV" & summaryRow).Value
    limitPrice = ws.Range("AS" & summaryRow).Value
    totalQuantity = ws.Range("AR" & summaryRow).Value
    
    ' Populate export row
    wsExport.Range("A" & exportRow).Value = orderID
    wsExport.Range("B" & exportRow).Value = destinationAccount
    wsExport.Range("C" & exportRow).Value = "LMT" ' Limit order
    wsExport.Range("D" & exportRow).Value = limitPrice
    
    ' Create option symbols and populate legs
    wsExport.Range("E" & exportRow).Value = CreateOptionSymbol(ws, sourceRow, 1)
    wsExport.Range("F" & exportRow).Value = totalQuantity
    wsExport.Range("G" & exportRow).Value = "OPEN"
    
    wsExport.Range("H" & exportRow).Value = CreateOptionSymbol(ws, sourceRow, 2)
    wsExport.Range("I" & exportRow).Value = totalQuantity
    wsExport.Range("J" & exportRow).Value = "OPEN"
    
    wsExport.Range("K" & exportRow).Value = CreateOptionSymbol(ws, sourceRow, 3)
    wsExport.Range("L" & exportRow).Value = totalQuantity
    wsExport.Range("M" & exportRow).Value = "OPEN"
    
    wsExport.Range("N" & exportRow).Value = CreateOptionSymbol(ws, sourceRow, 4)
    wsExport.Range("O" & exportRow).Value = totalQuantity
    wsExport.Range("P" & exportRow).Value = "OPEN"
End Sub

Public Function CreateOptionSymbol(ws As Worksheet, row As Long, legNumber As Long) As String
    ' Get leg data based on leg number
    Dim baseCol As Long
    Select Case legNumber
        Case 1: baseCol = 14 ' Column N
        Case 2: baseCol = 20 ' Column T
        Case 3: baseCol = 26 ' Column Z
        Case 4: baseCol = 32 ' Column AF
    End Select
    
    Dim expiryDate As Date
    Dim optionType As String
    Dim strike As Double
    
    expiryDate = ws.Range("K" & row).Value
    optionType = ws.Cells(row, baseCol + 2).Value ' Option type column
    strike = ws.Cells(row, baseCol + 3).Value ' Strike column
    
    ' Format: SPX DDMMMYY C/P Strike
    Dim formattedExpiry As String
    formattedExpiry = Format(expiryDate, "yymmdd")
    
    Dim optionLetter As String
    optionLetter = IIf(UCase(optionType) = "CALL", "C", "P")
    
    Dim strikeFormatted As String
    strikeFormatted = Format(strike, "00000")
    
    CreateOptionSymbol = ".SPX" & formattedExpiry & optionLetter & strikeFormatted
End Function

Public Function FindSourceRowForOrderID(ws As Worksheet, orderID As String) As Long
    Dim lastRow As Long
    lastRow = ws.Range("A1").CurrentRegion.Rows.Count
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Range("AL" & i).Value = orderID Then
            FindSourceRowForOrderID = i
            Exit Function
        End If
    Next i
    
    FindSourceRowForOrderID = 0
End Function



