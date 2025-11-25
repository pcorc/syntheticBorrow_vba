Attribute VB_Name = "modTradeProcessing"

' ===================================================================
' TRADE PROCESSING MODULE - Version 2
' Synthetic Borrow Trading System
' ===================================================================

Option Explicit

Public Sub ProcessOrdersForExecution()
    ' Wrapper for order processing - NO PARAMETERS for macro assignment
    On Error GoTo ErrorHandler
    
    If Not OrderGenerationComplianceChecks() Then
        Exit Sub
    End If
    
    ' Process the orders
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ordersGenerated As Long
    ordersGenerated = ProcessVerifiedTrades()
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If ordersGenerated > 0 Then
    
        Call ExportTradeTickets
        Call ExportToTradeStaging
        Call GenerateOrderFile

    Else
        MsgBox "No orders generated.", vbExclamation
    End If
    
    Exit Sub
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error processing orders: " & Err.description, vbCritical
End Sub


Public Function ProcessVerifiedTrades() As Long
    On Error GoTo ErrorHandler
    
    Dim wsOrderGen As Worksheet
    Set wsOrderGen = ThisWorkbook.Worksheets("OrderGen")
    wsOrderGen.Range("A2:O1000").ClearContents
    Call SetupOrderGenHeaders(wsOrderGen)
    
    Dim wsCompliance As Worksheet
    Set wsCompliance = ThisWorkbook.Worksheets("Compliance")
    
    Dim wsRaw As Worksheet
    Set wsRaw = ThisWorkbook.Worksheets("RawTradeImport")
    
    Dim wsBbg As Worksheet
    Set wsBbg = ThisWorkbook.Worksheets("BBG_Validation")
    
    Dim lastRowCompliance As Long
    lastRowCompliance = wsCompliance.Range("A1").CurrentRegion.Rows.Count
    
    Dim orderGenRow As Long
    orderGenRow = 2
    
    ' Simple loop - one order per approved trade
    Dim i As Long
    For i = 2 To lastRowCompliance
        If wsCompliance.Range("G" & i).Value = "APPROVED" Then
            Dim tradeID As String
            tradeID = wsCompliance.Range("A" & i).Value
            
            Dim rawRow As Long
            rawRow = FindTradeInRawImport(wsRaw, tradeID)
            
            Dim rawBbg As Long
            rawBbg = FindTradeInBbg(wsBbg, tradeID)
            
            If rawRow > 0 Then
                ' Write directly to OrderGen
                Call WriteIndividualTradeToOrderGen(wsRaw, rawRow, wsOrderGen, orderGenRow, tradeID, wsBbg, rawBbg)
                orderGenRow = orderGenRow + 1
            End If
        End If
    Next i
    
    ProcessVerifiedTrades = orderGenRow - 2
    Exit Function
    
ErrorHandler:
    ProcessVerifiedTrades = 0
    MsgBox "Error processing verified trades: " & Err.description, vbCritical
End Function

' ===================================================================
' EXPORT TRADE TICKETS TO SQL
' Exports order execution details from OrderGen
' ===================================================================

Public Sub ExportTradeTickets()
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
    
    Dim wsOrderGen As Worksheet
    Set wsOrderGen = ThisWorkbook.Worksheets("OrderGen")
    
    Dim wsRaw As Worksheet
    Set wsRaw = ThisWorkbook.Worksheets("RawTradeImport")
    
    Dim lastRow As Long
    lastRow = wsOrderGen.Range("A" & Rows.Count).End(xlUp).row
    
    If lastRow < 2 Then
        MsgBox "No orders found in OrderGen", vbExclamation
        Exit Sub
    End If
    
    Dim insertCount As Long
    insertCount = 0
    
    Dim i As Long
    For i = 2 To lastRow
        If wsOrderGen.Range("A" & i).Value <> "" Then
            
            ' Read order details directly from OrderGen
            Dim account As String, orderType As String, limitPrice As Double
            Dim leg1Symbol As String, leg1Qty As Long, leg1OpenClose As String
            Dim leg2Symbol As String, leg2Qty As Long, leg2OpenClose As String
            Dim leg3Symbol As String, leg3Qty As Long, leg3OpenClose As String
            Dim leg4Symbol As String, leg4Qty As Long, leg4OpenClose As String
            
            account = wsOrderGen.Range("A" & i).Value
            orderType = wsOrderGen.Range("B" & i).Value
            limitPrice = -Int(-(wsOrderGen.Range("C" & i).Value / 0.5)) * 0.5 ' this rounds up to the nearest 0.5 increment (conservative)
            
            leg1Symbol = wsOrderGen.Range("D" & i).Value
            leg1Qty = wsOrderGen.Range("E" & i).Value
            leg1OpenClose = wsOrderGen.Range("F" & i).Value
            
            leg2Symbol = wsOrderGen.Range("G" & i).Value
            leg2Qty = wsOrderGen.Range("H" & i).Value
            leg2OpenClose = wsOrderGen.Range("I" & i).Value
            
            leg3Symbol = wsOrderGen.Range("J" & i).Value
            leg3Qty = wsOrderGen.Range("K" & i).Value
            leg3OpenClose = wsOrderGen.Range("L" & i).Value
            
            leg4Symbol = wsOrderGen.Range("M" & i).Value
            leg4Qty = wsOrderGen.Range("N" & i).Value
            leg4OpenClose = wsOrderGen.Range("O" & i).Value
            
            ' Find matching trade ID from RawTradeImport
            Dim tradeID As String, rawRow As Long
            Dim leg1Strike As Double, leg2Strike As Double, expiryDate As String
            
            
            leg1Strike = ExtractStrikeFromTicker(leg1Symbol)
            leg2Strike = ExtractStrikeFromTicker(leg2Symbol)
            expiryDate = ExtractExpiryFromTicker(leg1Symbol)
            
            rawRow = FindFirstTradeForBox(wsRaw, leg1Strike, leg2Strike, expiryDate)
                        
            If rawRow > 0 Then
                tradeID = wsRaw.Range("A" & rawRow).Value
                
                ' Escape quotes
                tradeID = Replace(tradeID, "'", "''")
                account = Replace(account, "'", "''")
                orderType = Replace(orderType, "'", "''")
                leg1Symbol = Replace(leg1Symbol, "'", "''")
                leg2Symbol = Replace(leg2Symbol, "'", "''")
                leg3Symbol = Replace(leg3Symbol, "'", "''")
                leg4Symbol = Replace(leg4Symbol, "'", "''")
                leg1OpenClose = Replace(leg1OpenClose, "'", "''")
                leg2OpenClose = Replace(leg2OpenClose, "'", "''")
                leg3OpenClose = Replace(leg3OpenClose, "'", "''")
                leg4OpenClose = Replace(leg4OpenClose, "'", "''")
                
                ' Build simple SQL
                Dim sql As String
                sql = "INSERT INTO trade_ticket "
                sql = sql & "(synthetic_borrow_app_id, execution_date, account, order_type, limit_price, "
                sql = sql & "leg1_symbol, leg1_quantity, leg1_open_close, "
                sql = sql & "leg2_symbol, leg2_quantity, leg2_open_close, "
                sql = sql & "leg3_symbol, leg3_quantity, leg3_open_close, "
                sql = sql & "leg4_symbol, leg4_quantity, leg4_open_close, "
                sql = sql & "status, import_date) VALUES ("
                sql = sql & "'" & tradeID & "', '" & formattedDate & "', '" & account & "', '" & orderType & "', " & limitPrice & ", "
                sql = sql & "'" & leg1Symbol & "', " & leg1Qty & ", '" & leg1OpenClose & "', "
                sql = sql & "'" & leg2Symbol & "', " & leg2Qty & ", '" & leg2OpenClose & "', "
                sql = sql & "'" & leg3Symbol & "', " & leg3Qty & ", '" & leg3OpenClose & "', "
                sql = sql & "'" & leg4Symbol & "', " & leg4Qty & ", '" & leg4OpenClose & "', "
                sql = sql & "'READY', NOW())"
                
                dbConn.Execute sql
                insertCount = insertCount + 1
                
                Debug.Print "Exported ticket " & tradeID & " | Qty: " & Abs(leg1Qty)
            End If
        End If
    Next i
    
    If insertCount > 0 Then
        Debug.Print vbCrLf & "=== Exported " & insertCount & " trade tickets ==="
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error exporting trade tickets: " & Err.description, vbCritical
End Sub


Public Function ExtractStrikeFromTicker(ticker As String) As Double
    On Error Resume Next
    If Len(ticker) >= 16 Then ExtractStrikeFromTicker = CDbl(Right(ticker, 5))
End Function

Public Function ExtractExpiryFromTicker(ticker As String) As String
    ' Extract expiry date from format: .SPX251017C6000 -> 2025-10-17
    On Error Resume Next
    If Len(ticker) >= 13 Then
        Dim dateStr As String
        dateStr = Mid(ticker, 5, 6) ' YYMMDD portion
        
        Dim yy As String, mm As String, dd As String
        yy = "20" & Left(dateStr, 2)
        mm = Mid(dateStr, 3, 2)
        dd = Right(dateStr, 2)
        
        ExtractExpiryFromTicker = yy & "-" & mm & "-" & dd
    Else
        ExtractExpiryFromTicker = Format(Date, "yyyy-mm-dd")
    End If
End Function

Public Function FindFirstTradeForBox(ws As Worksheet, strike1 As Double, strike2 As Double, expiryDate As String) As Long
    On Error Resume Next
    Dim lastRow As Long
    lastRow = ws.Range("A" & Rows.Count).End(xlUp).row
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Range("A" & i).Value <> "" Then
            If Abs(ws.Range("Q" & i).Value - strike1) < 0.01 And _
               Abs(ws.Range("W" & i).Value - strike2) < 0.01 And _
               Format(ws.Range("K" & i).Value, "yyyy-mm-dd") = expiryDate Then
                FindFirstTradeForBox = i
                Exit Function
            End If
        End If
    Next i
    FindFirstTradeForBox = 0
End Function



Public Sub WriteIndividualTradeToOrderGen(wsRaw As Worksheet, rawRow As Long, wsOrderGen As Worksheet, orderGenRow As Long, tradeID As String, wsBbg As Worksheet, rawBbg As Long)
    ' Master account
    wsOrderGen.Range("A" & orderGenRow).Value = Range("vest_master_account").Value
    wsOrderGen.Range("B" & orderGenRow).Value = "LIMIT"
    wsOrderGen.Range("C" & orderGenRow).Value = wsBbg.Range("K" & rawBbg).Value + 0.1
    
    ' Map legs (existing code)
    Call MapLegToOrderGen(wsRaw, rawRow, wsOrderGen, orderGenRow, 1, "D")
    Call MapLegToOrderGen(wsRaw, rawRow, wsOrderGen, orderGenRow, 2, "G")
    Call MapLegToOrderGen(wsRaw, rawRow, wsOrderGen, orderGenRow, 3, "J")
    Call MapLegToOrderGen(wsRaw, rawRow, wsOrderGen, orderGenRow, 4, "M")
    
End Sub


Public Sub GenerateOrderFile()
    On Error GoTo ErrorHandler
    
    Dim wbExport As Workbook
    Dim wsOrderGen As Worksheet
    
    ' Get worksheet
    Set wsOrderGen = ThisWorkbook.Worksheets("OrderGen")
    
    ' Find last row of data
    Dim lastRow As Long
    lastRow = wsOrderGen.Range("A" & Rows.Count).End(xlUp).row
    
    If lastRow < 2 Then
        MsgBox "No orders to export", vbExclamation
        Exit Sub
    End If
    
    ' Check named ranges exist
    On Error Resume Next
    Dim testPath As String
    testPath = Range("order_template_path").Value
    If Err.Number <> 0 Then
        MsgBox "Named range 'order_template_path' not found", vbCritical
        Exit Sub
    End If
    
    Dim testFilename As String
    testFilename = Range("export_order_filename").Value
    If Err.Number <> 0 Then
        MsgBox "Named range 'export_order_filename' not found", vbCritical
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Build filename with timestamp
    Dim exportPath As String
    Dim fileName As String
    exportPath = Range("order_template_path").Value
    fileName = Replace(Range("export_order_filename").Value, "#DATE#", Format(Now(), "yyyymmdd"))
    
    ' Create new workbook for export
    Set wbExport = Workbooks.Add
    
    Dim wsExport As Worksheet
    Set wsExport = wbExport.Worksheets(1)
    
    ' Copy headers and data (columns A:O)
    wsOrderGen.Range("A1:O" & lastRow).Copy
    wsExport.Range("A1").PasteSpecial xlPasteValues
    wsExport.Range("A1").PasteSpecial xlPasteFormats
    
    Application.CutCopyMode = False
    
    ' Format the export sheet
    wsExport.Columns("A:O").AutoFit
    
    ' Ensure path ends with backslash
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    
    ' Check if directory exists
    If Len(Dir(exportPath, vbDirectory)) = 0 Then
        MsgBox "Directory does not exist: " & exportPath, vbCritical
        wbExport.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Full file path
    Dim fullPath As String
    fullPath = exportPath & fileName
    
    ' Save the file as CSV
    wbExport.SaveAs fileName:=fullPath, FileFormat:=xlCSV, CreateBackup:=False
    wbExport.Close SaveChanges:=False
    
    ' Send email with attachment
    Call SendOrderEmail(fullPath, lastRow - 1)
    
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "Error in GenerateOrderFile: " & Err.description & " (Error " & Err.Number & ")"
    
    ' Close workbook if it exists
    If Not wbExport Is Nothing Then
        On Error Resume Next
        wbExport.Close SaveChanges:=False
        On Error GoTo 0
    End If
    
    MsgBox errorMsg, vbCritical
End Sub

Public Sub ExportToExecutionStaging()
    On Error GoTo ErrorHandler
    
    ' Ensure database connection
    If Not modSQLConnections.EnsureConnection() Then
        MsgBox "Database connection failed", vbCritical
        Exit Sub
    End If
    
    Dim dbConn As ADODB.Connection
    Set dbConn = modSQLConnections.GetConnection()
    If dbConn Is Nothing Then Exit Sub
    
    ' Get today's date from named range
    Dim todayDate As Date
    todayDate = Range("today").Value
    Dim formattedDate As String
    formattedDate = Format(todayDate, "yyyy-mm-dd")
    
    ' Get OrderGen worksheet
    Dim wsOrderGen As Worksheet
    Set wsOrderGen = ThisWorkbook.Worksheets("OrderGen")
    
    ' Find last row of data
    Dim lastRow As Long
    lastRow = wsOrderGen.Range("A" & Rows.Count).End(xlUp).row
    
    If lastRow < 2 Then Exit Sub
    
    ' Loop through each order and insert into staging table
    Dim i As Long
    Dim sql As String
    Dim orderID As String
    Dim tradingSystemID As String
    Dim executedQty As Long
    Dim avgPrice As Double
    
    For i = 2 To lastRow
        ' Generate unique order ID for this staged order
        orderID = "ORD" & Format(Now(), "YYYYMMDDHHMMSS") & Right("000" & i, 3)
        tradingSystemID = wsOrderGen.Range("A" & i).Value ' Account number
        executedQty = Abs(wsOrderGen.Range("E" & i).Value) ' L1 Quantity (absolute value)
        avgPrice = wsOrderGen.Range("C" & i).Value ' Limit Price
        
        ' Escape single quotes
        orderID = Replace(orderID, "'", "''")
        
        ' Build INSERT statement - only insert required fields
        sql = "INSERT INTO trade_staging " & _
              "(order_id, trading_system_id, execution_date, executed_quantity, " & _
              "avg_execution_price, status, import_date, processed) " & _
              "VALUES (" & _
              "'" & orderID & "', " & _
              "'" & tradingSystemID & "', " & _
              "'" & formattedDate & "', " & _
              executedQty & ", " & _
              avgPrice & ", " & _
              "'STAGED', " & _
              "NOW(), " & _
              "'N')"
        
        ' Execute the insert
        dbConn.Execute sql
    Next i
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error exporting to trade_staging: " & Err.description, vbCritical
End Sub


Public Function CopyApprovedTradesToOrderGen(wsCompliance As Worksheet, wsRaw As Worksheet, _
                                            wsOrderGen As Worksheet, startRow As Long) As Long
    Dim lastRowCompliance As Long
    lastRowCompliance = wsCompliance.Range("A1").CurrentRegion.Rows.Count
    
    Dim orderGenRow As Long
    orderGenRow = startRow
    
    Dim i As Long
    For i = 2 To lastRowCompliance
        If wsCompliance.Range("G" & i).Value = "APPROVED" Then
            Dim tradeID As String
            tradeID = wsCompliance.Range("A" & i).Value
            
            Dim rawRow As Long
            rawRow = FindTradeInRawImport(wsRaw, tradeID)
            
            If rawRow > 0 Then
                Call MapTradeToOrderGen(wsRaw, rawRow, wsOrderGen, orderGenRow)
                orderGenRow = orderGenRow + 1
            End If
        End If
    Next i
    
    ' THIS LINE IS CRITICAL - Must return the final row number
    CopyApprovedTradesToOrderGen = orderGenRow
End Function

Public Function FindTradeInRawImport(wsRaw As Worksheet, tradeID As String) As Long
    Dim lastRowRaw As Long
    lastRowRaw = wsRaw.Range("A1").CurrentRegion.Rows.Count
    
    Dim j As Long
    Dim cellValue As Variant
    
    For j = 1 To lastRowRaw
        ' Get cell value with error handling
        On Error Resume Next
        cellValue = wsRaw.Range("A" & j).Value
        On Error GoTo 0
        
        ' Skip if cell has error or is empty
        If Not IsError(cellValue) And Not IsEmpty(cellValue) Then
            ' Convert both to strings for comparison
            If CStr(cellValue) = tradeID Then
                FindTradeInRawImport = j
                Exit Function
            End If
        End If
    Next j
    
    FindTradeInRawImport = 0  ' Return 0 if not found
End Function

Public Function FindTradeInBbg(wsBbg As Worksheet, tradeID As String) As Long
    Dim lastRowRaw As Long
    lastRowRaw = wsBbg.Range("A1").CurrentRegion.Rows.Count
    
    Dim j As Long
    Dim cellValue As Variant
    
    For j = 1 To lastRowRaw
        ' Get cell value with error handling
        On Error Resume Next
        cellValue = wsBbg.Range("A" & j).Value
        On Error GoTo 0
        
        ' Skip if cell has error or is empty
        If Not IsError(cellValue) And Not IsEmpty(cellValue) Then
            ' Convert both to strings for comparison
            If CStr(cellValue) = tradeID Then
                FindTradeInBbg = j
                Exit Function
            End If
        End If
    Next j
    
    FindTradeInBbg = 0  ' Return 0 if not found
End Function

Public Sub MapTradeToOrderGen(wsRaw As Worksheet, rawRow As Long, _
                             wsOrderGen As Worksheet, orderGenRow As Long)
    ' Fixed values
    wsOrderGen.Range("A" & orderGenRow).Value = 8471863
    wsOrderGen.Range("B" & orderGenRow).Value = "LIMIT"
    wsOrderGen.Range("C" & orderGenRow).Value = wsRaw.Range("J" & rawRow).Value
    
    ' Map each leg
    Call MapLegToOrderGen(wsRaw, rawRow, wsOrderGen, orderGenRow, 1, "D") ' Leg 1
    Call MapLegToOrderGen(wsRaw, rawRow, wsOrderGen, orderGenRow, 2, "G") ' Leg 2
    Call MapLegToOrderGen(wsRaw, rawRow, wsOrderGen, orderGenRow, 3, "J") ' Leg 3
    Call MapLegToOrderGen(wsRaw, rawRow, wsOrderGen, orderGenRow, 4, "M") ' Leg 4
End Sub

Public Sub MapLegToOrderGen(wsRaw As Worksheet, rawRow As Long, _
                           wsOrderGen As Worksheet, orderGenRow As Long, _
                           legNumber As Long, startColumn As String)
    Dim col As Long
    col = Range(startColumn & 1).Column
    
    wsOrderGen.Cells(orderGenRow, col).Value = CreateOptionSymbol(wsRaw, rawRow, legNumber)
    wsOrderGen.Cells(orderGenRow, col + 1).Value = GetLegQuantity(wsRaw, rawRow, legNumber)
    wsOrderGen.Cells(orderGenRow, col + 2).Value = "OPEN"
End Sub

Public Function GetLegQuantity(ws As Worksheet, row As Long, legNumber As Long) As Long
    ' Get quantity column for each leg
    Select Case legNumber
        Case 1: GetLegQuantity = ws.Range("O" & row).Value  ' L1 Qty
        Case 2: GetLegQuantity = ws.Range("U" & row).Value  ' L2 Qty
        Case 3: GetLegQuantity = ws.Range("AA" & row).Value  ' L3 Qty
        Case 4: GetLegQuantity = ws.Range("AG" & row).Value ' L4 Qty
    End Select
End Function


' ===================================================================
' ORDER BLOCKING LOGIC
' ===================================================================

' Add this function to modTradeProcessing.bas
Public Sub ExportOrderTrackingForBlock(blockedOrder As clsBlockedOrder, wsRaw As Worksheet, approvedTrades As Collection, boxKey As String, expiryDate As Date)
    On Error GoTo ErrorHandler
    
    ' Ensure database connection
    If Not modSQLConnections.EnsureConnection() Then Exit Sub
    
    Dim dbConn As ADODB.Connection
    Set dbConn = modSQLConnections.GetConnection()
    If dbConn Is Nothing Then Exit Sub
    
    ' Loop through all approved trades and find those that match this block
    Dim i As Long
    For i = 1 To approvedTrades.Count
        Dim tradeData As Variant
        tradeData = Split(approvedTrades(i), "|")
        Dim rawRow As Long
        Dim tradeID As String
        rawRow = CLng(tradeData(0))
        tradeID = tradeData(1)
        
        ' Check if this trade belongs to this block
        Dim tradeBoxKey As String
        tradeBoxKey = CreateBoxKeyFromRaw(wsRaw, rawRow)
        Dim tradeExpiry As Date
        tradeExpiry = wsRaw.Range("K" & rawRow).Value
        
        If tradeBoxKey = boxKey And tradeExpiry = expiryDate Then
            ' This trade is part of this block - create tracking record
            Dim sql As String
            Dim clientName As String
            Dim clientAccount As String
            Dim individualQty As Long
            Dim limitPct As Double
            
            clientName = wsRaw.Range("B" & rawRow).Value
            clientAccount = wsRaw.Range("C" & rawRow).Value
            individualQty = wsRaw.Range("O" & rawRow).Value
            limitPct = wsRaw.Range("J" & rawRow).Value
            
            ' Escape single quotes
            clientName = Replace(clientName, "'", "''")
            clientAccount = Replace(clientAccount, "'", "''")
            tradeID = Replace(tradeID, "'", "''")
            
            sql = "INSERT INTO " & Range("table_order_tracking").Value & " " & _
                  "(allocation_id, synthetic_borrow_app_id, client_name, client_account, " & _
                  "individual_quantity, total_blocked_quantity, " & _
                  "expiry_date, limit, created_at) " & _
                  "VALUES (" & _
                  "'" & blockedOrder.orderID & "', " & _
                  "'" & tradeID & "', " & _
                  "'" & clientName & "', " & _
                  "'" & clientAccount & "', " & _
                  individualQty & ", " & _
                  blockedOrder.totalQuantity & ", " & _
                  "'" & Format(expiryDate, "yyyy-mm-dd") & "', " & _
                  limitPct & ", " & _
                  "NOW())"
            
            dbConn.Execute sql
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error exporting order tracking: " & Err.description
End Sub



Public Function GetSignedQuantity(action As String, quantity As Long) As Long
    ' Apply sign based on action
    If UCase(action) = "BUY" Then
        GetSignedQuantity = quantity  ' Keep positive
    ElseIf UCase(action) = "SELL" Then
        GetSignedQuantity = -quantity  ' Make negative
    Else
        GetSignedQuantity = quantity  ' Default to positive if action unclear
    End If
End Function



Public Function CreateOptionSymbolFromBlock(blockedOrder As clsBlockedOrder, legNumber As Long) As String
    ' Create option symbol in format: .SPXYYMMDDCXXXXX
    ' Example: .SPX251017C6000
    
    Dim optionType As String
    Dim strike As Double
    
    Select Case legNumber
        Case 1
            optionType = blockedOrder.Leg1OptionType
            strike = blockedOrder.leg1Strike
        Case 2
            optionType = blockedOrder.Leg2OptionType
            strike = blockedOrder.leg2Strike
        Case 3
            optionType = blockedOrder.Leg3OptionType
            strike = blockedOrder.leg3Strike
        Case 4
            optionType = blockedOrder.Leg4OptionType
            strike = blockedOrder.leg4Strike
    End Select
    
    ' Format expiry date as YYMMDD
    Dim formattedExpiry As String
    formattedExpiry = Format(blockedOrder.expiryDate, "yymmdd")
    
    ' Option letter
    Dim optionLetter As String
    If UCase(optionType) = "CALL" Then
        optionLetter = "C"
    Else
        optionLetter = "P"
    End If
    
    ' Format strike - remove decimal and pad to 5 digits
    Dim strikeFormatted As String
    strikeFormatted = Format(strike, "00000")
    
    ' Build symbol: .SPXYYMMDDCXXXXX
    CreateOptionSymbolFromBlock = ".SPX" & formattedExpiry & optionLetter & strikeFormatted
End Function

