Attribute VB_Name = "modPostExecutionImportBackup"
' ===================================================================
' POST-EXECUTION PROCESSING MODULE
' Synthetic Borrow Trading System
' Handles execution file import, allocation generation, and client recap emails
' ===================================================================

Option Explicit

' ===================================================================
' MAIN POST-EXECUTION WORKFLOW
' ===================================================================

Public Sub ProcessExecutionAndAllocations()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Step 1: Paste execution file from Schwab
    Call PasteExecutionFile
    MsgBox "Execution Import complete. Run python script to ingest, allocate, and report."

    ' Step 2: Import execution file from Schwab
    ' Call ParseExecutionFile

    ' Step 2: Generate allocation file for trading system upload
    'Call GenerateAllocationFile

    ' Step 3: Send trade recap emails to each client
    'Call SendTradeRecapEmails

    ' Step 4: Export execution data to SQL
    'Call ExportExecutionsToSQL

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

'    MsgBox "Post-execution processing complete!" & vbCrLf & _
'           "- Execution file imported" & vbCrLf & _
'           "- Allocation file generated" & vbCrLf & _
'           "- Trade recap emails sent" & vbCrLf & _
'           "- Data exported to SQL", vbInformation
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error in post-execution processing: " & Err.description, vbCritical
End Sub


' ===================================================================
' STEP 1: PASTE RAW EXECUTION FILE
' ===================================================================

Public Sub PasteExecutionFile()
    On Error GoTo ErrorHandler

    Dim executionFilePath As String
    Dim executionFileName As String
    Dim fullPath As String
    
    executionFilePath = Range("trade_execution_path").Value
    executionFileName = Range("schwab_execution_filename").Value
    fullPath = executionFilePath & executionFileName

    ' If named range is empty, prompt user
    If executionFilePath = "" Or Dir(executionFilePath) = "" Then
        executionFilePath = Application.GetOpenFilename("CSV Files (*.csv),*.csv", , "Select Schwab Execution File")
        If executionFilePath = "False" Then Exit Sub
    End If


    ' Clear ExecutionResults tab
    Dim wsExecution As Worksheet
    Set wsExecution = ThisWorkbook.Worksheets("ExecutionResults")
    wsExecution.Cells.ClearContents
    wsExecution.Cells.Interior.colorIndex = xlNone

    ' Add section header
    wsExecution.Range("A1").Value = "RAW EXECUTION FILE DATA"
    wsExecution.Range("A1").Font.Bold = True
    wsExecution.Range("A1").Font.Size = 12
    wsExecution.Range("A1").Interior.Color = RGB(217, 217, 217)

    wsExecution.Range("A2").Value = "File: " & executionFilePath
    wsExecution.Range("A2").Font.Size = 9
    wsExecution.Range("A2").Font.Italic = True

    ' Import raw CSV data starting at row 4
    Call ImportRawCSV(fullPath, wsExecution, 4)

    ' Auto-fit columns
    wsExecution.Columns("A:Z").AutoFit

    ' Navigate to sheet
    wsExecution.Activate
    wsExecution.Range("A1").Select

    Exit Sub

ErrorHandler:
    MsgBox "Error pasting execution file: " & Err.description, vbCritical
End Sub



' ===================================================================
' STEP 1: PASTEEXECUTION FILE
' ===================================================================


Public Sub ImportRawCSV(filePath As String, ws As Worksheet, startRow As Long)
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    fileNum = FreeFile

    Open filePath For Input As #fileNum

    Dim lineText As String
    Dim currentRow As Long
    currentRow = startRow

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText

        ' Split by comma and paste into columns
        Dim dataArray As Variant
        dataArray = Split(lineText, ",")

        ' Paste the array into the row
        Dim col As Long
        For col = 0 To UBound(dataArray)
            ws.Cells(currentRow, col + 1).Value = dataArray(col)
        Next col

        currentRow = currentRow + 1
    Loop

    Close #fileNum

    ' Format the raw data area with light borders
    Dim lastCol As Long
    Dim lastRow As Long
    lastCol = ws.Cells(startRow, ws.Columns.Count).End(xlToLeft).Column
    lastRow = currentRow - 1

    If lastCol > 0 And lastRow >= startRow Then
        With ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, lastCol))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(200, 200, 200)
        End With

        ' Bold the first row (likely headers)
        ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, lastCol)).Font.Bold = True
        ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, lastCol)).Interior.Color = RGB(242, 242, 242)
    End If

    Exit Sub

ErrorHandler:
    If fileNum > 0 Then Close #fileNum
    MsgBox "Error importing raw CSV: " & Err.description, vbCritical
End Sub
'
'
'' ===================================================================
'' STEP 2: IMPORT EXECUTION FILE
'' ===================================================================
'
'Public Sub ParseExecutionFile()
'    On Error GoTo ErrorHandler
'
'    ' Prompt user to select execution file
'    Dim executionFilePath As String
'    executionFilePath = Range("schwab_execution_filename").Value
'
'    ' If named range is empty, prompt user
'    If executionFilePath = "" Or Dir(executionFilePath) = "" Then
'        executionFilePath = Application.GetOpenFilename("CSV Files (*.csv),*.csv", , "Select Schwab Execution File")
'        If executionFilePath = "False" Then Exit Sub
'    End If
'
'    ' Clear ExecutionResults tab
'    Dim wsExecution As Worksheet
'    Set wsExecution = ThisWorkbook.Worksheets("ExecutionResults")
'    wsExecution.Range("A4:Z1000").ClearContents
'
'    ' Parse execution file
'    Call ParseSchwabExecutionFile(executionFilePath, wsExecution)
'
'    ' Match with OrderGen data
'    Call MatchExecutionWithOrders(wsExecution)
'
'    MsgBox "Execution file imported successfully!", vbInformation
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "Error importing execution file: " & Err.description, vbCritical
'End Sub
'
'Private Sub ParseSchwabExecutionFile(filePath As String, ws As Worksheet)
'    ' Parse Schwab execution file and extract Fill Summary data
'
'    Dim fileNum As Integer
'    fileNum = FreeFile
'
'    Open filePath For Input As #fileNum
'
'    Dim lineText As String
'    Dim inFillSummary As Boolean
'    Dim headerParsed As Boolean
'    Dim rowNum As Long
'    rowNum = 4
'
'    ' Add headers to ExecutionResults
'    Call AddExecutionResultsHeaders(ws)
'
'    Do While Not EOF(fileNum)
'        Line Input #fileNum, lineText
'
'        ' Check if we've entered Fill Summary section
'        If InStr(lineText, "Fill Summary") > 0 Then
'            inFillSummary = True
'            headerParsed = False
'        ElseIf inFillSummary And Not headerParsed And InStr(lineText, "Symbol,Description") > 0 Then
'            headerParsed = True ' Skip header row
'        ElseIf inFillSummary And headerParsed And Trim(lineText) <> "" And Left(Trim(lineText), 1) <> "," Then
'            ' Parse fill summary data line
'            Dim dataArray As Variant
'            dataArray = Split(lineText, ",")
'
'            If UBound(dataArray) >= 7 Then
'                ' Extract relevant data
'                Dim symbol As String, description As String
'                Dim qtyBought As Variant, avgBought As Variant
'                Dim qtySold As Variant, avgSold As Variant
'
'                symbol = Trim(dataArray(0))
'                description = Trim(dataArray(1))
'                qtyBought = IIf(Trim(dataArray(2)) = "", 0, CDbl(Trim(dataArray(2))))
'                avgBought = IIf(Trim(dataArray(3)) = "", 0, CDbl(Trim(dataArray(3))))
'                qtySold = IIf(Trim(dataArray(4)) = "", 0, CDbl(Trim(dataArray(4))))
'                avgSold = IIf(Trim(dataArray(5)) = "", 0, CDbl(Trim(dataArray(5))))
'
'                If symbol <> "" Then
'                    ' Populate ExecutionResults row
'                    ws.Range("A" & rowNum).Value = symbol ' Symbol
'                    ws.Range("B" & rowNum).Value = description ' Description
'                    ws.Range("C" & rowNum).Value = qtyBought ' Qty Bought
'                    ws.Range("D" & rowNum).Value = avgBought ' Avg Price Bought
'                    ws.Range("E" & rowNum).Value = qtySold ' Qty Sold
'                    ws.Range("F" & rowNum).Value = avgSold ' Avg Price Sold
'                    ws.Range("G" & rowNum).Value = Format(Date, "yyyy-mm-dd") ' Execution Date
'
'                    rowNum = rowNum + 1
'                End If
'            End If
'        ElseIf inFillSummary And Trim(lineText) = "" Then
'            ' End of Fill Summary section
'            Exit Do
'        End If
'    Loop
'
'    Close #fileNum
'End Sub
'
'Private Sub AddExecutionResultsHeaders(ws As Worksheet)
'    Dim headers As Variant
'    headers = Array("Symbol", "Description", "Qty Bought", "Avg $ Bought", _
'                   "Qty Sold", "Avg $ Sold", "Execution Date", "Order ID", _
'                   "Client Account", "Client Name", "Leg Number", "Leg Action", _
'                   "Leg Strike", "Leg Type", "Net Premium", "Trade ID")
'
'    ws.Range("A3").Resize(1, UBound(headers) + 1).Value = headers
'    ws.Range("A3").Resize(1, UBound(headers) + 1).Font.Bold = True
'End Sub
'
'Private Sub MatchExecutionWithOrders(wsExecution As Worksheet)
'    ' Match execution data with OrderGen data to populate client details
'
'    Dim wsOrderGen As Worksheet
'    Set wsOrderGen = ThisWorkbook.Worksheets("OrderGen")
'
'    Dim lastExecRow As Long, lastOrderRow As Long
'    lastExecRow = wsExecution.Range("A4").End(xlDown).row
'    lastOrderRow = wsOrderGen.Range("A4").End(xlDown).row
'
'    If lastExecRow < 4 Or lastOrderRow < 4 Then Exit Sub
'
'    ' For each execution row, find matching order
'    Dim i As Long
'    For i = 4 To lastExecRow
'        Dim execDescription As String
'        execDescription = wsExecution.Range("B" & i).Value
'
'        ' Parse description to extract expiry date and strike
'        Dim expiryDate As Date, strike As Double, optionType As String
'        Call ParseOptionDescription(execDescription, expiryDate, strike, optionType)
'
'        ' Find matching order in OrderGen
'        Dim j As Long
'        For j = 4 To lastOrderRow
'            ' Check if strikes and expiry match
'            Dim orderExpiry As Date
'            orderExpiry = wsOrderGen.Range("K" & j).Value
'
'            If DateSerial(year(orderExpiry), Month(orderExpiry), day(orderExpiry)) = _
'               DateSerial(year(expiryDate), Month(expiryDate), day(expiryDate)) Then
'
'                ' Check if any leg matches this execution
'                If MatchesLeg(wsOrderGen, j, strike, optionType) Then
'                    ' Populate execution details from order
'                    wsExecution.Range("H" & i).Value = wsOrderGen.Range("AL" & j).Value ' Order ID
'                    wsExecution.Range("I" & i).Value = wsOrderGen.Range("C" & j).Value ' Client Account
'                    wsExecution.Range("J" & i).Value = wsOrderGen.Range("B" & j).Value ' Client Name
'                    wsExecution.Range("P" & i).Value = wsOrderGen.Range("A" & j).Value ' Trade ID (synthetic_borrow_app_id)
'
'                    Exit For
'                End If
'            End If
'        Next j
'    Next i
'End Sub
'
'Private Sub ParseOptionDescription(description As String, ByRef expiryDate As Date, ByRef strike As Double, ByRef optionType As String)
'    ' Parse "May15 26 5950 P" format
'    On Error Resume Next
'
'    Dim parts() As String
'    parts = Split(Trim(description), " ")
'
'    If UBound(parts) >= 3 Then
'        ' Extract expiry (e.g., "May15 26")
'        Dim monthDay As String, year As String
'        monthDay = parts(0) ' "May15"
'        year = "20" & parts(1) ' "2026"
'
'        ' Parse month and day
'        Dim monthName As String, day As String
'        monthName = Left(monthDay, Len(monthDay) - 2)
'        day = Right(monthDay, 2)
'
'        expiryDate = DateValue(monthName & " " & day & ", " & year)
'
'        ' Extract strike
'        strike = CDbl(parts(2))
'
'        ' Extract option type
'        optionType = UCase(Trim(parts(3)))
'    End If
'
'    On Error GoTo 0
'End Sub
'
'Private Function MatchesLeg(ws As Worksheet, row As Long, strike As Double, optionType As String) As Boolean
'    ' Check if any leg in this order matches the strike and type
'
'    Dim legStrikes As Variant
'    Dim legTypes As Variant
'
'    ' Leg 1: Column Q (strike), Column P (type)
'    ' Leg 2: Column V (strike), Column U (type)
'    ' Leg 3: Column Z (strike), Column Y (type)
'    ' Leg 4: Column AD (strike), Column AC (type)
'
'    legStrikes = Array(ws.Range("Q" & row).Value, ws.Range("V" & row).Value, _
'                      ws.Range("Z" & row).Value, ws.Range("AD" & row).Value)
'    legTypes = Array(ws.Range("P" & row).Value, ws.Range("U" & row).Value, _
'                    ws.Range("Y" & row).Value, ws.Range("AC" & row).Value)
'
'    Dim i As Long
'    For i = 0 To 3
'        If Abs(CDbl(legStrikes(i)) - strike) < 0.01 And _
'           UCase(Left(legTypes(i), 1)) = Left(optionType, 1) Then
'            MatchesLeg = True
'            Exit Function
'        End If
'    Next i
'
'    MatchesLeg = False
'End Function
'
'' ===================================================================
'' STEP 2: GENERATE ALLOCATION FILE
'' ===================================================================
'
'Public Sub GenerateAllocationFile()
'    On Error GoTo ErrorHandler
'
'    Application.ScreenUpdating = False
'
'    Dim wsOrderGen As Worksheet
'    Set wsOrderGen = ThisWorkbook.Worksheets("OrderGen")
'
'    Dim lastRow As Long
'    lastRow = wsOrderGen.Range("A4").End(xlDown).row
'
'    If lastRow < 4 Then
'        MsgBox "No order data found for allocation", vbExclamation
'        Exit Sub
'    End If
'
'    ' Load allocation template
'    Dim templatePath As String
'    templatePath = Range("allocation_template_path").Value
'
'    Dim wbAllocation As Workbook
'
'    If templatePath <> "" And Dir(templatePath) <> "" Then
'        Set wbAllocation = Workbooks.Open(templatePath)
'    Else
'        Set wbAllocation = Workbooks.Add
'    End If
'
'    Dim wsAllocation As Worksheet
'    Set wsAllocation = wbAllocation.Worksheets(1)
'
'    ' Clear existing data
'    wsAllocation.Cells.Clear
'
'    ' Add allocation headers
'    Call AddAllocationHeaders(wsAllocation)
'
'    ' Process each trade for allocation
'    Call PopulateAllocationData(wsOrderGen, wsAllocation)
'
'    ' Save allocation file
'    Dim fileName As String
'    fileName = Range("file_directory").Value & Range("export_allocation_filename").Value
'
'    wbAllocation.SaveAs fileName:=fileName, FileFormat:=xlCSV
'    wbAllocation.Close SaveChanges:=False
'
'    Application.ScreenUpdating = True
'
'    MsgBox "Allocation file generated: " & fileName, vbInformation
'    Exit Sub
'
'ErrorHandler:
'    Application.ScreenUpdating = True
'    If Not wbAllocation Is Nothing Then wbAllocation.Close SaveChanges:=False
'    MsgBox "Error generating allocation file: " & Err.description, vbCritical
'End Sub
'
'Private Sub AddAllocationHeaders(ws As Worksheet)
'    Dim headers As Variant
'    headers = Array("#ORDER IDS", "#ACCOUNT", "ORDER TYPE", "LIMIT PRICE", _
'                   "LEG 1 SYMBOL", "LEG 1 QUANTITY", "LEG 1 OPEN CLOSE", _
'                   "LEG 2 SYMBOL", "LEG 2 QUANTITY", "LEG 2 OPEN CLOSE", _
'                   "LEG 3 SYMBOL", "LEG 3 QUANTITY", "LEG 3 OPEN CLOSE", _
'                   "LEG 4 SYMBOL", "LEG 4 QUANTITY", "LEG 4 OPEN CLOSE")
'
'    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
'End Sub
'
'Private Sub PopulateAllocationData(wsOrderGen As Worksheet, wsAllocation As Worksheet)
'    ' Traverse OrderGen and create allocation rows
'    ' Unblock trades if they were blocked, otherwise 1-to-1 allocation
'
'    Dim lastRow As Long
'    lastRow = wsOrderGen.Range("A4").End(xlDown).row
'
'    Dim allocRow As Long
'    allocRow = 2
'
'    ' Dictionary to track processed block IDs
'    Dim processedBlocks As Object
'    Set processedBlocks = CreateObject("Scripting.Dictionary")
'
'    Dim i As Long
'    For i = 4 To lastRow
'        If wsOrderGen.Range("A" & i).Value <> "" Then
'            Dim blockID As String
'            blockID = wsOrderGen.Range("AL" & i).Value ' Block ID column
'
'            If blockID <> "" And Not processedBlocks.Exists(blockID) Then
'                ' This is a blocked order - need to allocate to multiple accounts
'                Call AllocateBlockedOrder(wsOrderGen, wsAllocation, blockID, allocRow)
'                processedBlocks.Add blockID, True
'            ElseIf blockID = "" Then
'                ' This is a standalone order - 1-to-1 allocation
'                Call AllocateSingleOrder(wsOrderGen, wsAllocation, i, allocRow)
'                allocRow = allocRow + 1
'            End If
'        End If
'    Next i
'End Sub
'
'Private Sub AllocateBlockedOrder(wsOrderGen As Worksheet, wsAllocation As Worksheet, blockID As String, ByRef allocRow As Long)
'    ' Find all trades with this block ID and create separate allocation rows
'
'    Dim lastRow As Long
'    lastRow = wsOrderGen.Range("A4").End(xlDown).row
'
'    Dim i As Long
'    For i = 4 To lastRow
'        If wsOrderGen.Range("AL" & i).Value = blockID Then
'            ' Create allocation row for this client
'            Call AllocateSingleOrder(wsOrderGen, wsAllocation, i, allocRow)
'            allocRow = allocRow + 1
'        End If
'    Next i
'End Sub
'
'Private Sub AllocateSingleOrder(wsOrderGen As Worksheet, wsAllocation As Worksheet, orderRow As Long, allocRow As Long)
'    ' Create allocation row for a single trade
'
'    ' #ORDER IDS - always master account
'    wsAllocation.Range("A" & allocRow).Value = "8471863"
'
'    ' #ACCOUNT - client account from OrderGen
'    wsAllocation.Range("B" & allocRow).Value = wsOrderGen.Range("C" & orderRow).Value
'
'    ' ORDER TYPE
'    wsAllocation.Range("C" & allocRow).Value = "LIMIT"
'
'    ' LIMIT PRICE - Maximum Limit %
'    wsAllocation.Range("D" & allocRow).Value = wsOrderGen.Range("J" & orderRow).Value
'
'    ' LEG 1
'    wsAllocation.Range("E" & allocRow).Value = CreateOCCSymbol(wsOrderGen, orderRow, 1)
'    wsAllocation.Range("F" & allocRow).Value = wsOrderGen.Range("O" & orderRow).Value ' Quantity
'    wsAllocation.Range("G" & allocRow).Value = "OPEN"
'
'    ' LEG 2
'    wsAllocation.Range("H" & allocRow).Value = CreateOCCSymbol(wsOrderGen, orderRow, 2)
'    wsAllocation.Range("I" & allocRow).Value = -wsOrderGen.Range("T" & orderRow).Value ' Negative for SELL
'    wsAllocation.Range("J" & allocRow).Value = "OPEN"
'
'    ' LEG 3
'    wsAllocation.Range("K" & allocRow).Value = CreateOCCSymbol(wsOrderGen, orderRow, 3)
'    wsAllocation.Range("L" & allocRow).Value = wsOrderGen.Range("X" & orderRow).Value ' Quantity
'    wsAllocation.Range("M" & allocRow).Value = "OPEN"
'
'    ' LEG 4
'    wsAllocation.Range("N" & allocRow).Value = CreateOCCSymbol(wsOrderGen, orderRow, 4)
'    wsAllocation.Range("O" & allocRow).Value = -wsOrderGen.Range("AB" & orderRow).Value ' Negative for SELL
'    wsAllocation.Range("P" & allocRow).Value = "OPEN"
'End Sub
'
'Private Function CreateOCCSymbol(ws As Worksheet, row As Long, legNumber As Long) As String
'    ' Create OCC option symbol format: .SPX YYMMDD C/P Strike
'    ' Example: .SPX260515C07000
'
'    Dim baseOffset As Long
'    Select Case legNumber
'        Case 1: baseOffset = 0  ' Columns N-R
'        Case 2: baseOffset = 5  ' Columns S-W
'        Case 3: baseOffset = 10 ' Columns X-AB
'        Case 4: baseOffset = 15 ' Columns AC-AG
'    End Select
'
'    Dim expiryDate As Date
'    Dim optionType As String
'    Dim strike As Double
'
'    expiryDate = ws.Range("K" & row).Value ' Expiry Date
'    optionType = ws.Cells(row, 16 + baseOffset).Value ' Option Type (Call/Put)
'    strike = ws.Cells(row, 17 + baseOffset).Value ' Strike
'
'    ' Format: .SPXYYMMDDCSTRIKE or .SPXYYMMDDPSTRIKE
'    Dim formattedExpiry As String
'    formattedExpiry = Format(expiryDate, "yymmdd")
'
'    Dim optionLetter As String
'    optionLetter = UCase(Left(optionType, 1))
'
'    ' Format strike with 5 digits (pad with zeros)
'    Dim formattedStrike As String
'    formattedStrike = Format(strike, "00000")
'
'    CreateOCCSymbol = ".SPX" & formattedExpiry & optionLetter & formattedStrike
'End Function
'
'' ===================================================================
'' STEP 3: SEND TRADE RECAP EMAILS
'' ===================================================================
'
'Public Sub SendTradeRecapEmails()
'    On Error GoTo ErrorHandler
'
'    Dim wsRawImport As Worksheet
'    Dim wsExecution As Worksheet
'    Dim wsOrderGen As Worksheet
'
'    Set wsRawImport = ThisWorkbook.Worksheets("RawTradeImport")
'    Set wsExecution = ThisWorkbook.Worksheets("ExecutionResults")
'    Set wsOrderGen = ThisWorkbook.Worksheets("OrderGen")
'
'    Dim lastRow As Long
'    lastRow = wsRawImport.Range("A4").End(xlDown).row
'
'    If lastRow < 4 Then
'        MsgBox "No trade data found for recap emails", vbExclamation
'        Exit Sub
'    End If
'
'    ' Email storage path
'    Dim emailStoragePath As String
'    emailStoragePath = Range("trade_recap_email_path").Value
'
'    ' Process each trade
'    Dim i As Long
'    For i = 4 To lastRow
'        If wsRawImport.Range("A" & i).Value <> "" Then
'            Call SendClientTradeRecap(wsRawImport, wsExecution, wsOrderGen, i, emailStoragePath)
'        End If
'    Next i
'
'    MsgBox "Trade recap emails sent to all clients!", vbInformation
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "Error sending trade recap emails: " & Err.description, vbCritical
'End Sub
'
'Private Sub SendClientTradeRecap(wsRaw As Worksheet, wsExec As Worksheet, wsOrder As Worksheet, row As Long, storagePath As String)
'    On Error GoTo ErrorHandler
'
'    ' Get client information
'    Dim clientEmail As String, clientName As String, accountNumber As String
'    Dim tradeID As String
'
'    clientEmail = wsRaw.Range("D" & row).Value
'    clientName = wsRaw.Range("B" & row).Value
'    accountNumber = wsRaw.Range("C" & row).Value
'    tradeID = wsRaw.Range("A" & row).Value
'
'    ' Calculate trade details from execution data
'    Dim creditToday As Double, paybackAmount As Double, impliedRate As Double
'    Dim expiryDate As Date
'
'    creditToday = wsRaw.Range("G" & row).Value ' Quoted Borrow Amount
'    paybackAmount = wsRaw.Range("H" & row).Value ' Payback Amount
'    impliedRate = wsRaw.Range("I" & row).Value ' Annualized Rate
'    expiryDate = wsRaw.Range("K" & row).Value ' Expiry Date
'
'    ' Build email body
'    Dim emailBody As String
'    emailBody = Range("email_allocation_body").Value
'
'    ' Replace placeholders
'    emailBody = Replace(emailBody, "#ACCOUNT_NUMBER#", accountNumber)
'    emailBody = Replace(emailBody, "#ACCOUNT_NAME#", clientName)
'    emailBody = Replace(emailBody, "#CREDIT_TODAY#", Format(creditToday, "$#,##0.00"))
'    emailBody = Replace(emailBody, "#PAY_BACK_AMOUNT#", Format(paybackAmount, "$#,##0.00"))
'    emailBody = Replace(emailBody, "#IMPLIED_RATE#", Format(impliedRate, "0.00%"))
'    emailBody = Replace(emailBody, "#EXPIRTY#", Format(expiryDate, "mm/dd/yyyy"))
'    emailBody = Replace(emailBody, "#DATE#", Format(Date, "mm/dd/yyyy"))
'
'    ' Build subject
'    Dim emailSubject As String
'    emailSubject = Range("email_allocation_subject").Value
'    emailSubject = Replace(emailSubject, "#DATE#", Format(Date, "mm/dd/yyyy"))
'
'    ' Send email
'    Call SendTradeRecapEmail( _
'        clientEmail, _
'        Range("email_allocation_cc").Value, _
'        emailSubject, _
'        emailBody, _
'        storagePath, _
'        tradeID _
'    )
'
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "Error sending recap for " & clientEmail & ": " & Err.description
'End Sub
'
'Private Sub SendTradeRecapEmail(toEmail As String, ccEmail As String, subject As String, body As String, storagePath As String, tradeID As String)
'    On Error GoTo ErrorHandler
'
'    Dim objOutlook As Object
'    Dim objMail As Object
'
'    Set objOutlook = CreateObject("Outlook.Application")
'    Set objMail = objOutlook.CreateItem(0)
'
'    With objMail
'        .To = toEmail
'        .CC = ccEmail
'        .subject = subject
'        .body = body
'
'        ' Save email to storage path
'        If storagePath <> "" Then
'            Dim fileName As String
'            fileName = storagePath & "\" & tradeID & "_" & Format(Date, "yyyymmdd") & ".msg"
'            .SaveAs fileName
'        End If
'
'        ' Display email (change to .Send to send automatically)
'        .Display
'    End With
'
'    Set objMail = Nothing
'    Set objOutlook = Nothing
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "Error sending email: " & Err.description
'End Sub
'
'' ===================================================================
'' STEP 4: EXPORT EXECUTIONS TO SQL
'' ===================================================================
'
'Public Sub ExportExecutionsToSQL()
'    On Error GoTo ErrorHandler
'
'    If Not modSQLConnections.EnsureConnection() Then Exit Sub
'
'    Dim wsExecution As Worksheet
'    Dim wsRawImport As Worksheet
'
'    Set wsExecution = ThisWorkbook.Worksheets("ExecutionResults")
'    Set wsRawImport = ThisWorkbook.Worksheets("RawTradeImport")
'
'    Dim lastRow As Long
'    lastRow = wsRawImport.Range("A4").End(xlDown).row
'
'    If lastRow < 4 Then
'        MsgBox "No execution data to export", vbExclamation
'        Exit Sub
'    End If
'
'    Dim conn As ADODB.Connection
'    Set conn = modSQLConnections.GetConnection()
'
'    Dim todayStr As String
'    todayStr = Format(Date, "yyyy-mm-dd")
'
'    ' Export each trade execution
'    Dim i As Long
'    For i = 4 To lastRow
'        If wsRawImport.Range("A" & i).Value <> "" Then
'            Dim sql As String
'            sql = BuildExecutionInsertSQL(wsRawImport, wsExecution, i, todayStr)
'
'            conn.Execute sql
'        End If
'    Next i
'
'    MsgBox "Execution data exported to SQL successfully!", vbInformation
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "Error exporting executions to SQL: " & Err.description, vbCritical
'End Sub
'
'Private Function BuildExecutionInsertSQL(wsRaw As Worksheet, wsExec As Worksheet, row As Long, executionDate As String) As String
'    ' Build INSERT statement for trade_executions table
'
'    Dim tradeID As String, clientName As String, clientAccount As String
'    Dim expiryDate As String, executedPremium As Double, paybackAmount As Double
'    Dim annualizedRate As Double
'
'    tradeID = wsRaw.Range("A" & row).Value ' synthetic_borrow_app_id
'    clientName = wsRaw.Range("B" & row).Value
'    clientAccount = wsRaw.Range("C" & row).Value
'    expiryDate = Format(wsRaw.Range("K" & row).Value, "yyyy-mm-dd")
'    executedPremium = wsRaw.Range("G" & row).Value
'    paybackAmount = wsRaw.Range("H" & row).Value
'    annualizedRate = wsRaw.Range("I" & row).Value
'
'    ' Get execution prices from ExecutionResults
'    Dim leg1Price As Double, leg2Price As Double, leg3Price As Double, leg4Price As Double
'    leg1Price = GetLegExecutionPrice(wsExec, tradeID, 1)
'    leg2Price = GetLegExecutionPrice(wsExec, tradeID, 2)
'    leg3Price = GetLegExecutionPrice(wsExec, tradeID, 3)
'    leg4Price = GetLegExecutionPrice(wsExec, tradeID, 4)
'
'    BuildExecutionInsertSQL = "INSERT INTO trade_executions " & _
'        "(synthetic_borrow_app_id, client_name, client_account, execution_date, " & _
'        "expiry_date, executed_premium, payback_amount, annualized_rate, " & _
'        "leg1_execution_price, leg2_execution_price, leg3_execution_price, leg4_execution_price, " & _
'        "status, created_at) VALUES (" & _
'        "'" & tradeID & "', " & _
'        "'" & clientName & "', " & _
'        "'" & clientAccount & "', " & _
'        "'" & executionDate & "', " & _
'        "'" & expiryDate & "', " & _
'        "'" & executedPremium & "', " & _
'        "'" & paybackAmount & "', " & _
'        "'" & annualizedRate & "', " & _
'        "'" & leg1Price & "', " & _
'        "'" & leg2Price & "', " & _
'        "'" & leg3Price & "', " & _
'        "'" & leg4Price & "', " & _
'        "'EXECUTED', NOW()) " & _
'        "ON DUPLICATE KEY UPDATE " & _
'        "execution_date = '" & executionDate & "', " & _
'        "executed_premium = '" & executedPremium & "', " & _
'        "status = 'EXECUTED'"
'End Function
'
'Private Function GetLegExecutionPrice(ws As Worksheet, tradeID As String, legNumber As Long) As Double
'    ' Find execution price for specific leg from ExecutionResults
'
'    Dim lastRow As Long
'    lastRow = ws.Range("A4").End(xlDown).row
'
'    Dim i As Long
'    For i = 4 To lastRow
'        If ws.Range("P" & i).Value = tradeID And ws.Range("K" & i).Value = legNumber Then
'            ' Return average price (bought or sold)
'            If ws.Range("C" & i).Value > 0 Then
'                GetLegExecutionPrice = ws.Range("D" & i).Value ' Bought
'            Else
'                GetLegExecutionPrice = ws.Range("F" & i).Value ' Sold
'            End If
'            Exit Function
'        End If
'    Next i
'
'    GetLegExecutionPrice = 0
'End Function
'


'
'Public Sub FormatExecutionResults(ws As Worksheet)
'    ' Format the ExecutionResults worksheet
'    Dim lastRow As Long
'    lastRow = ws.Range("A4").End(xlDown).row
'
'    If lastRow < 4 Then Exit Sub
'
'    ' Format price columns as currency
'    ws.Range("D4:D" & lastRow).NumberFormat = "$#,##0.00"
'    ws.Range("F4:F" & lastRow).NumberFormat = "$#,##0.00"
'    ws.Range("O4:O" & lastRow).NumberFormat = "$#,##0.00"
'
'    ' Format date column
'    ws.Range("G4:G" & lastRow).NumberFormat = "mm/dd/yyyy"
'
'    ' Auto-fit columns
'    ws.Columns("A:P").AutoFit
'
'    ' Add conditional formatting for matched vs unmatched
'    With ws.Range("I4:I" & lastRow)
'        .FormatConditions.Delete
'        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:=""""""
'        .FormatConditions(1).Interior.Color = RGB(255, 200, 200) ' Light red for unmatched
'    End With
'End Sub

