Attribute VB_Name = "modBloomergValidation"
' ===================================================================
' BLOOMBERG VALIDATION WITH MARGIN CHECK
' ===================================================================


' ===================================================================
' CALCULATE TENOR FOR BBG_VALIDATION COLUMN D
' ===================================================================

Public Sub CalculateTenorForBBGValidation()
    On Error GoTo ErrorHandler
    
    Dim wsBbg As Worksheet
    Set wsBbg = ThisWorkbook.Worksheets("BBG_Validation")
    
    Dim lastRow As Long
    lastRow = wsBbg.Range("A1").CurrentRegion.Rows.Count
    
    If lastRow < 2 Then Exit Sub
    
    Dim todayDate As Date
    todayDate = Range("today").Value ' From named range
    
    Dim i As Long
    For i = 2 To lastRow
        If wsBbg.Range("E" & i).Value <> "" Then ' Expiry Date is in column E
            If IsDate(wsBbg.Range("E" & i).Value) Then
                Dim expiryDate As Date
                expiryDate = wsBbg.Range("E" & i).Value
                
                ' Calculate tenor string
                wsBbg.Range("D" & i).Value = CalculateTenorString(todayDate, expiryDate)
            End If
        End If
    Next i
    
    ' Auto-fit the column
    wsBbg.Columns("D").AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error calculating tenor: " & Err.description, vbCritical
End Sub

Public Function CalculateTenorString(startDate As Date, endDate As Date) As String
    On Error GoTo ErrorHandler
    
    ' Calculate total days
    Dim totalDays As Long
    totalDays = endDate - startDate
    
    If totalDays < 0 Then
        CalculateTenorString = "Expired"
        Exit Function
    ElseIf totalDays = 0 Then
        CalculateTenorString = "Today"
        Exit Function
    End If
    
    ' Calculate months and remaining days
    Dim months As Long
    Dim weeks As Long
    Dim days As Long
    Dim tempDate As Date
    
    ' Calculate full months
    months = 0
    tempDate = startDate
    Do While DateAdd("m", months + 1, startDate) <= endDate
        months = months + 1
    Loop
    
    ' Calculate remaining days after full months
    tempDate = DateAdd("m", months, startDate)
    Dim remainingDays As Long
    remainingDays = endDate - tempDate
    
    ' Convert remaining days to weeks and days
    weeks = Int(remainingDays / 7)
    days = remainingDays Mod 7
    
    ' Build the string
    Dim result As String
    result = ""
    
    If months > 0 Then
        result = months & " month" & IIf(months > 1, "s", "")
    End If
    
    If weeks > 0 Then
        If result <> "" Then result = result & " and "
        result = result & weeks & " week" & IIf(weeks > 1, "s", "")
    End If
    
    ' Only show days if no months and weeks, or if specifically requested
    If months = 0 And weeks = 0 And days > 0 Then
        result = days & " day" & IIf(days > 1, "s", "")
    End If
    
    CalculateTenorString = result
    Exit Function
    
ErrorHandler:
    CalculateTenorString = "Error"
End Function


Public Function GetTreasuryRate() As Double

    On Error Resume Next
    GetTreasuryRate = Range("current_treasury_rate").Value
    
End Function



