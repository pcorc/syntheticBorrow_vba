Attribute VB_Name = "modArchive"
' ===================================================================
' DAILY MODEL ARCHIVING
' ===================================================================

Public Sub ArchiveModelAsValues()
    ' Archives the current model as values-only in a date-stamped file
    On Error GoTo ErrorHandler
    
    ' Confirm with user
    Dim response As VbMsgBoxResult
    response = MsgBox("This will save a values-only archive copy of the model." & vbCrLf & vbCrLf & _
                     "The original working file will be reopened with formulas intact." & vbCrLf & vbCrLf & _
                     "Do you want to proceed?", _
                     vbYesNo + vbQuestion, "Archive Model")
    
    If response = vbNo Then Exit Sub
    
    ' Store original workbook path before SaveAs
    Dim originalPath As String
    originalPath = ThisWorkbook.FullName
    
    ' Disable screen updating for speed
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Get archive path and filename from named ranges
    Dim archivePath As String
    Dim archiveFileName As String
    Dim fullArchivePath As String
    
    On Error Resume Next
    archivePath = Range("archive_model_path").Value
    archiveFileName = Range("archive_model_file").Value
    On Error GoTo ErrorHandler
    
    ' Validate named ranges exist
    If archivePath = "" Then
        MsgBox "Named range 'archive_model_path' not found or empty", vbCritical
        GoTo Cleanup
    End If
    
    If archiveFileName = "" Then
        MsgBox "Named range 'archive_model_file' not found or empty", vbCritical
        GoTo Cleanup
    End If
    
    ' Ensure path ends with backslash
    If Right(archivePath, 1) <> "\" Then
        archivePath = archivePath & "\"
    End If
    
    ' Check if directory exists
    If Len(Dir(archivePath, vbDirectory)) = 0 Then
        MsgBox "Archive directory does not exist: " & archivePath & vbCrLf & vbCrLf & _
               "Please create the directory first.", vbCritical, "Directory Not Found"
        GoTo Cleanup
    End If
    
    ' Build full file path
    fullArchivePath = archivePath & archiveFileName
    
    ' Check if file already exists
    If Len(Dir(fullArchivePath)) > 0 Then
        Dim overwriteResponse As VbMsgBoxResult
        overwriteResponse = MsgBox("Archive file already exists:" & vbCrLf & _
                                   fullArchivePath & vbCrLf & vbCrLf & _
                                   "Do you want to overwrite it?", _
                                   vbYesNo + vbExclamation, "File Exists")
        If overwriteResponse = vbNo Then GoTo Cleanup
        
        ' Delete existing file
        Kill fullArchivePath
    End If
    
    ' SaveAs to archive location (this makes the archive the active workbook)
    ThisWorkbook.SaveAs fileName:=fullArchivePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    ' Now convert all formulas to values in THIS workbook (which is now the archive)
    Call ConvertCurrentWorkbookToValues
    
    ' Save the values-only archive
    ThisWorkbook.Save
    
    ' Close the archive file
    ThisWorkbook.Close SaveChanges:=False
    
    ' Reopen the original working file
    Workbooks.Open originalPath
    
    ' Re-enable display
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Success message
    MsgBox "Model archived successfully!" & vbCrLf & vbCrLf & _
           "Archive Location: " & fullArchivePath & vbCrLf & vbCrLf & _
           "All formulas converted to values in archive." & vbCrLf & _
           "Original working file reopened with formulas intact.", _
           vbInformation, "Archive Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error archiving model: " & Err.description, vbCritical, "Archive Error"
    
Cleanup:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


Private Sub ConvertCurrentWorkbookToValues()
    ' Converts all formulas in all sheets of ThisWorkbook to values
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim usedRange As Range
    
    For Each ws In ThisWorkbook.Worksheets
        ' Skip any hidden sheets if needed
        If ws.Visible = xlSheetVisible Then
            
            ' Get the used range
            Set usedRange = ws.usedRange
            
            If Not usedRange Is Nothing Then
                ' Copy and paste values for the entire used range
                usedRange.Copy
                usedRange.PasteSpecial xlPasteValues
                Application.CutCopyMode = False
                
                ' Optional: Remove any data validation
                On Error Resume Next
                usedRange.Validation.Delete
                On Error GoTo 0
            End If
        End If
    Next ws
    
    ' Clear any remaining clipboard
    Application.CutCopyMode = False
    
    On Error GoTo 0
End Sub

