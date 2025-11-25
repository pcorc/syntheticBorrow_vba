Attribute VB_Name = "modSQLConnections"
' ===================================================================
' SQL DATABASE CONNECTIONS MODULE
' Synthetic Borrow Trading System - Version 2
' ===================================================================

Option Explicit

' Global connection object
Public conn As ADODB.Connection

' ===================================================================
' CONNECTION MANAGEMENT
' ===================================================================

Public Function GetConnection() As ADODB.Connection
    If EnsureConnection() Then
        Set GetConnection = conn
    Else
        Set GetConnection = Nothing
    End If
End Function

Public Function OpenSQLConnection() As Boolean
    On Error GoTo ErrorHandler
    
    ' Close existing connection if open
    Call CloseSQLConnection
    
    ' Create new connection
    Set conn = New ADODB.Connection
    
    ' Build connection string from named ranges
    Dim connectionString As String
    connectionString = "DRIVER={MySQL ODBC 8.0 Unicode Driver};" & _
                      "SERVER=" & Range("sql_server").Value & ";" & _
                      "DATABASE=" & Range("sql_database").Value & ";" & _
                      "UID=" & Range("sql_uid").Value & ";" & _
                      "PWD=" & Range("sql_pwd").Value & ";" & _
                      "PORT=" & Range("sql_port").Value & ";" & _
                      "OPTION=3"
    
    ' Open connection
    conn.connectionString = connectionString
    conn.Open
    
    OpenSQLConnection = True
    Exit Function
    
ErrorHandler:
    MsgBox "Failed to connect to SQL database: " & Err.description, vbCritical, "Database Connection Error"
    OpenSQLConnection = False
End Function

Public Sub CloseSQLConnection()
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
        Set conn = Nothing
    End If
End Sub

Public Function EnsureConnection() As Boolean
    If conn Is Nothing Then
        EnsureConnection = OpenSQLConnection()
    ElseIf conn.State <> adStateOpen Then
        EnsureConnection = OpenSQLConnection()
    Else
        EnsureConnection = True
    End If
End Function




Public Sub TestConnection()
    ' Wrapper for testing database connection - for macro assignment
    Dim result As String
    result = TestDatabaseConnection()
    MsgBox result, vbInformation, "Database Connection Test"
End Sub

' ===================================================================
' TESTING AND DIAGNOSTICS
' ===================================================================

Public Function TestDatabaseConnection() As String
    On Error GoTo ErrorHandler
    
    If OpenSQLConnection() Then
        Dim sql As String
        sql = "SELECT COUNT(*) as test_count FROM " & Range("table_synthetic_borrow").Value
        
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open sql, conn
        
        TestDatabaseConnection = "Connection successful!" & vbCrLf & _
                                "Found " & rs("test_count").Value & " records in synthetic_borrow table."
        
        rs.Close
        Set rs = Nothing
        Call CloseSQLConnection
    Else
        TestDatabaseConnection = "Connection failed! Please check database settings."
    End If
    
    Exit Function
    
ErrorHandler:
    TestDatabaseConnection = "Connection test error: " & Err.description
    Call CloseSQLConnection
End Function


' ===================================================================
' WORKBOOK EVENTS
' ===================================================================

Public Sub Workbook_BeforeClose(Cancel As Boolean)
    Call CloseSQLConnection
End Sub

