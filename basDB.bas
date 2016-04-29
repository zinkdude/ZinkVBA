Public conn As ADODB.Connection
Public rs As ADODB.Recordset
Public AsOfDate As String

Global Const DB_DEFAULT = "yourSQLdb"
Global Const QRY_SEC = &H0
Global Const START_ROW = 2
Global Const MAXROWS = 5000

Public Sub OpenDB()
    Set conn = New Connection
    conn.ConnectionString = "Driver={SQL Server};Server=" & "BPDB1" & ";Database=" & DB_DEFAULT & ";Uid=" & "xxx" & ";Pwd=" & "xxx" & ";"
    conn.Open
    conn.CommandTimeout = 120
   Set rs = New ADODB.Recordset
End Sub

Public Function doQry(qryType As Integer, Optional FName As String, Optional param1 As Variant, Optional param2 As Variant, Optional param3 As Variant) As Boolean
    Dim qrystr As String
    Dim q1 As String
    Dim q2 As String
    Dim q3 As String
    Dim q4 As String
    Dim q5 As String
    
    TailCond = " "
    On Error GoTo ErrHandler
                
    If conn Is Nothing Then
        Call OpenDB
    End If
    
    'Query the AccessDB
    Select Case qryType
        Case SEC
            q1 = ""
        Case Else
            qrystr = ""
    End Select
    
    qrystr = q1 + q2 + q3 + q4 + q5
    If qrystr <> "" Then
        
        Set rs = conn.Execute(qrystr, , adCmdText)
        
        If rs.EOF Then doQry = False Else doQry = True
    End If
    
    Exit Function
ErrHandler:
        doQry = False
End Function
Public Sub CloseDB()
    On Error Resume Next

    rs.Close
    Set rs = Nothing

End Sub


Public Function doStoredProc(qryType As Integer, Optional param1 As Variant, Optional param2 As Variant, Optional param3 As Variant, Optional param4 As Variant) As Boolean
    'do Stored Procedure into SQL
    
    Dim adoCmd As Command
    Dim adoParam1 As New ADODB.Parameter
    Dim adoParam2 As New ADODB.Parameter
    Dim QStr As String
    
    On Error GoTo ErrHandler
    
    If conn Is Nothing Then
        Call OpenDB
    End If
        
    Select Case qryType
        Case 1
                    
            Set adoCmd = New Command
            adoCmd.CommandText = "spNameOfStoredProc"
            adoCmd.CommandType = adCmdStoredProc
            adoCmd.ActiveConnection = conn  'Conn should have been previously declared
            adoCmd.Parameters.Refresh
            adoCmd.Parameters(3) = param1
            adoCmd.CommandTimeout = 300
            Set rs = adoCmd.Execute
            
        Case 2
            Set adoCmd = New Command
            adoCmd.CommandText = "spNameofStoredProc2"
            adoCmd.CommandType = adCmdStoredProc
            adoCmd.ActiveConnection = conn
            
            Set adoParam1 = adoCmd.CreateParameter("@sym", adVarChar, adParamInput, Len(param1), param1)
            Set adoParam2 = adoCmd.CreateParameter("@orig_cost_per_share", adSingle, adParamOutput, 25)
            adoCmd.Parameters.Append adoParam1
            adoCmd.Parameters.Append adoParam2
            adoCmd.CommandTimeout = 240
            Set rs = adoCmd.Execute
            
    End Select
        
    If rs.EOF Then doStoredProc = False Else doStoredProc = True
    
    doStoredProc = True
    Exit Function
ErrHandler:
    doStoredProc = False
End Function


Public Sub CloseDB()
    On Error Resume Next

    rs.Close
    db.Close
     
    'InBook.SaveCopyAs (SaveFile)
    
    Set rs = Nothing
    Set db = Nothing

End Sub
