Private Sub TerminateApp(exe_name As String)
     '---------------------------------------------------------------------------------------
     ' Terminates the exe process specified.
     ' Uses WMI (Windows Management Instrumentation) to query all running processes
     ' then terminates ALL instances of the exe process held in the variable strTerminateThis.
     '---------------------------------------------------------------------------------------
    On Error Resume Next
    
    Dim strTerminateThis As String
    'The variable to hold the process to terminate

    Dim objWMIcimv2 As Object, objProcess As Object, objList As Object
    Dim intError As Integer

    'Process to terminate – you could specify and .exe program name here
    strTerminateThis = exe_name

    'Connect to CIMV2 Namespace and then find the .exe process
    Set objWMIcimv2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set objList = objWMIcimv2.ExecQuery("select * from win32_process where name='" & strTerminateThis & "'")
    For Each objProcess In objList
        intError = objProcess.Terminate 'Terminates a process and all of its threads.
        'Return value is 0 for success. Any other number is an error.
        If intError <> 0 Then Exit For
    Next

    'ALL instances of exe (strTerminateThis) have been terminated
    Set objWMIcimv2 = Nothing
    Set objList = Nothing
    Set objProcess = Nothing

End Sub


