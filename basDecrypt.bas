Public Function DecryptFile(efile As String) As String
    '------------------------------------------------------------------------------------
    'Use gpg2 freeware decrpyion
    '   http://gpg4win.org/
    '   Using gnupg 2.0.17, gpgwin 2.1.0
    '   Example Output Line : 'C:\PROGRA~1\GNU\GnuPG\gpg --passphrase "my passphrase" --batch -o "myfile.csv" -d "c:\reports\myencryptedfile.asc"
    '
    '       Updated 9/29/14 to use the gpg2 command line with quotes around the file name
    '------------------------------------------------------------------------------------
    On Error GoTo ErrHandler
    Dim exec_line As String
    Dim exename As String
    Dim RetVal As Variant
    Const GPG_PASSPHRASE = "my passphrase"
        
    exename = "C:\gpg\gpg2.exe " 
    exec_line = exename & "--passphrase " & Chr$(34) & GPG_PASSPHRASE & Chr$(34) & " --batch -o " & Chr$(34) & Left$(efile, Len(efile) - 7) + "txt" & Chr$(34) & " -d " & Chr$(34) & efile & Chr$(34)
    
    RetVal = Shell(exec_line, vbNormalFocus)
    Call PAUSE(4)
    DecryptFile = Left$(efile, Len(efile) - 7) + "txt"
    Exit Function
ErrHandler:
    DecryptFile = ""
End Function
