Private Function TrueTrim(tt1 As String) As String
    'Take out leading "TAB" character
    Dim i As Integer
    
    i = 1
    While Mid$(tt1, i, 1) = Chr$(32) Or Mid$(tt1, i, 1) = Chr$(9)
        i = i + 1
    Wend
    
    TrueTrim = Right(tt1, Len(tt1) - i + 1)
End Function


Private Sub AddChart()
    '-----------------------------------------------
    'Add simple pie chart for  Equity %
    'Only works on Excel 2007 - need to check out.
    '-----------------------------------------------
    With wksCR
        .Range("C16:G16").Select
        .Shapes.AddChart.Select
        ActiveChart.ChartType = xlPie
        ActiveChart.SeriesCollection(1).XValues = "={""PB""}"
        ActiveChart.SetSourceData Source:=Range("C16,E16,G16")
        ActiveChart.SeriesCollection(1).XValues = "={""CS"",""GS"",""MS""}"
        ActiveChart.SeriesCollection(1).Name = "="" % Equity"""
   
        With ActiveChart.Parent
            .Left = 715
            .Width = 125
            .Top = 175
            .Height = 105
        End With
        .Range("A3").Select
    End With
End Sub


Public Function FileExists(fname) As Boolean
    '   Returns TRUE if the file exists
    Dim x As String
    On Error Resume Next
    x = Dir(fname)
    If x <> "" Then FileExists = True _
        Else FileExists = False
End Function


Public Function OpenXLSFile(infile As String, IsReadOnly As Boolean) As Boolean
    On Error GoTo ErrHandler
    
    Set InBook = Workbooks.Open(infile, False, IsReadOnly, , , , True, , , , , , False)
    OpenXLSFile = True
    Exit Function
ErrHandler:
    OpenXLSFile = False
    Call ErrProc(OPEN_XLS_ERR, OK_ONLY, Err.description)
    
End Function

Public Function OpenTextFile(infile As String, IsReadOnly As Boolean) As Boolean
    '--------------------------
    '   Load a text file
    '--------------------------
        
OpenFile:
    On Error GoTo ErrHandler
    
    'Open Spreadsheet to Import DataFrom
    'Set InBook = Workbooks.Open(InFile, , IsReadOnly)
    
    Workbooks.OpenText Filename:=infile, Origin:=437, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=True, _
        Space:=False, other:=False, FieldInfo:=Array(Array(1, 2), Array(2, 2), Array( _
        3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10 _
        , 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), _
        Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array( _
        23, 1), Array(24, 1), Array(25, 1), Array(26, 1), Array(27, 1), Array(28, 1), Array(29, 1), _
        Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), Array(35, 1), Array( _
        36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 1), Array(41, 1), Array(42, 1), _
        Array(43, 1), Array(44, 1), Array(45, 1), Array(46, 1), Array(47, 1), Array(48, 1), Array( _
        49, 1), Array(50, 1), Array(51, 1), Array(52, 1), Array(53, 1), Array(54, 1), Array(55, 1), _
        Array(56, 1), Array(57, 1), Array(58, 1), Array(59, 1)), TrailingMinusNumbers:=True
        
        
    Set InBook = Application.ActiveWorkbook
    
    OpenTextFile = True
    Exit Function
    
ErrHandler:
    Call ErrProc(OPEN_FILE_ERR, OK_ONLY, Err.description)
    OpenTextFile = False
Outahere:

    'Workbooks.Open Filename:="O:\MS Reports\options\options0115.csv"
End Function
