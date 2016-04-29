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
