Private Function TrueTrim(tt1 As String) As String
    'Take out leading "TAB" character
    Dim i As Integer
    
    i = 1
    While Mid$(tt1, i, 1) = Chr$(32) Or Mid$(tt1, i, 1) = Chr$(9)
        i = i + 1
    Wend
    
    TrueTrim = Right(tt1, Len(tt1) - i + 1)
End Function
