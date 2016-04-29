Public Function IsPreMarket() As Boolean
    '---------------------------------------------------------------
    '   Check to see if it is currently before trading hours.
    '   < 9:30
    '   Return True if it is
    '---------------------------------------------------------------
    Dim ThisTime As String
    Dim ThisHour As Integer
    
    ThisTime = Format(Time, "hh:mm")
    ThisHour = Val(Left$(ThisTime, 2))
    
    If ThisHour < 9 Then
        IsPreMarket = True
    ElseIf ThisHour = 9 Then
        If Val(Right$(ThisTime, 2)) > 30 Then IsPreMarket = False Else IsPreMarket = True
    Else
        IsPreMarket = False
    End If

End Function
