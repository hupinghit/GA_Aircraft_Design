Function wind_rating(V As Single) As Single
'根据风速来判断风速等级

    If V < 0 Then
        wind_rating = -1
    ElseIf V >= 0 And V < 0.3 Then
        wind_rating = 0
        
    ElseIf V >= 0.3 And V < 1.6 Then
        wind_rating = 1
        
    ElseIf V >= 1.6 And V < 3.4 Then
        wind_rating = 2
    
    ElseIf V >= 3.4 And V < 5.5 Then
        wind_rating = 3
        
    ElseIf V >= 5.5 And V < 8 Then
        wind_rating = 4
        
    ElseIf V >= 8 And V < 10.8 Then
        wind_rating = 5
        
    ElseIf V >= 10.8 And V < 13.9 Then
        wind_rating = 6
        
    ElseIf V >= 13.9 And V < 17.2 Then
        wind_rating = 7
        
    ElseIf V >= 17.2 And V < 20.8 Then
        wind_rating = 8
        
    ElseIf V >= 20.8 And V < 24.5 Then
        wind_rating = 9
        
    ElseIf V >= 24.5 And V < 28.5 Then
        wind_rating = 10
        
    ElseIf V >= 28.5 And V < 32.7 Then
        wind_rating = 11
        
    ElseIf V >= 32.7 And V < 37 Then
        wind_rating = 12
        
    ElseIf V >= 37 And V < 41.5 Then
        wind_rating = 13
        
    ElseIf V >= 41.5 And V < 46.2 Then
        wind_rating = 14
        
    ElseIf V >= 46.2 And V < 51 Then
        wind_rating = 15
        
    ElseIf V >= 51 And V < 56.1 Then
        wind_rating = 16
        
    ElseIf V >= 56.1 Then
        wind_rating = 17
        
        
    End If
End Function
