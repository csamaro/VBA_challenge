Sub FizzyBuzz():
    For i = 2 To 100
        If (Cells(i, 1).Value Mod 3 = 0 And Cells(i, 1).Value Mod 5 = 0) Then
            Cells(i, 2).Value = "FizzyBuzz"
        ElseIf (Cells(i, 1).Value Mod 3 = 0) Then
            Cells(i, 2).Value = "Fizz"
        ElseIf (Cells(i, 1).Value Mod 5 = 0) Then
            Cells(i, 2).Value = "Buzz"
        End If
    
    Next i
End Sub
