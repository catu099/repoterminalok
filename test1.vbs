Sub credit_charges()
 
 Dim bank As String
 Dim total As Double
 Dim summary As Integer
 
summary = 2
 
For b = 2 To 101

  If Cells(b, 1).Value <> Cells(b + 1, 1).Value Then
  
    bank = Cells(b, 1).Value
    total = total + Cells(b, 3).Value
    Cells(summary, 7).Value = bank
    Cells(summary, 8).Value = total
    
    summary = summary + 1
    
    total = 0
        
    
 Else
 
 total = total + Cells(b, 3).Value
    
 
 End If
 
 Next b


End Sub
