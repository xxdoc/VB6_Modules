Attribute VB_Name = "Bas_Math"
Public Function Ceiling(ByVal n As Double) As Double
    
    Dim f As Double

    On Error GoTo ErrorHandler
    
    n = CDbl(n)
    
    f = Floor(n)
    
    If f = n Then
        Ceiling = n
        Exit Function
    End If

    Ceiling = CInt(f + 1)

    Exit Function
    
ErrorHandler:
    
    Ceiling = n

End Function

Public Function Floor(ByVal n As Double) As Double
    
    Dim iTmp As Double
    
    On Error GoTo ErrorHandler
        
    n = CDbl(n)

    'Round() rounds up
    iTmp = Round(n)

    'test rounded value against the non rounded value
    'if greater, subtract 1
    If iTmp > n Then iTmp = iTmp - 1

    Floor = CInt(iTmp)

    Exit Function
    
ErrorHandler:

    Floor = n
    
End Function

Private Sub Combinations()
    
   'Subroutine template
   'Does not support value passing
    
   Dim Numbers(6) As Integer
    
  Dim f As Integer
  Dim c As Integer
  Dim b As Integer
  Dim p As Integer
   
  For c = LBound(Numbers()) To UBound(Numbers()) - 2   'combinations
        
        For p = LBound(Numbers()) To UBound(Numbers()) 'place control

            For idx5 = p To p + c 'get base
                If p + c < 6 Then
                    b = b & ", " & Numbers(idx5)
                End If
            Next
            
            For idx2 = idx5 To UBound(Numbers()) Step 1 ' The most inner loop
                
                MsgBox b & "," & Numbers(idx2)
            
            Next
            
            b = ""

        Next
        
    Next
    
End Sub
