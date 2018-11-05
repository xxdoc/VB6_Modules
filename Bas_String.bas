Attribute VB_Name = "Bas_String"
Private Const ModuleName As String = "Bas_String"

Public Function Read_Seperated_Text(ByVal Line As String, _
                                     ByVal Seperator As String, _
                                     ByVal Field As Integer) As String
    
    Dim Start As Integer
    Dim TempLine As String
    Dim idx As Integer
    
    TempLine = Line
    
    For idx = 0 To Field - 1
        
        Start = InStr(1, TempLine, Seperator)
        
        TempLine = Mid(TempLine, Start + 1)
        
        If Start = 0 Then
            Exit Function
        End If
    
    Next
    
    Start = InStr(1, TempLine, Seperator)
    If Start > 0 Then
        TempLine = Mid(TempLine, 1, Start - 1)
    End If
    
    Read_Seperated_Text = TempLine

End Function
                                                   
Public Function leading(ByVal data As String, ByVal length As Integer, ByVal leadingcharacter As String) As String
    
    Dim Counter As Integer
    If Len(data) >= length Then
        leading = data
        Exit Function
    End If
    
    Dim Fill As Integer
    Fill = CInt(length) - Len(data)
    For Counter = 1 To Fill
        data = leadingcharacter & data
    Next
    leading = data
    
End Function

Public Function SliceString(ByVal data As String, ByVal slicelength As Integer, ByVal Seperator As String) As String
    
    Dim ReturnString As String
    Dim Index As Integer
    
    If datalength >= Len(data) Then
        SliceString = data
        Exit Function
    End If
    
    For Index = 1 To Len(data) - slicelength Step slicelength
        ReturnString = ReturnString & Mid(data, Index, slicelength) & Seperator
    Next
    
    SliceString = ReturnString & Mid(data, Index)
    
End Function

Public Function FillString(ByVal Str As String, ByVal length As String, ByVal fillcharacter As String) As String
        
    Dim Counter As Integer
    If Len(Str) >= length Then
        FillString = Str
        Exit Function
    End If
    
    Dim Fill As Integer
    Fill = CInt(length) - Len(Str)
    For Counter = 1 To Fill
        Str = Str & fillcharacter
    Next
    FillString = Str
        
End Function

Public Function CenterString(ByVal Str As String, ByVal length As Integer) As String

    Dim Rtn As String
    Dim Fill As Integer
    
    Dim Modulus As Integer
    
    Rtn = LTrim(Str)
    Rtn = Trim(Rtn)
    
    If Len(Rtn) > length Then
        CenterString = Rtn
        Exit Function
    End If
    
    Fill = length - Len(Rtn)
    
    Modulus = Fill Mod 2
    
    Fill = Fill - Modulus
            
    Rtn = Space(Fill / 2) & Rtn & Space(Fill / 2 + Modulus)
        
    CenterString = Rtn
    
End Function

Public Function RemoveMeta(ByVal Str As String) As String
    
    Dim TempStr As String
    Dim Product As String
    Dim Loc As Integer

    TempStr = Str
    
    Loc = InStr(1, TempStr, "<")
    
    If Loc = 0 Then
        Product = TempStr
    End If
    
    Do Until Loc <= 0
        
        TempStr = Mid(TempStr, Loc + 1)
        
        Loc = InStr(1, TempStr, ">")
        
        If Loc <= 0 Then
                    
            Exit Do
        
        End If
        
        TempStr = Mid(TempStr, Loc + 1)
        
        
        Loc = InStr(1, TempStr, "<")
            
        If Loc > 0 Then
        
            Product = Product & Mid(TempStr, 1, Loc - 1)
        
        End If
            
    Loop
    
    RemoveMeta = Product
    
End Function
