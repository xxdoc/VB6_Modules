Attribute VB_Name = "Bas_DataType"
Private Const ModuleName As String = "Bas_DataType"

Public Function IsHex(HexValue As String) As Boolean
    
    Dim idx As Integer
    Dim Temp As String
    
    For idx = 1 To Len(HexValue)
            Temp = Mid(HexValue, idx, 1)
            If Not (Asc(Temp) >= 48 And Asc(Temp) <= 57) And Not (Asc(Temp) >= 65 And Asc(Temp) <= 70) Then
                Exit Function
            End If
    Next
    
    IsHex = True
    
End Function

Public Function HexToDec(ByVal HexValue As String) As Double
    
    'Convert Hex  to Dec
    'Max Hex: F (Dec: 15)
    
    Dim DecValue As Double
    
    If Len(HexValue) > 1 Then
        Exit Function
    End If
    
    If IsNumeric(HexValue) = True Then
        HexToDec = CDbl(HexValue)
        Exit Function
    End If
    
    DecValue = Asc(HexValue) - 55
    
    If DecValue > 15 Then
        Exit Function
    End If
    
    HexToDec = DecValue
    
End Function

Public Function HexCharcode(ByVal HexValue As String) As Double
    
    'Convert Hex ASCII code to Dec ASCII code (ASCII 1 - 255)
    
    'Max Value
    'Hex FF

    Dim TempValue As String
    TempValue = HexValue
    
    If Len(HexValue) > 2 Then
        Exit Function
    End If
    
    If Len(HexValue) > 1 Then
            charcode = HexToDec(Mid(HexValue, 1, 1)) * 16
    End If
    
    HexCharcode = charcode + HexToDec(Right(HexValue, 1))
    
End Function

Public Function HexToByteArray(ByVal HexValue As String, ByRef ByteArray() As Byte) As Boolean
    
    Dim idx As Double
    Dim ArrayIdx As Double
    
    If IsHex(HexValue) = False Then
        Exit Function
    End If
    
    If Not Len(HexValue) Mod 2 = 0 Then
        Exit Function
    End If
    
    ArraySize = Len(HexValue) / 2 - 1
    ReDim ByteArray(0 To ArraySize)
    For idx = 1 To Len(HexValue) Step 2
        ByteArray((idx - 1) / 2) = HexCharcode(Mid(HexValue, idx, 2))
    Next
    
    HexToByteArray = True
    
End Function

Public Function HexToString(ByVal HexValue As String) As String
    
    Dim idx As Double
    Dim Str As String
    
    If IsHex(HexValue) = False Then
        Exit Function
    End If
    
    If Not Len(HexValue) Mod 2 = 0 Then
        Exit Function
    End If
    
    For idx = 1 To Len(HexValue) Step 2
        Str = Str & HexCharcode(Mid(HexValue, idx, 2))
    Next
    
    HexToString = Str
    
End Function

Public Function DecimalToBinary(DecimalValue As Long, _
    MinimumDigits As Integer) As String

' Returns a string containing the binary
' representation of a positive integer

Dim result As String
Dim ExtraDigitsNeeded As Integer

' Make sure value is not negative
DecimalValue = Abs(DecimalValue)

' Construct the binary value

Do
    result = CStr(DecimalValue Mod 2) & result
    DecimalValue = DecimalValue \ 2
Loop While DecimalValue > 0

' Add leading zeros if needed
ExtraDigitsNeeded = MinimumDigits - Len(result)
If ExtraDigitsNeeded > 0 Then
    result = String(ExtraDigitsNeeded, "0") & result
End If

DecimalToBinary = result

End Function

Public Function HexToBinary(HexStr As String, _
    Optional MinimumDigits As Integer) As String
    
    Dim BinaryString As String
    Dim DecimalValue As Double
    
    For idx = 1 To Len(HexStr)
        BinaryString = BinaryString & DecimalToBinary(HexToDec(Mid(HexStr, idx, 1)), 4)
    Next
    
    BinaryString = leading(BinaryString, MinimumDigits, "0")
    
    HexToBinary = BinaryString
    
End Function

Public Function BinaryToDec(BinaryString As String) As Long

    Dim DecimalValue As Long
    Dim idx As Integer
    
    DecimalValue = 0
    
    For idx = 0 To Len(BinaryString) - 1
    
        DecimalValue = DecimalValue + (2 ^ idx) * Mid(BinaryString, Len(BinaryString) - idx, 1)
        
    Next
    
    BinaryToDec = DecimalValue
    
End Function

Public Function ToHexString(Str As String) As String
    
    Dim idx As Integer
    Dim HexString As String
    
    If Len(Trim(Str)) <= 0 Then
        Exit Function
    End If
    
    For idx = 1 To Len(Str)
        
        HexString = HexString & leading(Hex(Asc(Mid(Str, idx, 1))), 2, "0")
        
    Next
    
    ToHexString = HexString
    
End Function

