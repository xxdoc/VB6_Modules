Attribute VB_Name = "Bas_Internet"


Public Function IsIP(ByVal IP As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ErrorMessage As String
    
    IsIP = False
    'Eliminate all non-numeric character
    '.
    IP = Replace(IP, ".", "")
    IP = Replace(IP, ":", "")
    
    If Len(IP) = 0 Then
        Exit Function
    End If
    
    If IsNumeric(IP) = True Then
        If IP >= 0 Then
            IsIP = True
        End If
    End If
    
    Exit Function
ErrorHandler:
    IsIP = False
    ErrorMessage = Now() & vbNewLine
    ErrorMessage = ErrorMessage & "Function IsIP" & vbNewLine
    ErrorMessage = ErrorMessage & Err.Description
    LogError ErrorMessage
End Function

Public Function IsPort(ByVal PORT As String) As Boolean
    On Error GoTo ErrorHandler
    
    IsPort = False
    
    If Len(PORT) = 0 Then
        IsPort = False
        Exit Function
    End If
    
    If Not IsNumeric(PORT) = True Then
        IsPort = False
        Exit Function
    End If
    
    If CLng(PORT) < 0 Then
        IsPort = False
        Exit Function
    End If
    
    If Not CLng(PORT) <= 65535 Then
        IsPort = False
        Exit Function
    End If
    
    IsPort = True
    
    Exit Function
ErrorHandler:
        IsPort = False
End Function

