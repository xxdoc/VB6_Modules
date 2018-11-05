Attribute VB_Name = "Function"

Public Function AppPath(Filename As String) As String
    AppPath = App.Path
    If Not Right(AppPath, 1) = "\" Then
        AppPath = AppPath & "\"
    End If
    AppPath = AppPath & Filename
End Function

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
        IsIP = True
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
    
    If Not CLng(PORT) <= 65535 Then
        IsPort = False
        Exit Function
    End If
    
    IsPort = True
    
    Exit Function
ErrorHandler:
        IsPort = False
End Function

Public Function AppendFile(ByVal Path As String, ByVal Content As String) As Boolean
    On Error GoTo ErrorHandler
    
    AppendFile = False
    
    Dim Fso As New FileSystemObject
    Dim Ts As TextStream
    
    Set Ts = Fso.OpenTextFile(Path, ForAppending, True)
    Ts.WriteLine Content
    Ts.WriteBlankLines (1)
    Ts.Close
    
    AppendFile = True
    
    Exit Function

ErrorHandler:
    AppendFile = False
    
End Function

Public Function LogError(ByVal ErrorMessage As String)
    On Error GoTo ErrorHandler
        AppendFile AppPath(FILE_PROGRAM_ERROR), ErrorMessage
    Exit Function
ErrorHandler:
    Exit Function
    
End Function

Public Function WebContent(ByVal Inet1 As Inet, ByVal WebAddress As String) As String
    On Error GoTo ErrorHandler
    Dim ErrorMessage As String
        WebContent = Inet1.OpenURL(WebAddress)
    Exit Function
ErrorHandler:
    ErrorMessage = Now() & vbNewLine
    ErrorMessage = ErrorMessage & "Function WebContent" & vbNewLine
    ErrorMessage = ErrorMessage & Err.Description & vbNewLine
    ErrorMessage = ErrorMessage & WebAddress
    LogError ErrorMessage
End Function


