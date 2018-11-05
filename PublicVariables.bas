Attribute VB_Name = "PublicVariables"
'UserID
Public USERNAME As String

'Password
Public PASSWORD As String

'Save Password
Public IS_SAVE_PASSWORD As Boolean

'Ethernet Adapter
Public ADAPTER As Integer

'Use Manual Assign IP
Public IS_MANUALLY_ASSIGN As Boolean

'Manual Assign IP
Public MANUAL_ASSIGN_IP As String

'Port
Public PORT As String

'Auto Connect when start up
Public IS_AUTO_CONNECT As Boolean

'Web address for connecting server
Public WEB_ADDR_LOADIP As String

'Web address for disconnecting
Public WEB_ADDR_DISCONNECT As String


Public Function READ_FILE_USER_ACCOUNT(ByVal Path As String) As Boolean

Dim Fso As New FileSystemObject
Dim Ts As TextStream

Dim Readline As String
Dim Property As String
Dim Value As Variant

READ_FILE_USER_ACCOUNT = False

On Error GoTo ErrorHandler

'File does not exist
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If Len(Path) = 0 Then
    Exit Function
End If

Set Fso = CreateObject("Scripting.FileSystemObject")
Set Ts = Fso.OpenTextFile(Path, ForReading, False)
    
Do Until Ts.AtEndOfStream
    Readline = Ts.Readline
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Property = Mid(Readline, 1, InStr(1, Readline, "=") - 1)
    Value = Mid(Readline, InStr(1, Readline, "=") + 1)
    Select Case UCase(Property)
        Case "USERNAME"
            USERNAME = Value
        Case "PASSWORD"
            PASSWORD = Value
        Case "IS_SAVE_PASSWORD"
            IS_SAVE_PASSWORD = Value
    End Select
Loop

Ts.Close
Set Ts = Nothing

READ_FILE_USER_ACCOUNT = True

Exit Function
ErrorHandler:
    ErrorMessage = Now() & vbNewLine
    ErrorMessage = ErrorMessage & "Module Public Variables" & vbNewLine
    ErrorMessage = ErrorMessage & "Function: READ_FILE_USER_ACCOUNT" & vbNewLine
    ErrorMessage = ErrorMessage & Err.Description & vbCrLf
    ErrorMessage = ErrorMessage & Path
    LogError ErrorMessage

End Function

Public Function WRITE_FILE_USER_ACCOUNT(Path As String) As Boolean
    Dim Fso As New FileSystemObject
    Dim Ts As TextStream
    
    Dim ErrorMessage As String
    
    On Error GoTo ErrorHandler
    
    WRITE_FILE_USER_ACCOUNT = False
    
    'File does not exist
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If Len(Path) = 0 Then
        Exit Function
    End If

    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set Ts = Fso.OpenTextFile(Path, ForWriting, 1)
    
    'UserID
    Ts.WriteLine "USERNAME=" & USERNAME
    
    'Save Password
    Ts.WriteLine "IS_SAVE_PASSWORD=" & IS_SAVE_PASSWORD
    
    'Password
    If IS_SAVE_PASSWORD = True Then
        Ts.WriteLine "PASSWORD=" & PASSWORD
    End If
    
    Ts.Close
    Set Ts = Nothing
    
    WRITE_FILE_USER_ACCOUNT = True
    
    Exit Function
    
ErrorHandler:
    ErrorMessage = Now() & vbNewLine
    ErrorMessage = ErrorMessage & "Module Public Variables" & vbNewLine
    ErrorMessage = ErrorMessage & "Function: WRITE_FILE_USER_ACCOUNT" & vbNewLine
    ErrorMessage = ErrorMessage & Err.Description & vbCrLf
    ErrorMessage = ErrorMessage & Path
    LogError ErrorMessage
    
End Function

Public Function READ_FILE_CONFIGURATION(ByVal Path As String) As Boolean

Dim Fso As New FileSystemObject
Dim Ts As TextStream

Dim Readline As String
Dim Property As String

Dim Value As Variant

READ_FILE_CONFIGURATION = False

On Error GoTo ErrorHandler

'File does not exist
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If Len(Path) = 0 Then
    Exit Function
End If

Set Fso = CreateObject("Scripting.FileSystemObject")
Set Ts = Fso.OpenTextFile(Path, ForReading, False)
    
Do Until Ts.AtEndOfStream
    Readline = Ts.Readline
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Property = Mid(Readline, 1, InStr(1, Readline, "=") - 1)
    Value = Mid(Readline, InStr(1, Readline, "=") + 1)
    Select Case UCase(Property)
        Case "ADAPTER"
            ADAPTER = Value
        Case "IS_MANUALLY_ASSIGN"
            IS_MANUALLY_ASSIGN = Value
        Case "MANUAL_ASSIGN_IP"
            MANUAL_ASSIGN_IP = Value
        Case "PORT"
            PORT = Value
        Case "IS_AUTO_CONNECT"
            IS_AUTO_CONNECT = Value
    End Select
Loop

Ts.Close
Set Ts = Nothing

READ_FILE_CONFIGURATION = True

Exit Function
ErrorHandler:
    ErrorMessage = Now() & vbNewLine
    ErrorMessage = ErrorMessage & "Module Public Variables" & vbNewLine
    ErrorMessage = ErrorMessage & "Function: READ_FILE_CONFIGURATION" & vbNewLine
    ErrorMessage = ErrorMessage & Err.Description & vbCrLf
    ErrorMessage = ErrorMessage & Path
    LogError ErrorMessage

End Function
Public Function WRITE_FILE_CONFIGURATION(Path As String) As Boolean
    On Error GoTo ErrorHandler
    
    WRITE_FILE_CONFIGURATION = False
    
    Dim ErrorMessage As String
    
    Dim Fso As New FileSystemObject
    Dim Ts As TextStream
    
    'File does not exist
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If Len(Path) = 0 Then
        Exit Function
    End If

    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set Ts = Fso.OpenTextFile(Path, ForWriting, 1)
    
    'Adapter
    Ts.WriteLine "ADAPTER=" & ADAPTER
    'IS_MANUALLY_ASSIGN
    Ts.WriteLine "IS_MANUALLY_ASSIGN=" & IS_MANUALLY_ASSIGN
    'MANUAL_ASSIGN_IP
    If IsIP(MANUAL_ASSIGN_IP) = True Then
        Ts.WriteLine "MANUAL_ASSIGN_IP=" & MANUAL_ASSIGN_IP
    End If
    'Port
    If IsNumeric(PORT) = True Then
        Ts.WriteLine "PORT=" & PORT
    End If
    Ts.WriteLine "IS_AUTO_CONNECT=" & IS_AUTO_CONNECT
    
    Ts.Close
    
    Set Ts = Nothing
    
    WRITE_FILE_CONFIGURATION = True
    
    Exit Function

ErrorHandler:
    ErrorMessage = Now() & vbNewLine
    ErrorMessage = ErrorMessage & "Module Public Variables" & vbNewLine
    ErrorMessage = ErrorMessage & "Function: WRITE_FILE_CONFIGURATION" & vbNewLine
    ErrorMessage = ErrorMessage & Err.Description & vbCrLf
    ErrorMessage = ErrorMessage & Path
    LogError ErrorMessage
End Function

