Attribute VB_Name = "Bas_General"
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'FOR SETTING DEFAULT SHORT DATE FORMAT
'MSDN REF: Q168793
'DATE: 22 NOV 2002

Private Const LOCALE_SSHORTDATE = &H1F
Private Const LOCALE_STIMEFORMAT = &H1003
      Private Const WM_SETTINGCHANGE = &H1A
      'same as the old WM_WININICHANGE
      Private Const HWND_BROADCAST = &HFFFF&

      Private Declare Function SetLocaleInfo Lib "kernel32" Alias _
          "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As _
          Long, ByVal lpLCData As String) As Boolean
      Private Declare Function PostMessage Lib "user32" Alias _
          "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
          ByVal wParam As Long, ByVal lParam As Long) As Long
      Private Declare Function GetSystemDefaultLCID Lib "kernel32" _
          () As Long
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Declare Function GetComputerName Lib "kernel32" Alias _
    "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    
Private Declare Function GetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer, nSize As Long) As Long

Private Const ModuleName As String = "Bas_General"

Public Function AppPath(Optional FileName As String) As String

    Dim Path As String
    Path = App.Path
    If Not Right(Path, 1) = "\" Then
        Path = Path & "\"
    End If
    
    AppPath = Path & FileName
    
End Function

Public Sub Setvbshortdate(ByVal DateFormat As String)

    Dim dwLCID As Long
    dwLCID = GetSystemDefaultLCID()
    If SetLocaleInfo(dwLCID, LOCALE_SSHORTDATE, DateFormat) _
            = False Then
        Call LogFormError(ModuleName, "Setvbshortdate(" & DateFormat & ")", "Failed")
        Exit Sub
    End If
    PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0

End Sub

Public Sub SetTimeFormat(ByVal TimeFormat As String)
    Dim dwLCID As Long
    dwLCID = GetSystemDefaultLCID()

    If SetLocaleInfo(dwLCID, LOCALE_STIMEFORMAT, TimeFormat) _
            = False Then
        Call LogFormError(ModuleName, "SetTimeFormat(" & TimeFormat & ")", "Failed")
        Exit Sub
    End If
    PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
End Sub



Public Function GetDateTimeSerial(Optional ByVal FormatString As String) As String
    
    Dim FormatStr As String
    Dim Serial As String
    
    FormatStr = "yyyyMMdd"
    If Not Len(FormatString) = 0 Then
        FormatStr = FormatString
    End If
    
    Serial = Format(Now(), FormatStr)
    Serial = Replace(Serial, "/", "")
    Serial = Replace(Serial, ":", "")
    Serial = Replace(Serial, " ", "")
    
    GetDateTimeSerial = Serial
    
End Function

Public Function GetHexDateTimeSerial() As String

    Dim SerialID As String
    
    SerialID = leading(Hex(GetDateTimeSerial("yyyy")), 4, "0")
    SerialID = SerialID & leading(Hex(GetDateTimeSerial("MM")), 2, "0")
    SerialID = SerialID & leading(Hex(GetDateTimeSerial("dd")), 2, "0")
    SerialID = SerialID & leading(Hex(GetDateTimeSerial("HH")), 2, "0")
    SerialID = SerialID & leading(Hex(GetDateTimeSerial("mm")), 2, "0")
    SerialID = SerialID & leading(Hex(GetDateTimeSerial("ss")), 3, "0")
    
    
    GetHexDateTimeSerial = SerialID
    

End Function

Public Function SysGetComputerName() As String
    
    'Visual Basic Source Code Library
    'ISBN 0-672-31387-1
    'P.519
    'samspublishing.com
    
    Dim Computer As String
    Dim BufSize As Long
    Dim RetCode As Long
    Dim NullCharPos As Long
    
    Computer = Space(80)
    BufSize = Len(Computer)
    
    RetCode = GetComputerName(Computer, BufSize)
    
    NullCharPos = InStr(Computer, Chr(0))
    If NullCharPos > 0 Then
        Computer = Left(Computer, NullCharPos - 1)
    Else
        Computer = ""
    End If
    
    SysGetComputerName = Computer

End Function

Public Function SysGetUserName() As String
    
    'Visual Basic Source Code Library
    'ISBN 0-672-31387-1
    'P.518
    'samspublishing.com
    
    Dim UserName As String
    Dim BufSize As Long
    Dim RetCode As Long
    Dim NullCharPos As Long
    
    UserName = Space(80)
    BufSize = Len(UserName)
    
    RetCode = GetUserName(UserName, BufSize)
    
    NullCharPos = InStr(UserName, Chr(0))
    If NullCharPos > 0 Then
        UserName = Left(UserName, NullCharPos)
    Else
        UserName = ""
    End If
    
    SysGetUserName = UserName
    
End Function
