Attribute VB_Name = "Registry"
' API Function Declares

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal Reserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

' API Type Declares

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

' API Constant Declares

Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const SYNCHRONIZE = &H100000

Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const ERROR_SUCCESS = 0&


Public Function Get_Registered_Device(ByVal Registry_Path As String, ByRef Devices() As String, ByRef Number_Of_Registered_Device As Integer)
    Dim FT As FILETIME
    Dim KeyHandle As Long
    Dim Res As Long
    Dim Index As Long
    Dim KeyName As String
    Dim ClassName As String
    Dim KeyLen As Long
    Dim ClassLen As Long
    
    Res = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Registry_Path, 0, KEY_READ, KeyHandle)
    
    If Res <> ERROR_SUCCESS Then
        MsgBox ("Problem with the Windows Registry" & vbNewLine & "Unable to open key")
        Exit Function
    End If
    
    Number_Of__Registered_Device = 0
    
    Do
        KeyLen = 64
        ClassLen = 64
        KeyName = String$(KeyLen, 0)
        ClassName = String$(ClassLen, 0)
        Res = RegEnumKeyEx(KeyHandle, Index, KeyName, KeyLen, 0, ClassName, ClassLen, FT)
        Index = Index + 1
        
        If Res = ERROR_SUCCESS Then
            'Return number of registered device
            Number_Of_Registered_Device = Number_Of_Registered_Device + 1
            'Get device name
            Devices(Number_Of_Registered_Device) = (Left$(KeyName, KeyLen))
        End If
    Loop While Res = ERROR_SUCCESS
    
    Call RegCloseKey(KeyHandle)

End Function
