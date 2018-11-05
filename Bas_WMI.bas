Attribute VB_Name = "Bas_WMI"
Private Const ModuleName As String = "Bas_WMI"

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long 'For ini read
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long 'For ini write
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Public Function GetMachineID() As String
  Dim SWbemSet(2) As SWbemObjectSet 'Need to incl project reference 'Microsoft WMI Scripting Library'
  Dim SWbemObj As SWbemObject
  Dim varObjectToId(2) As String
  Dim varSerial(2) As String
  Dim i, j As Integer
  Dim varSerialHex As String
  Dim varSerialTemp(2) As String
  
  On Error Resume Next
      
  varObjectToId(1) = "Win32_Processor,ProcessorId"
  varObjectToId(2) = "Win32_OperatingSystem,SerialNumber"
  'We're using CPU & OS serials but can be any WMI class object property that returns a numeric value (or alpha if you use an algorithm to convert characters to numbers - LSB of asc(character) maybe...)
  'Refer to http://msdn.microsoft.com/library/en-us/wmisdk/wmi/retrieving_a_class.asp for more info
    
  For i = 1 To 2 'we're only using 2 objects in this example - can be more but code will have to me modified to suit
    Set SWbemSet(i) = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf(Split(varObjectToId(i), ",")(0))
    varSerial(i) = ""
    For Each SWbemObj In SWbemSet(i) 'Is buggy if querying 2nd similar device (eg 2nd CPU ID or MAC Address) but I ran out of time
      varSerialHex = ""
      varSerialTemp(i) = ""
      varSerial(i) = SWbemObj.Properties_(Split(varObjectToId(i), ",")(1)) 'Property value
      For j = 1 To Len(varSerial(i)) 'Strip out any non hex characters so we can do some simple maths later
        varSerialHex = Mid(varSerial(i), j, 1)
        If varSerialHex Like "[0-9A-Fa-f]" Then varSerialTemp(i) = varSerialTemp(i) & varSerialHex
      Next
    Next
    varSerial(i) = varSerialTemp(i)
    varSerial(i) = Right(varSerial(i), 4) 'Let's just use last 4 digits of each serial for simplicity
  Next
  varMachineId = ""
  If Len(varSerial(1)) = 0 Then
    varSerial(1) = Replace(FSO.Drives(Left(AppPath(), 1)).SerialNumber, "-", "")
    varSerial(2) = StrReverse(varSerial(1))
  End If
  
  For i = 1 To 4
    varMachineId = varMachineId & Mid(varSerial(1), i, 1) & Mid(varSerial(2), i, 1) 'Just a little obfuscation
  Next
  
  varMachineId = CLng("&H" & varMachineId)
  
  GetMachineID = varMachineId
  
End Function








