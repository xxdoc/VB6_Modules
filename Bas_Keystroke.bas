Attribute VB_Name = "Bas_KeyStroke"
Option Explicit

Type POINTAPI
        x As Long
        y As Long
End Type

Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Public Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Public Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move

Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Integer) As Long
Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long

Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

'  dwPlatformId defines:
'
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1

Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON

Public Const VK_BACK = &H8
Public Const VK_TAB = &H9

Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD

Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_PAUSE = &H13
Public Const VK_CAPITAL = &H14

Public Const VK_ESCAPE = &H1B

Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87

Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91

'
'   VK_L VK_R - left and right Alt, Ctrl and Shift virtual keys.
'   Used only as parameters to GetAsyncKeyState() and GetKeyState().
'   No other API or message will distinguish left and right keys in this way.
'  /
Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_LMENU = &HA4
Public Const VK_RMENU = &HA5

Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE


'
' Sends a single character using keybd_event
'   Note that this function does not set shift state
'   (By pressing down the shift key or setting the shift keys state)
'   and it doesn't handle extended keys.
'
Public Sub SendAKey(ByVal c$)
    Dim vk%
    Dim scan%
    Dim oemchar$
    Dim dl&
    ' Get the virtual key code for this character
    vk% = VkKeyScan(Asc(c$)) And &HFF
    oemchar$ = "  " ' 2 character buffer
    ' Get the OEM character - preinitialize the buffer
    CharToOem Left$(c$, 1), oemchar$
    ' Get the scan code for this key
    scan% = OemKeyScan(Asc(oemchar$)) And &HFF
    ' Send the key down
    keybd_event vk%, scan%, 0, 0
    ' Send the key up
    keybd_event vk%, scan%, KEYEVENTF_KEYUP, 0
End Sub

Public Sub MySendKeys(ByVal c$)
    Dim x&
    For x& = 1 To Len(c$)
        SendAKey Mid$(c$, x&)
    Next x&
End Sub
