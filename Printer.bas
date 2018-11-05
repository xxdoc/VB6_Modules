Attribute VB_Name = "PrinterSupport"
'
' Use of this source code is subject to the terms of the EULA
' under which you licensed this SOFTWARE PRODUCT.
' If you did not accept the terms of the EULA, you are not
' authorized to use this source code.
'

'
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
' ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
' THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
' PARTICULAR PURPOSE.
'

Option Explicit

Public Const SP_ERROR = (-1)

Public Declare Function CreateDC Lib "coredll" Alias "CreateDCW" _
    (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
     ByVal lpOutput As String, ByVal lpInitData As Long) As Long

Public Declare Function DeleteDC Lib "coredll" (ByVal hdc As Long) _
    As Long

Public Declare Function DrawText Lib "coredll" Alias "DrawTextW" _
    (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, _
     ByVal lpRect As String, ByVal uFormat As Long) As Long

Public Declare Function StartDoc Lib "coredll" Alias "StartDocW" _
    (ByVal hdc As Long, ByVal lpdi As String) As Long

Public Declare Function StartPage Lib "coredll" (ByVal hdc As Long) _
    As Long

Public Declare Function EndDoc Lib "coredll" (ByVal hdc As _
    Long) As Long

Public Declare Function EndPage Lib "coredll" (ByVal hdc As Long) _
    As Long

Public Declare Function AbortDoc Lib "coredll" (ByVal hdc As _
    Long) As Long

Public Declare Function GetLastError Lib "coredll" () As Long

Public Declare Function ExtEscape Lib "coredll" (ByVal hdc As Long, _
    ByVal code As Long, ByVal insz As Long, ByVal inp As String, _
    ByVal out As Long, ByVal outp As String) As Long

Public Declare Function SetBkMode Lib "coredll" (ByVal hdc As Long, _
    ByVal mode As Long) As Long

Public Declare Function CreateFontIndirect Lib "coredll" Alias _
    "CreateFontIndirectW" (ByVal LogFont As String) As Long

Public Declare Function SetTextColor Lib "coredll" ( _
    ByVal hdc As Long, ByVal color As Long) As Long

Public Declare Function GetDeviceCaps Lib "coredll" ( _
    ByVal hdc As Long, ByVal index As Long) As Long

Public Declare Function SelectObject Lib "coredll" ( _
    ByVal hdc As Long, ByVal obj As Long) As Long

Public Declare Function DeleteObject Lib "coredll" ( _
    ByVal obj As Long) As Long

' DONT use that function... how to use DEVNAMES structure???
'Public Declare Function PageSetupDlg Lib "commdlg" Alias "PageSetupDlgW" _
'    (ByVal LPPageSetupDlg As String) As Long

Public Declare Function PrintDlg Lib "commdlg" _
    (ByVal LPPrintDlg As String) As Long

Public Declare Function MessageBeep Lib "coredll" _
    (ByVal index As Long) As Long

Public Declare Function MessageBox Lib "coredll" Alias "MessageBoxW" _
    (ByVal hwnd As Long, ByVal text As String, ByVal title As String, _
     ByVal uType As Long) As Long

Public Declare Function LocalAlloc Lib "coredll" _
    (ByVal flags As Long, ByVal bytes As Long) As Long
    
Public Declare Function LocalFree Lib "coredll" _
    (ByVal hMem As Long) As Long

Public Declare Sub MoveMemory Lib "coredll" Alias "memmove" _
    (ByVal Destination As String, ByVal Source As Long, ByVal Length As Long)

Public Declare Function SHLoadImageFile Lib "aygshell" Alias "#75" _
    (ByVal ImageFile As String) As Long

Public Declare Function TransparentImage Lib "coredll" _
    (ByVal hdc As Long, ByVal DstX As Long, ByVal DstY As Long, _
     ByVal DstCx As Long, ByVal DstCy As Long, ByVal hSrc As Long, _
     ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcCx As Long, _
     ByVal SrcCy As Long, ByVal TransparentColor As Long) As Long

Public Declare Function GetObject Lib "coredll" Alias "GetObjectW" _
    (ByVal obj As Long, ByVal sz As Long, ByVal Result As String) As Long

Public Const LMEM_FIXED = &H0
Public Const LMEM_MOVEABLE = &H2
Public Const LMEM_NOCOMPACT = &H10         '/**** Used for Moveable Memory  ***/
Public Const LMEM_ZEROINIT = &H40

Public Const LHND = &H42 ' (LMEM_MOVEABLE + LMEM_ZEROINIT)
Public Const LPTR = &H40 ' (LMEM_FIXED + LMEM_ZEROINIT)


Public Const MB_OK = &H0
Public Const MB_OKCANCEL = &H1
Public Const MB_ABORTRETRYIGNORE = &H2
Public Const MB_YESNOCANCEL = &H3
Public Const MB_YESNO = &H4
Public Const MB_RETRYCANCEL = &H5

Public Const MB_ICONHAND = &H10
Public Const MB_ICONQUESTION = &H20
Public Const MB_ICONEXCLAMATION = &H30
Public Const MB_ICONASTERISK = &H40

Public Const MB_ICONWARNING = &H30       ' MB_ICONEXCLAMATION
Public Const MB_ICONERROR = &H10       ' MB_ICONHAND

Public Const MB_ICONINFORMATION = &H40       ' MB_ICONASTERISK
Public Const MB_ICONSTOP = &H10       ' MB_ICONHAND

Public Const MB_DEFBUTTON1 = &H0
Public Const MB_DEFBUTTON2 = &H100
Public Const MB_DEFBUTTON3 = &H200
Public Const MB_DEFBUTTON4 = &H300

Public Const MB_APPLMODAL = &H0
Public Const MB_SETFOREGROUND = &H10000

Public Const MB_TOPMOST = &H40000

' PrintDlg constants
' Out only
Public Const PD_SELECTALLPAGES = &H1
Public Const PD_SELECTSELECTION = &H2
Public Const PD_SELECTDRAFTMODE = &H8
Public Const PD_SELECTA4 = &H10
Public Const PD_SELECTLETTER = &H20
Public Const PD_SELECTINFRARED = &H40
Public Const PD_SELECTSERIAL = &H80

' In only
Public Const PD_DISABLEPAPERSIZE = &H100
Public Const PD_DISABLEPRINTRANGE = &H200
Public Const PD_DISABLEMARGINS = &H400
Public Const PD_DISABLEORIENTATION = &H800
Public Const PD_RETURNDEFAULTDC = &H2000
Public Const PD_ENABLEPRINTHOOK = &H4000
Public Const PD_ENABLEPRINTTEMPLATE = &H8000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_TITLE = &H20000

' In-Out
Public Const PD_SELECTPORTRAIT = &H40000
Public Const PD_SELECTLANDSCAPE = &H80000
Public Const PD_MARGINS = &H100000
Public Const PD_INTHOUSANDTHSOFINCHES = &H200000
Public Const PD_INHUNDREDTHSOFMILLIMETERS = &H400000
Public Const PD_MINMARGINS = &H800000

' PageSetupDlg constants
'Public Const PSD_DEFAULTMINMARGINS = &H0        ' default (printer's)
'Public Const PSD_INWININIINTLMEASURE = &H0        ' 1st of 4 possible

'Public Const PSD_MINMARGINS = &H1        ' use caller's
'Public Const PSD_MARGINS = &H2        ' use caller's
'Public Const PSD_INTHOUSANDTHSOFINCHES = &H4        ' 2nd of 4 possible
'Public Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8        ' 3rd of 4 possible
'Public Const PSD_DISABLEMARGINS = &H10
'Public Const PSD_DISABLEPRINTER = &H20
'' Public Const PSD_NOWARNING = &h00000080 ' not used in CE
'Public Const PSD_DISABLEORIENTATION = &H100
'Public Const PSD_DISABLEPAPER = &H200
'Public Const PSD_RETURNDEFAULT = &H400
'' Public Const PSD_SHOWHELP = &h00000800 ' not used in CE
'Public Const PSD_ENABLEPAGESETUPHOOK = &H2000
'Public Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000
'Public Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000
'' Public Const PSD_ENABLEPAGEPAINTHOOK = &H40000    ' not used in CE
'' Public Const PSD_DISABLEPAGEPAINTING = &H80000    ' not used in CE
'' Public Const PSD_NONETWORKBUTTON = &H200000   ' not used in CE

'' new in Win CE - print range flags
'Public Const PSD_DISABLEPRINTRANGE = &H10000000
'Public Const PSD_RANGESELECTION = &H20000000

' some of standard codes
Public Const ESCAPE_NEWFRAME = 1
Public Const ESCAPE_ABORTDOC = 2
Public Const ESCAPE_NEXTBAND = 3
Public Const ESCAPE_QUERYESCSUPPORT = 8
Public Const ESCAPE_SETABORTPROC = 9
Public Const ESCAPE_STARTDOC = 10
Public Const ESCAPE_ENDDOC = 11
Public Const ESCAPE_PASSTHROUGH = 19

' special printer codes
Public Const ESCAPE_READSCARD = 7000
Public Const ESCAPE_READSTATUS = 7001 ' 1 byte
Public Const ESCAPE_READINFO = 7002 ' 2 bytes
Public Const ESCAPE_READCAPABILITIES = 7003 ' 32 bytes
Public Const ESCAPE_READPRINTER = 7004 ' appropriate
Public Const ESCAPE_PRINTERDISCONNECT = 7005 ' close printer (force or not)
Public Const ESCAPE_PRINTERPASSTHROUGH = 7006 ' enable/disable passthrough

Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

Public Const PHYSICALWIDTH = 110 ' Physical Width in device units
Public Const PHYSICALHEIGHT = 111 ' Physical Height in device units
Public Const LOGPIXELSX = 88 ' Logical pixels/inch in X
Public Const LOGPIXELSY = 90 ' Logical pixels/inch in Y

Public Const DT_LEFT = 0
Public Const DT_CENTER = 1
Public Const DT_RIGHT = 2
Public Const DT_NOPREFIX = &H800
Public Const DT_WORDBREAK = &H10
Public Const DT_CALCRECT = &H400

Private Const DRAW_TEXT_DRAW = &H810 'DT_LEFT + DT_NOPREFIX + DT_WORDBREAK
Private Const DRAW_TEXT_CALC = &HC10 'DRAW_TEXT_DRAW + DT_CALCRECT

Public Const FW_DONTCARE = 0
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_HEAVY = 900

Public Const TRACK_1 = 1
Public Const TRACK_2 = 2
Public Const TRACK_3 = 4

Public Const PRINTER_DEACTIVATE = 1
Public Const PRINTER_ACTIVATE = 0

Public Const STATUS_NO_PAPER = 4 ' also cover open, user intervention required
Public Const STATUS_OVERHEAT = 8 ' printer head overheaten, wait until return to normal
Public Const STATUS_LOW_BATERY = 64 ' charge battery

Public Const CAPS_IRDA_ALLOW = 1  ' printer allows IrDA connection
Public Const CAPS_MAGSTRIPE = 2   ' magstripe reader installed
Public Const CAPS_THREE_HEAD = 4  ' three head MGR
Public Const CAPS_KATAKANA = 8    ' Katakana symbols support
Public Const CAPS_KANJI = 16      ' JIS & Shift-JIS support
Public Const CAPS_FAHRENHEIT = 32  ' Degree are measured in Fahrenheit (not in Celsius)
Public Const CAPS_UPGRADABLE = 256 ' Upgradable firmware

Public Const BARCODE_UPCA = 65
Public Const BARCODE_UPCE = 66
Public Const BARCODE_EAN13 = 67
Public Const BARCODE_EAN8 = 68
Public Const BARCODE_CODE39 = 69
Public Const BARCODE_ITF = 70
Public Const BARCODE_CODABAR = 71
Public Const BARCODE_CODE93 = 72
Public Const BARCODE_CODE128 = 73
Public Const BARCODE_WIDTH_TIGHT = 2
Public Const BARCODE_WIDTH_NORMAL = 3
Public Const BARCODE_WIDTH_WIDE = 4
Public Const BARCODE_FONT_A = 0
Public Const BARCODE_FONT_B = 1
Public Const BARCODE_TEXTS_NONE = 0
Public Const BARCODE_TEXTS_ABOVE = 1
Public Const BARCODE_TEXTS_BELLOW = 2
Public Const BARCODE_TEXTS_BOTH = 3

Private Function MemoryStringToLong(StringIn As String) As Long
  On Error Resume Next
  Dim hWorkVal As String

  Dim i As Long
  Dim Help As String
  For i = 4 To 1 Step -1 ' sizeof LONG = 4 bytes
    Help = Hex(AscB(MidB(StringIn, i, 1)))
    If Len(Help) <> 2 Then Help = "0" + Help
    hWorkVal = hWorkVal & Help
  Next i
  '
  ' Return Long Integer value.
  MemoryStringToLong = CLng("&H" & hWorkVal)
End Function

Private Function MemoryStringToInt(StringIn As String) As Integer
  On Error Resume Next
  Dim hWorkVal As String

  Dim i As Long
  Dim Help As String
  For i = 2 To 1 Step -1 ' sizeof INTEGER = 2 bytes
    Help = Hex(AscB(MidB(StringIn, i, 1)))
    If Len(Help) <> 2 Then Help = "0" + Help
    hWorkVal = hWorkVal & Help
  Next i
  '
  ' Return Integer value.
  MemoryStringToInt = CInt("&H" & hWorkVal)
End Function

Private Function MemoryStringToByte(StringIn As String) As Byte
  On Error Resume Next
  Dim hWorkVal As String
    
  hWorkVal = Hex(AscB(MidB(StringIn, 1, 1))) ' only ONE byte
  '
  ' Return Byte value.
  MemoryStringToByte = CByte("&H" & hWorkVal)
End Function

Private Function MemoryStringToString(StringIn As String, ByVal Size As Integer) As String
  Dim Pos As Long
  For Pos = 1 To Size Step 1
    MemoryStringToString = MemoryStringToString & Chr(AscB(MidB(StringIn, Pos, 1)))
  Next Pos
End Function

Private Function LongToMemoryString(ByVal lInputValue As Long) As String
  Dim hWorkVal As String
  Dim n As Long
  Dim i As Long
  '
  ' Convert to HEX value.
  hWorkVal = Hex(lInputValue)
  
  '
  ' Check to see if it is not zero.
  If hWorkVal <> "0" Then
    '
    ' Place leading zeros in 8 character sequence to
    ' maintain consistent character count
    n = Len(hWorkVal)
    If n < 8 Then
        hWorkVal = String(8 - n, "0") & hWorkVal
    End If
    '
    ' Use ChrB to rebuild Bytes.
    For i = 8 - 1 To 1 Step -2
        LongToMemoryString = LongToMemoryString & _
                             ChrB(CInt("&H" & Mid(hWorkVal, i, 2)))
    Next i
       
  Else
    ' Just return zeros.
    ' Use ChrB to build Bytes.
    LongToMemoryString = ChrB(CInt("&H00"))
    LongToMemoryString = LongToMemoryString & ChrB(CInt("&H00"))
    LongToMemoryString = LongToMemoryString & ChrB(CInt("&H00"))
    LongToMemoryString = LongToMemoryString & ChrB(CInt("&H00"))
  End If
End Function

Private Function IntToMemoryString(ByVal iInputValue As Integer) As String
  Dim hWorkVal As String
  Dim n As Long
  Dim i As Long
  '
  ' Convert to HEX value.
  
  hWorkVal = Hex(iInputValue)
  
  '
  ' Check to see if it is not zero.
  If hWorkVal <> "0" Then
    ' Place leading zeros in 4 character sequence to
    ' maintain consistent character count
    n = Len(hWorkVal)
    If n < 4 Then
        hWorkVal = String(4 - n, "0") & hWorkVal
    End If
    '
    ' Use ChrB to rebuild Bytes.
    For i = 4 - 1 To 1 Step -2
        IntToMemoryString = IntToMemoryString & _
                             ChrB(CInt("&H" & Mid(hWorkVal, i, 2)))
    Next i

  Else
    ' Just return zeros.
    ' Use ChrB to build Bytes.
    IntToMemoryString = ChrB(CInt("&H00"))
    IntToMemoryString = IntToMemoryString & ChrB(CInt("&H00"))
  End If
End Function

Private Function ByteToMemoryString(ByVal Val As Byte) As String
  ByteToMemoryString = ChrB(Val)
End Function

Private Function StringToMemoryString(StringIn As String) As String
  Dim Pos, Length As Integer
  Length = Len(StringIn)
  For Pos = 1 To Length Step 1
    StringToMemoryString = StringToMemoryString & ChrB(Asc(Mid(StringIn, Pos, 1)))
  Next Pos
End Function

Private Function doRect(ByVal Left As Long, ByVal Top As Long, _
    ByVal Right As Long, ByVal Bottom As Long) As String
  doRect = LongToMemoryString(Left) & _
           LongToMemoryString(Top) & _
           LongToMemoryString(Right) & _
           LongToMemoryString(Bottom)
End Function

Private Sub getRect(ByVal rect As String, ByRef Left As Long, _
    ByRef Top As Long, ByRef Right As Long, ByRef Bottom As Long)
  Left = MemoryStringToLong(MidB(rect, 1, 4))
  Top = MemoryStringToLong(MidB(rect, 5, 4))
  Right = MemoryStringToLong(MidB(rect, 9, 4))
  Bottom = MemoryStringToLong(MidB(rect, 13, 4))
End Sub

Private Function doPoint(ByVal px As Long, ByVal py As Long) As String
  doPoint = LongToMemoryString(px) & _
            LongToMemoryString(py)
End Function

Private Sub getPoint(ByVal point As String, ByRef px As Long, ByRef py As Long)
  px = MemoryStringToLong(MidB(point, 1, 4))
  py = MemoryStringToLong(MidB(point, 5, 4))
End Sub

Private Function DrawPageText(ByVal dc As Long, ByVal text As String, _
    ByVal count As Long, ByVal rect As String) As Boolean
  DrawPageText = False
  If StartPage(dc) <= 0 Then Exit Function
  
  Dim Result As Long
  Result = DrawText(dc, text, count, rect, DRAW_TEXT_DRAW)
  
  If EndPage(dc) <= 0 Then Exit Function
  
  DrawPageText = True
End Function

Public Function GetPrinterDC(ByVal NoChoice As Boolean) As Long
  If NoChoice Then
    GetPrinterDC = CreateDC("CMP-10.DLL", "CMP-10", "IRDA", 0) ' or "IrDA" or ...
    Exit Function
  End If
  
  GetPrinterDC = 0
  
  Dim PageSetup As String
  PageSetup = LongToMemoryString(68) ' sizeof (PRINTDLG)
  PageSetup = PageSetup & LongToMemoryString(Screen.ActiveForm.hwnd)
  PageSetup = PageSetup & LongToMemoryString(0)
  ' add here PD_RETURNDEFAULTDC so default printer to be got
  PageSetup = PageSetup & LongToMemoryString(PD_DISABLEMARGINS + PD_MARGINS + PD_MINMARGINS)
  ' min margins here
  PageSetup = PageSetup & doRect(0, 0, 0, 0)
  ' margins here
  PageSetup = PageSetup & doRect(0, 0, 0, 0)
  PageSetup = PageSetup & LongToMemoryString(App.hInstance)
  ' unusable hook & templates...
  PageSetup = PageSetup & LongToMemoryString(0)
  PageSetup = PageSetup & LongToMemoryString(0)
  PageSetup = PageSetup & LongToMemoryString(0)
  PageSetup = PageSetup & LongToMemoryString(0)

  If PrintDlg(PageSetup) = 0 Then Exit Function
  
  GetPrinterDC = MemoryStringToLong(MidB(PageSetup, 9, 4))
End Function

Public Function PrintText(ByVal text As String, ByVal points As Long, ByVal clr As Long, _
    ByVal IsBold As Boolean, ByVal IsItalic As Boolean, ByVal IsStrike As Boolean, _
    ByVal IsUnderline) As Boolean
  PrintText = False
  On Error Resume Next
  
  Dim font, weight As Long
  Dim dc As Long
  
  dc = GetPrinterDC(True)
  If dc = 0 Then Exit Function
  
  If IsBold Then weight = 700 Else weight = 400
  Dim Help As String
  Help = LongToMemoryString(-(points * GetDeviceCaps(dc, LOGPIXELSY)) / 72) ' font size
  Help = Help & LongToMemoryString(0) ' appropriate width
  Help = Help & LongToMemoryString(0) ' no escapment
  Help = Help & LongToMemoryString(0) ' no orientation
  Help = Help & LongToMemoryString(weight) ' some weight THIN=100,NORMAL=400,BOLD=700
  Help = Help & ByteToMemoryString(IsItalic) ' italic
  Help = Help & ByteToMemoryString(IsUnderline) ' underline
  Help = Help & ByteToMemoryString(IsStrike) ' strikeout
  Help = Help & ByteToMemoryString(0) ' default charset
  Help = Help & ByteToMemoryString(0) ' default out precision
  Help = Help & ByteToMemoryString(0) ' default clip precision
  Help = Help & ByteToMemoryString(0) ' default quality
  Help = Help & ByteToMemoryString(0) ' default pitch
  Help = Help & "Tahoma" ' most popular and existing TTF
  font = CreateFontIndirect(Help)
  
  If font = 0 Then
    DeleteDC (dc)
    Exit Function
  End If
  
  ' Prepare DC
  Dim oldFont As Long
  oldFont = SelectObject(dc, font)
  Dim Result As Long
  Result = SetBkMode(dc, TRANSPARENT)
  Result = SetTextColor(dc, clr)  ' 0 for black
  
  Dim PageLeft, PageRight, PageTop, PageBottom, CalcRight, CalcBottom As Long
  PageLeft = 0
  PageTop = 0
  PageRight = GetDeviceCaps(dc, PHYSICALWIDTH)
  PageBottom = GetDeviceCaps(dc, PHYSICALHEIGHT)
  Dim Modified As String
  Modified = doRect(PageLeft, PageTop, PageRight, PageBottom)
  Result = DrawText(dc, text, -1, Modified, DRAW_TEXT_CALC)
  Call getRect(Modified, PageLeft, PageTop, CalcRight, CalcBottom)
  
  Help = LongToMemoryString(5 * 4) & _
         LongToMemoryString(0) & _
         LongToMemoryString(0) & _
         LongToMemoryString(0) & _
         LongToMemoryString(0)
  If StartDoc(dc, Help) <= 0 Then
    Result = SelectObject(dc, oldFont)
    DeleteDC (dc)
    DeleteObject (font)
    Exit Function
  End If
  
  Dim Pages As Long
  If (CalcBottom >= PageBottom) Then
    For Pages = 0 To CalcBottom \ PageBottom - 1 Step 1
      If Not DrawPageText(dc, text, -1, doRect(PageLeft, -Pages * PageBottom, PageRight, PageBottom)) Then
        AbortDoc (dc)
        Result = SelectObject(dc, oldFont)
        DeleteDC (dc)
        DeleteObject (font)
        Exit Function
      End If
    Next
  End If
  
  Pages = CalcBottom Mod PageBottom
  If Pages <> 0 Then
    If Not DrawPageText(dc, text, -1, doRect(PageLeft, Pages - CalcBottom, PageRight, PageBottom)) Then
      AbortDoc (dc)
      Result = SelectObject(dc, oldFont)
      DeleteDC (dc)
      DeleteObject (font)
      Exit Function
    End If
  End If
  
  If EndDoc(dc) <= 0 Then
    Result = SelectObject(dc, oldFont)
    DeleteDC (dc)
    DeleteObject (font)
    Exit Function
  End If

  ' finally - cleanup
  Result = SelectObject(dc, oldFont)
  DeleteDC (dc)
  DeleteObject (font)
  PrintText = True
End Function


Public Function PrintGraphic(ByVal bmp As Long) As Boolean
  PrintGraphic = False
  On Error Resume Next

  If bmp = 0 Then Exit Function
  Dim Help As String
  Help = LongToMemoryString(0) ' must be 0
  Help = Help & LongToMemoryString(0) ' width
  Help = Help & LongToMemoryString(0) ' height
  Help = Help & LongToMemoryString(0) ' width bytes
  Help = Help & IntToMemoryString(0) ' planes
  Help = Help & IntToMemoryString(0) ' bpp
  Help = Help & LongToMemoryString(0) ' ptr

  If GetObject(bmp, 6 * 4, Help) <> 6 * 4 Then Exit Function
  
  Dim dc, bmpWidth, bmpHeight As Long
  
  dc = GetPrinterDC(True)
  If dc = 0 Then Exit Function

  bmpWidth = MemoryStringToLong(MidB(Help, 5, 4))
  bmpHeight = MemoryStringToLong(MidB(Help, 9, 4))

  Help = LongToMemoryString(5 * 4) & _
         LongToMemoryString(0) & _
         LongToMemoryString(0) & _
         LongToMemoryString(0) & _
         LongToMemoryString(0)
  If StartDoc(dc, Help) <= 0 Or _
     StartPage(dc) <= 0 Then
    DeleteDC (dc)
    Exit Function
  End If
  
  If TransparentImage(dc, 0, 0, GetDeviceCaps(dc, PHYSICALWIDTH), GetDeviceCaps(dc, PHYSICALHEIGHT), _
                      bmp, 0, 0, bmpWidth, bmpHeight, -1) = 0 Then
    AbortDoc (dc)
    DeleteDC (dc)
    Exit Function
  End If
  
  If EndPage(dc) <= 0 Or _
     EndDoc(dc) <= 0 Then
    DeleteDC (dc)
    Exit Function
  End If

  ' finally - cleanup
  DeleteDC (dc)
  PrintGraphic = True
End Function

Public Function ReadMagstripe(ByRef Track1 As String, ByRef Track2 As String, _
    ByRef Track3 As String) As Boolean
  ReadMagstripe = False
  On Error Resume Next

  Dim Size, Cmd, Ptr As Long
  Dim Want(2) As Boolean
  Size = 256
  Cmd = &H8 ' separate info required
  If Len(Track1) <> 0 Then Cmd = Cmd + TRACK_1: Want(0) = True Else Want(0) = False
  If Len(Track2) <> 0 Then Cmd = Cmd + TRACK_2: Want(1) = True Else Want(1) = False
  If Len(Track3) <> 0 Then Cmd = Cmd + TRACK_3: Want(2) = True Else Want(2) = False
  If Not Want(0) And Not Want(1) And Not Want(2) Then Exit Function ' at least one track must be selected
  Ptr = LocalAlloc(LPTR, Size)
  If Ptr = 0 Then Exit Function
  
  Dim dc As Long
  dc = GetPrinterDC(True)
  If dc = 0 Then
    LocalFree (Ptr)
    Exit Function
  End If

  Dim iEsc As String
  iEsc = LongToMemoryString(ESCAPE_READSCARD)

  If ExtEscape(dc, ESCAPE_QUERYESCSUPPORT, 4, iEsc, 0, vbNullptr) <= 0 Then
    DeleteDC (dc)
    LocalFree (Ptr)
    Exit Function
  End If
  
  Dim Help As String
  Help = LongToMemoryString(Ptr)
  Help = Help & LongToMemoryString(Size)
  ' Cmd = Cmd - &H8 ' for printer with default separators on...
  Help = Help & LongToMemoryString(Cmd)
  If ExtEscape(dc, ESCAPE_READSCARD, 3 * 4, Help, 0, vbNullptr) <= 0 Then
    DeleteDC (dc)
    LocalFree (Ptr)
    Exit Function
  End If

  DeleteDC (dc) ' nothing more to do with it

  Size = MemoryStringToLong(MidB(Help, 5, 4))
  If Size <= 1 Then
    LocalFree (Ptr)
    Exit Function ' nothing (error or empty) read...
  End If
  
  Dim InfoB As String
  InfoB = String(256, 0)
  Call MoveMemory(InfoB, Ptr, Size)
  LocalFree (Ptr)
  
  ' now separate wanted data...
  Dim Pos, Track As Long
  Track = 0
  For Pos = 1 To Size Step 1
    If AscB(MidB(InfoB, Pos, 1)) = &HF1 Then
      Track = 1
      Track1 = ""
    ElseIf AscB(MidB(InfoB, Pos, 1)) = &HF2 Then
      Track = 2
      Track2 = ""
    ElseIf AscB(MidB(InfoB, Pos, 1)) = &HF3 Then
      Track = 3
      Track3 = ""
    Else ' distribute character to the right place...
      Select Case Track
        Case 1
          If Want(0) Then Track1 = Track1 & Chr(AscB(MidB(InfoB, Pos, 1)))
        Case 2
          If Want(1) Then Track2 = Track2 & Chr(AscB(MidB(InfoB, Pos, 1)))
        Case 3
          If Want(2) Then Track3 = Track3 & Chr(AscB(MidB(InfoB, Pos, 1)))
        Case Else ' here is presumption of track separator!
          Exit Function
      End Select
    End If
  Next Pos
  
  ReadMagstripe = True
End Function

Public Function ReadPrinterStatus(ByRef status As Integer) As Boolean
  ReadPrinterStatus = False
  On Error Resume Next

  Dim dc As Long
  dc = GetPrinterDC(True)
  If dc = 0 Then Exit Function

  Dim iEsc As String
  iEsc = LongToMemoryString(ESCAPE_READSTATUS)

  If ExtEscape(dc, ESCAPE_QUERYESCSUPPORT, 4, iEsc, 0, vbNullptr) <= 0 Then
    DeleteDC (dc)
    Exit Function
  End If

  Dim Result As String
  Result = String(1, 0)

  If ExtEscape(dc, ESCAPE_READSTATUS, 0, vbNullptr, 1, Result) <= 0 Then
    DeleteDC (dc)
    Exit Function
  End If

  DeleteDC (dc) ' never forgot to close (i.e. disconnect) it!
  status = MemoryStringToByte(Result)
  ReadPrinterStatus = True
End Function

Public Function ReadPrinterInfo(ByRef Temperature As Integer, ByRef Voltage As Integer) As Boolean
  ReadPrinterInfo = False
  On Error Resume Next

  Dim dc As Long
  dc = GetPrinterDC(True)
  If dc = 0 Then Exit Function

  Dim iEsc As String
  iEsc = LongToMemoryString(ESCAPE_READINFO)

  If ExtEscape(dc, ESCAPE_QUERYESCSUPPORT, 4, iEsc, 0, vbNullptr) <= 0 Then
    DeleteDC (dc)
    Exit Function
  End If

  Dim Result As String
  Result = String(2, 0)

  If ExtEscape(dc, ESCAPE_READINFO, 0, vbNullptr, 2, Result) <= 0 Then
    DeleteDC (dc)
    Exit Function
  End If

  DeleteDC (dc) ' never forgot to close (i.e. disconnect) it!
  Voltage = CByte("&H" & Hex(AscB(MidB(Result, 1, 1)))) - &H20  ' first byte is voltage
  Temperature = CByte("&H" & Hex(AscB(MidB(Result, 2, 1)))) - &H20  ' second byte is temperature
  ReadPrinterInfo = True
End Function

Public Function ReadPrinterCaps(ByRef Caps As Integer, ByRef Name As String) As Boolean
  ReadPrinterCaps = False
  On Error Resume Next

  Dim dc As Long
  dc = GetPrinterDC(True)
  If dc = 0 Then Exit Function

  Dim iEsc As String
  iEsc = LongToMemoryString(ESCAPE_READCAPABILITIES)

  If ExtEscape(dc, ESCAPE_QUERYESCSUPPORT, 4, iEsc, 0, vbNullptr) <= 0 Then
    DeleteDC (dc)
    Exit Function
  End If

  Dim Result As String
  Result = String(32, 0)

  If ExtEscape(dc, ESCAPE_READCAPABILITIES, 0, vbNullptr, 32, Result) <= 0 Then
    DeleteDC (dc)
    Exit Function
  End If

  DeleteDC (dc) ' never forgot to close (i.e. disconnect) it!
  
  Name = MemoryStringToString(MidB(Result, 1, 27), 27)
  Caps = MemoryStringToInt(MidB(Result, 28, 2))
  ReadPrinterCaps = True
End Function

Public Function ForceClosePrinter(ByVal dc As Long) As Boolean
  ForceClosePrinter = False
  If dc = 0 Then Exit Function

  Dim iEsc As String
  iEsc = LongToMemoryString(ESCAPE_PRINTERDISCONNECT)

  If ExtEscape(dc, ESCAPE_QUERYESCSUPPORT, 4, iEsc, 0, vbNullptr) > 0 Then
    If ExtEscape(dc, ESCAPE_PRINTERDISCONNECT, True, vbNullptr, 0, vbNullptr) > 0 Then
      ForceClosePrinter = True
    End If
  End If

  DeleteDC (dc) ' anyway delete it!
End Function

Public Function WritePrinterDirect(ByVal dc As Long, ByVal data As String, _
    ByVal Length As Integer) As Boolean
  WritePrinterDirect = False
  If dc = 0 Then Exit Function

  Dim iEsc As String
  iEsc = LongToMemoryString(ESCAPE_PASSTHROUGH)

  If ExtEscape(dc, ESCAPE_QUERYESCSUPPORT, 4, iEsc, 0, vbNullptr) <= 0 Then Exit Function

  If ExtEscape(dc, ESCAPE_PASSTHROUGH, Length, data, 0, vbNullptr) <= 0 Then Exit Function

  WritePrinterDirect = True
End Function

Public Function ReadPrinterDirect(ByVal dc As Long, ByRef Size As Integer) As String
  ReadPrinterDirect = "" ' empty string is an error
  If dc = 0 Then Exit Function

  Dim iEsc As String
  iEsc = LongToMemoryString(ESCAPE_READPRINTER)

  If ExtEscape(dc, ESCAPE_QUERYESCSUPPORT, 4, iEsc, 0, vbNullptr) <= 0 Then Exit Function

  Dim Result As String
  Result = String(Size, 0)

  Size = ExtEscape(dc, ESCAPE_READPRINTER, 0, vbNullptr, Size, Result)
  If Size <= 0 Then Exit Function

  ReadPrinterDirect = Result
End Function

Public Function ReadPrinterDirectTest() As Boolean
  ReadPrinterDirectTest = False
  On Error Resume Next

  Dim dc As Long
  dc = GetPrinterDC(True)
  If dc = 0 Then Exit Function

  Dim Cmd As String
  ' build commands: print test, feed 80/203", read status
  Cmd = ByteToMemoryString(27) & _
        ByteToMemoryString(Asc("T")) & _
        ByteToMemoryString(27) & _
        ByteToMemoryString(Asc("J")) & _
        ByteToMemoryString(80) & _
        ByteToMemoryString(27) & _
        ByteToMemoryString(Asc("v"))

  If WritePrinterDirect(dc, Cmd, 7) Then
    Dim sz As Integer
    sz = 1
    Cmd = ReadPrinterDirect(dc, sz)
    If sz = 1 Then
      ' process MemoryStringToByte (Cmd) here
      ReadPrinterDirectTest = True
    End If
  End If

  DeleteDC (dc) ' anyway delete it
End Function

Public Function ControlPassthrough(ByVal dc As Long, ByVal activate As Boolean) _
    As Boolean
  ControlPassthrough = False
  If dc = 0 Then Exit Function

  Dim iEsc As String
  iEsc = LongToMemoryString(ESCAPE_PRINTERPASSTHROUGH)

  If ExtEscape(dc, ESCAPE_QUERYESCSUPPORT, 4, iEsc, 0, vbNullptr) <= 0 Then Exit Function

  If ExtEscape(dc, ESCAPE_PRINTERPASSTHROUGH, activate, vbNullptr, 0, vbNullptr) <= 0 Then Exit Function

  ControlPassthrough = True
End Function


Public Function PrintBarcode(ByVal data As String, ByVal bartype As Byte, ByVal width As Byte, _
    ByVal height As Byte, ByVal isFontA As Boolean, ByVal textpos As Byte, ByVal feed As Byte) As Boolean

  PrintBarcode = False
  On Error Resume Next

  If data = "" Then Exit Function
  If bartype < 65 Or bartype > 73 Then Exit Function ' (bartype < 0 Or bartype > 6) And
  If width < 2 Or width > 4 Then Exit Function
  'If height < 1 Then Exit Function
  If textpos < 0 Or textpos > 3 Then Exit Function
  'If feed < 1 Then Exit Function

  Dim dc As Long
  dc = GetPrinterDC(True)
  If dc = 0 Then Exit Function

  Dim Length As Integer
  Length = Len(data)
  Dim font As Byte
  If isFontA Then font = 0 Else font = 1
  Dim Cmd As String
  '1Dh77h[2-4] width
  '1Dh68h[1-255] height
  '1Dh66h[0-1] font A/B
  '1Dh48h[0-3] text pos
  '1Dh6Bh[65-73][len][data] - print barcode
  '1Bh4Ah[0-255] line feed
  Cmd = ByteToMemoryString(&H1D) & _
        ByteToMemoryString(&H77) & _
        ByteToMemoryString(width) & _
        ByteToMemoryString(&H1D) & _
        ByteToMemoryString(&H68) & _
        ByteToMemoryString(height) & _
        ByteToMemoryString(&H1D) & _
        ByteToMemoryString(&H66) & _
        ByteToMemoryString(font) & _
        ByteToMemoryString(&H1D) & _
        ByteToMemoryString(&H48) & _
        ByteToMemoryString(textpos) & _
        ByteToMemoryString(&H1D) & _
        ByteToMemoryString(&H6B) & _
        ByteToMemoryString(bartype) & _
        ByteToMemoryString(Length) & _
        StringToMemoryString(data) & _
        ByteToMemoryString(27) & _
        ByteToMemoryString(Asc("J")) & _
        ByteToMemoryString(80)

  PrintBarcode = WritePrinterDirect(dc, Cmd, 19 + Length)
  DeleteDC (dc)
End Function
