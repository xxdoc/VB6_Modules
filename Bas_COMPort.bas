Attribute VB_Name = "Bas_COMPort"
Public Const ConstIntMinPort = 1
Public Const ConstIntMaxPort = 256

Public Sub CreateCOMPortDropDown(ByRef Combo1 As ComboBox, Min As Integer, Max As Integer)
    Dim Counter As Integer
    
    With Combo1
        .Clear
        For Counter = Min To Max
            .AddItem "COM " & CStr(Counter)
        Next
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With

End Sub

Public Function SetMsComm(ByRef Ctrl_MsComm As MSComm, Baudrate As Double, Parity As String, DataBits As Integer, StopBit As Integer) As Boolean
    
    Dim Settings As String
    
    On Error Resume Next
    
    With Ctrl_MsComm
    
        If .PortOpen = True Then
           .PortOpen = False
        End If
        
        Settings = CStr(Baudrate) & ","
        Settings = Settings & CStr(Parity) & ","
        Settings = Settings & CStr(DataBits) & ","
        Settings = Settings & CStr(StopBit)
        
        .Settings = Settings
        
    End With
    
    SetMsComm = True
    
End Function

Public Function GetMsCommSettings(ByRef Ctrl_MsComm As MSComm, ByRef Baudrate As Double, ByRef Parity As String, ByRef DataBits As Integer, ByRef StopBit As Integer) As Boolean
    
    Dim Settings As String
    
    On Error GoTo ErrorHandler
    
    With Ctrl_MsComm
    
        Settings = .Settings
        
        Baudrate = CDbl(Mid(Settings, 1, InStr(1, Settings, ",")))
        
        Settings = Mid(Settings, InStr(1, Settings, ",") + 1)
        Parity = UCase(Mid(Settings, 1, 1))
        
        Settings = Mid(Settings, InStr(1, Settings, ",") + 1)
        DataBits = CInt(Mid(Settings, 1, 1))
        
        Settings = Mid(Settings, InStr(1, Settings, ",") + 1)
        StopBit = CInt(Mid(Settings, 1, 1))

End With
    
    GetMsCommSettings = True
    
    Exit Function
    
ErrorHandler:
    GetMsCommSettings = False
    
End Function

Public Function WritePrinter(H_Printer As Printer, CommandString() As Byte) As Boolean
    
    On Error GoTo ErrorHandler
    
    'Write Command to Printer
    'To call this function, a printer object and a byte array with commands should be passed
    Dim PortNo As Long
    Dim msgStr As String
    
    'Command will be write to printer port only
    WriteCommand = False
    If (UCase(Left(H_Printer.Port, 3)) = "LPT") And (Right(H_Printer.Port, 1) = ":") Then
        PortNo = FreeFile
        Open H_Printer.Port For Binary As PortNo
        Put PortNo, , CommandString
        Close PortNo
        WritePrinter = True
    End If
    
    Exit Function
    
ErrorHandler:
        MsgBox Err.Description, vbCritical
    
End Function

Public Function MSCommSend(Command As String, MSComm As MSComm) As String
' Send a command to the terminal and get the result.
' Parameters : 1. Command -- A character string that will be send.
'              2. MSComm  -- A VB communication object.
'              3. GetCOMPort -- A integer com port number.
' Return : Result -- A character string.

Dim Response As String

On Error GoTo ErrorHandler

Dim StartTime As Single

If Not Err.Number = 0 Then
    Exit Function
End If

StartTime = Timer

MSComm.Output = Command
                
Do Until Timer >= StartTime + 3
    
    DoEvents
    
    If MSComm.InBufferCount > 0 Then
    
        Response = Response + MSComm.Input
        
        If Right(Response, 1) = Chr(13) Then
            MSCommSend = Left(Response, Len(Response) - 1)
            Exit Function
        End If
        
    End If
    
Loop

Exit Function

ErrorHandler:

   
End Function

