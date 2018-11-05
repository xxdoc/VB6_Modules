Attribute VB_Name = "Bas_DirectIO"
Public Function DirectWrite(H_Printer As Printer, CommandString() As Byte) As Boolean
    
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
        WriteCommand = True
    End If
    
    Exit Function
    
ErrorHandler:
        MsgBox Err.Description, vbCritical
    
End Function
