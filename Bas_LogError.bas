Attribute VB_Name = "Bas_LogError"
Public Const FILE_PROGRAM_ERROR = "error.log"

Public Sub LogError(ByVal ErrorMessage As String)

    On Error GoTo ErrorHandler
        AppendFile AppPath(FILE_PROGRAM_ERROR), ErrorMessage
    Exit Sub
ErrorHandler:
    Exit Sub
    
End Sub

Public Function LogFormError(ByVal FrmName As String, ByVal Routine As String, ByVal ErrDesc As String) As Boolean
    
    Dim Details As String
    
    On Error GoTo ErrorHandler
    Details = "Date: " & Format(Now(), "yyyy/mm/dd") & vbNewLine
    Details = Details & "Time: " & Format(Now(), "hh:mm:ss") & vbNewLine
    Details = Details & "Form: " & FrmName & vbNewLine
    Details = Details & "Routine: " & Routine & vbNewLine
    Details = Details & "Error Information: " & ErrDesc & vbNewLine
    
    Call LogError(Details)
    
    LogFormError = True
    
    Exit Function
    
ErrorHandler:
    LogFormError = False
    
End Function

Public Function ClearLogFile() As Boolean
    
    On Error GoTo ErrorHandler
    Kill AppPath(FILE_PROGRAM_ERROR)
    
    ClearLogFile = True
    
    Exit Function

ErrorHandler:
    ClearLogFile = False
    
    
End Function
