Attribute VB_Name = "Bas_CommDialog"
Private Const ModuleName As String = "Bas_CommDialog"

Public Function OpenFileDialog(ByRef CDL As CommonDialog) As String
On Error Resume Next
With CDL
    .CancelError = True
    .ShowOpen
    If Not Err.Number = 0 Then
        Exit Function
    End If
    OpenFileDialog = .FileName
End With
End Function


Public Function OpenPrinterDialog(ByRef CDL As CommonDialog) As CommonDialog
On Error Resume Next

With CDL
    .CancelError = True
    
    .ShowPrinter
    
    If Not Err.Number = 0 Then
        Exit Function
    End If
    
    Set OpenPrinterDialog = CDL
    
End With

End Function

Public Function OpenFontDialog(ByRef CDL As CommonDialog) As CommonDialog

On Error Resume Next
With CDL
    .Flags = 0   'Required before setting FontName
    .FontName = "MS Sans Serif"
     .Flags = cdlCFBoth + cdlCFEffects
    .CancelError = True
    .ShowFont
    If Not Err.Number = 0 Then
        Exit Function
    End If
    Set OpenFontDialog = CDL
End With

End Function

Public Function OpenColorDialog(ByRef CDL As CommonDialog) As CommonDialog
On Error Resume Next
With CDL
    .ShowColor
    .CancelError = True
    If Not Err.Number = 0 Then
        Exit Function
    End If
    Set OpenColorDialog = CDL
End With

End Function


Public Function SaveFileDialog(ByRef CDL As CommonDialog) As String
On Error Resume Next

With CDL
    .CancelError = True
    .ShowSave
    If Not Err.Number = 0 Then
        Exit Function
    End If
    SaveFileDialog = .FileName
End With

End Function
