Attribute VB_Name = "Bas_Form"
Private Const ModuleName As String = "Bas_Form"

Public Sub ClearFormTextField(ByRef Frm As Form)
    Dim idx As Integer
    
    On Error GoTo ErrorHandler
    
    With Frm
        
        For idx = 0 To .txtfield.Count - 1
               .txtfield(idx).Text = ""
        Next

    End With
    
    Exit Sub
    
ErrorHandler:
    
    Call LogFormError(ModuleName, "ClearFormTextField(" & Frm.Name & ")", Err.Description)
    Resume Next

End Sub


