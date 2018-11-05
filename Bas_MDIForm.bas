Attribute VB_Name = "Bas_MDIForm"
Private Const ModuleName As String = "Bas_MDIForm"

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, _
ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_INVALIDATE = &H1
Private Const RDW_UPDATENOW = &H100

Public Sub SetMDIFormPicture(ByRef Frm As MDIForm, ByVal FilePath As String)
    
    
    
    'Assign picture
    
    On Error GoTo ErrorHandler
    
    With Frm
    
        If FSO.FileExists(FilePath) = False Then
            .Picture = Nothing
        Else
            .Picture = LoadPicture(FilePath)
        End If
    
    End With
    
    'Redraw Window
    Call RedrawWindow(Frm.hWnd, ByVal 0&, 0&, _
         RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)
    
    
    Exit Sub
    
ErrorHandler:
    
    Call LogFormError(ModuleName, "SetMDIFormPicture(" & Frm.Name & ")", Err.Description)

End Sub
