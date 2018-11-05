Attribute VB_Name = "Bas_ComboBox"
Private Const ModuleName As String = "Bas_POS_ComboBox"

Public Sub SelectComboBoxItem(ByRef Cbo As ComboBox, ByVal Value As Boolean)

    Dim idx As Double
    
    With Cbo
        
        For idx = 0 To .ListCount - 1
            If .ItemData(idx) = Value Then
                .ListIndex = idx
            End If
        Next
    
    End With

End Sub

Public Sub PrinterComboBox(ByRef Cbo As ComboBox)

    Dim Counter As Integer
    
    'Load all installed printers
    With Cbo
        For Counter = 0 To Printers.Count - 1
            .AddItem Printers(Counter).DeviceName
        Next Counter
        .ListIndex = 0
    End With

End Sub

