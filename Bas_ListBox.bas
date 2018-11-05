Attribute VB_Name = "Bas_ListBox"
Private Const ModuleName As String = "Bas_ListBox"

Public Sub ListBoxSelect(ByRef LstBox As ListBox, _
                                            ByVal ListBoxDataFieldRs As ADODB.Recordset, _
                                            ByVal DataFieldIdx As Integer, _
                                            ByVal SelectedFieldRs As ADODB.Recordset, _
                                            ByVal SelectedFieldIdx As Integer)
    
    Dim Criteria As String
    Dim delimiter As String
    
    If ListBoxDataFieldRs.RecordCount > 0 Then
            
        With SelectedFieldRs
            
            If .RecordCount > 0 Then
                     
                 delimiter = ADO_FieldDelimiter(ListBoxDataFieldRs(DataFieldIdx).Type)
                    
                    .MoveFirst
            
                    Do Until .EOF = True
                                                
                            Criteria = ListBoxDataFieldRs.Fields(DataFieldIdx).Name
                            Criteria = Criteria & "="
                            Criteria = Criteria & delimiter
                            Criteria = Criteria & .Fields(SelectedFieldIdx)
                            Criteria = Criteria & delimiter
                            
                            With ListBoxDataFieldRs
                            
                                .MoveFirst
                                    
                                .Find Criteria
                                If Not .EOF = True Then
                                    LstBox.Selected(.AbsolutePosition - 1) = True
                                End If
                                
                            End With
                           
                        
                         .MoveNext
                        
                    Loop
                    
             End If
    
        End With
        
        LstBox.ListIndex = 0
        
    End If
    
End Sub
Public Sub SelectLstItem(ByRef LstBox As ListBox, ByVal value As Boolean)

    Dim Idx As Double
    
    With LstBox
    
        For Idx = 0 To .ListCount - 1
            .Selected(Idx) = value
        Next
    
    End With

End Sub
