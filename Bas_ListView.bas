Attribute VB_Name = "Bas_ListView"
Private Const ModuleName As String = "Bas_ListView"
Public Sub SelectLstViewItem(ByRef LstView As ListView, ByVal value As Boolean)

    Dim idx As Double
    
    With LstView
    
        For idx = 1 To .ListItems.Count
            .ListItems(idx).Selected = value
        Next
    
    End With

End Sub

Public Sub CheckLstViewItem(ByRef LstView As ListView, ByVal value As Boolean)

    Dim idx As Double
    
    With LstView
    
        For idx = 1 To .ListItems.Count
            .ListItems(idx).Checked = value
        Next
    
    End With

End Sub

Public Function ListViewSelItem(ByRef LstView As ListView) As Double

    Dim idx As Double
    Dim cnt As Double
    
    cnt = 0
    
    With LstView
    
        For idx = 1 To .ListItems.Count
            If .ListItems(idx).Selected = True Then
                cnt = cnt + 1
            End If
        
        Next
    
    End With

    ListViewSelItem = cnt

End Function

Public Sub ListViewReverseSelection(ByRef LstView As ListView)

    Dim idx As Double
        
    With LstView
    
        For idx = 1 To .ListItems.Count
             .ListItems(idx).Selected = Not .ListItems(idx).Selected
        Next
    
    End With


End Sub

