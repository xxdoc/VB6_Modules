Attribute VB_Name = "Bas_DataGrid"
Private Const ModuleName As String = "Bas_DataGrid"
Public Sub SetDataGridColWidth(ByRef DG As DataGrid)
        
        Dim idx As Double
        Dim Cnt As Double
        
        With DG
                Cnt = 0
                For idx = 0 To .Columns.Count - 1
                    If .Columns(idx).Visible = True Then
                        Cnt = Cnt + 1
                    End If
                Next
                
                For idx = 0 To .Columns.Count - 1
                    
                     If .Columns(idx).Visible = True Then
                        .Columns(idx).Width = (.Width * 0.96) / Cnt
                    Else
                        .Columns(idx).Width = 0
                    End If
                Next
    
        End With

End Sub

Public Sub SetDataGridStyle(ByRef DG As DataGrid)

    With DG
                
        .RowDividerStyle = dbgLightGrayLine
        .BorderStyle = dbgFixedSingle
        
    End With
    
End Sub
