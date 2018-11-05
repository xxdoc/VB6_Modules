Attribute VB_Name = "Bas_CodeReady"
Private Const ModuleName As String = "Bas_CodeReady"

Public Sub OpenTable(FormGrid As FrmGrid, CN As ADODB.Connection, TableName As String)

    MousePointer = vbHourglass
    
    On Error GoTo ErrorHandler
    
    With FormGrid
    
    With .Adodc1

        .ConnectionString = CN.ConnectionString
        .CommandType = adCmdText
        .RecordSource = "select * from [" & TableName & "]"
        .Refresh
        .Caption = .Recordset.RecordCount & " records"
        
    End With
    
    With .DataGrid1
        Set .DataSource = FormGrid.Adodc1
        .Refresh
    End With
    
    .Caption = .Adodc1.RecordSource

End With
        
    MousePointer = vbDefault
    
    Exit Sub
    
ErrorHandler:
        MousePointer = vbDefault
        MsgBox Err.Description, vbExclamation
        Resume Next

End Sub
