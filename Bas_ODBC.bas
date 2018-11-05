Attribute VB_Name = "Bas_ODBC"
Private Const ModuleName As String = "Bas_ODBC"
Public Function GetDataLinks()

     'The OLE DB Service Component for displaying the Data Link Properties
    Dim MSDASCObj As MSDASC.DataLinks
    
    Dim connstr As String
    
    Set MSDASCObj = New MSDASC.DataLinks
    
    On Error Resume Next
    'if cancel is pressed, it returns error
    connstr = MSDASCObj.PromptNew
    
    GetDataLinks = connstr
    
End Function

