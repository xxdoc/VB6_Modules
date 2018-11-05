Attribute VB_Name = "Bas_JRO"
Public Function CompactDatabase(Filepath As String) As Boolean

    Dim Connstr As String
    Dim Jr As New JRO.JetEngine
    Dim TempFile As String
    
    TempFile = AppPath("jro.tmp")
    On Error Resume Next
    CompactDatabase = True
    Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Filepath & ";Persist Security Info=False"
    Jr.CompactDatabase Connstr, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & TempFile & " ;Jet OLEDB:Engine Type=4"
    If Not Err.Number = 0 Then
        CompactDatabase = False
        Exit Function
    End If
    FileCopy TempFile, compactfile
    Kill TempFile
    
End Function
