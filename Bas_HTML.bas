Attribute VB_Name = "Bas_HTML"
Private Const ModuleName As String = "Bas_HTML"

Public Function GetHTMLTitle(ByVal data As String) As String
    
    Dim StartLoc As Long
    Dim EndLoc As Long
    
    Dim Title As String
    
    data = Replace(data, Chr(10), "")
    data = Replace(data, Chr(13), "")
    data = Replace(data, Chr(8), "")
    
    StartLoc = InStr(1, LCase(data), "<title>")
    
    If Len(StartLoc) = 0 Then
        Exit Function
    End If
    
    Title = Mid(data, StartLoc + 7)
    
    EndLoc = InStr(1, LCase(Title), "</title>")
    
    If EndLoc = 0 Then
        EndLoc = 50
    End If
    
    Title = Mid(Title, 1, EndLoc - 1)
    
    GetHTMLTitle = Title
    
End Function

Public Function AddRootToAbsolutePath(ByVal HTML As String, ByVal Root As String) As String

    '3 August
    
    'Replace href = "/..." to  href = "http://local/..."
    
    Dim data As String
    
    data = HTML
    
    data = Replace(data & " ", """/", """" & Root & "/")
    data = Replace(data & " ", """../", """" & Root & "../")
    
    AddRootToAbsolutePath = data
    
End Function

Public Function GetRoot(ByVal URL As String) As String
    
    '3 August
    
    'Extract root from an URL
    
    'e.g.
    
    'http://www.yahoo.com/news/top.htm
    
    'http://www.yahoo.com/news
    
    Dim lastslash As Integer
    Dim Root As String
    
    Root = Replace(LCase(URL), "://", ":\\")
    
    Root = StrReverse(Root)
    
    lastslash = InStr(1, Root, "/")
    
    If lastslash > 0 Then
        
        Root = Mid(Root, lastslash + 1)
        Root = StrReverse(Root)
        Root = Replace(LCase(Root), ":\\", "://")
        
        GetRoot = Root
            
    Else
    
        GetRoot = URL
    
    End If
    
    
    
    
End Function

