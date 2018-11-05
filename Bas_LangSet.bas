Attribute VB_Name = "Bas_LangSet"
Public Lang As String

Public Function LangSet(ByVal VBName As String, ByVal SetID As Integer) As String
    
    Dim Rs As New ADODB.Recordset
    Dim SQL As String
    Dim SelLang As String
    
    Dim ConnectionString As String
    
    Static LangCN As New ADODB.Connection
    Static CnOpened As Boolean
    Static SetArray() As String
    Static cur_VBName As String
    Static cur_MaxSetID As Integer
    Static cur_Lang As String
    
    On Error GoTo ErrorHandler

    If CnOpened = False Then
        ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath("LangSet.cfg")
        LangCN.Open ConnectionString
        CnOpened = True
    End If
    
    Select Case UCase(Lang)
    
    Case "BIG5"
        SelLang = "BIG5"
    
    Case "GB"
        SelLang = "GB"
    
    Case Else
        SelLang = "ENG"
        
    End Select
    
    Lang = SelLang
    
    If Not cur_VBName = VBName Or _
      Not cur_Lang = SelLang Then
        SQL = "select setid, expression from languageset "
        SQL = SQL & " where vbname='" & VBName & "'"
        
        SQL = SQL & " and Lang='" & SelLang & "'"
        SQL = SQL & " order by setid"
        
        Rs.Open SQL, LangCN, 1, adLockReadOnly

        ReDim SetArray(0 To Rs.RecordCount - 1)
        
        Do Until Rs.EOF
                If Not Trim(Rs("Expression")) = 0 Then
                    SetArray(Rs("setid")) = Rs("expression")
                End If
                Rs.MoveNext
        Loop
        
        Rs.Close
        
        cur_VBName = VBName
        cur_Lang = SelLang
        
    End If
    
    If SetID <= UBound(SetArray) Then
        LangSet = SetArray(SetID)
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    
    LogFormError "LangSet", "LangSet", "(" & VBName & "," & CStr(SetID) & ")"

End Function

