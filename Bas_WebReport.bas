Attribute VB_Name = "Bas_WebReport"
Private Const ModuleName As String = "Bas_WebGrid"


Public Sub BuildWebGrid(ByVal ConnectionString As String, _
                                              ByVal Rs As ADODB.Recordset, _
                                              ByVal ASPFile As String)

    Dim WriteTs As TextStream
               
    Dim Fieldidx As Integer
    
    Dim SQL As String
    
    Set WriteTs = Fso.OpenTextFile(ASPFile, ForWriting, True)
    
        WriteTs.WriteLine "<%"
        WriteTs.WriteLine "Dim Cn"
        WriteTs.WriteLine "Dim Rs"
        WriteTs.WriteLine "Dim Connstr"
        WriteTs.WriteLine "Dim SQL"
        WriteTs.WriteLine "Connstr = """ & ConnectionString & """"
        WriteTs.WriteLine "Set Cn = server.CreateObject(""adodb.connection"")"
        WriteTs.WriteLine "Cn.Open Connstr"
        WriteTs.WriteLine ""
        WriteTs.WriteLine "Set Rs = server.CreateObject(""adodb.recordset"")"
        
        SQL = "SQL =""" & Rs.Source & """"
        
        SQL = Replace(LCase(SQL), " where ", """" & vbNewLine & "SQL = SQL & "" where ")
        SQL = Replace(LCase(SQL), " left outer join ", """" & vbNewLine & "SQL = SQL & "" left outer join ")
        
        SQL = LCase(SQL)
        
        WriteTs.WriteLine SQL
        WriteTs.WriteLine "Rs.Open SQL, Cn, 1, 1"
        WriteTs.WriteLine "%>"
        WriteTs.WriteLine "<html>"
        WriteTs.WriteLine "<head>"
        WriteTs.WriteLine "<!-- Code Ready " & App.Major & "." & App.Minor & "." & App.Revision & " " & Format(Now(), "yyyy/MM/dd HH:mm:ss") & "-->"
        WriteTs.WriteLine "<title>" & Rs.Source & "</title>"
        WriteTs.WriteLine "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
        WriteTs.WriteLine "</head>"
        WriteTs.WriteLine ""
        WriteTs.WriteLine "<body bgcolor=""#FFFFFF"" text=""#000000"">"
        WriteTs.WriteLine "<b>Source:</b>" & Rs.Source
        WriteTs.WriteLine "<br>"
        WriteTs.WriteLine "<b>Data Date Time:</b> <%=now()%>"
        WriteTs.WriteLine "<br>"
        WriteTs.WriteLine "<b>Record count:</b> <%=rs.recordcount%>"
        WriteTs.WriteLine "<table width=""100%"" border=""1"" cellspacing=""0"" cellpadding=""0"">"
        WriteTs.WriteLine "  <tr bgcolor=""#999999""> "
        
        For Fieldidx = 0 To Rs.Fields.Count - 1
            
            WriteTs.WriteLine "    <td >" & Rs(Fieldidx).Name & "&nbsp</td>"
        
        Next
        
        WriteTs.WriteLine "  </tr>"
        WriteTs.WriteLine "  <%"
        WriteTs.WriteLine "    Do Until Rs.eof"
        WriteTs.WriteLine "%>"
        WriteTs.WriteLine "  <tr> "
        
        For Fieldidx = 0 To Rs.Fields.Count - 1
            
            WriteTs.WriteLine "    <td><%=Rs(" & Fieldidx & ")%>&nbsp</td>"
        
        Next
        
        WriteTs.WriteLine "  </tr>"
        WriteTs.WriteLine "  <%"
        WriteTs.WriteLine "        rs.movenext"
        WriteTs.WriteLine "    Loop"
        WriteTs.WriteLine ""
        WriteTs.WriteLine "    rs.close"
        WriteTs.WriteLine "    cn.close"
        WriteTs.WriteLine "%>"
        WriteTs.WriteLine "</table>"
        WriteTs.WriteLine "</body>"
        WriteTs.WriteLine "</html>"

        WriteTs.Close
        
        Set WriteTs = Nothing

End Sub
