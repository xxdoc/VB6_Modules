Attribute VB_Name = "Bas_MultiLanguage"
Public Cn As New ADODB.Connection
Private FSO As New FileSystemObject
Public Function GetObjectName(ByVal Line As String) As String

    Dim TempLine As String
    Dim IDX As String
    
    TempLine = Line
    TempLine = LTrim(TempLine)
    
    TempLine = GetSeperatedText(TempLine, " ", 2)
    
    GetObjectName = TempLine
    
End Function
Public Function IsObjectBegin(ByVal Line As String) As Boolean
    
    Dim TempLine As String
    
    TempLine = Line
    
    TempLine = LTrim(TempLine)
    
    If Mid(TempLine, 1, 5) = "Begin" Then
        IsObjectBegin = True
    End If
    
End Function

Public Function IsExpression(ByVal Line As String, ByVal Expression As String) As Boolean
    
    Dim TempLine As String
    
    TempLine = Line
    
    TempLine = LTrim(TempLine)
    
    If Mid(TempLine, 1, Len(Expression)) = Expression Then
        IsExpression = True
    End If
    
End Function

Public Function IsObjectEnd(ByVal Line As String) As Boolean
    
    Dim TempLine As String
    
    TempLine = Line
    
    TempLine = LTrim(TempLine)
    
    If Mid(TempLine, 1, 3) = "End" Then
        IsObjectEnd = True
    End If
    
End Function

Public Function RemoveFormObjectsInit(ByVal FrmFile As String) As String
        
        Dim SourceTs As TextStream
        Dim TempFile As String
        Dim Content As String
        
        Dim Line As String
        Dim TempLine As String
        
        Dim FindString As String
        Dim CallSubStatement As String
        
         TempFile = AppPath(FSO.GetTempName)
        
        FileCopy FrmFile, TempFile
        
        Set SourceTs = FSO.OpenTextFile(TempFile, ForReading, False)
        
        Do Until SourceTs.AtEndOfStream
            DoEvents
            Line = SourceTs.ReadLine
            TempLine = Line
            
            If IsExpression(TempLine, "Private Sub FormObjectsInit()") = True Then
                
                Do Until IsExpression(TempLine, "End Sub") = True
                    
                    
                    Line = SourceTs.ReadLine
                    TempLine = Line
                    
                Loop
                TempLine = ""
                Line = ""
                
            End If
            Content = Content & Line & vbNewLine
        Loop
        
        SourceTs.Close
        FindString = "'MultiLanguages - Form Objects Initialize"
        Content = Replace(Content, FindString, "")
        
        CallSubStatement = "Call FormObjectsInit()"
        FindString = CallSubStatement
        Content = Replace(Content, FindString, "")
        
        RemoveFormObjectsInit = Content
        
        Kill TempFile
                
End Function

Public Function CreateFormObjectsInit(ByVal FrmFile As String) As String

        Dim TS As TextStream
        Dim TempFile As String
        
        Dim Line As String
        Dim TempLine As String
        Dim FileContent As String
        
        Dim SubContent As String
        Dim FrmType As String
        Dim LoadSubStatement As String
        Dim CallSubStatement As String
        Dim ReplaceString As String
        
        TempFile = AppPath(FSO.GetTempName)
         
         FileCopy FrmFile, TempFile
         
        Set TS = FSO.OpenTextFile(TempFile, ForReading, False)
        
        Do Until TS.AtEndOfStream
            DoEvents
            Line = TS.ReadLine
            TempLine = Line
            
            If IsObjectBegin(TempLine) = True Then
                    SubContent = SubContent & CreateObjectInit(TS, Line)
            End If

        Loop
        TS.Close

        
        Set TS = FSO.OpenTextFile(TempFile, ForReading, False)
        
        Do Until TS.AtEndOfStream
                DoEvents
                Line = TS.ReadLine
                TempLine = Line
                FileContent = FileContent & Line & vbNewLine
                'Assume The first narrated object is form object
                If IsObjectBegin(TempLine) = True And Len(FrmType) = 0 Then
                
                    'Form Type : MDIForm, Form
                    FrmType = GetSeperatedText(TempLine, " ", 1)
                    FrmType = GetSeperatedText(FrmType, ".", 1)
                    
                End If
                
        Loop
        
        TS.Close
            
        'Add Call Sub to Form_Load
        If Not Len(SubContent) = 0 Then

            SubContent = vbNewLine & SubContent
            SubContent = "Private Sub FormObjectsInit()" & vbNewLine & SubContent
            SubContent = SubContent & vbNewLine & vbNewLine & "End Sub" & vbNewLine

            LoadSubStatement = "Private Sub " & FrmType & "_Load()"
            CallSubStatement = "Call FormObjectsInit()"
            ReplaceString = LoadSubStatement & vbNewLine
            ReplaceString = ReplaceString & vbTab & "' MultiLanguages - Form Objects Initialize" & vbNewLine
            ReplaceString = ReplaceString & vbTab & CallSubStatement & vbNewLine

            FileContent = Replace(FileContent, LoadSubStatement, ReplaceString)

            If InStr(1, FileContent, CallSubStatement) = 0 Then
                FileContent = FileContent & vbNewLine & LoadSubStatement & vbNewLine
                FileContent = FileContent & vbTab & CallSubStatement & vbNewLine
                FileContent = FileContent & "End Sub" & vbNewLine
            End If
            
        End If
        
        CreateFormObjectsInit = FileContent & SubContent
        
        Kill TempFile

End Function

Public Function CreateObjectInit(ByRef TS As TextStream, ByVal ThisObjectLine As String) As String
    
    'This funciton is a recursive function
    
    Dim TempStr As String
    Dim Line As String
    Dim SubContent As String
    Dim WithContent As String
    Dim StartWith As String
    Dim ObjectName As String
    
    ObjectName = GetObjectName(ThisObjectLine)
    
    StartWith = "With " & ObjectName
    
    Do Until IsObjectEnd(TempStr) = True Or TS.AtEndOfStream = True

            Line = TS.ReadLine
            TempStr = Line
            If IsObjectBegin(TempStr) = True Then
                SubContent = SubContent & CreateObjectInit(TS, Line)
                
            Else
                
                If Len(GetValue(TempStr, "Caption")) > 0 Then
                    WithContent = WithContent & vbTab & vbTab & ".caption = " & GetValue(TempStr, "Caption") & vbNewLine
                End If
                
                
                If Len(GetValue(TempStr, "ToolTipText")) > 0 Then
                    WithContent = WithContent & vbTab & vbTab & ".tooltiptext = " & GetValue(TempStr, "ToolTipText") & vbNewLine
                End If
                
                If Len(GetValue(TempStr, "Text")) > 0 Then
                    WithContent = WithContent & vbTab & vbTab & ".text = " & GetValue(TempStr, "Text") & vbNewLine
                End If
                
                
                If Len(GetValue(TempStr, "Index")) > 0 Then
                    StartWith = "With " & ObjectName & "(" & GetValue(TempStr, "Index") & ")"
                End If
                
            End If
    Loop
    
    If Len(WithContent) > 0 Then
        SubContent = SubContent & vbTab & StartWith & vbNewLine & WithContent & vbTab & "End With" & vbNewLine
    End If

    CreateObjectInit = SubContent
        
End Function

Public Sub AddLanguageSet(ByVal Filename As String, ByVal VBName As String, ByVal SetID As Integer, ByVal Expression As String, ByVal Lang As String)
    
    Dim SQL As String
    Dim Fields As String
    Dim Values As String
    Dim TempStr As String
    
    SQL = "insert into languageset "
    
    Fields = "("
    Values = " values ("
        
    Fields = Fields & "VBname,"
    Values = Values & "'" & VBName & "',"
    
    Fields = Fields & "SetID,"
    Values = Values & SetID & ", "
    
    TempStr = Replace(Expression, "'", "''")
    Fields = Fields & "Expression,"
    Values = Values & "'" & TempStr & "',"
    
    Fields = Fields & "Lang,"
    Values = Values & "'" & Lang & "',"
    
    Fields = Mid(Fields, 1, Len(Fields) - 1)
    Values = Mid(Values, 1, Len(Values) - 1)
    
    Fields = Fields & ")"
    Values = Values & ")"

    SQL = SQL & Fields & Values

    Cn.Execute SQL
    
End Sub

Public Sub CreateConversionList(ByVal Path As String)

    Dim Line As String
    
    Dim TempStr As String
    Dim StartQuoteLoc As Integer
    Dim EndQuoteLoc As Integer
    Dim ExtractedStr As String
    
    Dim Counter As Double
    Dim Response As String
    Dim Message As String
    Dim ObjectOpened As Integer
    
    
    Dim SQL As String
    Dim StringIDX As Double
    
    Set TS = FSO.OpenTextFile(Path, ForReading, False)
    
    StringIDX = 0
    ObjectOpened = 0
    
    SQL = "delete * from conversionlist"
    SQL = SQL & " where filename='" & Path & "'"
    
    Cn.Execute SQL
    
    Do Until TS.AtEndOfStream
        
        DoEvents
        Line = TS.ReadLine
        TempStr = Line
        
        'Form Object
        If IsObjectBegin(TempStr) = True Then
            ObjectOpened = ObjectOpened + 1
            Do Until ObjectOpened = 0
                    Line = TS.ReadLine
                    TempStr = Line
                    If IsObjectBegin(TempStr) = True Then
                        ObjectOpened = ObjectOpened + 1
                    ElseIf IsObjectEnd(TempStr) = True Then
                        ObjectOpened = ObjectOpened - 1
                    End If
            Loop
            Line = TS.ReadLine
            TempStr = Line
        End If
        
        If Not IsExpression(TempStr, "Object") = True _
           And Not IsExpression(TempStr, "Attribute") = True Then
        
            StartQuoteLoc = -1
        
            Do Until StartQuoteLoc = 0
            
                    ExtractedStr = ""
                    
                    StartQuoteLoc = InStr(1, TempStr, """")
                    If StartQuoteLoc > 0 Then
                        
                        TempStr = Mid(TempStr, StartQuoteLoc + 1)
                        EndQuoteLoc = InStr(1, TempStr, """")
                        
                        If EndQuoteLoc > 0 Then
                            
                            Counter = Counter + 1
                            
                            ExtractedStr = Mid(TempStr, 1, EndQuoteLoc - 1)
                            TempStr = Mid(TempStr, EndQuoteLoc + 1)
                                            
                            If Len(ExtractedStr) > 0 Then
                                  StringIDX = StringIDX + 1
                                    Call AddConversionItem(Path, StringIDX, Line, TS.Line, ExtractedStr)
                            End If
                                        
                        End If
                        
                    End If
            Loop
        
        End If
        
    Loop
    TS.Close
    
End Sub


Public Sub AddConversionItem(ByVal Filename As String, ByVal StringIDX As Integer, ByVal LineContent As String, ByVal LineNo As Double, ByVal Str As String)

    Dim SQL As String
    Dim Fields As String
    Dim Values As String
    
    Dim TempStr As String
    
    SQL = "insert into conversionlist "
    
    Fields = "("
    Values = " values ("
    
    Fields = Fields & "Filename,"
    Values = Values & "'" & Filename & "',"
    
    Fields = Fields & "StringIDX,"
    Values = Values & StringIDX & ","
    
    TempStr = Replace(LineContent, "'", "''")
    TempStr = Mid(TempStr, 1, 255)
    Fields = Fields & "FileLineContent,"
    Values = Values & "'" & TempStr & "',"
    
    Fields = Fields & "Line,"
    Values = Values & "'" & LineNo & "',"
    
    TempStr = Replace(Str, "'", "''")
    TempStr = Mid(TempStr, 1, 255)
    Fields = Fields & "Str,"
    Values = Values & "'" & TempStr & "',"
    
    Fields = Mid(Fields, 1, Len(Fields) - 1)
    Values = Mid(Values, 1, Len(Values) - 1)
    
    Fields = Fields & ")"
    Values = Values & ")"
    
    SQL = SQL & Fields & Values
    
    Cn.Execute SQL
    
End Sub

Public Function ReplaceStringToVariable(ByVal Path As String) As String

    Dim Line As String
    Dim TempStr As String
    Dim StartQuoteLoc As Integer
    Dim EndQuoteLoc As Integer
    Dim ExtractedStr As String
    Dim ReplaceStr As String
    Dim FileContent As String
    
    Dim Counter As Double
    Dim Response As String
    Dim Message As String
    Dim ObjectOpened As Integer
    
    Dim Rs As New ADODB.Recordset
    Dim LangSetRs As New ADODB.Recordset
    Dim TS As TextStream
    Dim SQL As String
    Dim TempFile As String
    Dim StringIDX As Double
    Dim VBName As String
    Dim ReplaceCnt As Integer
    
    TempFile = AppPath(FSO.GetTempName)
    FileCopy Path, TempFile
    
    Set TS = FSO.OpenTextFile(TempFile, ForReading, False)
    
    SQL = "select * from conversionlist "
    SQL = SQL & " where filename='" & Path & "'"
    'SQL = SQL & " and IsCaption = True"
    SQL = SQL & " order by StringIDX"
    
    Rs.Open SQL, Cn, 1, 3
        
    StringIDX = 0
    ObjectOpened = 0
    
    Do Until TS.AtEndOfStream
        
        DoEvents
        Line = TS.ReadLine
        TempStr = Line
        
        'Form Object
        If IsObjectBegin(TempStr) = True Then
            ObjectOpened = ObjectOpened + 1
            FileContent = FileContent & Line & vbNewLine
            Do Until ObjectOpened = 0
                    Line = TS.ReadLine
                    TempStr = Line
                    FileContent = FileContent & Line & vbNewLine
                    If IsObjectBegin(TempStr) = True Then
                        ObjectOpened = ObjectOpened + 1
                    ElseIf IsObjectEnd(TempStr) = True Then
                        ObjectOpened = ObjectOpened - 1
                    End If
            Loop
            Line = TS.ReadLine
            TempStr = Line
        End If
        
        'Get VB Name
        If IsExpression(TempStr, "Attribute VB_Name") = True Then
            VBName = Replace(GetValue(TempStr, "Attribute VB_Name"), """", "")
            SQL = "select max(setid) from languageset where vbname='" & VBName & "'"
            LangSetRs.Open SQL, Cn, 1, 1
            If LangSetRs.RecordCount > 0 Then
                ReplaceCnt = LangSetRs(0)
            End If
            LangSetRs.Close
        End If
        
        
        If Not IsExpression(TempStr, "Object") = True _
           And Not IsExpression(TempStr, "Attribute") = True Then
            
            StartQuoteLoc = -1
            ReplaceString = ""
            Do Until StartQuoteLoc = 0
                    DoEvents
                    MousePointer = vbHourglass
                    ExtractedStr = ""
                    
                    StartQuoteLoc = InStr(1, TempStr, """")
                    If StartQuoteLoc > 0 Then
                    
                        ReplaceString = ReplaceString & Mid(TempStr, 1, StartQuoteLoc - 1)
                        
                        TempStr = Mid(TempStr, StartQuoteLoc + 1)
                        
                        EndQuoteLoc = InStr(1, TempStr, """")
                        
                        If EndQuoteLoc > 0 Then
                            
                            Counter = Counter + 1
                            
                            ExtractedStr = Mid(TempStr, 1, EndQuoteLoc - 1)
                            
                                            
                            If Len(ExtractedStr) > 0 Then
                                   StringIDX = StringIDX + 1
                                   If Rs("Line") = TS.Line And _
                                      Rs("FileLineContent") = Mid(Line, 1, 255) And _
                                      Rs("StringIDX") = StringIDX And _
                                      Rs("IsCaption") = True Then
                                      ReplaceString = ReplaceString & "  LangSet(""" & VBName & """," & ReplaceCnt & ")  "
                                      
                                      Call AddLanguageSet(Path, "CashInvoice", ReplaceCnt, Rs("str"), "ENG")
                                      ReplaceCnt = ReplaceCnt + 1
                                  Else
                                      ReplaceString = ReplaceString & """" & ExtractedStr & """"
                                   End If
                                   Rs.MoveNext
                            
                            Else
                                        
                                    ReplaceString = ReplaceString & """" & ExtractedStr & """"
                            
                            End If
                            
                            TempStr = Mid(TempStr, EndQuoteLoc + 1)
                        
                        Else
                        
                            ReplaceString = ReplaceString & TempStr
                                        
                        End If
                    
                    Else
                        
                            ReplaceString = ReplaceString & TempStr
                        
                    End If
            Loop

            
            FileContent = FileContent & ReplaceString & vbNewLine
            
        Else
        
            FileContent = FileContent & Line & vbNewLine
        
            
        End If
        
    Loop
    TS.Close
    Rs.Close
    
    MousePointer = vbDefault
    
    ReplaceStringToVariable = FileContent
    
    Kill TempFile
    
End Function


Public Function ReplaceVariableToString(ByVal Path As String) As String

    Dim Line As String
    Dim TempStr As String
    Dim FindStr As String
    Dim ReplaceStr As String
    Dim FileContent As String
    Dim SetID As String
    
    Dim StartLoc As Double
    Dim EndLoc As Double
        
    Dim Rs As New ADODB.Recordset
    Dim TS As TextStream
    Dim SQL As String
    Dim TempFile As String
    Dim VBName As String
    
    TempFile = AppPath(FSO.GetTempName)
    FileCopy Path, TempFile
    
    Set TS = FSO.OpenTextFile(TempFile, ForReading, False)
        
    SQL = "select * from languageset"
    
    Rs.Open SQL, Cn, 1, 1
    
    Set TS = FSO.OpenTextFile(TempFile, ForReading, False)
    
    FileContent = ""
    
    Do Until TS.AtEndOfStream
        Line = TS.ReadLine
        TempStr = Line
        DoEvents
        MousePointer = vbHourglass
        StartLoc = -1
        ReplaceStr = ""
        Do Until StartLoc = 0
            
                SetID = -1
                FindStr = "LangSet$("
                StartLoc = InStr(1, TempStr, FindStr)
                If StartLoc > 0 Then
                    ReplaceStr = ReplaceStr & Mid(TempStr, 1, StartLoc - 1)
                    TempStr = Mid(TempStr, StartLoc + Len(FindStr))

                    EndLoc = InStr(1, TempStr, ")")
                    If EndLoc > 0 Then
                            SetID = CDbl(Mid(TempStr, 1, EndLoc - 1))
                            Rs.MoveFirst
                            Rs.Find "SetID =" & SetID
                            If Not Rs.EOF Then
                                ReplaceStr = ReplaceStr & """" & Rs("expression") & """"
                            Else
                                ReplaceStr = ReplaceStr & FindStr & CStr(SetID) & ")"
                            End If
                            TempStr = Mid(TempStr, EndLoc + 1)
                            
                    Else
                        ReplaceStr = ReplaceStr & TempStr
                    End If
                                    
                Else
                
                    ReplaceStr = ReplaceStr & TempStr
                    
                End If
                
        Loop

        FileContent = FileContent & ReplaceStr & vbNewLine
    Loop
    TS.Close
    Rs.Close
    MousePointer = vbDefault
    ReplaceVariableToString = FileContent
    
    Kill TempFile
    
End Function

