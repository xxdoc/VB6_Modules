Attribute VB_Name = "Bas_LabelGadget"
Private Const ModuleName As String = "Bas_LabelGadget"

Public Type LabelSetting
    
    Height As Double
    Width  As Double
    Gap As Double
    RowsPerPage As Double
    ColumnsPerPage As Double
    
End Type
Public CurrentMachine As String
Public CurrentMachineID As String

Public Cn As New ADODB.Connection
Public POSCN As New ADODB.Connection
Public PictureCN As New ADODB.Connection

Public Function GetLabelSetting() As LabelSetting
    
    Dim MyLabelSetting As LabelSetting
    
    On Error Resume Next
    
    With MyLabelSetting
        
        .Height = GetSetting(App.Title, App.Title & "\LabelSetting", "Height")
        .Width = GetSetting(App.Title, App.Title & "\LabelSetting", "Width")
        .Gap = GetSetting(App.Title, App.Title & "\LabelSetting", "Gap")
        .RowsPerPage = GetSetting(App.Title, App.Title & "\LabelSetting", "Row")
        .ColumnsPerPage = GetSetting(App.Title, App.Title & "\LabelSetting", "Col")
    
    End With
    
    GetLabelSetting = MyLabelSetting
    
End Function

Public Function SaveLabelSetting(ByRef MyLabelSetting As LabelSetting) As Boolean
    
    On Error GoTo ErrorHandler
    
    With MyLabelSetting
        
        SaveSetting App.Title, App.Title & "\LabelSetting", "Height", .Height
        SaveSetting App.Title, App.Title & "\LabelSetting", "Width", .Width
        SaveSetting App.Title, App.Title & "\LabelSetting", "Gap", .Gap
        SaveSetting App.Title, App.Title & "\LabelSetting", "Row", .RowsPerPage
        SaveSetting App.Title, App.Title & "\LabelSetting", "Col", .ColumnsPerPage
    
    End With
    
    SaveLabelSetting = True
    
    Exit Function
    
ErrorHandler:
    
    SaveLabelSetting = False
    
End Function

Public Function GetKey(ByVal SerialNo) As String

    Dim Key As String
    Dim KeyAsc As Integer
    Dim idx As Integer
    Dim Segment(2) As String

    Dim CharSet(3) As String

    CharSet(0) = "B6MGNJ3ZY4YFRGS9ASJLDUI2RE3QLMQIF"
    CharSet(1) = "WNQRTNJWH8DY3RK1ZVPFIEVKMAPGXIMTE"
    CharSet(2) = "KDUI76VYM1HXJKUJSQ45WPEF8MV2S3DBF"
    CharSet(3) = "SPBLZWVMCGKFISJPT7X9DN1I3BQ48MV8C"
    
    Segment(1) = Mid(SerialNo, 1, 5)
    Segment(2) = Mid(SerialNo, 6)
    
    Segment(1) = StrReverse(Segment(1))
    Segment(2) = Segment(2)
    
    For idx = 1 To Len(Segment(1))
        KeyAsc = Asc(Mid(Segment(1), idx, 1)) Mod 32
        Key = Key & Mid(CharSet(idx Mod 4), KeyAsc + 1, 1)
    Next
    
    For idx = 1 To Len(Segment(2))
        KeyAsc = Asc(Mid(Segment(2), idx, 1)) Mod 32
        Key = Key & Mid(CharSet(idx Mod 4), KeyAsc + 1, 1)
    Next
    
    GetKey = Key
    
End Function

Public Sub ConnectPOSDatabase()
    
    Dim connstr As String
    Dim DataBasePath As String
    Dim PictureDatabasePath As String
    
    
    Dim Message As String
    
    On Error GoTo ErrorHandler
    
    If Len(GetSetting(App.Title, App.Title, "Database")) > 0 Then
        DataBasePath = GetSetting(App.Title, App.Title, "Database")
    End If
        
    If Len(DataBasePath) = 0 Then
        DataBasePath = AppPath("pos.mdb")
    End If
    
    If FSO.FileExists(DataBasePath) = False Then
        
        Message = "Database does not exists"
        MsgBox Message, vbExclamation
        
        Exit Sub
        
    End If
    
    connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBasePath & ";Jet OLEDB:Database Password=048520300032"

    If Not POSCN.State = 0 Then
        
        If POSCN.ConnectionString <> connstr Then
            POSCN.Close
            POSCN.Open connstr
        End If
    
    Else
    
        POSCN.Open connstr
    
    End If
        
    PictureDatabasePath = FSO.GetParentFolderName(DataBasePath) & "\Picture.mdb"
    
    connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PictureDatabasePath '& ";Jet OLEDB:Database Password=048520300032"
    
    If Not PictureCN.State = 0 Then
        
          If PictureCN.ConnectionString <> connstr Then
            PictureCN.Close
            PictureCN.Open connstr
        End If
    
    Else
    
        PictureCN.Open connstr
    
    End If
    
    Exit Sub
    
ErrorHandler:
    Message = Err.Description
    Message = Replace(Message & " ", "048520300032", "??????????")
    MsgBox Message, vbCritical

End Sub

