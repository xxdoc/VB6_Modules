Attribute VB_Name = "Bas_CDO"
Private Const ModuleName As String = "Bas_CDO"

Public Sub CreateCDOCharSetDropDown(ByRef CBO As ComboBox)
    
    'CDO for windows 2000
    'charset property
    
    With CBO
        .AddItem "big5"
        .AddItem "euc-jp"
        .AddItem "euc-jp"
        .AddItem "euc-kr"
        .AddItem "gb2312"
        .AddItem "iso-2022-jp"
        .AddItem "iso-2022-kr"
        .AddItem "iso-8859-1"
        .AddItem "iso-8859-2"
        .AddItem "iso-8859-3"
        .AddItem "iso-8859-4"
        .AddItem "iso-8859-5"
        .AddItem "iso-8859-6"
        .AddItem "iso-8859-7"
        .AddItem "iso-8859-8"
        .AddItem "iso-8859-9"
        .AddItem "koi8-r"
        .AddItem "shift-jis"
        .AddItem "us-ascii"
        .AddItem "utf-7"
        .AddItem "utf-8"

    End With
    


End Sub
