Attribute VB_Name = "Bas_PrinterObject"
Private Const ModuleName As String = "Bas_PrinterObject"
Public Sub PrintAlignedText(s As String, Alignment As AlignmentConstants)
    
    Select Case Alignment
    Case vbCenter
        Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(s)) \ 2
    Case vbLeftJustify
        Printer.CurrentX = 0
    Case vbRightJustify
        Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(s)
    End Select
    Printer.Print s
    
End Sub

