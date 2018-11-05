Attribute VB_Name = "Bas_Code39"
Private Const ModuleName As String = "Bas_Code39"

Private Type Barcode39
    Character As String * 1
    barcode As String * 10
End Type

Private code39(1 To 44)  As Barcode39

Private Sub InitBarCode39()
    
  Static IsInit As Boolean
  
  If IsInit = True Then
    Exit Sub
  End If
    
  code39(1).Character = "1": code39(1).barcode = "LSSLSSSSLI"
  code39(2).Character = "2": code39(2).barcode = "SSLLSSSSLI"
  code39(3).Character = "3": code39(3).barcode = "LSLLSSSSSI"
  code39(4).Character = "4": code39(4).barcode = "SSSLLSSSLI"
  code39(5).Character = "5": code39(5).barcode = "LSSLLSSSSI"
  code39(6).Character = "6": code39(6).barcode = "SSLLLSSSSI"
  code39(7).Character = "7": code39(7).barcode = "SSSLSSLSLI"
  code39(8).Character = "8": code39(8).barcode = "LSSLSSLSSI"
  code39(9).Character = "9": code39(9).barcode = "SSLLSSLSSI"
  code39(10).Character = "0": code39(10).barcode = "SSSLLSLSSI"
  code39(11).Character = "A": code39(11).barcode = "LSSSSLSSLI"
  code39(12).Character = "B": code39(12).barcode = "SSLSSLSSLI"
  code39(13).Character = "C": code39(13).barcode = "LSLSSLSSSI"
  code39(14).Character = "D": code39(14).barcode = "SSSSLLSSLI"
  code39(15).Character = "E": code39(15).barcode = "LSSSLLSSSI"
  code39(16).Character = "F": code39(16).barcode = "SSLSLLSSSI"
  code39(17).Character = "G": code39(17).barcode = "SSSSSLLSLI"
  code39(18).Character = "H": code39(18).barcode = "LSSSSLLSSI"
  code39(19).Character = "I": code39(19).barcode = "SSLSSLLSSI"
  code39(20).Character = "J": code39(20).barcode = "SSSSLLLSSI"
  code39(21).Character = "K": code39(21).barcode = "LSSSSSSLLI"
  code39(22).Character = "L": code39(22).barcode = "SSLSSSSLLI"
  code39(23).Character = "M": code39(23).barcode = "LSLSSSSLSI"
  code39(24).Character = "N": code39(24).barcode = "SSSSLSSLLI"
  code39(25).Character = "O": code39(25).barcode = "LSSSLSSLSI"
  code39(26).Character = "P": code39(26).barcode = "SSLSLSSLSI"
  code39(27).Character = "Q": code39(27).barcode = "SSSSSSLLLI"
  code39(28).Character = "R": code39(28).barcode = "LSSSSSLLSI"
  code39(29).Character = "S": code39(29).barcode = "SSLSSSLLSI"
  code39(30).Character = "T": code39(30).barcode = "SSSSLSLLSI"
  code39(31).Character = "U": code39(31).barcode = "LLSSSSSSLI"
  code39(32).Character = "V": code39(32).barcode = "SLLSSSSSLI"
  code39(33).Character = "W": code39(33).barcode = "LLLSSSSSSI"
  code39(34).Character = "X": code39(34).barcode = "SLSSLSSSLI"
  code39(35).Character = "Y": code39(35).barcode = "LLSSLSSSSI"
  code39(36).Character = "Z": code39(36).barcode = "SLLSLSSSSI"
  code39(37).Character = "-": code39(37).barcode = "SLSSSSLSLI"
  code39(38).Character = ".": code39(38).barcode = "LLSSSSLSSI"
  code39(39).Character = " ": code39(39).barcode = "SLLSSSLSSI"
  code39(40).Character = "*": code39(40).barcode = "SLSSLSLSSI"
  code39(41).Character = "$": code39(41).barcode = "SLSLSLSSSI"
  code39(42).Character = "/": code39(42).barcode = "SLSLSSSLSI"
  code39(43).Character = "+": code39(43).barcode = "SLSSSLSLSI"
  code39(44).Character = "%": code39(44).barcode = "SSSLSLSLSI"
  
  IsInit = True
  
End Sub

Public Sub PrintCode39(ByVal Thetext As String, _
                                         ByVal Smallbar As Integer, _
                                         ByVal Theheight As Integer, _
                                         ByVal obj As Object, _
                                         Optional ByVal Alignment As AlignmentConstants)

Call InitBarCode39

'!!!no comments on declaring variables please !!!
' to print to the Obj or a picturebox use Obj.Line or Obj.line instead of Obj.Line
' or add a variable specivying the device
' you can of course remove or add variables as you like
' this is the point you can tailor this to your EXACT needs
' this is just code to illustrate the principle

'Smallbar is the width of a small bar
'Theheight is the height of the barcode

Largebar = Smallbar * 2.1  'the ratio between large and small bars
Intercharacterspace = Smallbar * 1 'the ratio between Intercharacterspacing and small bars
Colour = &HFFFFFF
            
ThetextWithStartAndStopChar = "*" & Thetext & "*" 'if start and stop chars are not needed dont use em

Select Case Alignment
        
Case vbCenter
        
        BarcodeWidth = Code39Width(Thetext, Smallbar, Theheight, obj)
        startX = (obj.ScaleWidth - BarcodeWidth) / 2
    
Case vbRightJustify
        
        BarcodeWidth = Code39Width(Thetext, Smallbar, Theheight, obj)
        startX = obj.ScaleWidth - BarcodeWidth
        
Case Else

        startX = obj.CurrentX  'where printing should start

End Select

Newposition = startX
starty = obj.CurrentY

For x = 1 To Len(ThetextWithStartAndStopChar)

    Chartosearch = Mid(ThetextWithStartAndStopChar, x, 1)
    Foundchar = ""
    For z = 1 To UBound(code39)
        If code39(z).Character = Chartosearch Then
                Foundchar = code39(z).barcode
        End If
    Next z
    
    If Len(Foundchar) Then
    
        For y = 1 To 10
            Onechar = Mid(Foundchar, y, 1)
            
            If Colour = &HFFFFFF Then       ' White ,to make this thing more generic we even draw white bars
                Colour = &H0                ' Black
            Else
                Colour = &HFFFFFF   'White
            End If
                 
            Select Case Onechar
            
            Case "L" ' a large bar
                obj.Line (Newposition, starty)-Step(Largebar, Theheight), Colour, BF
                Newposition = Newposition + Largebar
            
            Case "S" ' a small bar
                obj.Line (Newposition, starty)-Step(Smallbar, Theheight), Colour, BF
                Newposition = Newposition + Smallbar
                        
            Case "I" ' the Intercharacterspacing
                obj.Line (Newposition, starty)-Step(Intercharacterspace, Theheight), Colour, BF
                Newposition = Newposition + Intercharacterspace
                        
            End Select
    
        Next y
            
    Else
        
        obj.Print "?";
        
    End If
   
Next x
   
End Sub

Public Function Code39Width(ByVal Thetext As String, _
                                         ByVal Smallbar As Integer, _
                                         ByVal Theheight As Integer, _
                                         ByVal obj As Object)
    
    Dim BarcodeWidth As Double
    
    Call InitBarCode39
    
    Largebar = Smallbar * 2.1  'the ratio between large and small bars
    Intercharacterspace = Smallbar * 1 'the ratio between Intercharacterspacing and small bars

    ThetextWithStartAndStopChar = "*" & Thetext & "*" 'if start and stop chars are not needed dont use em

    For x = 1 To Len(ThetextWithStartAndStopChar)
        
            Chartosearch = Mid(ThetextWithStartAndStopChar, x, 1)
            Foundchar = ""
            For z = 1 To UBound(code39)
                If code39(z).Character = Chartosearch Then
                    Foundchar = code39(z).barcode
                End If
            Next z
            
            If Len(Foundchar) > 0 Then
                
                For y = 1 To 10
                
                    Onechar = Mid(Foundchar, y, 1)
                    
                    Select Case Onechar
                        
                        Case "L" ' a large bar
                            BarcodeWidth = BarcodeWidth + Largebar
                        
                        Case "S" ' a small bar
                            BarcodeWidth = BarcodeWidth + Smallbar
                                    
                        Case "I" ' the Intercharacterspacing
                            BarcodeWidth = BarcodeWidth + Intercharacterspace
                            
                    End Select
                    
                Next y
                
            Else
            
                    Exit Function
                    
            End If
                    
        Next x
        
        Code39Width = BarcodeWidth
        
End Function
