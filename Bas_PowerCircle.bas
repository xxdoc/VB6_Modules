Attribute VB_Name = "Bas_PowerCircle"
Private Const ModuleName As String = "Bas_PowerCircle"

Public CurrentMachine As String
Public CurrentMachineID As String

Public Function GetKey(ByVal SerialNo) As String

    Dim Key As String
    Dim KeyAsc As Integer
    Dim idx As Integer
    Dim Segment(2) As String

    Dim CharSet(3) As String

    CharSet(0) = "9ETISMVQYMW4CLAXJ45UD9HGTEAXQHVRL"
    CharSet(1) = "JZM7IB5Y9782QSYJXKB4WT6NPAU5NERE1"
    CharSet(2) = "UDX4MZYWPUI5TJFBTM1N8AD9ENGAVFWKG"
    CharSet(3) = "4VMEB5QRUWHYJP1MCI6CEV9BL4RFQITSA"
    
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
