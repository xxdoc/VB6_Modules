Attribute VB_Name = "Bas_KeyCode"
Public Const ModuleName As String = "Bas_CodeReady"

Public RegisterName As String
Public LicenseKey As String

Public IsLicensed As Boolean

Public Sub SaveLicense(ByVal RegisterNameString As String, _
                                    ByVal UserCodeKey As String)
                                    
                                        
    Dim Ts As TextStream
    
    Set Ts = Fso.OpenTextFile(AppPath("license.dat"), ForWriting, True)
    
    Ts.WriteLine "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    
    Ts.WriteLine "Date: " & Format(Now(), "YYYY/MM/DD")
    
    Ts.WriteLine "Register Name: " & RegisterNameString
    
    Ts.WriteLine "License Key: " & UserCodeKey
    
    Ts.Close
    
End Sub

Public Function ValidateLicense(ByVal RegisterNameString As String, _
                                            ByVal UserCodeKey As String) As Boolean
                                            
    Dim Code(0 To 16) As Integer
    Dim RegisterName(0 To 16) As Integer
        
    Dim Pointer(0 To 7) As Boolean
    Dim Pointers As String
        
    Dim Idx As Integer
    
    Dim CodeString As String
    Dim KeyString As String
    Dim KeyLength As String
    
    Dim UserKey As String
    
    If Len(UserCodeKey) < 7 Then
        
        Exit Function
        
    End If
    
    CodeString = Mid(UserCodeKey, 1, 7)
    
    For Idx = 0 To 6
        
        Code(Idx) = Asc(Mid(CodeString, Idx + 1, 1))
    
    Next
    
    For Idx = 0 To 8
    
        RegisterName(Idx) = 32
        
        If Len(RegisterNameString) > Idx Then
        
            Select Case Asc(Mid(UCase(RegisterNameString), Idx + 1, 1))
        
            Case 48 To 57
            
                RegisterName(Idx) = Asc(Mid(UCase(RegisterNameString), Idx + 1, 1))
            
            Case 65 To 90
            
                RegisterName(Idx) = Asc(Mid(UCase(RegisterNameString), Idx + 1, 1))
            
            Case Else
            
                RegisterName(Idx) = 32
            
            End Select
        
        End If
    
    Next
    
    Pointers = DecimalToBinary(CLng(RegisterName(3)), 7)
    
    For Idx = 0 To 6
        
        Select Case CInt(Mid(Pointers, Idx + 1, 1))
        
        Case 0
        
            Pointer(Idx) = True
            
        Case 1
        
            Pointer(Idx) = False
            
        End Select
        
    Next
    
    
    For Idx = 0 To 6
        
        If Pointer(Idx) = True Then
            
            Select Case Code(Idx) + Idx
            
            Case 0 To 47
                    
                    Key = Key & Chr(49)
            
            Case 58 To 64
            
                    Key = Key & Chr(Code(Idx) + Idx + 7)
            
            Case Is > 90
            
                    Key = Key & Chr(Code(Idx) + Idx - 90 + 65)
                    
            Case Else
            
                    Key = Key & Chr(Code(Idx) + Idx)
                    
            End Select
            
        End If
    
    Next
    
    KeyLength = Asc(Right(UserCodeKey, 1))
    
    KeyLength = CInt(Right(CStr(KeyLength), 1))
    
    UserKey = Right(UserCodeKey, KeyLength + 1)
    
    UserKey = Left(UserKey, KeyLength)
    
    If Key = UserKey Then
    
        ValidateLicense = True
    
    End If
    
End Function

Public Sub ReadLicenseFile(ByRef RegisterName As String, ByRef LicenseKey As String)
    
    Dim Ts As TextStream
    Dim ReadLine As String
    
    If Fso.FileExists(AppPath("license.dat")) = False Then
        
        Exit Sub
        
    End If
    
    Set Ts = Fso.OpenTextFile(AppPath("license.dat"), ForReading, False)
    
    Do Until Ts.AtEndOfStream
        
        ReadLine = Ts.ReadLine
        
        Select Case True
        
        
        Case InStr(1, " " & ReadLine, "Register Name:") > 0
            
            RegisterName = Mid(ReadLine, Len("Register Name: ") + 1)
        
        Case InStr(1, " " & ReadLine, "License Key: ") > 0
            
            
            LicenseKey = Mid(ReadLine, Len("License Key: ") + 1)
            
        End Select
         
        
    Loop
    
    
    
    
    
End Sub
