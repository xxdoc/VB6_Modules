Attribute VB_Name = "Bas_CDROM"
'Declare the following in the declarations section of your code
Private Declare Function GetVolumeInformation Lib "kernel32.dll" _
 Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
 ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, _
 lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
 lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
 ByVal nFileSystemNameSize As Long) As Long


' **********************************************************************
   '
   ' FUNCTION:
   '    GetFirstCdRomDriveLetter()
   '
   ' PURPOSE:
   '    Finds the first CD-ROM device and then returns its drive letter.
   '
   ' ARGUMENTS:
   '    None
   '
   ' RETURNS:
   '    A string that represents the first CD-ROM drive letter. If the
   '    function fails for any reason, it returns vbNullString.
   '
   ' **********************************************************************
   Declare Function GetDriveType Lib "kernel32" Alias _
      "GetDriveTypeA" (ByVal nDrive As String) As Long

   Declare Function GetLogicalDriveStrings Lib "kernel32" Alias _
      "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
      ByVal lpBuffer As String) As Long

   Public Const DRIVE_CDROM As Long = 5

'GetSerialNumber Procedure - Put this in the module or form where it is called.
Function GetSerialNumber(strDrive As String) As Long
    Dim SerialNum As Long
    Dim Res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    'initialise the strings
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    'call the API function
    Res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
End Function



Public Function GetFirstCdRomDriveLetter() As String

      ' Declare variables.
      Dim lDriveType As Long
      Dim strDrive As String
      Dim lStart As Long: lStart = 1

      ' Create a string to hold the logical drives.
      Dim strDrives As String
      strDrives = Space(150)

      ' Get the logial drives on the system.
      ' If the function fails it returns zero.
      Dim lRetVal As Long
      lRetVal = GetLogicalDriveStrings(150, strDrives)

      ' Check to see if GetLogicalDriveStrings() worked.
      If lRetVal = 0 Then

         ' Get GetLogicalDriveStrings() failed.
         GetFirstCdRomDriveLetter = vbNullString
         Exit Function
      End If

      ' Get the string that represents the first drive.
      strDrive = Mid(strDrives, lStart, 3)

      Do

         ' Test the first drive.
         lDriveType = GetDriveType(strDrive)

         ' Check if the drive type is a CD-ROM.
         If lDriveType = DRIVE_CDROM Then

            ' Found the first CD-ROM drive on the system.
            GetFirstCdRomDriveLetter = strDrive
            Exit Function
         End If

         ' Increment lStart to next drive in the string.
         lStart = lStart + 4

         ' Get the string that represents the first drive.
         strDrive = Mid(strDrives, lStart, 3)

      Loop While (Mid(strDrives, lStart, 1) <> vbNullChar)

   End Function

