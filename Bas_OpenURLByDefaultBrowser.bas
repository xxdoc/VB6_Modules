Attribute VB_Name = "Bas_OpenURLByDefaultBrowser"
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
         "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
         String, ByVal lpFile As String, ByVal lpParameters As String, _
         ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function FindExecutable Lib "shell32.dll" Alias _
         "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
         String, ByVal lpResult As String) As Long
         
Public Sub OpenBrowser(URL As String)
      
      'This function open an URL by the default browser of the machine.
      'Aware, there will be an void browser appeared while passing a mailto: url.
      
      Dim FileName As String
      Dim Dummy As String
      Dim BrowserExec As String * 255
      Dim RetVal As Long
      Dim FileNumber As Integer
      
      Const SW_SHOW = 5       ' Displays Window in its current size
                                      ' and position
      Const SW_SHOWNORMAL = 1 ' Restores Window if Minimized or
                                      ' Maximized
       
      ' First, create a known, temporary HTML file
      BrowserExec = Space(255)
      FileName = "C:\temphtm.HTM"
      FileNumber = FreeFile                    ' Get unused file number
      Open FileName For Output As #FileNumber  ' Create temp HTML file
          Write #FileNumber, "<HTML> <\HTML>"  ' Output text
      Close #FileNumber                        ' Close file
      ' Then find the application associated with it
      RetVal = FindExecutable(FileName, Dummy, BrowserExec)
      BrowserExec = Trim(BrowserExec)
      ' If an application is found, launch it!
      If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
          MsgBox "Could not find associated Browser", vbExclamation, _
            "Browser Not Found"
      Else
          RetVal = ShellExecute(hwnd, "open", BrowserExec, _
            URL, Dummy, SW_SHOWNORMAL)
          If RetVal <= 32 Then        ' Error
              MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
          End If
      End If
      Kill FileName
      'delete temp HTML file
End Sub

