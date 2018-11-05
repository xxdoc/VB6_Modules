Attribute VB_Name = "Bas_GetLocalIP"
   Private Const WS_VERSION_REQD = &H101
   Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
   Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
   Private Const MIN_SOCKETS_REQD = 1
   Private Const SOCKET_ERROR = -1
   Private Const WSADescription_Len = 256
   Private Const WSASYS_Status_Len = 128

   Private Type HOSTENT
       hName As Long
       hAliases As Long
       hAddrType As Integer
       hLength As Integer
       hAddrList As Long
   End Type

   Private Type WSADATA
       wversion As Integer
       wHighVersion As Integer
       szDescription(0 To WSADescription_Len) As Byte
       szSystemStatus(0 To WSASYS_Status_Len) As Byte
       iMaxSockets As Integer
       iMaxUdpDg As Integer
       lpszVendorInfo As Long
   End Type

   Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
   Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
   Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
   
   Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
   Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
   Private Declare Sub RtlMoveMemory Lib "KERNEL32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
    
   Public Sub SocketsInitialize()
   Dim WSAD As WSADATA
   Dim iReturn As Integer
   Dim sLowByte As String, sHighByte As String, sMsg As String

       iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

       If iReturn <> 0 Then
           MsgBox "Winsock.dll is not responding."
           End
       End If

       If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
           WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

           sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
           sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
           sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
           sMsg = sMsg & " is not supported by winsock.dll "
           MsgBox sMsg
           End
       End If

       'iMaxSockets is not used in winsock 2. So the following check is only
       'necessary for winsock 1. If winsock 2 is requested,
       'the following check can be skipped.

       If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
           sMsg = "This application requires a minimum of "
           sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
           MsgBox sMsg
           End
       End If

   End Sub
   
   Private Function hibyte(ByVal wParam As Integer)

       hibyte = wParam \ &H100 And &HFF&

   End Function

   Private Function lobyte(ByVal wParam As Integer)

       lobyte = wParam And &HFF&

   End Function

   Private Sub SocketsCleanup()
   Dim lReturn As Long

       lReturn = WSACleanup()

       If lReturn <> 0 Then
           MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
           End
       End If

   End Sub
   
   Public Function GetLocalIP(ByRef LOCAL_IP() As String, ByRef MaxAdapter As Integer)
   
   Dim hostname As String * 256
   Dim hostent_addr As Long
   Dim host As HOSTENT
   Dim hostip_addr As Long
   Dim temp_ip_address() As Byte
   Dim i As Integer
   Dim ip_address As String
   Dim Message As String

   Dim Counter As Integer
   
   SocketsInitialize
   
       If gethostname(hostname, 256) = SOCKET_ERROR Then
           MsgBox "Windows Sockets error " & Str(WSAGetLastError())
           Exit Function
       Else
           hostname = Trim$(hostname)
       End If

       hostent_addr = gethostbyname(hostname)

       If hostent_addr = 0 Then
           MsgBox "Winsock.dll is not responding."
           Exit Function
       End If

       RtlMoveMemory host, hostent_addr, LenB(host)
       RtlMoveMemory hostip_addr, host.hAddrList, 4
       
       'get all of the IP address if machine is  multi-homed
       MAX_ADAPTER = 0
       Counter = 0
       Do
           ReDim temp_ip_address(1 To host.hLength)
           RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

           For i = 1 To host.hLength
               ip_address = ip_address & temp_ip_address(i) & "."
           Next
           ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
           
           Counter = Counter + 1
           LOCAL_IP(Counter) = ip_address

           ip_address = ""
           host.hAddrList = host.hAddrList + LenB(host.hAddrList)
           RtlMoveMemory hostip_addr, host.hAddrList, 4
        
        Loop While (hostip_addr <> 0)
        MaxAdapter = Counter
        
        SocketsCleanup
        
    End Function

