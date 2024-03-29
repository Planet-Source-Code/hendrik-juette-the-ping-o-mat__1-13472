Attribute VB_Name = "modping"
Option Explicit

Private Const IP_SUCCESS As Long = 0
Private Const IP_STATUS_BASE As Long = 11000
Private Const IP_BUF_TOO_SMALL As Long = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE As Long = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE As Long = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE As Long = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE As Long = (11000 + 5)
Private Const IP_NO_RESOURCES As Long = (11000 + 6)
Private Const IP_BAD_OPTION As Long = (11000 + 7)
Private Const IP_HW_ERROR As Long = (11000 + 8)
Private Const IP_PACKET_TOO_BIG As Long = (11000 + 9)
Private Const IP_REQ_TIMED_OUT As Long = (11000 + 10)
Private Const IP_BAD_REQ As Long = (11000 + 11)
Private Const IP_BAD_ROUTE As Long = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT As Long = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM As Long = (11000 + 14)
Private Const IP_PARAM_PROBLEM As Long = (11000 + 15)
Private Const IP_SOURCE_QUENCH As Long = (11000 + 16)
Private Const IP_OPTION_TOO_BIG As Long = (11000 + 17)
Private Const IP_BAD_DESTINATION As Long = (11000 + 18)
Private Const IP_ADDR_DELETED As Long = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE As Long = (11000 + 20)
Private Const IP_MTU_CHANGE As Long = (11000 + 21)
Private Const IP_UNLOAD As Long = (11000 + 22)
Private Const IP_ADDR_ADDED As Long = (11000 + 23)
Private Const IP_GENERAL_FAILURE As Long = (11000 + 50)
Private Const MAX_IP_STATUS As Long = (11000 + 50)
Private Const IP_PENDING As Long = (11000 + 255)
Private Const PING_TIMEOUT As Long = 500
Private Const WS_VERSION_REQD As Long = &H101
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long
   'Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Private Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
   
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
    
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Private Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long
    
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Function gethostname Lib "WSOCK32.DLL" _
   (ByVal szHost As String, _
    ByVal dwHostLen As Long) As Long
    
Private Declare Function gethostbyname Lib "WSOCK32.DLL" _
   (ByVal szHost As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)
   
Private Declare Function inet_addr Lib "WSOCK32.DLL" _
   (ByVal s As String) As Long
    

Public Function GetStatusCode(status As Long) As String

   Dim msg As String
   
   Select Case status
      Case IP_SUCCESS:               msg = "ip found"
      Case INADDR_NONE:              msg = "invalid ip format"
      Case IP_BUF_TOO_SMALL:         msg = "ip error, too small?"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip target unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip target port blocked"
      Case IP_NO_RESOURCES:          msg = "ip too few resources"
      Case IP_BAD_OPTION:            msg = "ip wrong option"
      Case IP_HW_ERROR:              msg = "ip hardwareerror"
      Case IP_PACKET_TOO_BIG:        msg = "ip datapacket too big"
      Case IP_REQ_TIMED_OUT:         msg = "ip request timed out"
      Case IP_BAD_REQ:               msg = "ip unknown error"
      Case IP_BAD_ROUTE:             msg = "ip invalid route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip packets died"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip packets died"
      Case IP_PARAM_PROBLEM:         msg = "ip error on pinging"
      Case IP_SOURCE_QUENCH:         msg = "ip error - check system"
      Case IP_OPTION_TOO_BIG:        msg = "ip too big"
      Case IP_BAD_DESTINATION:       msg = "ip target unreachable"
      Case IP_ADDR_DELETED:          msg = "ip address couldn´t be found"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip error"
      Case IP_MTU_CHANGE:            msg = "ip unknown host-error"
      Case IP_UNLOAD:                msg = "ip can not be resolved"
      Case IP_ADDR_ADDED:            msg = "ip address unknown on host"
      Case IP_GENERAL_FAILURE:       msg = "ip common error"
      Case IP_PENDING:               msg = "ip wrong"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case Else:                     msg = "unknown host error"
   End Select
   
   GetStatusCode = CStr(status) & "   [ " & msg & " ]"
   
End Function



Public Function ping(sAddress As String, _
                     sDataToSend As String, _
                     ECHO As ICMP_ECHO_REPLY) As Long

   
   Dim hPort As Long
   Dim dwAddress As Long
   
   dwAddress = inet_addr(sAddress)
   
     If dwAddress <> INADDR_NONE Then
   
        hPort = IcmpCreateFile()
      
        If hPort Then
      
                 Call IcmpSendEcho(hPort, _
                           dwAddress, _
                           sDataToSend, _
                           Len(sDataToSend), _
                           0, _
                           ECHO, _
                           Len(ECHO), _
                           PING_TIMEOUT)

                 ping = ECHO.status
         Call IcmpCloseHandle(hPort)
      
      End If
      
   Else:
                 ping = INADDR_NONE
         
   End If
  
End Function
   

Public Sub SocketsCleanup()
   
   If WSACleanup() <> 0 Then
       MsgBox "Windows Socket Error on Socket Close", vbExclamation
   End If
    
End Sub


Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function

