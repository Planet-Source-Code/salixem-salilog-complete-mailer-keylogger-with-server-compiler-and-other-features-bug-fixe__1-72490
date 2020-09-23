Attribute VB_Name = "getip"
'                    SaLiLoG 1.0
'                  By SaLar Zeynali
'        Salixem@Gmail.Com - S4LiX3M@Yahoo.Com
'  _________      .____    .______  ___        _____
' /   _____/____  |    |   |__\   \/  /____   /     \
' \_____  \\__  \ |    |   |  |\     // __ \ /  \ /  \
' /        \/ __ \|    |___|  |/     \  ___//    Y    \
'/_______  (____  /_______ \__/___/\  \___  >____|__  /
'        \/     \/        \/        \_/   \/        \/

Private Const IP_SUCCESS As Long = 0
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1

Private Type WSADATA
wVersion As Integer
wHighVersion As Integer
szDescription(0 To MAX_WSADescription) As Byte
szSystemStatus(0 To MAX_WSASYSStatus) As Byte
wMaxSockets As Long
wMaxUDPDG As Long
dwVendorInfo As Long
End Type
Private Declare Function gethostbyname Lib "wsock32" _
                                        (ByVal hostname As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
                                     Alias "RtlMoveMemory" _
                                    (xDest As Any, _
                                     xSource As Any, _
                                     ByVal nbytes As Long)
Private Declare Function WSAStartup Lib "wsock32" _
                                    (ByVal wVersionRequired As Long, _
                                     lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Function SocketsInitialize() As Boolean
Dim WSAD As WSADATA
Dim success As Long
SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
End Function

Public Function GIFNHost(ByVal sHostName As String) As String
Dim nbytes As Long
Dim ptrHosent As Long
Dim ptrName As Long
Dim ptrAddress As Long
Dim ptrIPAddress As Long
Dim sAddress As String
If SocketsInitialize() Then
sAddress = Space$(4)
ptrHosent = gethostbyname(sHostName & vbNullChar)
If ptrHosent <> 0 Then
ptrAddress = ptrHosent + 12
'Get the IP address
CopyMemory ptrAddress, ByVal ptrAddress, 4
CopyMemory ptrIPAddress, ByVal ptrAddress, 4
CopyMemory ByVal sAddress, ByVal ptrIPAddress, 4
GIFNHost = IPToText(sAddress)
End If
Else
GIFNHost = ""
End If
End Function

Private Function IPToText(ByVal IPAddress As String) As String
IPToText = CStr(Asc(IPAddress)) & "." & _
CStr(Asc(Mid$(IPAddress, 2, 1))) & "." & _
CStr(Asc(Mid$(IPAddress, 3, 1))) & "." & _
CStr(Asc(Mid$(IPAddress, 4, 1)))
End Function


