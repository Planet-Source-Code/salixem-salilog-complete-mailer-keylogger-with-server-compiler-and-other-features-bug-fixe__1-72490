Attribute VB_Name = "SysInf"
'                    SaLiLoG 1.0
'                  By SaLar Zeynali
'        Salixem@Gmail.Com - S4LiX3M@Yahoo.Com
'  _________      .____    .______  ___        _____
' /   _____/____  |    |   |__\   \/  /____   /     \
' \_____  \\__  \ |    |   |  |\     // __ \ /  \ /  \
' /        \/ __ \|    |___|  |/     \  ___//    Y    \
'/_______  (____  /_______ \__/___/\  \___  >____|__  /
'        \/     \/        \/        \_/   \/        \/

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg key root types
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
Const ERROR_SUCCESS = 0
Const REG_SZ = 1 ' Unicode null terminated string
Const REG_DWORD = 4 ' 32-bit number
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
'Getting memory
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Type MEMORYSTATUS
dwLength As Long
dwMemoryLoad As Long
dwTotalPhys As Long
dwAvailPhys As Long
dwTotalPageFile As Long
dwAvailPageFile As Long
dwTotalVirtual As Long
dwAvailVirtual As Long
End Type

Const RAS95_MaxEntryName = 256
Const RAS95_MaxDeviceType = 16
Const RAS95_MaxDeviceName = 32

Private Type RASCONN95
dwSize As Long
hRasCon As Long
szEntryName(RAS95_MaxEntryName) As Byte
szDeviceType(RAS95_MaxDeviceType) As Byte
szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Private Type RASCONNSTATUS95
dwSize As Long
RasConnState As Long
dwError As Long
szDeviceType(RAS95_MaxDeviceType) As Byte
szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Private Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO ' 148 bytes
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Private Function LPSTRToVBString$(ByVal s$)
Dim nullpos&
nullpos& = InStr(s$, Chr$(0))
If nullpos > 0 Then
LPSTRToVBString = Left$(s$, nullpos - 1)
Else
LPSTRToVBString = ""
End If
End Function

Public Function GetOperate(OsName As String, OsVer As String, srvPAck As String)
Dim dl&, s$
Dim myVer As OSVERSIONINFO
myVer.dwOSVersionInfoSize = 148
dl& = GetVersionEx&(myVer)
srvPAck = LPSTRToVBString(myVer.szCSDVersion)
OsName = myVer.dwPlatformId
Select Case OsName
Case 0
OsName = "Unidentified"
Case 1
OsName = "Windows 95/98/ME"
Case 2
OsName = "Windows NT/2000/XP"
End Select
OsVer = myVer.dwMajorVersion & "." & myVer.dwMinorVersion & " Build " & (myVer.dwBuildNumber And &HFFFF&)
End Function

Public Function FSF32() As String
Dim strFolder As String * 255
Dim intLenght As Integer
'Get system directory
intLenght = GetSystemDirectory(strFolder, 255)
FSF32 = Left(strFolder, intLenght)
End Function

Public Function CheckInternetConnection() As Boolean
On Error GoTo ErrHandler
Dim TRasCon(255) As RASCONN95
Dim lg As Long
Dim lpcon As Long
Dim RetVal As Long
Dim Tstatus As RASCONNSTATUS95
TRasCon(0).dwSize = 412
lg = 256 * TRasCon(0).dwSize
RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)
If RetVal <> 0 Then Exit Function
Tstatus.dwSize = 160
RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)
If Tstatus.RasConnState = &H2000 Then
CheckInternetConnection = True
Else
CheckInternetConnection = False
End If
ExitCheckConnection:
Exit Function
ErrHandler:
CheckInternetConnection = False
Resume ExitCheckConnection
End Function

Public Function RGGetKeyValue(hKey As Long, SubKey As String, ValueName As String, Optional Default As String = "")
Dim lngRet As Long
Dim lngResult As Long
Dim sData As String
lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
If lngRet = ERROR_SUCCESS Then
sData = String(128, vbNullChar)
lngRet = RegQueryValueEx(lngResult, ValueName, 0, REG_SZ, ByVal sData, Len(sData))
If Not lngRet = ERROR_SUCCESS Then RGGetKeyValue = Default: Exit Function
RGGetKeyValue = Left(sData, InStr(1, sData, vbNullChar) - 1)
RegCloseKey lngResult
Else
RGGetKeyValue = Default
End If
End Function


