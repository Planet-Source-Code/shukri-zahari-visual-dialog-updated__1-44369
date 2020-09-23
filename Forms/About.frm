VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Visual Dialog++"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "About.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6360
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4950
      TabIndex        =   0
      Top             =   3810
      Width           =   1305
   End
   Begin VB.Image MyPic 
      Height          =   2100
      Left            =   150
      Picture         =   "About.frx":2CFA
      ToolTipText     =   " Hi, my name is, (what?), my name is (huh?), my name is, <SCRATCH> Shukri Zahari!!!! "
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "The next generation of Form Designer (TM)"
      Height          =   195
      Index           =   7
      Left            =   2010
      TabIndex        =   8
      Top             =   750
      Width           =   3495
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "You may use part or whole of this program in your own program but you must include me in your credit."
      Height          =   495
      Index           =   6
      Left            =   2040
      TabIndex        =   7
      Top             =   2220
      Width           =   3855
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "All right reserved"
      Height          =   225
      Index           =   5
      Left            =   2040
      TabIndex        =   6
      Top             =   1710
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003, Shukri Zahari"
      Height          =   225
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   1500
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0.0 Build(0002)"
      Height          =   225
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      Top             =   1290
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Win32 Visual Dialog++ for Microsoft Windows"
      Height          =   225
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Dialog++"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   585
      Index           =   0
      Left            =   1860
      TabIndex        =   1
      Top             =   240
      Width           =   3465
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   150
      Picture         =   "About.frx":5138
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Dialog++"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   585
      Index           =   1
      Left            =   1470
      TabIndex        =   2
      Top             =   300
      Width           =   4305
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Form_Load()
Beep
End Sub
