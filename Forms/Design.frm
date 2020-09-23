VERSION 5.00
Begin VB.Form Design 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Dialog++ Form Designer 1.0 BETA 2"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   855
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Design.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   8985
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   7470
      TabIndex        =   33
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Frame frResize 
      Caption         =   "Resize Control:"
      Height          =   2535
      Left            =   5520
      TabIndex        =   26
      Top             =   30
      Width           =   3375
      Begin VB.TextBox txtNewWidth 
         Height          =   285
         Left            =   1290
         TabIndex        =   32
         Top             =   1140
         Width           =   1935
      End
      Begin VB.TextBox txtNewHeight 
         Height          =   285
         Left            =   1290
         TabIndex        =   31
         Top             =   810
         Width           =   1935
      End
      Begin VB.TextBox txtCTLName 
         Height          =   285
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   390
         Width           =   1935
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "New Width:"
         Height          =   225
         Index           =   9
         Left            =   180
         TabIndex        =   30
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "New Height:"
         Height          =   225
         Index           =   8
         Left            =   180
         TabIndex        =   29
         Top             =   870
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Control Name:"
         Height          =   225
         Index           =   7
         Left            =   180
         TabIndex        =   27
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   3690
      TabIndex        =   19
      Top             =   2070
      Width           =   315
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      ScaleHeight     =   285
      ScaleWidth      =   1485
      TabIndex        =   5
      Top             =   2670
      Width           =   1515
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Designer Area"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   1425
      End
   End
   Begin VB.PictureBox Design 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   5865
      Left            =   90
      ScaleHeight     =   5835
      ScaleWidth      =   8775
      TabIndex        =   4
      Top             =   2790
      Width           =   8805
      Begin VB.FileListBox File 
         Height          =   1455
         Index           =   0
         Left            =   1860
         TabIndex        =   25
         Top             =   1920
         Width           =   1905
      End
      Begin VB.DirListBox Dir 
         Height          =   1440
         Index           =   0
         Left            =   1860
         TabIndex        =   24
         Top             =   390
         Width           =   1905
      End
      Begin VB.ListBox Listbox 
         Height          =   645
         Index           =   0
         Left            =   330
         TabIndex        =   23
         Top             =   2460
         Width           =   1395
      End
      Begin VB.OptionButton Options 
         Caption         =   "Optionbox #1"
         Height          =   225
         Index           =   0
         Left            =   330
         TabIndex        =   22
         Top             =   2160
         Width           =   1425
      End
      Begin VB.CheckBox Checkbox 
         Caption         =   "Checkbox #1"
         Height          =   225
         Index           =   0
         Left            =   330
         TabIndex        =   21
         Top             =   1890
         Width           =   1425
      End
      Begin VB.TextBox TextBox 
         Height          =   285
         Index           =   0
         Left            =   330
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   20
         Text            =   "TextBox #1"
         Top             =   1560
         Width           =   1395
      End
      Begin VB.PictureBox PictureBox 
         Height          =   675
         Index           =   0
         Left            =   330
         ScaleHeight     =   615
         ScaleWidth      =   1335
         TabIndex        =   16
         Top             =   810
         Width           =   1395
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Command #1"
         Height          =   375
         Index           =   0
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   390
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   525
      Left            =   4170
      TabIndex        =   0
      Top             =   390
      Width           =   1005
   End
   Begin VB.Frame frAdd 
      Caption         =   "Add Control:"
      Height          =   2535
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   5295
      Begin VB.TextBox txtPic 
         Height          =   285
         Left            =   1530
         TabIndex        =   18
         Top             =   2040
         Width           =   2025
      End
      Begin VB.ComboBox cboValue 
         Height          =   315
         ItemData        =   "Design.frx":2CFA
         Left            =   1530
         List            =   "Design.frx":2D0D
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1680
         Width           =   2385
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   1530
         TabIndex        =   12
         Text            =   "1500"
         Top             =   1350
         Width           =   2385
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   1530
         TabIndex        =   11
         Text            =   "400"
         Top             =   1020
         Width           =   2385
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1530
         TabIndex        =   8
         Text            =   "ControlCaption"
         Top             =   690
         Width           =   2385
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "Design.frx":2D7E
         Left            =   1530
         List            =   "Design.frx":2D9A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   330
         Width           =   2385
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Control Picture:"
         Height          =   225
         Index           =   6
         Left            =   240
         TabIndex        =   17
         Top             =   2070
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Control Type:"
         Height          =   225
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   1740
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Control Width:"
         Height          =   225
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Control Height:"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Control Caption:"
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Control Type:"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   390
         Width           =   1005
      End
   End
   Begin VB.Image Grid 
      Height          =   165
      Left            =   8190
      Picture         =   "Design.frx":2DEB
      Stretch         =   -1  'True
      Top             =   2610
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuvdphelp 
         Caption         =   "Visual Dialog++ &Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnusep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnueula 
         Caption         =   "Developer &EULA"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuupdinfo 
         Caption         =   "Update &Info"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About Visual Dialog++"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "Design"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Dim PicBox, CMD, TBox, Chk, Opt, List, DirBox, FileBox As Long

Private Sub cboType_Click()
Select Case cboType.ListIndex
Case 0
txtCaption.Enabled = False
txtPic.Enabled = True
cmdBrowse.Enabled = True
cboValue.Enabled = False
Case 1
txtCaption.Enabled = True
txtPic.Enabled = False
cmdBrowse.Enabled = False
cboValue.Enabled = False
Case 2
txtCaption.Enabled = True
txtPic.Enabled = False
cmdBrowse.Enabled = False
cboValue.Enabled = False
Case 3
txtCaption.Enabled = True
txtPic.Enabled = False
cmdBrowse.Enabled = False
cboValue.Enabled = True
Case 4
txtCaption.Enabled = True
txtPic.Enabled = False
cmdBrowse.Enabled = False
cboValue.Enabled = True
Case 5
txtCaption.Enabled = False
txtPic.Enabled = False
cmdBrowse.Enabled = False
cboValue.Enabled = False
Case 6
txtCaption.Enabled = False
txtPic.Enabled = False
cmdBrowse.Enabled = False
cboValue.Enabled = False
Case 7
txtCaption.Enabled = False
txtPic.Enabled = False
cmdBrowse.Enabled = False
cboValue.Enabled = False
Case Else
MsgBox "Control is not available.", vbCritical, "Error": Exit Sub
End Select
End Sub

Private Sub Checkbox_Click(Index As Integer)
'Nothing...
End Sub

Private Sub Checkbox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Checkbox(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  SendCTLProp Checkbox(Index)
End Sub

Private Sub Checkbox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Checkbox(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Else
      Checkbox(Index).ZOrder 0
  End If
End Sub

Private Sub cmdAdd_Click()
MakeCTL cboType.ListIndex, txtHeight.Text, txtWidth.Text, txtCaption.Text, cboValue.ListIndex, cboValue.ListIndex, txtPic.Text
End Sub

Private Sub cmdApply_Click()
'This doesn't work, so anybody help?
'See the ResizeCTL function for more...
'ResizeCTL FocusCTL
End Sub

Private Sub CommandButton_Click(Index As Integer)
'Nothing...
End Sub

Private Sub CommandButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(CommandButton(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  SendCTLProp CommandButton(Index)
End Sub

Private Sub CommandButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(CommandButton(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Else
      CommandButton(Index).ZOrder 0
  End If
End Sub

Private Sub Design_Click()
txtCTLName.Text = ""
txtCTLName.Tag = ""
txtNewHeight.Text = ""
txtNewWidth.Text = ""
cmdApply.Tag = ""
End Sub

Private Sub Dir_Click(Index As Integer)
'Nothing...
End Sub

Private Sub Dir_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Dir(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  SendCTLProp Dir(Index)
End Sub

Private Sub Dir_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Dir(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Else
      Dir(Index).ZOrder 0
  End If
End Sub

Private Sub File_Click(Index As Integer)
'Nothing...
End Sub

Private Sub File_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(File(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  SendCTLProp File(Index)
End Sub

Private Sub File_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(File(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Else
      File(Index).ZOrder 0
  End If
End Sub

Private Sub Form_Click()
'MsgBox txtCTLName.Tag
End Sub

Private Sub Form_Load()
SendMessage CommandButton(Index).hWnd, &HF4&, &H0&, 0&
cboType.ListIndex = 0
cboValue.ListIndex = 0
PicBox = 0
CMD = 0
TBox = 0
Chk = 0
Opt = 0
List = 0
DirBox = 0
FileBox = 0
Design.PaintPicture Grid.Picture, 0, 0
'Design.BackColor = Me.BackColor
Me.Top = (Screen.Height - (Me.Height + 450)) / 2
Me.Left = (Screen.Width - (Me.Width)) / 2
CTLCount = 8
End Sub

Private Sub Picture_Click(Index As Integer)
'Nothing...
End Sub

Private Sub Listbox_Click(Index As Integer)
'Nothing...
End Sub

Private Sub Listbox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Listbox(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  SendCTLProp Listbox(Index)
End Sub

Private Sub Listbox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Listbox(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Else
      Listbox(Index).ZOrder 0
  End If
End Sub

Private Sub mnuabout_Click()
About.Show vbModal, Me
End Sub

Private Sub mnuback_Click()
SetZOrder GetCTLN, False
End Sub

Private Sub mnueula_Click()
Dim E As String
E = E & "You may use my codes in your application" & vbCrLf
E = E & "but you must include me in the credit."
MsgBox E, vbInformation, "EULA"
End Sub

Private Sub mnufront_Click()
SetZOrder GetCTLN, True
End Sub

Private Sub mnuupdinfo_Click()
Dim Upd As String
Upd = Upd & "Version 1.0.0 Build (0002) - 29/03/2003" & vbCrLf
Upd = Upd & "  - Fixed the flickering background by using PaintPicture" & vbCrLf
Upd = Upd & "    method instead of put picture" & vbCrLf
Upd = Upd & "  - Use the SendMessage API to disable the blue caret" & vbCrLf
Upd = Upd & "    on CommandButton (WinXP)" & vbCrLf
Upd = Upd & "  - Added capabilities to add controls dynamically at runtime" & vbCrLf
Upd = Upd & "  - Try to make support to resize control (not working for now :-P)" & vbCrLf
Upd = Upd & "  - Right-click on control will set the selected control's ZOrder to 0" & vbCrLf
Upd = Upd & "  - Sleek & clean new GUI compared to v1 BETA 1" & vbCrLf & vbCrLf
Upd = Upd & "See the included Readme file for more info..."
MsgBox Upd, vbInformation, "Update Info"
End Sub

Private Sub mnuvdphelp_Click()
Dim Hlp As String
Hlp = Hlp & "Just drag the control created on the form" & vbCrLf
Hlp = Hlp & "around the designer area. Or, if you want" & vbCrLf
Hlp = Hlp & "you can create the control with your own" & vbCrLf
Hlp = Hlp & "properties." & vbCrLf & vbCrLf
Hlp = Hlp & "If you're pro developer, help me to  fixes" & vbCrLf
Hlp = Hlp & "some unsolved problems.  Any  help  will" & vbCrLf
Hlp = Hlp & "be appreciated."
MsgBox Hlp, vbInformation, "Help"
End Sub

Private Sub Options_Click(Index As Integer)
'Nothing...
End Sub

Private Sub Options_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Options(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  SendCTLProp Options(Index)
End Sub

Private Sub Options_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(Options(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Else
      Options(Index).ZOrder 0
  End If
End Sub

Private Sub PictureBox_Click(Index As Integer)
'Nothing...
End Sub

Private Sub PictureBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(PictureBox(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  SendCTLProp PictureBox(Index)
End Sub

Private Sub PictureBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(PictureBox(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Else
      PictureBox(Index).ZOrder 0
  End If
End Sub

Private Sub TextBox_Click(Index As Integer)
'Nothing...
End Sub

Private Sub TextBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(TextBox(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  End If
  SendCTLProp TextBox(Index)
End Sub

Private Sub TextBox_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim lngReturnValue As Long
   If Button = 1 Then
      Call ReleaseCapture
      lngReturnValue = SendMessage(TextBox(Index).hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
  Else
      TextBox(Index).ZOrder 0
  End If
End Sub

Private Function SendCTLProp(CTLName As Control)
txtCTLName.Text = CTLName.Name & "(" & CTLName.Index & ")"
txtCTLName.Tag = CTLName.Name
txtNewHeight.Text = CTLName.Height
txtNewWidth.Text = CTLName.Width
cmdApply.Tag = CTLName.Index
End Function

Private Function ResizeCTL(ByVal CTLName As Control) 'Not working...
CTLName(CTLName.Index).Height = txtNewHeight.Text
CTLName(CTLName.Index).Width = txtNewWidth.Text
End Function

Private Function MakeCTL(CTLType As Integer, CTLHeight As Long, CTLWidth As Long, Optional CTLCaption, Optional ChkValue, Optional OptValue, Optional CTLPic)
Select Case CTLType
Case 0 'PictureBox
Load PictureBox(PicBox + 1)
PictureBox(PicBox + 1).Height = CTLHeight
PictureBox(PicBox + 1).Width = CTLWidth
If CTLPic <> "" Then: PictureBox(PicBox + 1).Picture = LoadPicture(CTLPic)
PictureBox(PicBox).Top = 0
PictureBox(PicBox).Left = 2000
PictureBox(PicBox + 1).Visible = True
PicBox = PicBox + 1
CTLCount = CTLCount + 1
Case 1 'Command Button
Load CommandButton(CMD + 1)
CommandButton(CMD + 1).Height = CTLHeight
CommandButton(CMD + 1).Caption = CTLCaption
CommandButton(CMD + 1).Width = CTLWidth
CommandButton(CMD + 1).Top = 120
CommandButton(CMD + 1).Left = 2000
CommandButton(CMD + 1).Visible = True
SendMessage CommandButton(CMD + 1).hWnd, &HF4&, &H0&, 0&
CMD = CMD + 1
CTLCount = CTLCount + 1
Case 2 'Textbox
Load TextBox(TBox + 1)
TextBox(TBox + 1).Height = CTLHeight
TextBox(TBox + 1).Text = CTLCaption
TextBox(TBox + 1).Width = CTLWidth
TextBox(TBox + 1).Top = 240
TextBox(TBox + 1).Left = 2000
TextBox(TBox + 1).Visible = True
TBox = TBox + 1
Case 3 'Checkbox
Load Checkbox(Chk + 1)
Checkbox(Chk + 1).Height = CTLHeight
Checkbox(Chk + 1).Width = CTLWidth
Checkbox(Chk + 1).Caption = CTLCaption
Checkbox(Chk + 1).Value = ChkValue
Checkbox(Chk + 1).Top = 360
Checkbox(Chk + 1).Left = 2000
Checkbox(Chk + 1).Visible = True
Chk = Chk + 1
CTLCount = CTLCount + 1
Case 4 'Optionbox
Load Options(Opt + 1)
Options(Opt + 1).Height = CTLHeight
Options(Opt + 1).Width = CTLWidth
Options(Opt + 1).Caption = CTLCaption
If OptValue = 3 Then: Options(Opt + 1).Value = True
Options(Opt + 1).Value = False
Options(Opt + 1).Top = 480
Options(Opt + 1).Left = 2000
Options(Opt + 1).Visible = True
Opt = Opt + 1
CTLCount = CTLCount + 1
Case 5 'Listbox
Load Listbox(List + 1)
Listbox(List + 1).Height = CTLHeight
Listbox(List + 1).Width = CTLWidth
Listbox(List + 1).Top = 520
Listbox(List + 1).Left = 2000
Listbox(List + 1).Visible = True
List = List + 1
CTLCount = CTLCount + 1
Case 6 'Dir
Load Dir(DirBox + 1)
Dir(DirBox + 1).Height = CTLHeight
Dir(DirBox + 1).Width = CTLWidth
Dir(DirBox + 1).Top = 640
Dir(DirBox + 1).Left = 2000
Dir(DirBox + 1).Visible = True
DirBox = DirBox + 1
CTLCount = CTLCount + 1
Case 7 'File
Load File(FileBox + 1)
File(FileBox + 1).Height = CTLHeight
File(FileBox + 1).Width = CTLWidth
File(FileBox + 1).Top = 760
File(FileBox + 1).Left = 2000
File(FileBox + 1).Visible = True
FileBox = FileBox + 1
CTLCount = CTLCount + 1
Case Else
MsgBox "The control is not available.", vbCritical, "Error"
End Select
End Function
