VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form index 
   BackColor       =   &H00262626&
   BorderStyle     =   0  'None
   Caption         =   "Home | Zharudar 1.8"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10275
   Icon            =   "dddddddddddd.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "dddddddddddd.frx":BEF6
   ScaleHeight     =   6375
   ScaleMode       =   0  'User
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer msg5 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   6120
      Top             =   3360
   End
   Begin VB.Timer msg2 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   6240
      Top             =   960
   End
   Begin VB.Timer msg1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   10920
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   240
      TabIndex        =   0
      Top             =   6480
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image Image13 
      Height          =   990
      Left            =   240
      MouseIcon       =   "dddddddddddd.frx":19593
      MousePointer    =   99  'Custom
      Picture         =   "dddddddddddd.frx":196E5
      ToolTipText     =   "Automaticly shutdown your system... "
      Top             =   4500
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   9480
      MouseIcon       =   "dddddddddddd.frx":23809
      MousePointer    =   99  'Custom
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image12 
      Appearance      =   0  'Flat
      Height          =   960
      Left            =   6930
      MouseIcon       =   "dddddddddddd.frx":2395B
      MousePointer    =   99  'Custom
      Picture         =   "dddddddddddd.frx":23AAD
      ToolTipText     =   "Remove autorun.inf from your Flash Drive or Hard Drive..."
      Top             =   2115
      Visible         =   0   'False
      Width           =   3090
   End
   Begin VB.Image Image11 
      Appearance      =   0  'Flat
      Height          =   1170
      Left            =   6930
      MouseIcon       =   "dddddddddddd.frx":2D5EF
      MousePointer    =   99  'Custom
      Picture         =   "dddddddddddd.frx":2D741
      ToolTipText     =   "Encrypt or Decrypt your personal files..."
      Top             =   960
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image Image3 
      Height          =   2175
      Left            =   6960
      MouseIcon       =   "dddddddddddd.frx":395A5
      MousePointer    =   99  'Custom
      ToolTipText     =   "Privacy and Security tool..."
      Top             =   960
      Width           =   3015
   End
   Begin VB.Image Image4 
      Height          =   1170
      Left            =   240
      MouseIcon       =   "dddddddddddd.frx":396F7
      MousePointer    =   99  'Custom
      Picture         =   "dddddddddddd.frx":39849
      ToolTipText     =   "Completely uninstall your unnecessary programs more easily..."
      Top             =   3330
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.Image imgmsg2 
      Height          =   2115
      Left            =   11640
      MousePointer    =   14  'Arrow and Question
      Picture         =   "dddddddddddd.frx":456AD
      ToolTipText     =   "Maximum temporary files cleaned successfully!"
      Top             =   3480
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   360
      ToolTipText     =   "Zharudar 1.8 | BornnoLab"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   0
      MouseIcon       =   "dddddddddddd.frx":4D1C0
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image imgmsg5 
      Height          =   2130
      Left            =   13680
      MousePointer    =   14  'Arrow and Question
      Picture         =   "dddddddddddd.frx":4D312
      ToolTipText     =   "Connecting...."
      Top             =   1920
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image imgmsg1 
      Appearance      =   0  'Flat
      Height          =   2115
      Left            =   10320
      MousePointer    =   14  'Arrow and Question
      Picture         =   "dddddddddddd.frx":62D76
      ToolTipText     =   "Maximum dust files cleaned successfully!"
      Top             =   1560
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   9840
      MouseIcon       =   "dddddddddddd.frx":6AAB3
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image6 
      Height          =   2055
      Left            =   6960
      MouseIcon       =   "dddddddddddd.frx":6AC05
      MousePointer    =   99  'Custom
      ToolTipText     =   "About Me !!!"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Image Image5 
      Height          =   2055
      Left            =   3600
      MouseIcon       =   "dddddddddddd.frx":6AD57
      MousePointer    =   99  'Custom
      ToolTipText     =   "Check for update!"
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   2115
      Left            =   3600
      MouseIcon       =   "dddddddddddd.frx":6AEA9
      MousePointer    =   99  'Custom
      ToolTipText     =   "System Optimizer"
      Top             =   960
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   2115
      Left            =   240
      MouseIcon       =   "dddddddddddd.frx":6AFFB
      MousePointer    =   99  'Custom
      ToolTipText     =   "Run Zharudar"
      Top             =   960
      Width           =   3135
   End
   Begin VB.Image Image14 
      Height          =   2055
      Left            =   240
      MouseIcon       =   "dddddddddddd.frx":6B14D
      MousePointer    =   99  'Custom
      ToolTipText     =   "Utility tools..."
      Top             =   3360
      Width           =   3135
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   ' Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
 '   Const SHERB_NOCONFIRMATION As Long = &H1    '1

Dim XX As Integer
Dim YY As Integer

Private Sub Command1_Click()
    uninstaller.Show



End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    XX = x
    YY = y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
    Me.Left = Me.Left - XX + x
    Me.Top = Me.Top - YY + y
    End If
End Sub

Private Sub Image1_Click()
Form1.Show
Unload Me
End Sub

Private Sub Image11_Click()
    crypter.Show
    Unload Me
End Sub

Private Sub Image12_Click()
    autorunRemover.Show
  Unload Me
End Sub

Private Sub Image13_Click()
autoShutdown.Show
   Unload Me
End Sub

Private Sub Image14_Click()
    Image4.Visible = True
    Image13.Visible = True
End Sub

Private Sub Image2_Click()
Unload Me
frmwintools.Show

End Sub

Private Sub Image3_Click()
    Image3.Visible = False
    Image11.Visible = True
    Image12.Visible = True
End Sub

Private Sub Image4_Click()
    uninstaller.Show
    Index.Hide
End Sub

Private Sub Image5_Click()
  
    Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"
End Sub
Private Sub Image6_Click()
    about.Show
about.w.URL = App.Path & "/bg.mp3"
End Sub
Private Sub Image7_Click()
    End
End Sub

Private Sub Image8_Click()
'    If Me.WindowState <> vbMinimized Then
   ' Me.WindowState = vbMinimized
    'Else
   ' Me.WindowState = vbNormal
   ' End If
  frmSysTrayIcon.Enabled = False
Unload Me
End Sub

Private Sub Image9_Click()
    Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"
End Sub
Private Sub msg1_Timer()
    imgmsg1.Visible = True
End Sub

Private Sub msg2_Timer()
    imgmsg2.Visible = True
End Sub

Private Sub msg4_Timer()

End Sub

Private Sub msg5_Timer()
    imgmsg5.Visible = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = True
On Error Resume Next
    ProgressBar1.Value = ProgressBar1.Value + 2
    If ProgressBar1.Value = 10 Then
    ProgressBar1.Value = ProgressBar1 + 10
    If ProgressBar1.Value >= ProgressBar1.Max Then
    End If
    End If
End Sub

Private Sub Timer4_Timer()
    imgmsg5.Visible = True
End Sub

Private Sub Timer5_Timer()

End Sub

