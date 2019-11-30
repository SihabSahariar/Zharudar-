VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10275
   LinkTopic       =   "Form3"
   Picture         =   "ram.frx":0000
   ScaleHeight     =   6360
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   1320
      Top             =   3960
   End
   Begin Zharudar.ctlTrickKnob t 
      Height          =   2655
      Left            =   4080
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   4683
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   2160
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   830
      Left            =   5040
      MouseIcon       =   "ram.frx":11B3C4
      MousePointer    =   99  'Custom
      Picture         =   "ram.frx":11B516
      ScaleHeight     =   795
      ScaleWidth      =   705
      TabIndex        =   1
      ToolTipText     =   "Clean Ram..."
      Top             =   4680
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   240
      TabIndex        =   0
      Top             =   5640
      Width           =   9800
      _ExtentX        =   17277
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image i 
      Height          =   2475
      Left            =   3720
      Picture         =   "ram.frx":11DDEC
      Top             =   1320
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Clean Ram and Optimize PC Speed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   240
      ToolTipText     =   "Zharudar 1.8 | BornnoLab"
      Top             =   360
      Width           =   1935
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   0
      MouseIcon       =   "ram.frx":138792
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   9840
      MouseIcon       =   "ram.frx":1388E4
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      MouseIcon       =   "ram.frx":138A36
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   2040
      Width           =   495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim XX As Integer
Dim YY As Integer
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
Unload Me
Form1.Show
End Sub

Private Sub Image7_Click()
Unload Me
index.Show
End Sub

Private Sub Image9_Click()
Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"
End Sub

Private Sub Picture1_Click()
Timer1.Enabled = True
Timer2.Enabled = True
t.Visible = True
End Sub

Private Sub Timer1_Timer()
 Timer1.Enabled = True
On Error Resume Next
    ProgressBar1.Value = ProgressBar1.Value + 2

    If ProgressBar1.Value = ProgressBar1.Max Then
       Timer2.Enabled = False
   t.Visible = False
   i.Visible = True
    End If
End Sub

Private Sub Timer2_Timer()
t.Value = t.Value + 2
If t.Value = 5 Then
Timer1.Enabled = False
End If
End Sub
