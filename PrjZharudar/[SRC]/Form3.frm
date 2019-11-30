VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form autorunRemover 
   BorderStyle     =   0  'None
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":BEF6
   ScaleHeight     =   6360
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   600
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      DragMode        =   1  'Automatic
      Height          =   135
      Left            =   120
      TabIndex        =   0
      Top             =   5640
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   9840
      MouseIcon       =   "Form3.frx":1272BA
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      MouseIcon       =   "Form3.frx":12740C
      MousePointer    =   99  'Custom
      ToolTipText     =   "Back"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   4320
      MouseIcon       =   "Form3.frx":12755E
      MousePointer    =   99  'Custom
      ToolTipText     =   "Click here to remove autorun.inf..."
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      MouseIcon       =   "Form3.frx":1276B0
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   240
      MouseIcon       =   "Form3.frx":127802
      MousePointer    =   99  'Custom
      ToolTipText     =   "Zharudar 1.8 | BornnoLab"
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "autorunRemover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
    Timer1.Enabled = True
    ProgressBar1.Value = ProgressBar1.Min
On Error Resume Next
    MkDir "A:\autorun.inf"
    MkDir "B:\autorun.inf"
    MkDir "C:\autorun.inf"
    MkDir "D:\autorun.inf"
    MkDir "E:\autorun.inf"
    MkDir "F:\autorun.inf"
    MkDir "G:\autorun.inf"
    MkDir "H:\autorun.inf"
    MkDir "I:\autorun.inf"
    MkDir "J:\autorun.inf"
    MkDir "K:\autorun.inf"
    MkDir "L:\autorun.inf"
    MkDir "M:\autorun.inf"
    MkDir "N:\autorun.inf"
    MkDir "O:\autorun.inf"
    MkDir "P:\autorun.inf"
    MkDir "Q:\autorun.inf"
    MkDir "R:\autorun.inf"
    MkDir "S:\autorun.inf"
    MkDir "T:\autorun.inf"
    MkDir "U:\autorun.inf"
    MkDir "V:\autorun.inf"
    MkDir "W:\autorun.inf"
    MkDir "X:\autorun.inf"
    MkDir "Y:\autorun.inf"
    MkDir "Z:\autorun.inf"
End Sub

Private Sub Image2_Click()
    Unload Me
    index.Show
End Sub

Private Sub Image3_Click()
Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"
End Sub

Private Sub Image4_Click()
    Unload Me
    index.Show
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
