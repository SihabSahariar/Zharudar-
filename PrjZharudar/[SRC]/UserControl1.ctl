VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UserControl1 
   BackColor       =   &H80000007&
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   Picture         =   "UserControl1.ctx":0000
   ScaleHeight     =   6345
   ScaleWidth      =   10260
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   240
      TabIndex        =   0
      Top             =   5640
      Width           =   9800
      _ExtentX        =   17277
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
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
      MouseIcon       =   "UserControl1.ctx":11B3C4
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   9840
      MouseIcon       =   "UserControl1.ctx":11B516
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      MouseIcon       =   "UserControl1.ctx":11B668
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   2040
      Width           =   495
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim XX As Integer
Dim YY As Integer
  Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
    Const SHERB_NOCONFIRMATION As Long = &H1    '1
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
