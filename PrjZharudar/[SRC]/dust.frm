VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10275
   LinkTopic       =   "Form2"
   Picture         =   "dust.frx":0000
   ScaleHeight     =   6345
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   830
      Left            =   4920
      MouseIcon       =   "dust.frx":11B3C4
      MousePointer    =   99  'Custom
      Picture         =   "dust.frx":11B516
      ScaleHeight     =   795
      ScaleWidth      =   705
      TabIndex        =   0
      ToolTipText     =   "Clean Junk Files..."
      Top             =   2880
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   9800
      _ExtentX        =   17277
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      MouseIcon       =   "dust.frx":121F58
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clean unused and temporary files."
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
      Left            =   3600
      TabIndex        =   1
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   9840
      MouseIcon       =   "dust.frx":1220AA
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   0
      MouseIcon       =   "dust.frx":1221FC
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   240
      ToolTipText     =   "Zharudar 1.8 | BornnoLab"
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
    Const SHERB_NOCONFIRMATION As Long = &H1    '1
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


Private Sub EmptyRec(para As Long)
    Dim nRet As Long
    nRet = SHEmptyRecycleBin(index.hwnd, vbNullString, para)
End Sub
Private Sub Picture1_Click()
    Timer1.Enabled = True
   ' msg1.Enabled = True
    ProgressBar1.Value = ProgressBar1.Min
On Error Resume Next
    Kill "C:\Intel\Logs\*.*"
    Kill "C:\WINDOWS\Temp\*.*"
    Kill "C:\WINDOWS\system32\1054\*.*"
    Kill "C:\Program Files\Uninstall Information\*.*"
    Kill "C:\WINDOWS\Offline Web Pages\*.*"
    Kill "C:\WINDOWS\Prefetch\*.*"
    Kill "C:\Windows\MEMORY.DMP"
    Kill "C:\Windows\MiniDump\*.*"
    Kill "C:\WINDOWS\*.log"
    Kill "C:\Windows\DirectX.log"
    Kill "C:\Windows\DtcInstall.log"
    Kill "C:\Windows\PFRO.log"
    Kill "C:\Windows\setupact.log"
    Kill "C:\Windows\setuperr.log"
    Kill "C:\Windows\TSSysprep.log"
    Kill "C:\Windows\Debug\sammui.log"
    Kill "C:\Windows\security\logs\*.*"
Dim ret As Variant
    Dim nn As Long
    nn = 1
    Call EmptyRec(nn)
End Sub
