VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   Picture         =   "ui.frx":0000
   ScaleHeight     =   6360
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image3 
      Height          =   615
      Left            =   0
      MouseIcon       =   "ui.frx":6FDF
      MousePointer    =   99  'Custom
      ToolTipText     =   "Back"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   1935
      Left            =   5400
      MouseIcon       =   "ui.frx":7131
      MousePointer    =   99  'Custom
      ToolTipText     =   "Run Ram Cleaner..."
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   2640
      MouseIcon       =   "ui.frx":7283
      MousePointer    =   99  'Custom
      ToolTipText     =   "Run Junk Cleaner..."
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   9840
      MouseIcon       =   "ui.frx":73D5
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   0
      MouseIcon       =   "ui.frx":7527
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BCM_SETSHIELD As Long = &H160C&
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
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
Form2.Show
Me.Hide
End Sub

Private Sub Image2_Click()
Me.Hide
Form3.Show
End Sub

Private Sub Image3_Click()
   Unload Me
   index.Show
End Sub

Private Sub Image4_Click()
 '   ShellExecute hwnd, "runas", "StartupClean.exe", "", CurDir$(), vbNormalFocus
End Sub

Private Sub Image5_Click()
   ' ShellExecute hwnd, "runas", "FileShred.exe", "", CurDir$(), vbNormalFocus
End Sub

Private Sub Image7_Click()
Unload Me
index.Show
End Sub

Private Sub Image9_Click()
Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"
End Sub
