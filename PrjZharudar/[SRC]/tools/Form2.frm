VERSION 5.00
Begin VB.Form frmwintools 
   BackColor       =   &H001F1710&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6330
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H001C1C1C&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   720
      ScaleHeight     =   4305
      ScaleWidth      =   9345
      TabIndex        =   0
      Top             =   1200
      Width           =   9375
      Begin VB.Image Image2 
         Height          =   1095
         Left            =   120
         MouseIcon       =   "Form2.frx":4811
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":4963
         ToolTipText     =   "CMD window."
         Top             =   1560
         Width           =   2865
      End
      Begin VB.Image Image3 
         Height          =   1095
         Left            =   120
         MouseIcon       =   "Form2.frx":12383
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":124D5
         ToolTipText     =   "Control panel"
         Top             =   3000
         Width           =   2865
      End
      Begin VB.Image Image4 
         Height          =   1095
         Left            =   3240
         MouseIcon       =   "Form2.frx":1FEF5
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":20047
         ToolTipText     =   "Disk Defragment for better performance."
         Top             =   240
         Width           =   2865
      End
      Begin VB.Image Image5 
         Height          =   1095
         Left            =   3240
         MouseIcon       =   "Form2.frx":2DA67
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":2DBB9
         ToolTipText     =   "Edit Registry Keys."
         Top             =   1560
         Width           =   2865
      End
      Begin VB.Image Image6 
         Height          =   1095
         Left            =   3240
         MouseIcon       =   "Form2.frx":3B5D9
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":3B72B
         ToolTipText     =   "Security Center "
         Top             =   3000
         Width           =   2865
      End
      Begin VB.Image Image7 
         Height          =   1095
         Left            =   6360
         MouseIcon       =   "Form2.frx":4914B
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":4929D
         ToolTipText     =   "Restore your operating system."
         Top             =   240
         Width           =   2865
      End
      Begin VB.Image Image8 
         Height          =   1095
         Left            =   6360
         MouseIcon       =   "Form2.frx":56CBD
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":56E0F
         ToolTipText     =   "Task Manager."
         Top             =   1560
         Width           =   2865
      End
      Begin VB.Image Image9 
         Height          =   1095
         Left            =   6360
         MouseIcon       =   "Form2.frx":6482F
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":64981
         ToolTipText     =   "Configure your system."
         Top             =   3000
         Width           =   2865
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   120
         MouseIcon       =   "Form2.frx":723A1
         MousePointer    =   99  'Custom
         Picture         =   "Form2.frx":724F3
         ToolTipText     =   "Windows Cleanup Manager"
         Top             =   240
         Width           =   2865
      End
   End
   Begin VB.Image Image11 
      Height          =   375
      Left            =   0
      MouseIcon       =   "Form2.frx":7FF13
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image Image10 
      Height          =   615
      Left            =   0
      MouseIcon       =   "Form2.frx":80065
      MousePointer    =   99  'Custom
      ToolTipText     =   "Back"
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image i2 
      Height          =   225
      Left            =   9840
      MouseIcon       =   "Form2.frx":801B7
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":80309
      ToolTipText     =   "EXIT"
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmwintools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
      
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long '1





Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub i2_Click()
Unload Me
index.Show
End Sub

Private Sub Image1_Click()
 ShellExecute (hwnd), vbNullString, "cleanmgr.exe", vbNullString, "C:\", 1
End Sub

Private Sub Image10_Click()

Unload Me
index.Show
End Sub

Private Sub Image11_Click()
    Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"

End Sub

Private Sub Image2_Click()
ShellExecute (hwnd), vbNullString, "cmd.exe", vbNullString, "C:\", 1
End Sub

Private Sub Image3_Click()
 ShellExecute (hwnd), vbNullString, "control.exe", vbNullString, "C:\", 1
End Sub

Private Sub Image4_Click()
On Error Resume Next
  ShellExecute (hwnd), vbNullString, "dfrgui.exe", vbNullString, "C:\", 1
End Sub

Private Sub Image5_Click()
 ShellExecute (hwnd), vbNullString, "regedit.exe", vbNullString, "C:\", 1
End Sub

Private Sub Image6_Click()
  ShellExecute (hwnd), vbNullString, "wscui.cpl", vbNullString, "C:\", 1
End Sub

Private Sub Image7_Click()
ShellExecute (hwnd), vbNullString, GetSystem32Path & "rstrui.exe", vbNullString, "C:\", 1
End Sub

Private Sub Image8_Click()
 ShellExecute (hwnd), vbNullString, "taskmgr.exe", vbNullString, "C:\", 1
End Sub

Private Sub Image9_Click()
  ShellExecute (hwnd), vbNullString, "msconfig.exe", vbNullString, "C:\", 1
End Sub
