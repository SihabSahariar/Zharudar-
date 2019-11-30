VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form about 
   BorderStyle     =   0  'None
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   FillStyle       =   2  'Horizontal Line
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":BEF6
   ScaleHeight     =   6345
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin WMPLibCtl.WindowsMediaPlayer w 
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   7200
      Width           =   735
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1296
      _cy             =   1085
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   3360
      MouseIcon       =   "Form2.frx":18007
      MousePointer    =   99  'Custom
      ToolTipText     =   "Facebook Profile"
      Top             =   5160
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   3240
      MouseIcon       =   "Form2.frx":18159
      MousePointer    =   99  'Custom
      ToolTipText     =   "Facebook Profile"
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   120
      ToolTipText     =   "Zharudar 1.8 | BornnoLab"
      Top             =   360
      Width           =   1935
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   9840
      MouseIcon       =   "Form2.frx":182AB
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      MouseIcon       =   "Form2.frx":183FD
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3960
      MouseIcon       =   "Form2.frx":1854F
      MousePointer    =   99  'Custom
      ToolTipText     =   "BornoLab on Facebook"
      Top             =   2880
      Width           =   2655
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XX As Integer
Dim YY As Integer

Private Sub Form_Load()
On Error Resume Next
w.URL = App.Path & "/bg.mp3"

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
    Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"
End Sub


Private Sub Image3_Click()
    Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"
End Sub

Private Sub Image4_Click()
    Shell "explorer.exe" & "http://facebook.com/profile.jar/"
    
End Sub

Private Sub Image5_Click()
     Shell "explorer.exe " & "http://facebook.com/sihabsahariarsizan/"

End Sub
Private Sub Image6_Click()
Unload Me
End Sub
