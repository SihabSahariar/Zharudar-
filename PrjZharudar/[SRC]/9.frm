VERSION 5.00
Begin VB.Form uninstallerInfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   10260
   ControlBox      =   0   'False
   Icon            =   "9.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "9.frx":BEF6
   ScaleHeight     =   6360
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox XPFrame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   720
      ScaleHeight     =   4545
      ScaleWidth      =   9345
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   9375
      Begin VB.TextBox TxtRName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   7575
      End
      Begin VB.TextBox TxtContact 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3960
         Width           =   7575
      End
      Begin VB.TextBox TxtUIAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3600
         Width           =   7575
      End
      Begin VB.TextBox TxtHLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3240
         Width           =   7575
      End
      Begin VB.TextBox TxtDVersion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2400
         Width           =   7575
      End
      Begin VB.TextBox TxtPublisher 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2040
         Width           =   7575
      End
      Begin VB.TextBox TxtUString 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1680
         Width           =   7575
      End
      Begin VB.TextBox TxtDName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   7575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Registry Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "URL Info About :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Help Link :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Display Version :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblprogname 
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Uninstall String :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Display Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   0
      MouseIcon       =   "9.frx":1272BA
      MousePointer    =   99  'Custom
      ToolTipText     =   "Back"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   9840
      MouseIcon       =   "9.frx":12740C
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   0
      MouseIcon       =   "9.frx":12755E
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      MouseIcon       =   "9.frx":1276B0
      MousePointer    =   99  'Custom
      ToolTipText     =   "Zharudar 1.8 | BornnoLab"
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "uninstallerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim XX As Integer
Dim YY As Integer
Private Declare Function ShellExecute Lib _
   "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
    
Private Const SW_SHOWNORMAL = 1
Private Sub Form_Load()
uninstaller.GetInformasi
End Sub

Private Sub Image2_Click()
    End
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    XX = X
    YY = Y
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
    Me.Left = Me.Left - XX + X
    Me.Top = Me.Top - YY + Y
    End If
End Sub

Private Sub Image3_Click()
    Unload Me
    Index.Show
End Sub

Private Sub Image4_Click()
Shell "explorer.exe " & "http://facebook.com/bornnoLab"
End Sub
