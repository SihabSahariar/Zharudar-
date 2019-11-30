VERSION 5.00
Begin VB.Form uninstallerInfo 
   BackColor       =   &H001F1710&
   BorderStyle     =   0  'None
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10215
   FillColor       =   &H001F1710&
   LinkTopic       =   "Form1"
   Picture         =   "uninstallerInfo.frx":0000
   ScaleHeight     =   6360
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Top             =   1080
      Width           =   9375
      Begin VB.TextBox TxtDName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   7575
      End
      Begin VB.TextBox TxtUString 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   7575
      End
      Begin VB.TextBox TxtPublisher 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2040
         Width           =   7575
      End
      Begin VB.TextBox TxtDVersion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2400
         Width           =   7575
      End
      Begin VB.TextBox TxtHLink 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3240
         Width           =   7575
      End
      Begin VB.TextBox TxtUIAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3600
         Width           =   7575
      End
      Begin VB.TextBox TxtContact 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3960
         Width           =   7575
      End
      Begin VB.TextBox TxtRName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   7575
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
         TabIndex        =   16
         Top             =   600
         Width           =   1215
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
         TabIndex        =   15
         Top             =   1680
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
         TabIndex        =   14
         Top             =   2040
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
         TabIndex        =   13
         Top             =   2400
         Width           =   1215
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
         TabIndex        =   12
         Top             =   3240
         Width           =   975
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
         TabIndex        =   11
         Top             =   3600
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
         TabIndex        =   10
         Top             =   3960
         Width           =   855
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
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   -120
      MouseIcon       =   "uninstallerInfo.frx":4811
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   -120
      MouseIcon       =   "uninstallerInfo.frx":4963
      MousePointer    =   99  'Custom
      ToolTipText     =   "Back"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image i2 
      Height          =   225
      Left            =   9840
      MouseIcon       =   "uninstallerInfo.frx":4AB5
      MousePointer    =   99  'Custom
      Picture         =   "uninstallerInfo.frx":4C07
      ToolTipText     =   "EXIT"
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "uninstallerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Sub Form_Load()
uninstaller.GetInformasi
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub




Private Sub Image5_Click()
Unload Me
End Sub

Private Sub Image9_Click()
    Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"

End Sub
