VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form crypter 
   BorderStyle     =   0  'None
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   Icon            =   "xeb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "xeb.frx":BEF6
   ScaleHeight     =   6360
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6360
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      DragMode        =   1  'Automatic
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      TabIndex        =   1
      Text            =   "Password !"
      ToolTipText     =   "Password !"
      Top             =   3480
      Width           =   2550
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   2760
      Width           =   3350
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   9840
      MouseIcon       =   "xeb.frx":1272BA
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   0
      MouseIcon       =   "xeb.frx":12740C
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   240
      MouseIcon       =   "xeb.frx":12755E
      MousePointer    =   99  'Custom
      ToolTipText     =   "Zharudar 1.8 | BornnoLab"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   0
      MouseIcon       =   "xeb.frx":1276B0
      MousePointer    =   99  'Custom
      ToolTipText     =   "Back"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   1575
      Left            =   2640
      MouseIcon       =   "xeb.frx":127802
      MousePointer    =   99  'Custom
      ToolTipText     =   "Select your file..."
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   6240
      MouseIcon       =   "xeb.frx":127954
      MousePointer    =   99  'Custom
      ToolTipText     =   "Click here to Decrypt..."
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   6240
      MouseIcon       =   "xeb.frx":127AA6
      MousePointer    =   99  'Custom
      ToolTipText     =   "Click here to Encrypt..."
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   4725
   End
End
Attribute VB_Name = "crypter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOCONFIRMATION As Long = &H1

Dim XX As Integer
Dim YY As Integer

Private Sub cmdBrowse_Click()

End Sub

Private Sub cmdDecrypt_Click()

End Sub

Private Sub cmdEncrypt_Click()

End Sub



Private Sub DoEncrypt()
        Dim csCrypt       As New clsCrypto
    Dim strFile       As String
    Dim lFileLength   As Long
    

    lFileLength = FileLen(txtFileName.text)
   
    strFile = String(lFileLength, vbNullChar)
      Open txtFileName.text For Binary Access Read As #1
    
    Get 1, , strFile
    Close #1
    
    csCrypt.Password = txtPassword.text
    csCrypt.InBuffer = strFile
    
    If Not csCrypt.HashFile Then Exit Sub
    
    If Not csCrypt.GeneratePasswordKey Then Exit Sub
   
    If Not csCrypt.EncryptFileData Then Exit Sub
    
    csCrypt.DestroySessionKey
    
    If csCrypt.OutBuffer <> "" Then
        
        Kill txtFileName.text
        
        Open txtFileName.text For Binary Access Write As #2
        
        Put 2, , csCrypt.OutBuffer
        Close #2
    End If
End Sub

Private Sub DoDecrypt()
    
    Dim csCrypt     As New clsCrypto
    Dim strFile     As String
    Dim lFileLength As String
    
    
    lFileLength = FileLen(txtFileName.text)
   
    strFile = String(lFileLength, vbNullChar)
    
    Open txtFileName.text For Binary Access Read As #1
    
    Get 1, , strFile
    Close #1

    csCrypt.Password = txtPassword.text
    csCrypt.InBuffer = strFile

    If Not csCrypt.GeneratePasswordKey Then Exit Sub

    If Not csCrypt.DecryptFileData Then Exit Sub
    csCrypt.DestroySessionKey
    

    If csCrypt.OutBuffer <> "" Then

        Kill txtFileName.text

        Open txtFileName.text For Binary Access Write As #2
        Put 2, , csCrypt.OutBuffer
        Close #2
    End If
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
Timer1.Enabled = True
ProgressBar1.Visible = True
ProgressBar1.Value = ProgressBar1.Min

    Me.MousePointer = 11
    If txtFileName.text = "" Then
      
        MsgBox "Please specify file to encrypt!", vbCritical, "No File"
        Me.MousePointer = 0
        lblMessage.Caption = ""
        txtFileName.SetFocus
        Exit Sub
    ElseIf Dir(txtFileName.text, vbNormal) = "" Then
      
        MsgBox "Invalid file name, or file missing!", vbCritical, "Invalid File"
        txtFileName.SetFocus
        Me.MousePointer = 0
        lblMessage.Caption = ""
        Exit Sub
    ElseIf txtPassword.text = "" Then
      
        MsgBox "Password required for encrypting file!", vbCritical, "No Password"
        txtPassword.SetFocus
        Me.MousePointer = 0
        lblMessage.Caption = ""
        Exit Sub
    Else
        DoEncrypt
    End If
    Me.MousePointer = 0
End Sub

Private Sub Image2_Click()
Timer1.Enabled = True
ProgressBar1.Visible = True
ProgressBar1.Value = ProgressBar1.Min
    Me.MousePointer = 11
    If txtFileName.text = "" Then
            MsgBox "Please specify file to encrypt!", vbCritical, "No File"
        Me.MousePointer = 0
        lblMessage.Caption = ""
        txtFileName.SetFocus
        Exit Sub
    ElseIf Dir(txtFileName.text, vbNormal) = "" Then
                
                MsgBox "Invalid file name, or file missing!", vbCritical, "Invalid File"
        txtFileName.SetFocus
        Me.MousePointer = 0
        lblMessage.Caption = ""
        Exit Sub
    ElseIf txtPassword.text = "" Then
        
        MsgBox "Password required for decrypting file!", vbCritical, "No Password"
        txtPassword.SetFocus
        Me.MousePointer = 0
        lblMessage.Caption = ""
        Exit Sub
    Else
        DoDecrypt
    End If
    Me.MousePointer = 0
End Sub

Private Sub Image3_Click()
    CommonDialog1.Filter = "All Files|*.*"
    CommonDialog1.ShowOpen
    txtFileName = CommonDialog1.FileName
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Image4_Click()
    Unload Me
    index.Show
End Sub

Private Sub Image6_Click()
Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"
End Sub

Private Sub Image7_Click()
   Unload Me
    index.Show
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 2
If ProgressBar1.Value = 10 Then
ProgressBar1.Value = ProgressBar1 + 10
If ProgressBar1.Value >= ProgressBar1.Max Then
End If
End If
End Sub
