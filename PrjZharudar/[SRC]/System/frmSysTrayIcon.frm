VERSION 5.00
Begin VB.Form frmSysTrayIcon 
   Caption         =   "Zharudar"
   ClientHeight    =   900
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   1740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSysTrayIcon.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSysTrayIcon.frx":BEF6
   ScaleHeight     =   900
   ScaleWidth      =   1740
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu ui 
         Caption         =   "Zharudar Dashbord"
      End
      Begin VB.Menu ff 
         Caption         =   "-"
      End
      Begin VB.Menu rtuyru 
         Caption         =   "About"
         Index           =   15
      End
      Begin VB.Menu mnuFileArray 
         Caption         =   "Exit"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmSysTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'User-defined variable to pass to the Shell_NotiyIcon function
Private Type NOTIFYICONDATA
    cbSize              As Long
    hwnd                As Long
    uId                 As Long
    uFlags              As Long
    uCallBackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type

'Constants for the Shell_NotifyIcon function
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

'Declare the API function call
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    
Dim nid As NOTIFYICONDATA

Public Sub AddIcon(ByVal ToolTip As String)

    On Error GoTo ErrorHandler
    
    'Add icon to system tray
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = ToolTip & vbNullChar
    End With
    Call Shell_NotifyIcon(NIM_ADD, nid)
    
Exit Sub
ErrorHandler:   'Display error message
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbInformation, App.ProductName & " - " & Me.Caption

End Sub

Private Sub About2_Click(Index As Integer)

End Sub



Private Sub Form_Load()

    Call AddIcon("Zharudar 1.9")
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim msg             As Long

    On Error GoTo ErrorHandler
    
    'Respond to user interaction
    msg = x / Screen.TwipsPerPixelX
    Select Case msg
            
        Case WM_LBUTTONDBLCLK
            'nothing
    
        Case WM_LBUTTONDOWN
            'nothing
        
        Case WM_LBUTTONUP
            If Me.WindowState = vbMinimized Then
                Me.WindowState = vbNormal
                Me.Show
            Else
                Me.WindowState = vbMinimized
                Me.Hide
            End If
            
        Case WM_RBUTTONDBLCLK
            'nothing
        
        Case WM_RBUTTONDOWN
            'nothing
        
        Case WM_RBUTTONUP
            Call PopupMenu(mnuFile, vbPopupMenuRightAlign)
            
    End Select
    
Exit Sub
ErrorHandler:   'Display error message
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbInformation, App.ProductName & " - " & Me.Caption

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Remove icon from system tray
    Call Shell_NotifyIcon(NIM_DELETE, nid)

End Sub

Private Sub mnuFileArray_Click(Index As Integer)
End
End Sub

Private Sub rtuyru_Click(Index As Integer)
about.Show
about.w.URL = App.Path & "/bg.mp3"

End Sub





Private Sub ui_Click()
Index.Show
End Sub
