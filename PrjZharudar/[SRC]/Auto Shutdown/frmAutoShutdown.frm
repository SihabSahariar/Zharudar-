VERSION 5.00
Begin VB.Form autoShutdown 
   BorderStyle     =   0  'None
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutoShutdown.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAutoShutdown.frx":BEF6
   ScaleHeight     =   6330.141
   ScaleMode       =   0  'User
   ScaleWidth      =   10290.09
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optShutdown 
      Appearance      =   0  'Flat
      BackColor       =   &H00479725&
      Caption         =   "Shutdown"
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   2160
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.OptionButton optRestart 
      Appearance      =   0  'Flat
      BackColor       =   &H00479725&
      Caption         =   "Restart"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   2880
      Width           =   3015
   End
   Begin VB.OptionButton optLogoff 
      Appearance      =   0  'Flat
      BackColor       =   &H00479725&
      Caption         =   "Log off"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   3525
      Width           =   3015
   End
   Begin VB.ComboBox cmbHour 
      Height          =   330
      Left            =   6480
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   3315
      Width           =   1215
   End
   Begin VB.ComboBox cmbMin 
      Height          =   330
      Left            =   7920
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3315
      Width           =   975
   End
   Begin VB.ComboBox cmbDay 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6480
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2460
      Width           =   2415
   End
   Begin VB.ComboBox cmbMonth 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   6480
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2640
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   840
   End
   Begin VB.Image Image6 
      Height          =   1125
      Left            =   5640
      Picture         =   "frmAutoShutdown.frx":1272BA
      Top             =   4380
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Image Image5 
      Height          =   1125
      Left            =   2040
      Picture         =   "frmAutoShutdown.frx":12D58F
      Top             =   4381
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   9840
      MouseIcon       =   "frmAutoShutdown.frx":1336E5
      MousePointer    =   99  'Custom
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   0
      MouseIcon       =   "frmAutoShutdown.frx":133837
      MousePointer    =   99  'Custom
      ToolTipText     =   "Back"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      MouseIcon       =   "frmAutoShutdown.frx":133989
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmAutoShutdown.frx":133ADB
      MousePointer    =   99  'Custom
      ToolTipText     =   "Zharudar 1.8 | BornnoLab"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Image cmdCancel 
      Height          =   1095
      Left            =   5640
      MouseIcon       =   "frmAutoShutdown.frx":133C2D
      MousePointer    =   99  'Custom
      ToolTipText     =   "Stop timer..."
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Image cmdSet 
      Height          =   1095
      Left            =   2040
      MouseIcon       =   "frmAutoShutdown.frx":133D7F
      MousePointer    =   99  'Custom
      ToolTipText     =   "Start timer..."
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   0
      Top             =   3720
      Width           =   3135
   End
End
Attribute VB_Name = "autoShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private formHeight As Integer
Private formWidth As Integer


Private Sub ChangeDay(Optional ByVal dateNow As Boolean)
    Dim oldVal As Integer
    If Not dateNow Then oldVal = cmbDay.text
    cmbDay.Clear
    Dim dd As Integer
    Select Case cmbMonth.ListIndex + 1
        Case 1, 3, 5, 7, 8, 10, 12
            dd = 31
        Case 2
            If Val(Format(Now, "YYYY")) Mod 4 = 0 Then
                dd = 29
            Else
                dd = 28
            End If
        Case Else
            dd = 30
    End Select
    Dim i As Integer
    For i = 1 To dd
        cmbDay.AddItem i
        If i = Val(Format(Now, "DD")) And dateNow Then
            cmbDay.text = i
        End If
    Next i
    If Not dateNow Then
        If oldday <= dd Then
            cmbDay.text = oldVal
        Else
            cmbDay.text = dd
        End If
    End If
End Sub

Public Function GetMonthNumber(ByVal monStr As String) As Integer
    Dim retInt As Integer
    Dim month As String
    month = "January, February, March, April, May, June, July, August, September, October, November, December"
    Dim monArr() As String
    monArr = Split(month, ",")
    Dim i As Integer
    For i = 0 To UBound(monArr)
        If UCase(Trim(monArr(i))) = UCase(Trim(monStr)) Then
            retInt = i + 1
        End If
    Next i
    GetMonthNumber = retInt
End Function
Private Sub InitComponents()
    formHeight = Me.height
    formWidth = Me.width
    Dim month As String
    month = "January, February, March, April, May, June, July, August, September, October, November, December"
    Dim monArr() As String
    monArr = Split(month, ",")
    Dim i As Integer
    For i = 0 To UBound(monArr)
        cmbMonth.AddItem Trim(monArr(i))
        If Val(Format(Now, "MM")) = i + 1 Then
            cmbMonth.text = Trim(monArr(i))
        End If
    Next i
    ChangeDay True
    For i = 0 To 23
        cmbHour.AddItem i
        If i = Val(Format(Now, "hh")) Then
            cmbHour.text = i
        End If
    Next i
    For i = 0 To 59
        cmbMin.AddItem i
        If i = Val(Split(Format(Now, "hh:mm"), ":")(1)) Then
            cmbMin.text = i
        End If
    Next i
    lblDate.Caption = "Date: " & Format(Now, "MMMM dd, YYYY hh:mm:ss AM/PM")
End Sub

Private Sub cmbDay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then KeyCode = 0
End Sub

Private Sub cmbDay_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbHour_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then KeyCode = 0
End Sub

Private Sub cmbHour_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbMin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then KeyCode = 0
End Sub

Private Sub cmbMin_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbMonth_Click()
    ChangeDay
End Sub

Private Sub cmbMonth_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then KeyCode = 0
End Sub

Private Sub cmbMonth_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
Timer1.Enabled = False
    Image6.Visible = True
    Call Shell("Shutdown /a") 'to Abort
            
End Sub
Private Sub cmdSet_Click()
    cmdSet.Enabled = False
    cmdCancel.Enabled = True
    Timer1.Enabled = True
    Image5.Visible = True
End Sub

Private Sub Form_Load()
    InitComponents
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        Me.height = formHeight
        Me.width = formWidth
    End If
End Sub

Private Sub optAbort_Click()

End Sub

Private Sub Image1_Click()
    cmdSet.Enabled = True
    cmdCancel.Enabled = False
    Timer1.Enabled = False
End Sub

Private Sub Image2_Click()
Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"
End Sub

Private Sub Image3_Click()
    Unload Me
    index.Show
End Sub

Private Sub Image4_Click()
  Me.Hide
  index.Show
End Sub

Private Sub Timer1_Timer()
    If Val(Format(Now, "MM")) = GetMonthNumber(cmbMonth.text) And Val(Format(Now, "DD")) = Val(cmbDay.text) _
            And Val(Format(Now, "hh")) = Val(cmbHour.text) And Val(Split(Format(Now, "hh:mm"), ":")(1)) = Val(cmbMin.text) Then
        If optShutdown.Value Then
            Call Shell("Shutdown /s") 'to shutdown
            Timer1.Enabled = False
        ElseIf optRestart.Value Then
            Call Shell("Shutdown /r") 'to restart
            Timer1.Enabled = False
        ElseIf optLogoff.Value Then
            Call Shell("Shutdown /l") 'to log off
            Timer1.Enabled = False
        'ElseIf optAbort.Value Then
            
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    lblDate.Caption = "Date: " & Format(Now, "MMMM dd, YYYY hh:mm:ss AM/PM")
End Sub
