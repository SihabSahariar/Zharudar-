VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form uninstaller 
   BackColor       =   &H001F1414&
   BorderStyle     =   0  'None
   ClientHeight    =   6330
   ClientLeft      =   105
   ClientTop       =   -75
   ClientWidth     =   10260
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":BEF6
   ScaleHeight     =   6330
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10707
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10B62
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstview 
      Height          =   4695
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Select program on the list and then double click list view for run Uninstall program"
      Top             =   960
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   8281
      View            =   3
      Arrange         =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList1"
      ForeColor       =   2036756
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Display Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Uninstall String"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Registry Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Publisher"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Display Version"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Help Link"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "URL Info About "
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Contact"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   0
      MouseIcon       =   "Form1.frx":10EB4
      MousePointer    =   99  'Custom
      ToolTipText     =   "Join us on Facebook"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image i2 
      Height          =   225
      Left            =   9840
      MouseIcon       =   "Form1.frx":11006
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":11158
      ToolTipText     =   "EXIT"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   0
      MouseIcon       =   "Form1.frx":1160E
      MousePointer    =   99  'Custom
      ToolTipText     =   "Back"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      MouseIcon       =   "Form1.frx":11760
      MousePointer    =   99  'Custom
      ToolTipText     =   "Zharudar 1.8 | BornnoLab"
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "uninstaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FormFlag As Boolean, fx As Long, FY As Long
Public FormFirst As Boolean, AX As Long, AY As Long
Dim XX As Integer
Dim YY As Integer
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Dim GetlocString, RegName, UString, Dname, Publisher, DVersion, HelpLink, UIAbout, Contact As String
Dim iKetetapan As Integer
Dim fTimer

Sub GetInformasi()
RegName = lstview.SelectedItem.Key
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayName")
UString = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "UninstallString")
Publisher = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "Publisher"))
DVersion = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayVersion"))
HelpLink = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "HelpLink"))
UIAbout = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "URLInfoAbout"))
Contact = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "Contact"))

If Len(RegName) = 0 Then
uninstallerInfo.TxtRName = "(Not available)"
Else
uninstallerInfo.TxtRName = RegName
End If

If Len(Dname) = 0 Then
uninstallerInfo.TxtDName.text = "(Not available)"
Else
uninstallerInfo.TxtDName.text = Dname
End If

If Len(UString) = 0 Then
uninstallerInfo.TxtUString.text = "(Not available)"
Else
uninstallerInfo.TxtUString.text = UString
End If

If Len(Publisher) = 0 Then
uninstallerInfo.TxtPublisher.text = "(Not available)"
Else
uninstallerInfo.TxtPublisher.text = Publisher
End If

If Len(DVersion) = 0 Then
uninstallerInfo.TxtDVersion.text = "(Not available)"
Else
uninstallerInfo.TxtDVersion.text = DVersion
End If

If Len(HelpLink) = 0 Then
uninstallerInfo.TxtHLink.text = "(Not available)"
Else
uninstallerInfo.TxtHLink.text = HelpLink
End If

If Len(UIAbout) = 0 Then
uninstallerInfo.TxtUIAbout.text = "(Not available)"
Else
uninstallerInfo.TxtUIAbout.text = UIAbout
End If

If Len(Contact) = 0 Then
uninstallerInfo.TxtContact.text = "(Not available)"
Else
uninstallerInfo.TxtContact.text = Contact
End If
    
End Sub

Private Sub GetKetReg()
GetlocString = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
ModRegistry.GetKeyNames HKEY_LOCAL_MACHINE, GetlocString
End Sub

Private Sub ShowUninstallList()
On Error Resume Next
Dim LokasiItem As ListItem
Call GetKetReg
For iKetetapan = 1 To sKeys.count - 0
    Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "DisplayName")
    UString = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "UninstallString")
    Publisher = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "Publisher")
    DVersion = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "DisplayVersion")
    HelpLink = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "HelpLink")
    UIAbout = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "URLInfoAbout")
    Contact = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "Contact")

    If Len(Dname) > 0 Then
        Set LokasiItem = lstview.ListItems.Add(, sKeys(iKetetapan), Dname, 1, 1)
        
        If Len(UString) = 0 Then
           LokasiItem.SubItems(1) = "(Not available)"
        Else
           LokasiItem.SubItems(1) = UString
        End If
            
        If Len(sKeys(iKetetapan)) = 0 Then
           LokasiItem.SubItems(2) = "(Not available)"
        Else
           LokasiItem.SubItems(2) = sKeys(iKetetapan)
        End If
           
        If Len(Publisher) = 0 Then
           LokasiItem.SubItems(3) = "(Not available)"
        Else
           LokasiItem.SubItems(3) = Publisher
        End If
           
        If Len(DVersion) = 0 Then
           LokasiItem.SubItems(4) = "(Not available)"
        Else
           LokasiItem.SubItems(4) = DVersion
        End If
        
        If Len(HelpLink) = 0 Then
           LokasiItem.SubItems(5) = "(Not available)"
        Else
           LokasiItem.SubItems(5) = HelpLink
        End If
        
        If Len(UIAbout) = 0 Then
           LokasiItem.SubItems(6) = "(Not available)"
        Else
           LokasiItem.SubItems(6) = UIAbout
        End If
        
        If Len(Contact) = 0 Then
           LokasiItem.SubItems(7) = "(Not available)"
        Else
           LokasiItem.SubItems(7) = Contact
        End If
End If
Next iKetetapan
    Set sKeys = Nothing
   
End Sub

Sub Show_FormUninstall()
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayName")
RegName = lstview.SelectedItem.Key
frmUninstall.TxtDName.text = Dname
frmUninstall.TxtRegname.text = RegName
frmUninstall.Show vbModal, uninstaller
Get_Uninstall
End Sub

Sub Get_Uninstall()
Dim strRemove As String
strRemove = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "UninstallString")
WinExec strRemove, 1
End Sub









Private Sub cmddelete_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
            
            On Error GoTo 0

End Sub


Private Sub cmdinfo_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next

    On Error GoTo 0

End Sub






Private Sub cmdtweak_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
    On Error GoTo 0

End Sub


Private Sub cmduninstall_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
        
    On Error GoTo 0

End Sub







Private Sub delete_Click()

End Sub









Private Sub Form_Load()
On Error Resume Next
    Set sKeys = New Collection
    lstview.Refresh
    ShowUninstallList
    lstview.View = lvwReport
   
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    GetlocString = ""
    Dim Form As Form
    For Each Form In Forms
        If Form.Name <> Me.Name Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
End Sub


Private Sub ImBantu_Click()

End Sub



Private Sub ImgAbout_Click()
End Sub

Private Sub ImgDetailInfo_Click()
uninstallerInfo.Show vbModal, uninstaller
End Sub
Private Sub ImgKeluar_Click()
Unload Me
End Sub

Private Sub ImgLaporan_Click()
End Sub
Private Sub ImgSetings_Click()

End Sub
Private Sub ImgUninstall_Click()
Call Show_FormUninstall
End Sub

Sub Backup_Registry()
On Error Resume Next
Dim fName As String
fName = App.Path & "\" & "temp" & ".tmp"
SaveKey "HKEY_LOCAL_MACHINE" & "\" & GetlocString & lstview.SelectedItem.Key, fName
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayName")
End Sub
Private Sub Lblbantu_Click()

End Sub
Private Sub LblEdit_Click()

End Sub
Private Sub LblFile_Click()

End Sub
Private Sub LblMaximized_Click()
Me.WindowState = vbMaximized
End Sub

Private Sub LblMinimized_Click()
Me.WindowState = vbNormal
End Sub

Private Sub LblRestore_Click()

Me.WindowState = vbMinimized
End Sub

Sub MnuHapusEntry_Click()
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.Key, "DisplayName")
RegName = lstview.SelectedItem.Key
End Sub


Private Sub LblTampilan_Click()

End Sub



Private Sub LblTools_Click()

End Sub


Private Sub LblUninstall_Click()

End Sub







Private Sub Information_Click()

End Sub



Private Sub large_Click()

End Sub



Private Sub list_Click()

End Sub

Private Sub Image2_Click()
    End
End Sub



Private Sub Image4_Click()
    Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"
End Sub

Private Sub i2_Click()
Unload Me
End Sub

Private Sub Image5_Click()
    Unload Me
    index.Show
End Sub

Private Sub Image9_Click()
    Shell "explorer.exe " & "https://bornno-lab.blogspot.com/"

End Sub

Private Sub lstview_DblClick()
Call Show_FormUninstall
End Sub
Private Sub lstview_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
            On Error GoTo 0

End Sub

Private Sub lstview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu FrmPopup.Menu
End Sub


Private Sub MnuInformation_Click()
uninstallerInfo.Show vbModal, uninstaller
End Sub

Sub New_Refresh()
lstview.ListItems.Clear
Set sKeys = New Collection
ShowUninstallList
lstview.Refresh
lstview.View = lvwReport

End Sub













Private Sub Readmeindo_Click()

End Sub







Private Sub Report_Click()

End Sub

Private Sub Restore_Click()

End Sub





Private Sub tweak_Click()

End Sub



Private Sub XPFrame1_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
       
    On Error GoTo 0

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
