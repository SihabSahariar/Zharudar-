VERSION 5.00
Begin VB.Form FrmPopup 
   BorderStyle     =   0  'None
   ClientHeight    =   1260
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmPopup.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu MnuUninstall 
         Caption         =   "Uninstall"
      End
      Begin VB.Menu MnuInformasi 
         Caption         =   "Information Details"
      End
      Begin VB.Menu MnuRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "FrmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub MnuInformasi_Click()
uninstallerInfo.Show vbModal, uninstaller
End Sub
Private Sub MnuPG_Click()
On Error Resume Next
Shell ("explorer C:\program files\"), vbNormalFocus
End Sub
Private Sub MnuPG2_Click()
Shell ("explorer C:\program files\"), vbNormalFocus
End Sub
Private Sub MnuRefresh_Click()
uninstaller.New_Refresh
End Sub
Private Sub MnuUninstall_Click()
uninstaller.Show_FormUninstall
End Sub

