VERSION 5.00
Begin VB.Form frmUninstall 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Merry Uninstaller 2005 - Uninstal Program"
   ClientHeight    =   2190
   ClientLeft      =   -45
   ClientTop       =   -225
   ClientWidth     =   3855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmUninstall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmUninstall.frx":000C
   ScaleHeight     =   2190
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Uninstaller2005.XPButton LblBatal 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Uninstaller2005.XPButton LblUninstall 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Uninstall"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Uninstaller2005.XPButton LblBersihkan 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtRegname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox TxtDname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label LblInfomasi 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmUninstall.frx":1B896
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label LblInfomasi2 
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure to uninstall this program ?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmUninstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : FrmUninstall.frm
' Tanggal    : 8/29/2005 22:29
' Programmer : Rusman Indradi (rusman@olivault.com)
' Lokasi     : Bogor, INDONESIA
' Catatan    : Rusman Indradi ekeur stres Gw Teh euY... untuk sapa yach program ini..
'              ok deCh untuk Temen gw saudara gw yayang GW CroTZ selalu.... :)
'              tHanKz tO Rizki Priatna, Abby, Ronny, pon-pon, Maryam thaNk's for
'              yOur support Euy..... Hapy CodinG and dont forGEt me Ok....
'              unTuk mAryam And pon-pon kapan Ceng-Ceng lg euY......
'
' Website    : wwww.olivault.com
' Contact HP : ?
' E-mail     : intouch@olivault.com
'
'                                  Roes Love Maryam
'
'Note       : This Code Source is destined to You which wish to learn
'             programming.by using is Visual Basic 6.0. If You use this code source,
'             expect that remain to mention the name of me in part of Your About
'             application( Credit Title) as well as in part of Your place code source
'             using it ( IDEA of VB6). Usage of code source for the purpose of is
'             commercial / profit, HAVE TO PERMIT OF its OWNER.
'             Trespasser- an of this thing can be ensnared by penalization
'             related to misdemeanour of Copyrights and [Code/Law] Rights Of Intellectual.
'---------------------------------------------------------------------------------------
Option Explicit

Private Sub Label1_Click()

End Sub

Private Sub LblBatal_Click()
Unload Me
End Sub


Private Sub LblBersihkan_Click()
Call DeleteKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + TxtRegname.text)
Unload Me
frmmain.New_Refresh
End Sub


Private Sub LblUninstall_Click()
LblInfomasi2.Visible = False
TxtDname.Visible = False
LblUninstall.Visible = False
LblInfomasi.Visible = True
LblBersihkan.Visible = True
frmmain.Get_Uninstall
End Sub


Private Sub XPFrame1_GotFocus()

End Sub
