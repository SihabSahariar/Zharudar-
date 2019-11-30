VERSION 5.00
Begin VB.Form frmUninstall 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   2190
   ClientLeft      =   -45
   ClientTop       =   -225
   ClientWidth     =   3855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "msgUnstaller.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "msgUnstaller.frx":BEF6
   ScaleHeight     =   2190
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Zharudar.XPButton XPButton3 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Zharudar.XPButton XPButton2 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Uninstall"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Zharudar.XPButton XPButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Refresh"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      Caption         =   $"msgUnstaller.frx":27780
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
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "frmUninstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub XPButton1_Click()
Call DeleteKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + TxtRegname.text)
uninstaller.New_Refresh
End Sub

Private Sub XPButton2_Click()
LblInfomasi2.Visible = False
TxtDname.Visible = False
XPButton2.Visible = False
LblInfomasi.Visible = True
XPButton1.Visible = False
uninstaller.Get_Uninstall
End Sub

Private Sub XPButton3_Click()
    Unload Me
End Sub
