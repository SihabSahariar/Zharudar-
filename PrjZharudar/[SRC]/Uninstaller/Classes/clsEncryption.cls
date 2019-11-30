VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncryption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Function Cryption(strSource As String, _
                          strPassword As String, _
                          EnDeCrypt As Boolean) As String

Dim intPassword As Long
Dim intCrypt    As Long
Dim x           As Long

    For x = 1 To Len(strPassword)
        intPassword = intPassword + Asc(Mid$(strPassword, x, 1))
    Next x
    For x = 1 To Len(strSource)
        If EnDeCrypt Then
            intCrypt = Asc(Mid$(strSource, x, 1)) + intPassword + x
            Do Until intCrypt <= 255
                intCrypt = intCrypt - 255
            Loop
        Else 'ENDECRYPT = FALSE/0
            intCrypt = Asc(Mid$(strSource, x, 1)) - intPassword - x
            Do Until intCrypt > 0
                intCrypt = intCrypt + 255
            Loop
        End If
        Cryption = Cryption & Chr$(intCrypt)
    Next x

End Function

