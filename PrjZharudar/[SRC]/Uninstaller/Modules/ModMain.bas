Attribute VB_Name = "ModMain"
'---------------------------------------------------------------------------------------
' Nama File  : ModMain.bas
' Tanggal    : 8/29/2005 22:26
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
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                                                ByVal lpOperation As String, _
                                                                                ByVal lpFile As String, _
                                                                                ByVal lpParameters As String, _
                                                                                ByVal lpDirectory As String, _
                                                                                ByVal nShowCmd As Long) As Long



Public Sub SaveFile1(ff As Form)
On Error Resume Next
    Dim fName As String

    Dim i As Long
    fName = App.Path & "\" & "temp" & ".tmp"
    Close #1
       
            Open fName For Output As #1
            Print #1, "--> Create Report On : " & Now
            Print #1, "--> MERRY UNINSTALLER 2004 v 1.0 - Copyright © 2005 Olivault Software"
            Print #1, "--> By Rusman Indradi - Olivault Software"
            Print #1, "--> http://olivault.com"
            Print #1, "--> Find : " & ff.lstview.ListItems.Count & " Installed Programs"
            Print #1, ""
            Print #1, "INSTALLED PROGRAMS :"
            Print #1, ""
            For i = 1 To ff.lstview.ListItems.Count
            Print #1, ff.lstview.ListItems(i).text
        Next i
    
End Sub

Public Sub SaveFile2(ff As Form)
On Error Resume Next
    Dim fName As String

    Dim i As Long
fName = App.Path & "\" & "temp" & ".tmp"
        Close #1
       
            Open fName For Output As #1
            Print #1, "--> Create Report On : " & Now
            Print #1, "--> MERRY UNINSTALLER 2004 v 1.0 - Copyright © 2005 Olivault Software"
            Print #1, "--> By Rusman Indradi - Olivault Software"
            Print #1, "--> http://olivault.com"
            Print #1, "--> Find : " & ff.lstview.ListItems.Count & " Installed Programs"
            Print #1, ""
            Print #1, "INSTALLED PROGRAMS :"
            Print #1, ""
            For i = 1 To ff.lstview.ListItems.Count
            Print #1, ff.lstview.ListItems(i).text
            Print #1, ff.lstview.ListItems(i).ListSubItems(1).text
            Print #1, ""
        Next i
    
End Sub

Public Sub SaveFile3(ff As Form)
On Error Resume Next
    Dim fName As String

    Dim i As Long
fName = App.Path & "\" & "temp" & ".tmp"
        Close #1
       
            Open fName For Output As #1
            Print #1, "--> Create Report On : " & Now
            Print #1, "--> MERRY UNINSTALLER 2004 V1.0 - Copyright © 2005 Olivault Software"
            Print #1, "--> By Rusman Indradi - Olivault Software"
            Print #1, "--> http://olivault.com"
            Print #1, "--> Find : " & ff.lstview.ListItems.Count & " Installed Programs"
            Print #1, ""
            Print #1, "INSTALLED PROGRAMS :"
            Print #1, ""
            For i = 1 To ff.lstview.ListItems.Count
            Print #1, "Display Name --> " & ff.lstview.ListItems(i).text
            Print #1, "Uninstall String --> " & ff.lstview.ListItems(i).ListSubItems(1).text
            Print #1, "Registry Name --> " & ff.lstview.ListItems(i).ListSubItems(2).text
            Print #1, "Publisher --> " & ff.lstview.ListItems(i).ListSubItems(3).text
            Print #1, "Display Version --> " & ff.v.ListItems(i).ListSubItems(4).text
            Print #1, "HelpLink --> " & ff.v.ListItems(i).ListSubItems(5).text
            Print #1, "URL Info About --> " & ff.lstview.ListItems(i).ListSubItems(6).text
            Print #1, "Contact --> " & ff.lstview.ListItems(i).ListSubItems(7).text
            Print #1, ""
        Next i
    
End Sub

Public Function HyperJump(ByVal url As String) As Long
On Error Resume Next
HyperJump = ShellExecute(0&, vbNullString, url, vbNullString, vbNullString, vbNormalFocus)
End Function
