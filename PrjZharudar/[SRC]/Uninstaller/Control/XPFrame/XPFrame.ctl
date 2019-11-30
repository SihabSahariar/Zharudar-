VERSION 5.00
Begin VB.UserControl XPFrame 
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ControlContainer=   -1  'True
   FontTransparent =   0   'False
   ForeColor       =   &H00D54600&
   ScaleHeight     =   1545
   ScaleWidth      =   3795
End
Attribute VB_Name = "XPFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Caption As String
Private m_BorderColor As OLE_COLOR
Private m_TextColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_RightToLeft As Boolean
Private GradientBackground As Boolean
Private Const m_TopColor As Long = &HBFD0D0
Private Const m_BottomColor As Long = &HD54600
Private Sub UserControl_InitProperties()
On Error Resume Next
m_Caption = UserControl.Name
m_RightToLeft = False
m_BorderColor = &HBFD0D0
m_TextColor = &HD54600
m_BackColor = &H8000000F
End Sub
Private Sub UserControl_Paint()
On Error Resume Next
Redraw
End Sub
Private Sub Redraw()
Static SubIsDoing As Boolean
Dim Wid As Long, Hei As Long
Dim LineWidth As Long, LineHeight As Long
Dim ChamWid As Long, ChamHei As Long
Dim Px1 As Long, Px2 As Long
Dim Py1 As Long, Py2 As Long
Dim StartY As Long
If SubIsDoing Then Exit Sub
SubIsDoing = True
UserControl.Backcolor = m_BackColor
Cls
'RedrawBackground
Px1 = TwipX(1): Px2 = 2 * Px1
Py1 = TwipY(1): Py2 = 2 * Py1
Wid = ScaleWidth
Hei = ScaleHeight
ChamWid = TwipX(2)
ChamHei = TwipX(2)
LineWidth = Wid - ChamWid
LineHeight = Hei - ChamHei
StartY = TextHeight(m_Caption) / 2
ForeColor = m_BorderColor
'###### Draw Vertical & Horizontal Lines #######
Line (ChamWid, StartY)-(LineWidth, StartY) ' ----- Upper line
Line (Wid - Px1, ChamHei + StartY)-(Wid - Px1, LineHeight) '||||| Right line
Line (0, ChamHei + StartY)-(0, LineHeight) '|||| Left Line
Line (ChamWid, Hei - Py1)-(LineWidth, Hei - Py1) ' ----- Botton Line
'######### Draw Corner Pixels ############
' Top Left Corner
PSet (Px2, Py1 + StartY)
PSet (Py1, Py2 + StartY)
PSet (Py1, Py1 + StartY)
' Top Right Corner
PSet (LineWidth - Px1, Px1 + StartY)
PSet (LineWidth, Px1 + StartY)
PSet (LineWidth, Px2 + StartY)
' Bottom Left Corner
PSet (Px1, LineHeight - Px1)
PSet (Px2, LineHeight)
PSet (Px1, LineHeight)
' Bottom Right Corner
PSet (LineWidth, LineHeight)
PSet (LineWidth - Px1, LineHeight)
PSet (LineWidth, LineHeight - Px1)
'############# Draw Text! ###############
ForeColor = m_TextColor
If Not m_RightToLeft Then
CurrentX = 7 * Px1
Else
CurrentX = Wid - TextWidth(m_Caption) - (7 * Px1)
End If
CurrentY = 0
Print m_Caption
SubIsDoing = False
End Sub
Private Function TwipX(lngPixel As Long) As Long
TwipX = ScaleX(lngPixel, vbPixels, vbTwips)
End Function
Private Function TwipY(lngPixel As Long) As Long
TwipY = ScaleY(lngPixel, vbPixels, vbTwips)
End Function
Private Sub RedrawBackground()
GradientBackground = True
If GradientBackground Then
Dim i As Single, Steps As Single
Dim r1 As Single, g1 As Single, b1 As Single
Dim r2 As Single, g2 As Single, b2 As Single
Dim r As Single, G As Single, B As Single
Dim rs As Single, gs As Single, bs As Single
Dim c As Long
Dim Hei As Long, Wid As Long
Wid = ScaleWidth
Hei = ScaleHeight
ColorToRGB m_TopColor, r1, g1, b1
ColorToRGB m_BottomColor, r2, g2, b2
rs = (r2 - r1) / Hei
gs = (g2 - g1) / Hei
bs = (b2 - b1) / Hei
r = r1: G = g1: B = b1
For i = 1 To Hei
r = r + rs: G = G + gs: B = B + bs
c = r + 256& * G + 65536 * B
Line (0, i)-(Wid, i), c
Next
End If
End Sub
Private Sub ColorToRGB(Color As Long, r As Single, G As Single, B As Single)
B = ((Color \ &H10000) Mod &H100)
G = ((Color \ &H100) Mod &H100)
r = (Color And &HFF)
End Sub
Property Let RightToLeft(NewVal As Boolean)
m_RightToLeft = NewVal
PropertyChanged "RightToLeft"
Redraw
End Property
Property Get RightToLeft() As Boolean
RightToLeft = m_RightToLeft
End Property
Property Let Caption(NewVal As String)
m_Caption = NewVal
PropertyChanged "Caption"
Redraw
End Property
Property Get Caption() As String
Caption = m_Caption
End Property
Property Set Font(NewVal As Font)
Set UserControl.Font = NewVal
PropertyChanged "Font"
Redraw
End Property
Property Get Font() As Font
Set Font = UserControl.Font
End Property
Property Let FontItalic(NewVal As Boolean)
UserControl.FontItalic = NewVal
PropertyChanged "FontItalic"
Redraw
End Property
Property Get FontItalic() As Boolean
FontItalic = UserControl.FontItalic
End Property
Property Let FontBold(NewVal As Boolean)
UserControl.FontBold = NewVal
PropertyChanged "FontBold"
Redraw
End Property
Property Get FontBold() As Boolean
FontBold = UserControl.FontBold
End Property
Property Let FontName(NewVal As String)
UserControl.FontName = NewVal
PropertyChanged "FontName"
Redraw
End Property
Property Get FontName() As String
FontName = UserControl.FontName
End Property
Property Let FontSize(NewVal As Long)
UserControl.FontSize = NewVal
PropertyChanged "FontSize"
Redraw
End Property
Property Get FontSize() As Long
FontSize = UserControl.FontSize
End Property
Property Let BorderColor(NewVal As OLE_COLOR)
m_BorderColor = NewVal
PropertyChanged "BorderColor"
Redraw
End Property
Property Get BorderColor() As OLE_COLOR
BorderColor = m_BorderColor
End Property
Property Let TextColor(NewVal As OLE_COLOR)
m_TextColor = NewVal
PropertyChanged "TextColor"
Redraw
End Property
Property Get TextColor() As OLE_COLOR
TextColor = m_TextColor
End Property
Property Let Backcolor(NewVal As OLE_COLOR)
m_BackColor = NewVal
PropertyChanged "BackColor"
Redraw
End Property
Property Get Backcolor() As OLE_COLOR
Backcolor = m_BackColor
End Property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
PropBag.WriteProperty "Caption", m_Caption, vbNullString
PropBag.WriteProperty "RightToLeft", m_RightToLeft, False
PropBag.WriteProperty "Font", UserControl.Font
PropBag.WriteProperty "BorderColor", m_BorderColor, &HBFD0D0
PropBag.WriteProperty "TextColor", m_TextColor, &HD54600
PropBag.WriteProperty "BackColor", m_BackColor, &H8000000F
PropBag.WriteProperty "FontName", UserControl.FontName
PropBag.WriteProperty "FontSize", UserControl.FontSize
PropBag.WriteProperty "FontBold", UserControl.FontBold
PropBag.WriteProperty "FontItalic", UserControl.FontItalic
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
m_Caption = PropBag.ReadProperty("Caption", vbNullString)
m_RightToLeft = PropBag.ReadProperty("RightToLeft", False)
Set UserControl.Font = PropBag.ReadProperty("Font")
m_BorderColor = PropBag.ReadProperty("BorderColor", &HBFD0D0)
m_TextColor = PropBag.ReadProperty("TextColor", &HD54600)
m_BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
UserControl.FontName = PropBag.ReadProperty("FontName")
UserControl.FontSize = PropBag.ReadProperty("FontSize")
UserControl.FontBold = PropBag.ReadProperty("FontBold")
UserControl.FontItalic = PropBag.ReadProperty("FontItalic")
End Sub
