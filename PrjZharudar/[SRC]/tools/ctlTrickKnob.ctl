VERSION 5.00
Begin VB.UserControl ctlTrickKnob 
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Malgun Gothic"
      Size            =   6
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ctlTrickKnob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

' // When value has been changed
Public Event Change()

Private Const pi    As Double = 3.14159265358979

Private mMax        As Long             ' // Maximum value
Private mMin        As Long             ' // Minimum value
Private mValue      As Long             ' // Current value
Private mCaption    As String           ' // Label

Dim gdipToken   As Long                 ' // GDI+ token
Dim mBufGraph   As Long                 ' // Graphics of buffer
Dim mSelected   As Boolean              ' // Selection flag
Dim mRedrawBack As Boolean              ' // Background has been changed flag
Dim hBufferBmp  As Long                 ' // Buffer bitmap
Dim hBackBmp    As Long                 ' // Background bitmap
Dim tmpDC       As Long                 ' // Temporary device context
Dim bufDC       As Long                 ' // Buffer device context
Dim mLineBrush  As Long                 ' // Gradient brush
Dim mKnobBrush  As Long                 ' // Knob brush
Dim mKnobSize   As Long                 ' // Size of knob
Dim mOldMouse   As POINTF

' // Current value
Public Property Let Value(ByVal v As Long)

    If v > mMax Or v < mMin Then
    
'        Err.Raise 5
        'Exit Property
        
    End If
    
    If mValue <> v Then
    
        mValue = v
        RaiseEvent Change
        Redraw
        PropertyChanged "Value"
        
    End If
    
End Property
Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0

    Value = mValue

End Property

' // Maximum
Public Property Let Max(ByVal v As Long)

    If v < mMin Then
    
        Err.Raise 5
        Exit Property
        
    End If
    
    mMax = v
    
    If mValue > v Then
    
        mValue = v
        RaiseEvent Change
        
    End If
    
    Redraw
    
    PropertyChanged "Max"
End Property
Public Property Get Max() As Long

    Max = mMax

End Property

' // Minimum
Public Property Let Min(ByVal v As Long)

    If v > mMax Then
    
        Err.Raise 5
        Exit Property
        
    End If

    mMin = v
    
    If mValue < v Then
    
        mValue = v
        RaiseEvent Change
        
    End If
    
    Redraw
    
    PropertyChanged "Min"
End Property
Public Property Get Min() As Long

    Min = mMin

End Property

' // Caption/label
Public Property Let Caption(ByVal v As String)

    mCaption = v
    
    PropertyChanged "Caption"
End Property
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518

    Caption = mCaption

End Property

' // Redraw
Public Sub Redraw()
    ' // Save context state
    SaveDC tmpDC
    ' // If bcakground has been changed
    If mRedrawBack Then

        Dim tmpFont As IFont
        Dim rc      As RECT
        Dim txtH    As Long
        Dim sz      As Size
        ' // Query IFont interface to get font handle
        Set tmpFont = UserControl.Font
        '// If buffer bitmaps exist then delete them
        If CBool(hBufferBmp) Or CBool(hBackBmp) Then
        
            RestoreDC bufDC, -1
            GdipDeleteGraphics mBufGraph
            GdipDeleteBrush mLineBrush
            DeleteObject hBufferBmp
            DeleteObject hBackBmp
            
        End If
        ' // Create buffer and background bitmap
        hBufferBmp = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)
        hBackBmp = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)
        ' // Save context state
        SaveDC bufDC
        SelectObject bufDC, hBufferBmp
        SelectObject bufDC, tmpFont.hFont
        SetBkMode bufDC, TRANSPARENT
        ' // Calculate height of font
        GetTextExtentPoint32 bufDC, StrPtr("0"), 1, sz
        sz.cy = sz.cy + 3
        
        mKnobSize = IIf(UserControl.ScaleWidth > UserControl.ScaleHeight - sz.cy, UserControl.ScaleHeight - sz.cy, UserControl.ScaleWidth - 1)
        ' // Create buffer graphics
        GdipCreateFromHDC bufDC, mBufGraph
        ' // Enable antialiasing
        GdipSetSmoothingMode mBufGraph, SmoothingModeAntiAlias
        ' // Create gradient brush
        GdipCreateLineBrush GdipPointF(0, 0), GdipPointF(0, mKnobSize), _
                            Colors.col_GradB1, Colors.col_GradB2, WrapModeTile, mLineBrush
        ' // Redraw parent background
        SelectObject tmpDC, hBackBmp
        GetClientRect UserControl.hwnd, rc
        MapWindowPoints UserControl.hwnd, UserControl.ContainerHwnd, rc, 2
        SetViewportOrgEx tmpDC, -rc.Left, -rc.Top, ByVal 0&
        SendMessageA UserControl.ContainerHwnd, WM_PAINT, tmpDC, ByVal 0&
        SetViewportOrgEx tmpDC, 0, 0, ByVal 0&
        mRedrawBack = False
        
    End If
    ' // Redraw background
    SelectObject tmpDC, hBackBmp
    BitBlt bufDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
           tmpDC, 0, 0, vbSrcCopy
    ' // Enclosure
    GdipFillEllipse mBufGraph, mLineBrush, 0, 0, mKnobSize, mKnobSize
    ' // Knob
    Dim posX    As Single
    Dim posY    As Single
    Dim ang     As Single
    Dim ofst    As Single
    Dim label   As String
    
    ang = (mValue - mMin) / (mMax - mMin) * (pi * 1.5) + (pi * 0.75)
    ofst = mKnobSize / 5
    
    posX = Cos(ang) * (mKnobSize / 2 - ofst / 1.3) + mKnobSize / 2
    posY = Sin(ang) * (mKnobSize / 2 - ofst / 1.3) + mKnobSize / 2
    
    If mSelected Then
        GdipSetSolidFillColor mKnobBrush, Colors.col_Knob
    Else
        GdipSetSolidFillColor mKnobBrush, Colors.col_KnobInactive
    End If
    
    GdipFillEllipse mBufGraph, mKnobBrush, posX - ofst / 2, posY - ofst / 2, ofst, ofst
    ' // Draw caption
    SetRect rc, 0, mKnobSize + 3, UserControl.ScaleWidth, UserControl.ScaleHeight
    SetTextColor bufDC, Colors.col_TextPanel And &HFFFFFF
    DrawText bufDC, StrPtr(mCaption), Len(mCaption), rc, DT_CENTER
    ' // Draw from buffer to control
    BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
           bufDC, 0, 0, vbSrcCopy
    ' // Restore device context
    RestoreDC tmpDC, -1
    
End Sub

' // Control initialization
Private Sub UserControl_Initialize()
    Dim ret     As Long
    Dim gpInput As GdiplusStartupInput
    ' // GDI+ initialization
    gpInput.GdiplusVersion = 1
    ret = GdiplusStartup(gdipToken, gpInput)
    
    If ret Then
        MsgBox "Gdi+ startup error"
        Exit Sub
    End If
    ' // Resources
    bufDC = CreateCompatibleDC(UserControl.hdc)
    tmpDC = CreateCompatibleDC(UserControl.hdc)
    GdipCreateSolidFill Colors.col_KnobInactive, mKnobBrush
    
End Sub

Private Sub UserControl_InitProperties()
    mMin = 0
    mMax = 100
    mValue = 50
    mCaption = vbNullString
End Sub

' // Press mouse within control
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Dim rad As Single
   ' Dim px  As Long
  '  Dim py  As Long
    
   ' px = x - mKnobSize / 2
   ' py = y - mKnobSize / 2
    
   ' rad = px * px + py * py
    
    'If rad <= (mKnobSize * mKnobSize / 4) Then
    
       ' mSelected = True
        'mOldMouse = GdipPointF(x, y)
        'Redraw
        
    'End If
    
End Sub

' // Move mouse
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If mSelected Then
        
        Dim dx  As Single
        Dim dy  As Single
        Dim v   As Long
        
        dx = x - mOldMouse.x
        dy = y - mOldMouse.y
        
        v = mValue - (dy * (mMax - mMin) / 300)
        
        If v > mMax Then v = mMax
        If v < mMin Then v = mMin
        
        If v <> mValue Then
            
            mOldMouse = GdipPointF(x, y)

            mValue = v
            RaiseEvent Change
            Redraw
            
        End If
        
    End If
    
End Sub

' // Release mouse
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If mSelected Then
    
        mSelected = False
        Redraw
        
    End If
    
End Sub

' // Redraw
Private Sub UserControl_Paint()
    
    If mRedrawBack Then
        Redraw
    End If
    
    BitBlt UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, _
           bufDC, 0, 0, vbSrcCopy
           
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    mMin = PropBag.ReadProperty("Min", 0)
    mMax = PropBag.ReadProperty("Max", 100)
    mValue = PropBag.ReadProperty("Value", 50)
    mCaption = PropBag.ReadProperty("Caption", vbNullString)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "Min", mMin, 0
    PropBag.WriteProperty "Max", mMax, 100
    PropBag.WriteProperty "Value", mValue, 50
    PropBag.WriteProperty "Caption", mCaption, vbNullString
    
End Sub

' // Resize
Private Sub UserControl_Resize()
    
    mRedrawBack = True
    
End Sub

' // Uninitialization
Private Sub UserControl_Terminate()

    If gdipToken Then
        
        If CBool(hBufferBmp) Or CBool(hBackBmp) Then
        
            RestoreDC bufDC, -1
            GdipDeleteGraphics mBufGraph
            GdipDeleteBrush mLineBrush
            DeleteObject hBufferBmp
            DeleteObject hBackBmp
            
        End If
        
        DeleteDC tmpDC
        DeleteDC bufDC
        GdipDeleteBrush mKnobBrush
        GdiplusShutdown gdipToken
        
    End If
    
End Sub

