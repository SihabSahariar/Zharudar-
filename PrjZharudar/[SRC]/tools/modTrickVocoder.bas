Attribute VB_Name = "modTrickVocoder"
Option Explicit

Public Enum Colors
    col_BackColor = &HFF000000
    col_FrameColor = &HFF330099
    col_GradA1 = &HFF0D0024
    col_GradA2 = &HFF030010
    col_GradB1 = &HFF333333
    col_GradB2 = &HFF191919
    col_Curve = &HFF250084
    col_Marker = &H306699FF
    col_MarkerSelection = &HFFFF9966
    col_Grid = &H80330099
    col_GridInactive = &H40330099
    col_Text = &HFF333399
    col_TextInactive = &H80333399
    col_TextPanel = &HFF7F7F7F
    col_Knob = &HFFFF9966
    col_KnobInactive = &HFF4C4C4C
End Enum

Public Type POINTF
    x                           As Single
    y                           As Single
End Type

Public Type RECTF
    Left                        As Single
    Top                         As Single
    Right                       As Single
    Bottom                      As Single
End Type

Public Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Public Type RGBQUAD
    rgbBlue                     As Byte
    rgbGreen                    As Byte
    rgbRed                      As Byte
    rgbReserved                 As Byte
End Type

Public Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

Public Type Size
    cx                          As Long
    cy                          As Long
End Type

Public Type COLOR16
    Value                       As Integer
End Type
Public Type COLOR32
    Value                       As Long
End Type

Public Type TRIVERTEX
    x                           As Long
    y                           As Long
    red                         As COLOR16
    green                       As COLOR16
    blue                        As COLOR16
    alpha                       As COLOR16
End Type
 
Public Type GRADIENT_RECT
    upperLeft                   As Long
    lowerRight                  As Long
End Type

Private Type OPENFILENAME
    lStructSize                 As Long
    hwndOwner                   As Long
    hInstance                   As Long
    lpstrFilter                 As Long
    lpstrCustomFilter           As Long
    nMaxCustFilter              As Long
    nFilterIndex                As Long
    lpstrFile                   As Long
    nMaxFile                    As Long
    lpstrFileTitle              As Long
    nMaxFileTitle               As Long
    lpstrInitialDir             As Long
    lpstrTitle                  As Long
    Flags                       As Long
    nFileOffset                 As Integer
    nFileExtension              As Integer
    lpstrDefExt                 As Long
    lCustData                   As Long
    lpfnHook                    As Long
    lpTemplateName              As Long
End Type

Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, Graphics As Long) As Long
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As Long
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal width As Single, ByVal unit As Long, Pen As Long) As Long
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal Pen As Long) As Long
Public Declare Function GdipSetPenColor Lib "gdiplus" (ByVal Pen As Long, ByVal ARGB As Long) As Long
Public Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMd As Long) As Long
Public Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal Graphics As Long, ByVal lColor As Long) As Long
Public Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
Public Declare Function GdipCreateLineBrush Lib "gdiplus" (point1 As POINTF, point2 As POINTF, ByVal color1 As Long, ByVal color2 As Long, ByVal WrapMd As Long, lineGradient As Long) As Long
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Public Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
Public Declare Function GdipDrawLineI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
Public Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As Long, Graphics As Long) As Long
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Public Declare Function GdipCreateFontFromDC Lib "gdiplus" (ByVal hdc As Long, createdfont As Long) As Long
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As Long
Public Declare Function GdipDrawString Lib "gdiplus" (ByVal Graphics As Long, ByVal str As Long, ByVal Length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal Brush As Long) As Long
Public Declare Function GdipStringFormatGetGenericDefault Lib "gdiplus" (StringFormat As Long) As Long
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As Long
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, Brush As Long) As Long
Public Declare Function GdipDrawLines Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As POINTF, ByVal count As Long) As Long
Public Declare Function GdipFillEllipse Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
Public Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal Brush As Long, ByVal ARGB As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutW" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Public Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, vertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Public Declare Function SetDCPenColor Lib "gdi32" (ByVal hdc As Long, ByVal colorref As Long) As Long
Public Declare Function SetDCBrushColor Lib "gdi32" (ByVal hdc As Long, ByVal colorref As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Public Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As Any) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As Any) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameW" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As Any) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpsz As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub memset Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal numBytes As Long, ByVal Value As Byte)

Public Const HTCAPTION              As Long = 2
Public Const DT_CALCRECT            As Long = &H400
Public Const DT_CENTER              As Long = &H1
Public Const OPAQUE                 As Long = 2
Public Const TRANSPARENT            As Long = 1
Public Const DC_PEN                 As Long = 19
Public Const DC_BRUSH               As Long = 18
Public Const NULL_PEN               As Long = 8
Public Const NULL_BRUSH             As Long = 5
Public Const UnitPixel              As Long = 2
Public Const SmoothingModeAntiAlias As Long = 4
Public Const WrapModeTile           As Long = 0
Public Const DIB_RGB_COLORS         As Long = 0
Public Const GRADIENT_FILL_RECT_H   As Long = 0
Public Const GRADIENT_FILL_RECT_V   As Long = 1
Public Const WM_PAINT               As Long = &HF
Public Const WM_NCLBUTTONDOWN       As Long = &HA1
Public Const WM_SETICON             As Long = &H80
Public Const LB_GETITEMRECT         As Long = &H198
Public Const AB_32Bpp255            As Long = &H1FF0000
Public Const GWL_STYLE              As Long = &HFFFFFFF0
Public Const WS_BORDER              As Long = &H800000
Public Const SWP_FRAMECHANGED       As Long = &H20
Public Const SWP_NOMOVE             As Long = &H2
Public Const SWP_NOSIZE             As Long = &H1&
Public Const SWP_NOZORDER           As Long = &H4
Public Const ICON_SMALL             As Long = 0
Public Const ICON_BIG               As Long = 1
Public Const IMAGE_ICON             As Long = 1
Public Const SM_CXICON              As Long = 11
Public Const SM_CYICON              As Long = 12
Public Const SM_CXSMICON            As Long = 49
Public Const SM_CYSMICON            As Long = 50
Public Const LR_DEFAULTSIZE         As Long = &H40
Public Const LR_SHARED              As Long = &H8000&
Public Const SampleRate             As Long = 44100 ' Частота дискретизации

' // Set 32bpp icon to window
Public Sub SetIcon(ByVal hwnd As Long)
    Dim hIcon   As Long
    Dim cx      As Long
    Dim cy      As Long
    
    hIcon = LoadImage(App.hInstance, StrPtr("#101"), IMAGE_ICON, cx, cy, LR_SHARED)
    
    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
   
    SendMessage hwnd, WM_SETICON, ICON_BIG, ByVal hIcon
    
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    
    hIcon = LoadImage(App.hInstance, StrPtr("#101"), IMAGE_ICON, cx, cy, LR_SHARED)
    
    SendMessage hwnd, WM_SETICON, ICON_SMALL, ByVal hIcon
    
End Sub

' // Create POINTF variable
Public Function GdipPointF(ByVal x As Single, ByVal y As Single) As POINTF
    GdipPointF.x = x: GdipPointF.y = y
End Function

' // Create RECTF variable
Public Function GdipRectF(ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single) As RECTF
    GdipRectF.Left = Left:  GdipRectF.Top = Top: GdipRectF.Right = Right: GdipRectF.Bottom = Bottom
End Function

' // Create TRIVERTEX variable
Public Function GdiTrivertex(ByVal x As Long, ByVal y As Long, ByVal Color As Colors) As TRIVERTEX
    Dim col As COLOR32
    
    With GdiTrivertex
        .x = x
        .y = y
        col.Value = (Color And &HFF&) * &H100:    LSet .blue = col
        col.Value = Color And &HFF00&:            LSet .green = col
        col.Value = Color \ &H100 And &HFF00&:    LSet .red = col
        col.Value = Color \ &H10000 And &HFF00&:  LSet .alpha = col
    End With
    
End Function

' // Create GRADIENT_RECT variable
Public Function GdiGradientRect(ByVal ulIndex As Long, ByVal lrIndex As Long) As GRADIENT_RECT

    GdiGradientRect.lowerRight = ulIndex
    GdiGradientRect.upperLeft = lrIndex
    
End Function

' // Translate from ARGB to BGR
Public Function toVBColor(ByVal Color As Colors) As Long
    Dim c32 As COLOR32
    Dim cc  As RGBQUAD
    
    c32.Value = Color
    
    LSet cc = c32
    
    cc.rgbRed = cc.rgbRed Xor cc.rgbBlue
    cc.rgbBlue = (cc.rgbRed Xor cc.rgbBlue)
    cc.rgbRed = (cc.rgbRed Xor cc.rgbBlue)
    ' // Alpha multiplication
    cc.rgbRed = CLng(cc.rgbRed) * cc.rgbReserved \ &H100
    cc.rgbGreen = CLng(cc.rgbGreen) * cc.rgbReserved \ &H100
    cc.rgbBlue = CLng(cc.rgbBlue) * cc.rgbReserved \ &H100
    
    toVBColor = cc.rgbBlue Or (CLng(cc.rgbGreen) * &H100) Or (CLng(cc.rgbRed) * &H10000)
End Function



