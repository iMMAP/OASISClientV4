Attribute VB_Name = "GDIFx"
Public Enum PenStyles
    PS_SOLID = 0
    PS_DASH = 1
    PS_DOT = 2
    PS_DASHDOT = 3
    PS_DASHDOTDOT = 4
    PS_NULL = 5
    PS_INSIDEFRAME = 6
    PS_USERSTYLE = 7
    PS_ALTERNATE = 8
    PS_STYLE_MASK = &HF
    PS_ENDCAP_ROUND = &H0
    PS_ENDCAP_SQUARE = &H100
    PS_ENDCAP_FLAT = &H200
    PS_ENDCAP_MASK = &HF00
    PS_JOIN_ROUND = &H0
    PS_JOIN_BEVEL = &H1000
    PS_JOIN_MITER = &H2000
    PS_JOIN_MASK = &HF000
    PS_COSMETIC = &H0
    PS_GEOMETRIC = &H10000
    PS_TYPE_MASK = &HF0000
End Enum

Public Enum BrushStyles
    HS_HORIZONTAL = 0
    HS_VERTICAL = 1
    HS_FDIAGONAL = 2
    HS_BDIAGONAL = 3
    HS_CROSS = 4
    HS_DIAGCROSS = 5
    HS_BDIAGONAL1 = 7
End Enum

'System Color Brushs
Enum SysColBrush
    COLOR_BTNHIGHLIGHT = 20
    COLOR_BACKGROUND = 1
    COLOR_BTNHILIGHT = COLOR_BTNHIGHLIGHT
    COLOR_BTNSHADOW = 16
    COLOR_BTNTEXT = 18
    COLOR_CAPTIONTEXT = 9
    COLOR_DESKTOP = COLOR_BACKGROUND
    COLOR_BTNFACE = 15
    COLOR_ACTIVEBORDER = 10
    COLOR_3DDKSHADOW = 21
    COLOR_3DFACE = COLOR_BTNFACE
    COLOR_3DHIGHLIGHT = COLOR_BTNHIGHLIGHT
    COLOR_3DLIGHT = 22
    COLOR_3DSHADOW = COLOR_BTNSHADOW
    COLOR_ACTIVECAPTION = 2
    COLOR_APPWORKSPACE = 12
    COLOR_BLUE = 708
    COLOR_BLUEACCEL = 728
End Enum

Public Enum PolyFillMode
    Alternate = 1
    WINDING = 2
    BLACKBRUSH = 4
End Enum

Public Enum ArcDirection
    AD_CLOCKWISE = 2
    AD_COUNTERCLOCKWISE = 1
End Enum

Public Enum Bmp_Compression
    BI_RGB = 0&
    BI_RLE4 = 2&
    BI_RLE8 = 1&
End Enum

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

Public Enum GRADIENT_DIR
    Horizontal = &H0
    Vertical = &H1
End Enum

Public Enum giPaths
    pBeginPath = 0
    pEndPath = 1
    pFlattenPath = 2
    pAbortPath = 3
    pWidenPath = 4
    pStrokePath = 5
    pStrokeAndFillPath = 6
    pFillPath = 7
    pPathToRegion = 8
End Enum

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Public Type BITMAP_INFO
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Bmp_Compression
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
    bfType As Integer
    bfSize As Long
    bfOffBits As Long
End Type

'Public Type POINTAPI
'    x As Long
'    y As Long
'End Type
'
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RECT_AREA
    Width  As Long
    Height As Long
End Type

Public Type lRGB
    Red As Integer
    Blue As Integer
    Green As Integer
End Type

Public Type lHSV
    h As Integer
    s As Integer
    v As Integer
End Type

Public Const CLR_INVALID = &HFFFF
Public Const DC_PEN As Long = 19

Public Declare Function GDI_InflateRect Lib "user32.dll" Alias "InflateRect" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GDI_GetSysColorBrush Lib "user32.dll" Alias "GetSysColorBrush" (ByVal nIndex As SysColBrush) As Long
Public Declare Function GDI_DrawEdge Lib "user32.dll" Alias "DrawEdge" (ByVal hdc As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Public Declare Function GDI_GetDC Lib "user32.dll" Alias "GetDC" (ByVal Hwnd As Long) As Long
Public Declare Function GDI_GetPixel Lib "gdi32.dll" Alias "GetPixel1" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Private Declare Function BeginPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function FlattenPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function AbortPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function WidenPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function StrokePath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function StrokeAndFillPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function FillPath Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function PathToRegion Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long

Public Declare Function GDI_CreateCompatibleBitmap Lib "gdi32.dll" Alias "CreateCompatibleBitmap" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GDI_CreateCompatibleDC Lib "gdi32.dll" Alias "CreateCompatibleDC" (ByVal hdc As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Public Declare Function GDI_MoveToEx Lib "gdi32.dll" Alias "MoveToEx" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal Hwnd As Long, ByRef lpRect As RECT) As Long
Public Declare Function GDI_FillRect Lib "user32.dll" Alias "FillRect" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GDI_PtInRect Lib "user32" Alias "PtInRect" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function Ellipse Lib "gdi32.dll" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Chord Lib "gdi32.dll" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Pie Lib "gdi32.dll" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Polyline Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function ArcTo Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function SetArcDirection Lib "gdi32" (ByVal hdc As Long, ByVal ArcDirection As Long) As Long
Private Declare Function GetArcDirection Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long

Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function LoadBitmap Lib "user32.dll" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GDI_SetPixel Lib "gdi32" Alias "SetPixel" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GDI_SetPixelV Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GDI_CreateBitmap Lib "gdi32.dll" Alias "CreateBitmap" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long

Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public Declare Function GDI_StretchBlt Lib "gdi32.dll" Alias "StretchBlt" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As RasterOpConstants) As Long

Private Sub setTriVertexColor(tTV As TRIVERTEX, oColor As Long)
    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long
    
    lRed = (oColor And &HFF&) * &H100&
    lGreen = (oColor And &HFF00&)
    lBlue = (oColor And &HFF0000) \ &H100&
    
    setTriVertexColorComponent tTV.Red, lRed
    setTriVertexColorComponent tTV.Green, lGreen
    setTriVertexColorComponent tTV.Blue, lBlue
End Sub

Private Sub setTriVertexColorComponent(ByRef oColor As Integer, ByVal lComponent As Long)
    If (lComponent And &H8000&) = &H8000& Then
        oColor = (lComponent And &H7F00&)
        oColor = oColor Or &H8000
    Else
        oColor = lComponent
    End If
End Sub

Public Sub RGB2HSV(lpRGB As lRGB, lpHSV As lHSV)
Dim r As Double, g As Double, b As Double
Dim iMin As Double, iMax As Double
Dim h As Double, s As Double, v As Double, delta As Double

    r = lpRGB.Red / 255
    g = lpRGB.Green / 255
    b = lpRGB.Blue / 255
    
    iMin = Min(Min(r, g), b)
    iMax = Max(Max(r, g), b)
    
    v = iMax
    delta = (iMax - iMin)
    
    If (iMax = 0) Or (delta = 0) Then
        s = 0
        h = 0
    Else
        s = delta / iMax
        If (r = iMax) Then
            h = (g - b) / delta
        ElseIf (g = iMax) Then
            h = 2 + (b - r) / delta
        Else
            h = 4 + (r - g) / delta
        End If
        
        h = h * 60
        If (h < 0) Then
            h = h + 360
        End If
    End If
    
    lpHSV.h = (h / 360 * 255)
    lpHSV.s = (s * 255)
    lpHSV.v = (v * 255)
    
    r = 0: g = 0: b = 0
    h = 0: s = 0: v = 0
    delta = 0
    iMin = 0: iMax = 0
    
End Sub

Public Sub LongToRGB(lpLngColor As OLE_COLOR, lpRGB As lRGB)
Dim ColByte(2) As Byte
    CopyMemory ColByte(0), lpLngColor, Len(lpLngColor)
    
    With lpRGB
        .Red = ColByte(0)
        .Green = ColByte(1)
        .Blue = ColByte(2)
    End With
    
    Erase ColByte
End Sub

Function Min(ValA As Double, ValB As Double) As Double
    If (ValA < ValB) Then
        Min = ValA
    Else
        Min = ValB
    End If
End Function

Function Max(ValA As Double, ValB As Double) As Double
    If (ValA > ValB) Then
        Max = ValA
    Else
        Max = ValB
    End If
End Function

Function GDI_BitBlt(ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As RasterOpConstants)
    
    GDI_BitBlt = BitBlt(hDestDC, x, y, nWidth, nHeight, hSrcDC, xSrc, ySrc, dwRop)
    
End Function

Function GDI_CreateDC(Height As Long, Width As Long) As Long
Dim nDc As Long, hBmp As Long

    'Function to Crate a New Dc
    nDc = GDI_CreateCompatibleDC(GDI_GetDC(0))
    hBmp = GDI_CreateCompatibleBitmap(GDI_GetDC(0), Width, Height)
    DeleteObject SelectObject(nDc, hBmp)
    GDI_CreateDC = nDc
End Function

Function GDI_DeleteDC(ByVal hdc As Long) As Long
    GDI_DeleteDC = DeleteDC(hdc)
End Function

Public Sub GDI_LineTo(hdc As Long, x1 As Long, y1 As Long, X2 As Long, Y2 As Long, hPen As Long, Optional DrawEx As Boolean = True)
    DeleteObject SelectObject(hdc, hPen)   ' select the DC to draw onto
    If DrawEx Then
        If x1 >= 0 Then GDI_MoveToEx hdc, x1, y1, 0
        LineTo hdc, X2, Y2  'Draw the line
    Else
        LineTo hdc, x1, y1
    End If
End Sub

Private Function GDI_TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, GDI_TranslateColor) Then
        GDI_TranslateColor = &HFFFF&
    End If
End Function

Function GDI_CreatePen(ByVal nPenStyle As PenStyles, ByVal nWidth As Long, ByVal crColor As Long)
    GDI_CreatePen = CreatePen(nPenStyle, nWidth, GDI_TranslateColor(crColor))
End Function

Function GDI_GetWndRECT(Hwnd As Long, lRect As RECT) As RECT
    'Get a windows Rect
    GetClientRect Hwnd, lRect
End Function

Function GDI_GetBitmapInfo(lFilename As String, bi As BITMAP_INFO) As Long
Dim iFile As Long
Dim bmp_i As BITMAPINFOHEADER
Dim bmp_f As BITMAPFILEHEADER

    If Not LenB(Dir(lFilename)) <> 0 Or Len(lFilename) = 0 Then
        Exit Function
    End If
    
    iFile = FreeFile
    Open lFilename For Binary As #iFile
        If LOF(iFile) = 0 Then
            Close #iFile
            Exit Function
        Else
            Get #iFile, , bmp_f
            Get #iFile, , bmp_i
        End If
    Close #iFile
    
    With bi
        'Bitmap Fileinfo
        .bfOffBits = bmp_f.bfOffBits
        .bfSize = bmp_f.bfSize
        .bfType = bmp_f.bfType
        'Bitmap info
        .biBitCount = bmp_i.biBitCount
        .biClrUsed = bmp_i.biClrUsed
        .biClrImportant = bmp_i.biClrImportant
        .biCompression = bmp_i.biCompression
        .biHeight = bmp_i.biHeight
        .biPlanes = bmp_i.biPlanes
        .biSize = bmp_i.biSize
        .biSizeImage = bmp_i.biSizeImage
        .biWidth = bmp_i.biWidth
        .biXPelsPerMeter = bmp_i.biXPelsPerMeter
        .biYPelsPerMeter = bmp_i.biYPelsPerMeter
    End With
    
    ZeroMemory bmp_f, Len(bmp_f)
    ZeroMemory bmp_f, Len(bmp_i)
    
    GDI_GetBitmapInfo = 1
    
End Function

Function GDI_GetArcDirection(hdc As Long) As ArcDirection
    GDI_GetArcDirection = GetArcDirection(hdc)
End Function

Public Sub GDI_SetArcDirection(hdc As Long, lArcDirection As ArcDirection)
    SetArcDirection hdc, lArcDirection
End Sub

Public Sub GDI_RectToArea(lRect As RECT, lAreaRect As RECT_AREA)
    lAreaRect.Width = (lRect.Right - lRect.Left) 'Width
    lAreaRect.Height = (lRect.Bottom - lRect.Top) 'Height
End Sub

Public Sub GDI_AreaToRect(Width As Long, Height As Long, lRect As RECT)
    lRect.Left = 0
    lRect.Top = 0
    lRect.Right = Width
    lRect.Bottom = Height
End Sub

Function GDI_CreateSoildBrush(bColor As OLE_COLOR) As Long
    GDI_CreateSoildBrush = CreateSolidBrush(GDI_TranslateColor(bColor))
End Function

Function GDI_CreatePatternBrush(ByVal hBitmap As Long) As Long
   GDI_CreatePatternBrush = CreatePatternBrush(hBitmap)
End Function

Function GDI_CreateEllipticRgn(ByVal x1 As Long, ByVal y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long
    GDI_CreateEllipticRgn = CreateEllipticRgn(x1, y1, X2, Y2)
End Function

Function GDI_CreatePolygonRgn(PolyPoints() As POINTAPI, FillMode As PolyFillMode)
    GDI_CreatePolygonRgn = CreatePolygonRgn(PolyPoints(0), UBound(PolyPoints) + 1, FillMode)
'lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long
End Function

Function GDI_FillRgn(ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
    GDI_FillRgn = FillRgn(hdc, hRgn, hBrush)
End Function

Function GDI_FrameRect(ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long)
    FrameRect hdc, lpRect, hBrush
End Function

Function GDI_CLS(hdc As Long, Optional Width As Long, _
    Optional Height As Long, Optional EraseCol As OLE_COLOR = vbBlack)
    
    Dim mRect As RECT
    
    GDI_AreaToRect Width, Height, mRect
    GDI_FillRect hdc, mRect, GDI_CreateSoildBrush(EraseCol)
End Function

Function GDI_ArcTo(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, _
    ByVal X4 As Long, ByVal Y4 As Long) As Long
    
    GDI_ArcTo = ArcTo(hdc, x1, y1, X2, Y2, X3, Y3, X4, Y4)
    
End Function
Function GDI_Arc(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, _
    ByVal X4 As Long, ByVal Y4 As Long) As Long
    GDI_Arc = Arc(hdc, x1, y1, X2, Y2, X3, Y3, X4, Y4)
End Function

Function GDI_Rectangle(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long
    GDI_Rectangle = Rectangle(hdc, x1, y1, X2, Y2)
End Function

Function GDI_RoundRect(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    GDI_RoundRect = RoundRect(hdc, x1, y1, X2, Y2, X3, Y3)
End Function

Function GDI_Ellipse(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long
    GDI_Ellipse = Ellipse(hdc, x1, y1, X2, Y2)
End Function

Function GDI_Chord(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, _
    ByVal X4 As Long, ByVal Y4 As Long)
    GDI_Chord = Chord(hdc, x1, y1, X2, Y2 _
    , X3, Y3, X4, Y4)
End Function

Function GDI_Pie(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, _
    ByVal X4 As Long, ByVal Y4 As Long) As Long
    GDI_Pie = Pie(hdc, x1, y1, X2, Y2, X3, Y3, X4, Y4)
End Function

Public Sub GDI_Polygon(hdc As Long, PolyPoints() As POINTAPI)
    Polygon hdc, PolyPoints(0), UBound(PolyPoints) + 1
End Sub

Public Sub GDI_Polyline(hdc As Long, PolyPoints() As POINTAPI)
    Polyline hdc, PolyPoints(0), UBound(PolyPoints) + 1
End Sub

Public Function GDI_SelectObject(hdc As Long, hObject As Long) As Long
    GDI_SelectObject = SelectObject(hdc, hObject)
End Function

Function GDI_TextOut(hdc As Long, x As Long, y As Long, Optional lpString As String = "")
    GDI_TextOut = TextOut(hdc, x, y, lpString, Len(lpString))
End Function

Public Sub GDI_LoadImageToDC(hdc As Long, lFile As String)
    If Not LenB(Dir(lFile)) <> 0 Or Len(lFile) = 0 Then
        Exit Sub
    End If
    
    If hdc = 0 Then Exit Sub
    DeleteObject SelectObject(hdc, LoadPicture(lFile))
End Sub

Function GDI_CopyRect(lpDestRect As RECT, lpSourceRect As RECT)
    GDI_CopyRect = CopyRect(lpDestRect, lpSourceRect)
End Function

Function GDI_IsRectEmpty(lpRect As RECT) As Long
    GDI_IsRectEmpty = IsRectEmpty(lpRect)
End Function

Sub GDI_SetRectZero(lpRect As RECT)
    SetRectEmpty lpRect
End Sub

Function GDI_SetTextColor(ByVal hdc As Long, ByVal crColor As OLE_COLOR)
    SetTextColor hdc, GDI_TranslateColor(crColor)
End Function

Sub GDI_DeleteObj(hObj As Long)
    DeleteObject hObj
End Sub

Sub GDI_CopyPixel(hdc As Long, x1 As Long, y1 As Long, X2 As Long, Y2 As Long)
    'SetPixel hdc, x1, y1, GetPixel1(hdc, X2, Y2)
End Sub

Function GDI_GetStockObject(ByVal nIndex As Long) As Long
    GDI_GetStockObject = GetStockObject(nIndex)
End Function

Function GDI_Paths(hdc As Long, vPaths As giPaths) As Long
    Select Case vPaths
        Case pBeginPath
            'Begin Path
            GDI_Paths = BeginPath(hdc)
        Case pEndPath
            'End Path
            GDI_Paths = EndPath(hdc)
        Case pFlattenPath
            'FlattenPath
            GDI_Paths = FlattenPath(hdc)
        Case pAbortPath
            'Abort Path
            GDI_Paths = AbortPath(hdc)
        Case pWidenPath
            'WidenPath
            GDI_Paths = WidenPath(hdc)
        Case pStrokePath
            'StrokePath
            GDI_Paths = StrokePath(hdc)
        Case pStrokeAndFillPath
            'StrokeAndFillPath
            GDI_Paths = StrokeAndFillPath(hdc)
        Case pFillPath
            'FillPath
            GDI_Paths = FillPath(hdc)
        Case pPathToRegion
            GDI_Paths = PathToRegion(hdc)
        Case Else
            GDI_Paths = 0
    End Select
End Function

Function GDI_LoadBitmap(ByVal hInstance As Long, ByVal lpBitmapName As String)
    GDI_LoadBitmap = LoadBitmap(hInstance, lpBitmapName)
End Function

Public Sub GDI_GradientFill(hdc As Long, mRect As RECT, mStartColor As OLE_COLOR, mEndColor As OLE_COLOR, gDir As GRADIENT_DIR)
Dim gRect As GRADIENT_RECT
Dim tTV(0 To 1) As TRIVERTEX
    'Function used to paint a Gradient effect on the listbox
    
    setTriVertexColor tTV(1), GDI_TranslateColor(mEndColor)
    tTV(0).x = mRect.Left
    tTV(0).y = mRect.Top
    
    setTriVertexColor tTV(0), GDI_TranslateColor(mStartColor)
    tTV(1).x = mRect.Right
    tTV(1).y = mRect.Bottom
    
    gRect.UpperLeft = 0
    gRect.LowerRight = 1
    
    GradientFill hdc, tTV(0), 2, gRect, 1, gDir
    
End Sub

Function GDI_CreateHatchBrush(ByVal bStyle As BrushStyles, ByVal crColor As OLE_COLOR)
    GDI_CreateHatchBrush = CreateHatchBrush(bStyle, crColor)
End Function
