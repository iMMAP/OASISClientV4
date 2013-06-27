Attribute VB_Name = "ColorFunctions"
Option Explicit


Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function InvertRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
    
    Private Const PS_SOLID = 0
    'This function will brighten or darken a
    '     color
    'Example: Picture1.BackColor = AdjustBri
    '     ghtness(Picture1.BackColor, -50)
    
    'Bitmap file format structures
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
Type BITMAPINFOHEADER
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
Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type

Public Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long, ByVal RasterOp As Long) As Long


Public Function AdjustBrightness(ByVal Color As Long, ByVal Amount As Single) As Long
    On Error Resume Next
    
    Dim r(1) As Integer, g(1) As Integer, b(1) As Integer
    
    'get red, green, and blue values
    GetRGB r(0), g(0), b(0), Color
    
    'add/subtract the amount to/from the ori
    '     ginal RGB values
    r(1) = SetBound(r(0) + Amount, 0, 255)
    g(1) = SetBound(g(0) + Amount, 0, 255)
    b(1) = SetBound(b(0) + Amount, 0, 255)
    
    'convert RGB back to Long value
    AdjustBrightness = RGB(r(1), g(1), b(1))
End Function
'This function will blend two colors tog
'     ether at any percentage 0 - 100
'Example: Picture1.BackColor = BlendColo
'     rs(vbRed, vbBlue, 50)


Public Function BlendColors(ByVal Color1 As Long, ByVal Color2 As Long, ByVal Percentage As Single) As Long
    On Error Resume Next
    
    Dim r(2) As Integer, g(2) As Integer, b(2) As Integer
    Dim fPercentage(2) As Single
    Dim DAmt(2) As Single
    
    'make sure Percentage is between 0 and 1
    '     00
    Percentage = SetBound(Percentage, 0, 100)
    
    'extract the RGB values from Color1 and
    '     Color2
    GetRGB r(0), g(0), b(0), Color1
    GetRGB r(1), g(1), b(1), Color2
    
    '1st part: get the positive or negative
    '     amount between the 2 colors
    '2nd part: calculate how much needs to b
    '     e added to Color1
    '(Difference divided by 100 multiplied b
    '     y the percentage)
    DAmt(0) = r(1) - r(0): fPercentage(0) = (DAmt(0) / 100) * Percentage
    DAmt(1) = g(1) - g(0): fPercentage(1) = (DAmt(1) / 100) * Percentage
    DAmt(2) = b(1) - b(0): fPercentage(2) = (DAmt(2) / 100) * Percentage
    
    'add/subtract each percentage to RGB val
    '     ues
    r(2) = r(0) + fPercentage(0)
    g(2) = g(0) + fPercentage(1)
    b(2) = b(0) + fPercentage(2)
    
    'convert RGB back to Long value
    BlendColors = RGB(r(2), g(2), b(2))
End Function
'This will draw Verticle/Horizontal grad
'     ient very quickly
'Example: DrawGradient Picture1.hDC, 0,
'     0, Picture1.ScaleWidth, Picture1.ScaleHe
'     ight, vbRed, vbBlue, True
'(note: if the picture1 is set to autore
'     draw it must be refreshed after this fun
'     ction)


Public Sub DrawGradient(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color1 As Long, ByVal Color2 As Long, Optional Horizontal As Boolean = False)
    Dim i As Long
    Dim CCol As Long
    Dim StepAmt As Single
    
    'if its a horizontal gradient use the X
    '     values, if not use Y values


    If Horizontal = False Then
        'Get the total lines between Y1 and Y2


        If Y2 > y1 Then
            StepAmt = 100 / (Y2 - y1)
        Else
            StepAmt = 100 / (y1 - Y2)
        End If
        'Draw a Line from X1 to X2 blending the
        '     Colors as Y increases


        For i = y1 To Y2
            DrawLine hdc, x1, i, X2, i, BlendColors(Color1, Color2, i * StepAmt)
        Next i
    Else
        'Get the total lines between X1 and X2


        If X2 > x1 Then
            StepAmt = 100 / (X2 - x1)
        Else
            StepAmt = 100 / (x1 - X2)
        End If
        'Draw a line from Y1 to Y2 blending the
        '     Colors as X increases


        For i = x1 To X2
            DrawLine hdc, i, y1, i, Y2, BlendColors(Color1, Color2, i * StepAmt)
        Next i
    End If
End Sub
'Fast way to draw a line (VB's built in
'     Line function is too slow)
'Example: DrawLine Picture1.hDC, 0, 0, 5
'     0, 50, vbRed


Public Sub DrawLine(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    Dim pt As POINTAPI
    Dim oldPen As Long, hPen As Long
    
    'Create a solid brush (to draw a solid l
    '     ine) with the color
    hPen = CreatePen(PS_SOLID, 1, Color)
    'Bind the brush to the control's DC
    oldPen = SelectObject(hdc, hPen)
    
    'Set the start point for the line
    MoveToEx hdc, x1, y1, pt
    'Set the end point for the line and draw
    '     it
    LineTo hdc, X2, Y2
    
    'delete the brush from memory


    SelectObject hdc, oldPen
        DeleteObject hPen
    End Sub
'Extract the red, green, and blue values
'     from a color
'Example: GetRGB R, G, B, vbMagenta


Public Sub GetRGB(r As Integer, g As Integer, b As Integer, ByVal Color As Long)
    Dim TempValue As Long
    
    'First translate the color from a long v
    '     alue to a short value
    TranslateColor Color, 0, TempValue
    
    'Calculate the red, green, and blue valu
    '     es from the short value
    r = TempValue And &HFF&
    g = (TempValue And &HFF00&) / 2 ^ 8
    b = (TempValue And &HFF0000) / 2 ^ 16
End Sub
'Invert colors (Negative image)
'Example: InvertColor Picture1.hDC, 0, 0
'     , Picture1.ScaleWidth, Picture1.ScaleHei
'     ght


Public Function InvertColor(ByVal hdc As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal X2 As Single, ByVal Y2 As Single)
    Dim hRect As RECT
    
    'Set the RECT shape to match X1, Y1, X2,
    '     and Y2 values
    SetRect hRect, x1, y1, X2, Y2
    
    'Use quick API function to Invert the co
    '     lors
    InvertRect hdc, hRect
End Function
'This ensures a variable is between 2 va
'     lues
'This is to support the functions above


Private Function SetBound(ByVal Num As Single, ByVal MinNum As Single, ByVal MaxNum As Single) As Single


    If Num < MinNum Then
        'if less that min value make it the min
        '     value
        SetBound = MinNum
    ElseIf Num > MaxNum Then
        'if more than max value make it the max
        '     value
        SetBound = MaxNum
    Else
        'if between min and max then leave it al
        '     one
        SetBound = Num
    End If
End Function

' Restrict the form to its "transparent" pixels.
Public Sub TrimPicture(ByVal pic As PictureBox, ByVal transparent_color As Long)
'Const RGN_OR = 2
'Dim bitmap_info As BITMAPINFO
'Dim pixels() As Byte
'Dim bytes_per_scanLine As Integer
'Dim pad_per_scanLine As Integer
'Dim transparent_r As Byte
'Dim transparent_g As Byte
'Dim transparent_b As Byte
'Dim wid As Integer
'Dim hgt As Integer
'Dim x As Integer
'Dim y As Integer
'Dim start_x As Integer
'Dim stop_x As Integer
'Dim combined_rgn As Long
'Dim new_rgn As Long
'
'    ' Prepare the bitmap description.
'    With bitmap_info.bmiHeader
'        .biSize = 40
'        .biWidth = picShape.ScaleWidth
'        ' Use negative height to scan top-down.
'        .biHeight = -picShape.ScaleHeight
'        .biPlanes = 1
'        .biBitCount = 32
'        .biCompression = 0 'BI_RGB
'        bytes_per_scanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
'        pad_per_scanLine = bytes_per_scanLine - (((.biWidth * .biBitCount) + 7) \ 8)
'        .biSizeImage = bytes_per_scanLine * Abs(.biHeight)
'    End With
'
'    ' Load the bitmap's data.
'    wid = picShape.ScaleWidth
'    hgt = picShape.ScaleHeight
'    ReDim pixels(1 To 4, 0 To wid - 1, 0 To hgt - 1)
'    GetDIBits picShape.hDC, picShape.Image, _
'        0, picShape.ScaleHeight, pixels(1, 0, 0), _
'        bitmap_info, 0 'DIB_RGB_COLORS
'
'    ' Break the transparent color into its components.
'    UnRGB transparent_color, transparent_r, transparent_g, transparent_b
'
'    ' Create the PictureBox's regions.
'    For y = 0 To hgt - 1
'        ' Create a region for this row.
'        x = 1
'        Do While x < wid
'            start_x = 0
'            stop_x = 0
'
'            ' Find the next non-transparent column.
'            Do While x < wid
'                If pixels(pixR, x, y) <> transparent_r Or _
'                   pixels(pixG, x, y) <> transparent_g Or _
'                   pixels(pixB, x, y) <> transparent_b _
'                Then
'                    Exit Do
'                End If
'                x = x + 1
'            Loop
'            start_x = x
'
'            ' Find the next transparent column.
'            Do While x < wid
'                If pixels(pixR, x, y) = transparent_r And _
'                   pixels(pixG, x, y) = transparent_g And _
'                   pixels(pixB, x, y) = transparent_b _
'                Then
'                    Exit Do
'                End If
'                x = x + 1
'            Loop
'            stop_x = x
'
'            ' Make a region from start_x to stop_x.
'            If start_x < wid Then
'                If stop_x >= wid Then stop_x = wid - 1
'
'                ' Create the region.
'                new_rgn = CreateRectRgn( _
'                    start_x, y, stop_x, y + 1)
'
'                ' Add it to what we have so far.
'                If combined_rgn = 0 Then
'                    combined_rgn = new_rgn
'                Else
'                    CombineRgn combined_rgn, _
'                        combined_rgn, new_rgn, RGN_OR
'                    DeleteObject new_rgn
'                End If
'            End If
'        Loop
'    Next y
'
'    ' Restrict the PictureBox to the region.
'    SetWindowRgn pic.hwnd, combined_rgn, True
'    DeleteObject combined_rgn
End Sub

Public Sub ShadeForm(frm As Form)

    Dim iLoop As Integer
    Dim NumberOfRects As Integer
    Dim GradColor As Long
    Dim GradValue As Integer
    frm.ScaleMode = 3
    frm.DrawStyle = 6
    frm.DrawWidth = 2
    frm.AutoRedraw = True
    NumberOfRects = 64
    
    For iLoop = 1 To 64
        GradValue = 255 - (iLoop * 4 - 1)
        
        GradColor = RGB(GradValue, GradValue, GradValue)
        frm.Line (0, frm.ScaleHeight * (iLoop - 1) / 64)-(frm.ScaleWidth, frm.ScaleHeight * iLoop / 64), GradColor, BF
        
    Next iLoop

    frm.Refresh
End Sub
