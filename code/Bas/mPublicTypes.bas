Attribute VB_Name = "mPublicTypes"
Option Explicit

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Public Const LF_FACESIZE = 32

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public Const LB_FINDSTRING As Long = &H18F
Public Const LB_FINDSTRINGEXACT As Long = &H1A2
Public Const CB_ERR As Long = (-1)
Public Const LB_ERR As Long = (-1)
Public Const WM_USER As Long = &H400
Public Const CB_FINDSTRING As Long = &H14C
Public Const CB_SHOWDROPDOWN As Long = &H14F


Public Declare Function SendMessageStr Lib _
    "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long


'Used to move around by caption
Public Declare Function ReleaseCapture Lib "user32" () As Long

'Used to draw the ellipse on the form
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, _
    ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
    ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'Used to create the regiod around the form to shape it
Public Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long, ByVal EllipseWidth As Long, _
    ByVal EllipseHeight As Long) As Long

Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
        lpPoint As POINTAPI) As Long 'Also used for getting positions of
                                     'objects/forms we want to place the
                                     'balloons by
Public mlWidth As Long
Public mlHeight As Long

Public Type BalloonCoords 'Used to store X and Y coordinates of balloon
    x As Long 'after using API and math operations to figure exact
    y As Long 'coordinates regarding where to place itself
End Type

Public Sub FindIndexStrEx(ctlSource As Control, _
    ByVal Str As String)
    
    Dim lngIdx As Long


    If TypeName(ctlSource) = "ComboBox" Then
        lngIdx = SendMessageStr(ctlSource.hwnd, _
        CB_FINDSTRING, -1, Str)
    ElseIf TypeName(ctlSource) = "ListBox" Then
        lngIdx = SendMessageStr(ctlSource.hwnd, _
        LB_FINDSTRING, -1, Str)
    Else
        Exit Sub
    End If

    

    If lngIdx <> -1 Then
        ctlSource.ListIndex = lngIdx
    End If

End Sub


Public Sub FindIndexStr(ctlSource As Control, _
    ByVal Str As String, intKey As Integer, _
    Optional ctlTarget As Variant)
    
    Dim lngIdx As Long
    Dim FindString As String
    If (intKey < 32 Or intKey > 127) And _
    (Not (intKey = 13 Or intKey = 8)) Then Exit Sub


    If Not intKey = 13 Or intKey = 8 Then


        If Len(ctlSource.Text) = 0 Then
            FindString = Str & Chr$(intKey)
        Else
            FindString = Left$(Str, ctlSource.SelStart) & Chr$(intKey)
        End If

    End If



    If intKey = 8 Then
        If Len(ctlSource.Text) = 0 Then Exit Sub
        Dim numChars As Integer
        numChars = ctlSource.SelStart - 1
        'FindString = Left(str, numChars)
        If numChars > 0 Then FindString = Left(Str, numChars)
    End If



    If IsMissing(ctlTarget) And TypeName(ctlSource) = "ComboBox" Then
        Set ctlTarget = ctlSource


        If intKey = 13 Then
            Call SendMessageStr(ctlTarget.hwnd, _
            CB_SHOWDROPDOWN, True, 0&)
            Exit Sub
        End If

        lngIdx = SendMessageStr(ctlTarget.hwnd, _
        CB_FINDSTRING, -1, FindString)
    ElseIf TypeName(ctlTarget) = "ListBox" Then
        If intKey = 13 Then Exit Sub '???
        lngIdx = SendMessageStr(ctlTarget.hwnd, _
        LB_FINDSTRING, -1, FindString)
    Else
        Exit Sub
    End If

    

    If lngIdx <> -1 Then
        ctlTarget.ListIndex = lngIdx
        If TypeName(ctlSource) = "TextBox" Then ctlSource.Text = ctlTarget.List(lngIdx)
        ctlSource.SelStart = Len(FindString)
        ctlSource.SelLength = Len(ctlSource.Text) - ctlSource.SelStart
    End If

    intKey = 0
End Sub

