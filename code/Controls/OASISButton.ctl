VERSION 5.00
Begin VB.UserControl OASISButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1365
   ScaleHeight     =   315
   ScaleWidth      =   1365
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   630
      ScaleHeight     =   555
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "OASISButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'api calls
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


'events
Event Click()
Event DblClick()
Event MouseEnter()
Event MouseExit()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDownOnDropdown()

'types
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'enumerations
Public Enum enBs
  bsRegular
  bsMenuDropdown
End Enum

Public Enum enTextAlign
    taLeft
    taMiddle
    taRight
End Enum

Public Enum enIf
   ifWYSWYG
   ifDistort
End Enum

'Default Property Values:
Const m_def_ImageLeft = 3
Const m_def_ImageTop = 3
Const m_def_ImageWidth = 16
Const m_def_ImageHeight = 16
Const m_def_TextColorOnMouseover = &HFF8080
Const m_def_ImageFormat = 0
Const m_def_Multiline = False
Const m_def_TextAlign = 1
Const m_def_TextColor = vbBlue
Const m_def_Text = "&caption"
Const m_def_ButtonStyle = 0

'Property Variables:
Dim m_ImageLeft As Long
Dim m_ImageTop As Long
Dim m_ImageWidth As Long
Dim m_ImageHeight As Long
Dim m_TextColorOnMouseover As OLE_COLOR
Dim m_ImageFormat As enIf
Dim m_Multiline As Boolean
Dim m_TextAlign As enTextAlign
Dim m_TextColor As OLE_COLOR
Dim m_Text As String
Dim m_ButtonStyle As enBs

'variables
Dim mR1                 As RECT
Dim mR2                 As RECT
Dim mR3                 As RECT
Dim mTxtRectNotFocused1 As RECT
Dim mTxtRectNotFocused2 As RECT
Dim mTxtRectFocused1    As RECT
Dim mTxtRectFocused2    As RECT
Dim mFocusBoxRect1      As RECT
Dim mFocusBoxRect2      As RECT
Dim mbFocused           As Boolean
Dim mbMouseIsDown       As Boolean
Dim mbMouseEntered      As Boolean


Private Sub subSetRects()

Dim x1  As Long
Dim y1  As Long
Dim x2  As Long
Dim y2  As Long
  
  With UserControl
    'this rect is used if
    '[m_ButtonStyle]= bsRegular
    x1 = 0
    y1 = 0
    x2 = (.Width \ Screen.TwipsPerPixelX)
    y2 = (.Height \ Screen.TwipsPerPixelY)
    SetRect mR1, x1, y1, x2, y2
  
    'these rects is used if
    '[m_ButtonStyle]= bsMenuDropdown
    'first, the left side
    x1 = 0
    y1 = 0
    x2 = ((.Width - 200) \ Screen.TwipsPerPixelX)
    y2 = (.Height \ Screen.TwipsPerPixelY)
    SetRect mR2, x1, y1, x2, y2

    'now, the right side
    x1 = ((.Width - 200) \ Screen.TwipsPerPixelX)
    y1 = 0
    x2 = (.Width \ Screen.TwipsPerPixelX)
    y2 = (.Height \ Screen.TwipsPerPixelY)
    SetRect mR3, x1, y1, x2, y2
    
    
    'define the rects for drawing text on the
    'button when the control DOESNT have focus
    '----------------------------------------
    
    'If [m_ButtonStyle] = bsRegular
    CopyRect mTxtRectNotFocused1, mR1
    InflateRect mTxtRectNotFocused1, -2, -2
    
    'If [m_ButtonStyle]=bsMenuDropdown
    CopyRect mTxtRectNotFocused2, mR2
    InflateRect mTxtRectNotFocused2, -2, -2


    'define the rects for drawing text on the
    'button when the control DOES have focus
    '----------------------------------------
    
    'If [m_ButtonStyle] = bsRegular
    CopyRect mTxtRectFocused1, mR1
    InflateRect mTxtRectFocused1, -4, -4
    
    'If [m_ButtonStyle]=bsMenuDropdown
    CopyRect mTxtRectFocused2, mR2
    InflateRect mTxtRectFocused2, -4, -4
    
    
    'define rects for drawing focus rect
    '----------------------------------------
    
    'If [m_ButtonStyle] = bsRegular
    CopyRect mFocusBoxRect1, mR1
    InflateRect mFocusBoxRect1, -3, -3
    
    'If [m_ButtonStyle]=bsMenuDropdown
    CopyRect mFocusBoxRect2, mR2
    InflateRect mFocusBoxRect2, -3, -3
  End With
  
End Sub

Private Sub subPaintButton(buttonIsUp As Boolean)
Dim CNST               As Long
Dim pt                 As POINTAPI
Const BF_TOP           As Long = &H2
Const BF_RIGHT         As Long = &H4
Const BF_BOTTOM        As Long = &H8
Const BF_LEFT          As Long = &H1
Const BF_RECT          As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Const BF_FLAT          As Long = &H4000
Const BDR_SUNKEN       As Long = &HA
Const BDR_RAISED       As Long = &H5

'clear previous graphics images
UserControl.Cls
 
If m_ButtonStyle = bsRegular Then
  If buttonIsUp Then
    'draws button in up state
    DrawEdge hdc, mR1, BDR_RAISED, BF_RECT
  Else
    'draws button in down state
    DrawEdge hdc, mR1, BDR_RAISED, BF_RECT Or BF_FLAT
  End If
 
ElseIf m_ButtonStyle = bsMenuDropdown Then
  If buttonIsUp Then
     'draw the left side up
     DrawEdge hdc, mR2, BDR_RAISED, BF_RECT
     'draw the right side up
     DrawEdge hdc, mR3, BDR_RAISED, BF_RECT
  Else
    'get where the cursor is now and convert
    'the coodinates, which are relative to
    'upper left corner of screen, and convert
    'to the upper left corner of this control
    GetCursorPos pt
    ScreenToClient hWnd, pt
   'if the mouse is down on left part...
   If PtInRect(mR2, pt.X, pt.Y) Then
      'draw left side down
      DrawEdge hdc, mR2, BDR_RAISED, BF_RECT Or BF_FLAT
      'draw right side up
      DrawEdge hdc, mR3, BDR_RAISED, BF_RECT
   'if the mouse is down on right part...
   ElseIf PtInRect(mR3, pt.X, pt.Y) Then
      'draw left side up
      DrawEdge hdc, mR2, BDR_RAISED, BF_RECT
      'right side down
      DrawEdge hdc, mR3, BDR_RAISED, BF_RECT Or BF_FLAT
   End If
  End If
End If
  
subDrawImage

End Sub
 

Private Sub subDrawImage()
Const SRCCOPY As Long = &HCC0020

With Picture1
  If .Picture <> 0 Then
    'paint the image the exact height and width
    'of the image of picture1
    If m_ImageFormat = ifWYSWYG Then
      StretchBlt hdc, m_ImageLeft, m_ImageTop, _
              (.Width \ Screen.TwipsPerPixelX), _
              (.Height \ Screen.TwipsPerPixelY), _
              .hdc, 0, 0, (.Width \ Screen.TwipsPerPixelX), _
              (.Height \ Screen.TwipsPerPixelY), SRCCOPY
    ElseIf m_ImageFormat = ifDistort Then
      StretchBlt hdc, m_ImageLeft, m_ImageTop, _
               m_ImageWidth, m_ImageHeight, _
              .hdc, 0, 0, (.Width \ Screen.TwipsPerPixelX), _
              (.Height \ Screen.TwipsPerPixelY), SRCCOPY
    End If
  End If
End With

subDrawText

End Sub

Private Sub subDrawText()
Const DT_CENTER      As Long = &H1
Const DT_LEFT        As Long = &H0
Const DT_RIGHT       As Long = &H2
Const DT_MULTILINE   As Long = (&H1)
Const DT_WORDBREAK   As Long = &H10
Const DT_VCENTER     As Long = &H4
Const DT_SINGLELINE  As Long = &H20
Dim DT_VAL1          As Long
Dim DT_VAL2          As Long
Dim txtRect          As RECT
  
 If m_Multiline Then
    DT_VAL2 = DT_WORDBREAK Or DT_MULTILINE
 Else
    DT_VAL2 = DT_VCENTER Or DT_SINGLELINE
 End If
 
 'set the texts alignment
 If m_TextAlign = taLeft Then
   DT_VAL1 = DT_LEFT Or DT_VAL2
 ElseIf m_TextAlign = taMiddle Then
   DT_VAL1 = DT_CENTER Or DT_VAL2
 ElseIf m_TextAlign = taRight Then
   DT_VAL1 = DT_RIGHT Or DT_VAL2
 End If
 
 If mbMouseEntered Then
    'set the text color to the val of [m_TextColorOnMouseover]
    SetTextColor hdc, m_TextColorOnMouseover
 Else
    'set the text color to the val of [m_TextColor]
    SetTextColor hdc, m_TextColor
 End If
 
 'draw the text on the button
 If m_ButtonStyle = bsRegular Then
   'draw the focus rect if we have focus
   If mbFocused Then
     DrawText hdc, m_Text, Len(m_Text), mTxtRectFocused1, DT_VAL1
     'set the forecolor back to black for focus rect
     SetTextColor hdc, 0&
     DrawFocusRect hdc, mFocusBoxRect1
   Else
     DrawText hdc, m_Text, Len(m_Text), mTxtRectNotFocused1, DT_VAL1
   End If
 ElseIf m_ButtonStyle = bsMenuDropdown Then
    'draw the focus rect if we have focus
   If mbFocused Then
     DrawText hdc, m_Text, Len(m_Text), mTxtRectFocused2, DT_VAL1
     'set the forecolor back to black for focus rect
     SetTextColor hdc, 0&
     DrawFocusRect hdc, mFocusBoxRect2
   Else
     DrawText hdc, m_Text, Len(m_Text), mTxtRectNotFocused2, DT_VAL1
   End If
   
   subDrawArrow
 End If
 
End Sub

Private Sub subDrawArrow()
Dim loldscale As Long
Dim lcentX    As Long
Dim lcentY    As Long

With UserControl
  'store the original scalemode
  loldscale = .ScaleMode
  'change to pixels
  .ScaleMode = vbPixels
  'get the center point of the right button
  'assuming the buttonstyle=[bsMenuDropdown]
  lcentX = ((.Width - 100) \ Screen.TwipsPerPixelX)
  lcentY = (.ScaleHeight * 0.5)
  'now draw lines in a way that we create arrow
  UserControl.Line ((lcentX - 3), lcentY)-((lcentX + 3), lcentY), vbBlack
  UserControl.Line ((lcentX - 2), (lcentY + 1))-((lcentX + 2), (lcentY + 1)), vbBlack
  UserControl.Line ((lcentX - 1), (lcentY + 2))-((lcentX + 1), (lcentY + 2)), vbBlack
  'restore the original scalemode
  .ScaleMode = loldscale
End With

End Sub
 

Private Sub subFindHotKey(stringToInspect As String)
Dim lLen  As Long
Dim i     As Integer
Dim schr  As String
Dim schr2 As String

'inspect each character looking for the "&"
lLen = Len(stringToInspect)
For i = 1 To lLen
  schr = Mid$(stringToInspect, i, 1)
  'if we find a "&", then check to make sure
  'the next character isnt also a "&"
  If schr = "&" Then
    schr2 = Mid$(stringToInspect, (i + 1), 1)
    If schr2 <> "&" Then
       UserControl.AccessKeys = schr2
       Exit For
    End If
  End If
Next i

End Sub

 

 

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
   mbFocused = True
  
   If mbMouseIsDown Then
      subPaintButton False
   Else
      subPaintButton True
   End If
End Sub
 
Private Sub UserControl_LostFocus()
   mbFocused = False
   'must repaint to remove the focus rect
   subPaintButton True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

mbMouseIsDown = True
subPaintButton False

'if this a regular style command button
'then raise the mouse down event.
'If this is the MenuDropdown style of
'button then we only want to raise the
'Mousedown event if the mouse is down on
'the left part of the button. When its
'down on the right part the [MouseDownOnDropdown]
'event is raised
If m_ButtonStyle = bsRegular Then
   RaiseEvent MouseDown(Button, Shift, X, Y)
Else
   Dim px   As Long
   Dim py   As Long
   
   With Screen
     px = (X \ .TwipsPerPixelX)
     py = (Y \ .TwipsPerPixelY)
     
     If PtInRect(mR2, px, py) Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
     End If
   End With
End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'this timer looks for when the mouse leaves the control
 With Timer1
   If .Interval = 0 Then
      .Interval = 100
      .Enabled = True
      mbMouseEntered = True
      subPaintButton True
      RaiseEvent MouseEnter
   End If
 End With
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim px  As Long
Dim py  As Long
  
'if were using dropdown style of button
'then we want to able to pass an event to
'the user is were mousing down on the
'right part of the button that has the
'drop down arrow so user can display a
'menu or whatever. This events isnt raised
'in the [Usercontrol_Mousedown] event because
'the resulting lose of focus prevents proper
'repainting of the button when the menu
'loses focus, for some strange reason

If m_ButtonStyle = bsMenuDropdown Then
  With Screen
    'convert x and y to pixels
    px = (X \ .TwipsPerPixelX)
    py = (Y \ .TwipsPerPixelY)
  End With
  'mouse is down in the dropdown part
  If PtInRect(mR3, px, py) Then
     RaiseEvent MouseDownOnDropdown
  End If
End If

mbMouseIsDown = False
subPaintButton True
RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Show()
  subPaintButton True
End Sub
 
Private Sub UserControl_Resize()
  'set the rect[mR1]
  subSetRects
  subPaintButton True
End Sub
 
Private Sub Timer1_Timer()
Dim pt As POINTAPI
  
  GetCursorPos pt
  'if the mouse has left the boundries of this control...
  If WindowFromPoint(pt.X, pt.Y) <> hWnd Then
     mbMouseEntered = False
     Timer1.Interval = 0
     subPaintButton True
     RaiseEvent MouseExit
  End If
End Sub

 
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    subPaintButton True
End Property

Public Property Get ButtonStyle() As enBs
Attribute ButtonStyle.VB_Description = "If ButtonStyle= [bsRegular] then control looks and acts as  regular command button. If ButtonStyle = [bsMenuDropdown] then right 200 twips of button is a seperate button area that has down arrow on it. Normally used in conjunction with a menu."
    ButtonStyle = m_ButtonStyle
End Property
Public Property Let ButtonStyle(ByVal New_ButtonStyle As enBs)
    m_ButtonStyle = New_ButtonStyle
    PropertyChanged "ButtonStyle"
    subPaintButton True
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    subPaintButton True
End Property

Public Property Get Image() As Picture
Attribute Image.VB_Description = "The image to display in the button."
    Set Image = Picture1.Picture
End Property
Public Property Set Image(ByVal New_Image As Picture)

 With Picture1
    Set .Picture = New_Image
    PropertyChanged "Image"
    'update ImageWidth and ImageHeight
    'properties to reflect the image
    'just loaded into picture1
    ImageWidth = .Width
    ImageHeight = .Height
 End With
 
 subPaintButton True
End Property

Public Property Get ImageFormat() As enIf
Attribute ImageFormat.VB_Description = "If ImageFormat=[ifWYSWYG] then the images height+ width will be the same as the image supplied. If ImageFormat=[ifDistort] then you can alter the ImageHeight and ImageWidth properties. ImageLeft and ImageTop can be altered either way."
    ImageFormat = m_ImageFormat
End Property
Public Property Let ImageFormat(ByVal New_ImageFormat As enIf)
    m_ImageFormat = New_ImageFormat
    PropertyChanged "ImageFormat"
    subPaintButton True
End Property

Public Property Get ImageLeft() As Long
    'convert the pixels to twips for the user
    ImageLeft = (m_ImageLeft * Screen.TwipsPerPixelX)
End Property
Public Property Let ImageLeft(ByVal New_ImageLeft As Long)
    'convert user supplied twips to pixels for API
    m_ImageLeft = (New_ImageLeft \ Screen.TwipsPerPixelX)
    PropertyChanged "ImageLeft"
    subPaintButton True
End Property

Public Property Get ImageTop() As Long
    'save the supplied twips val to pixels for API
    ImageTop = (m_ImageTop * Screen.TwipsPerPixelY)
End Property
Public Property Let ImageTop(ByVal New_ImageTop As Long)
   'convert user supplied twips to pixels for API
    m_ImageTop = (New_ImageTop \ Screen.TwipsPerPixelY)
    PropertyChanged "ImageTop"
    subPaintButton True
End Property

Public Property Get ImageWidth() As Long
    'convert the pixels val to twips for user convenience
    ImageWidth = (m_ImageWidth * Screen.TwipsPerPixelX)
End Property
Public Property Let ImageWidth(ByVal New_ImageWidth As Long)
    'convert the twips user provides to pixels for API
    m_ImageWidth = (New_ImageWidth \ Screen.TwipsPerPixelX)
    PropertyChanged "ImageWidth"
    subPaintButton True
End Property

Public Property Get ImageHeight() As Long
    'convert the stored pixels val to twips for the user
    ImageHeight = (m_ImageHeight * Screen.TwipsPerPixelY)
End Property
Public Property Let ImageHeight(ByVal New_ImageHeight As Long)
    'convert user supplied twips to pixels for API
    m_ImageHeight = (New_ImageHeight \ Screen.TwipsPerPixelY)
    PropertyChanged "ImageHeight"
    subPaintButton True
End Property

Public Property Get Multiline() As Boolean
Attribute Multiline.VB_Description = "Determines if the button can support multiple lines of text as opposed to single line display. When Multiline=[True] the [TextAlign]  property tends not to work well"
    Multiline = m_Multiline
End Property
Public Property Let Multiline(ByVal New_Multiline As Boolean)
    m_Multiline = New_Multiline
    PropertyChanged "Multiline"
    subPaintButton True
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "The text that is displayed in the button. This is the default property of this control."
Attribute Text.VB_MemberFlags = "200"
    Text = m_Text
End Property
Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    PropertyChanged "Text"
    subPaintButton True
End Property

Public Property Get TextAlign() As enTextAlign
Attribute TextAlign.VB_Description = "The horizontal alignment of the text in this control"
    TextAlign = m_TextAlign
End Property
Public Property Let TextAlign(ByVal New_TextAlign As enTextAlign)
    m_TextAlign = New_TextAlign
    PropertyChanged "TextAlign"
    subPaintButton True
End Property

Public Property Get TextColor() As OLE_COLOR
Attribute TextColor.VB_Description = "The color of the text in the control"
    TextColor = m_TextColor
End Property
Public Property Let TextColor(ByVal New_TextColor As OLE_COLOR)
    m_TextColor = New_TextColor
    PropertyChanged "TextColor"
    subPaintButton True
End Property

Public Property Get TextColorOnMouseover() As OLE_COLOR
Attribute TextColorOnMouseover.VB_Description = "The color of the text in the control when the mouse is hovering over it"
    TextColorOnMouseover = m_TextColorOnMouseover
End Property
Public Property Let TextColorOnMouseover(ByVal New_TextColorOnMouseover As OLE_COLOR)
    m_TextColorOnMouseover = New_TextColorOnMouseover
    PropertyChanged "TextColorOnMouseover"
End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ButtonStyle = m_def_ButtonStyle
    m_Text = m_def_Text
    m_TextColor = m_def_TextColor
    m_TextAlign = m_def_TextAlign
    m_Multiline = m_def_Multiline
    m_ImageFormat = m_def_ImageFormat
    m_TextColorOnMouseover = m_def_TextColorOnMouseover
    Set UserControl.Font = Ambient.Font
    m_ImageLeft = m_def_ImageLeft
    m_ImageTop = m_def_ImageTop
    m_ImageWidth = m_def_ImageWidth
    m_ImageHeight = m_def_ImageHeight
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ButtonStyle = PropBag.ReadProperty("ButtonStyle", m_def_ButtonStyle)
    m_Text = PropBag.ReadProperty("Text", m_def_Text)
    m_TextColor = PropBag.ReadProperty("TextColor", m_def_TextColor)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    m_Multiline = PropBag.ReadProperty("Multiline", m_def_Multiline)
    Set Picture1.Picture = PropBag.ReadProperty("Image", Nothing)
    m_ImageFormat = PropBag.ReadProperty("ImageFormat", m_def_ImageFormat)
    m_TextColorOnMouseover = PropBag.ReadProperty("TextColorOnMouseover", m_def_TextColorOnMouseover)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ImageLeft = PropBag.ReadProperty("ImageLeft", m_def_ImageLeft)
    m_ImageTop = PropBag.ReadProperty("ImageTop", m_def_ImageTop)
    m_ImageWidth = PropBag.ReadProperty("ImageWidth", m_def_ImageWidth)
    m_ImageHeight = PropBag.ReadProperty("ImageHeight", m_def_ImageHeight)

    'lets look for and access key "&"
    subFindHotKey m_Text
 End Sub
  
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ButtonStyle", m_ButtonStyle, m_def_ButtonStyle)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("TextColor", m_TextColor, m_def_TextColor)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
    Call PropBag.WriteProperty("Multiline", m_Multiline, m_def_Multiline)
    Call PropBag.WriteProperty("Image", Picture1.Picture, Nothing)
    Call PropBag.WriteProperty("ImageFormat", m_ImageFormat, m_def_ImageFormat)
    Call PropBag.WriteProperty("TextColorOnMouseover", m_TextColorOnMouseover, m_def_TextColorOnMouseover)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ImageLeft", m_ImageLeft, m_def_ImageLeft)
    Call PropBag.WriteProperty("ImageTop", m_ImageTop, m_def_ImageTop)
    Call PropBag.WriteProperty("ImageWidth", m_ImageWidth, m_def_ImageWidth)
    Call PropBag.WriteProperty("ImageHeight", m_ImageHeight, m_def_ImageHeight)
End Sub
   
