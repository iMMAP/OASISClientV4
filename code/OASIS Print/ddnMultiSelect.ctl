VERSION 5.00
Begin VB.UserControl ddnMultiSelect 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   ToolboxBitmap   =   "ddnMultiSelect.ctx":0000
   Begin VB.Timer tmrLostMouse 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2160
      Top             =   960
   End
End
Attribute VB_Name = "ddnMultiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAGIC_END_EDIT_IGNORE_WINDOW_PROP As String = "VBAL:SGRID:EDITOR"

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" _
   (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal lhDC As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pRect As RECT, pClipRect As RECT) As Long
    
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function ImageList_GetImageRect Lib "comctl32.dll" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT _
    ) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" ( _
        ByVal hIml As Long, ByVal i As Long, _
        ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, _
        ByVal fStyle As Long _
    ) As Long
Private Const ILD_NORMAL = 0
Private Const ILD_TRANSPARENT = 1
Private Const ILD_BLEND25 = 2
Private Const ILD_SELECTED = 4
Private Const ILD_FOCUS = 4
Private Const ILD_MASK = &H10&
Private Const ILD_IMAGE = &H20&
Private Const ILD_ROP = &H40&
Private Const ILD_OVERLAYMASK = 3840
Private Const ILC_COLOR = &H0
Private Const ILC_COLOR32 = &H20
Private Const ILC_MASK = &H1&

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID = 0
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
   Private Const GWL_STYLE = (-16)
   Private Const WS_BORDER = &H800000
   Private Const WS_CHILD = &H40000000
   Private Const WS_DISABLED = &H8000000
   Private Const WS_VISIBLE = &H10000000
   Private Const WS_TABSTOP = &H100000
   Private Const WS_HSCROLL = &H100000
   Private Const GWL_EXSTYLE = (-20)
   Private Const WS_EX_TOPMOST = &H8&
   Private Const WS_EX_CLIENTEDGE = &H200&
   Private Const WS_EX_STATICEDGE = &H20000
   Private Const WS_EX_WINDOWEDGE = &H100&
   Private Const WS_EX_APPWINDOW = &H40000
   Private Const WS_EX_TOOLWINDOW = &H80&
   Private Const WS_EX_LAYERED As Long = &H80000

Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
   Private Const SW_HIDE = 0

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
   Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
   Private Const SWP_NOACTIVATE = &H10
   Private Const SWP_NOMOVE = &H2
   Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
   Private Const SWP_NOREDRAW = &H8
   Private Const SWP_NOSIZE = &H1
   Private Const SWP_NOZORDER = &H4
   Private Const SWP_SHOWWINDOW = &H40
   Private Const HWND_DESKTOP = 0
   Private Const HWND_NOTOPMOST = -2
   Private Const HWND_TOP = 0
   Private Const HWND_TOPMOST = -1
   Private Const HWND_BOTTOM = 1

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    Private Const DT_LEFT = &H0&
    Private Const DT_TOP = &H0&
    Private Const DT_CENTER = &H1&
    Private Const DT_RIGHT = &H2&
    Private Const DT_VCENTER = &H4&
    Private Const DT_BOTTOM = &H8&
    Private Const DT_WORDBREAK = &H10&
    Private Const DT_SINGLELINE = &H20&
    Private Const DT_EXPANDTABS = &H40&
    Private Const DT_TABSTOP = &H80&
    Private Const DT_NOCLIP = &H100&
    Private Const DT_EXTERNALLEADING = &H200&
    Private Const DT_CALCRECT = &H400&
    Private Const DT_NOPREFIX = &H800
    Private Const DT_INTERNAL = &H1000&
    Private Const DT_WORD_ELLIPSIS = &H40000

Private Type OSVERSIONINFO
   dwVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion(0 To 127) As Byte
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_NT = 2


Private Type tItem
   sKey As String
   lIcon As Long
   bChecked As Boolean
   bEnabled As Boolean
   sText As String
   rcItem As RECT
   bMouseOver As Boolean
   bMouseDown As Boolean
End Type

Private m_tItems() As tItem
Private m_iItemCount As Long
Private m_hIml As Long
Private m_hWnd As Long
Private m_lWidth As Long
Private m_lMinWidth As Long
Private m_lHeight As Long
Private m_bIsNt As Boolean
Private m_bIsXp As Boolean
Private m_lItemHeight As Long
Private m_lIconWidth As Long
Private m_lIconHeight As Long
Private m_sDelimiter As String
Private m_bCheckBoxes As Boolean
Private m_bEnabled As Boolean

Private m_bDropDownMode As Boolean
Private m_ptrOwner As Long
Private m_bShowingPopup As Boolean
Private m_ptrPopup As Long

Private m_rcButton As RECT
Private m_bMouseOverButton As Boolean
Private m_bMouseDownButton As Boolean

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event CheckChange(ByVal lIndex As Long, ByVal bCancel As Boolean)
Public Event RequestDropDownInstance(ctl As ddnMultiSelect)

Public Property Get Enabled() As Boolean
        '<EhHeader>
        On Error GoTo Enabled_Err
        '</EhHeader>
100    Enabled = m_bEnabled
        '<EhFooter>
        Exit Property

Enabled_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.Enabled", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let Enabled(ByVal value As Boolean)
        '<EhHeader>
        On Error GoTo Enabled_Err
        '</EhHeader>
100    m_bEnabled = value
102    UserControl.Enabled = value
104    If Not (m_bDropDownMode) Then
106       pPaint
       End If
108    PropertyChanged "Enabled"
        '<EhFooter>
        Exit Property

Enabled_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.Enabled", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property

Public Property Get CheckBoxes() As Boolean
        '<EhHeader>
        On Error GoTo CheckBoxes_Err
        '</EhHeader>
100    CheckBoxes = m_bCheckBoxes
        '<EhFooter>
        Exit Property

CheckBoxes_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.CheckBoxes", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let CheckBoxes(ByVal value As Boolean)
        '<EhHeader>
        On Error GoTo CheckBoxes_Err
        '</EhHeader>
100    m_bCheckBoxes = value
102    EvalSize
104    PropertyChanged "CheckBoxes"
        '<EhFooter>
        Exit Property

CheckBoxes_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.CheckBoxes", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property

Public Sub EndEdit()
        '<EhHeader>
        On Error GoTo EndEdit_Err
        '</EhHeader>
100    If (m_bShowingPopup) Then
102       fPopupHide
       End If
        '<EhFooter>
        Exit Sub

EndEdit_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.EndEdit", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Public Property Get Delimiter() As String
        '<EhHeader>
        On Error GoTo Delimiter_Err
        '</EhHeader>
100    Delimiter = m_sDelimiter
        '<EhFooter>
        Exit Property

Delimiter_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.Delimiter", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let Delimiter(ByVal value As String)
        '<EhHeader>
        On Error GoTo Delimiter_Err
        '</EhHeader>
100    m_sDelimiter = value
102    If Not (m_bDropDownMode) Then
104       pPaint
       End If
106    PropertyChanged "Delimiter"
        '<EhFooter>
        Exit Property

Delimiter_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.Delimiter", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property

Public Property Get Selection() As String
        '<EhHeader>
        On Error GoTo Selection_Err
        '</EhHeader>
    Dim sSel As String
    Dim i As Long
100    For i = 1 To m_iItemCount
102       If m_tItems(i).bChecked Then
104          If Len(sSel) > 0 Then
106             sSel = sSel & m_sDelimiter & " "
             End If
108          sSel = sSel & m_tItems(i).sText
          End If
110    Next i
112    Selection = sSel
        '<EhFooter>
        Exit Property

Selection_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.Selection", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let Selection(ByVal value As String)
        '<EhHeader>
        On Error GoTo Selection_Err
        '</EhHeader>
    Dim iPos As Long
    Dim iNextPos As Long
    Dim iItem As Long
    Dim sItem As String
100    For iItem = 1 To m_iItemCount
102       m_tItems(iItem).bChecked = False
104    Next iItem
106    iPos = 1
108    iNextPos = InStr(iPos, value, m_sDelimiter)
110    Do While (iNextPos > 0)
112       sItem = Trim(Mid(value, iPos, iNextPos - iPos))
114       iItem = IndexForText(sItem)
116       If (iItem > 0) Then
118          m_tItems(iItem).bChecked = True
          End If
120       iPos = iNextPos + Len(m_sDelimiter)
122       iNextPos = InStr(iPos, value, m_sDelimiter)
       Loop
124    If (iPos < Len(value)) Then
126       sItem = Trim(Mid(value, iPos))
128       iItem = IndexForText(sItem)
130       If (iItem > 0) Then
132          m_tItems(iItem).bChecked = True
          End If
       End If
134    pPaint
        '<EhFooter>
        Exit Property

Selection_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.Selection", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get hWnd() As Long
        '<EhHeader>
        On Error GoTo hWnd_Err
        '</EhHeader>
100    hWnd = m_hWnd
        '<EhFooter>
        Exit Property

hWnd_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.hWnd", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get hIml() As Long
        '<EhHeader>
        On Error GoTo hIml_Err
        '</EhHeader>
100    hIml = m_hIml
        '<EhFooter>
        Exit Property

hIml_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.hIml", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let hIml(ByVal value As Long)
        '<EhHeader>
        On Error GoTo hIml_Err
        '</EhHeader>
    Dim rc As RECT
100    m_hIml = value
102    ImageList_GetImageRect m_hIml, 0, rc
104    m_lIconWidth = rc.right - rc.left
106    m_lIconHeight = rc.bottom - rc.top
108    EvalSize
        '<EhFooter>
        Exit Property

hIml_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.hIml", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get ItemCount() As Long
        '<EhHeader>
        On Error GoTo ItemCount_Err
        '</EhHeader>
100    ItemCount = m_iItemCount
        '<EhFooter>
        Exit Property

ItemCount_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemCount", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get ItemEnabled(ByVal nIndex As Long) As Boolean
        '<EhHeader>
        On Error GoTo ItemEnabled_Err
        '</EhHeader>
100    ItemEnabled = m_tItems(nIndex).bEnabled
        '<EhFooter>
        Exit Property

ItemEnabled_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemEnabled", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let ItemEnabled(ByVal nIndex As Long, ByVal value As Boolean)
        '<EhHeader>
        On Error GoTo ItemEnabled_Err
        '</EhHeader>
100    m_tItems(nIndex).bEnabled = value
        '<EhFooter>
        Exit Property

ItemEnabled_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemEnabled", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get ItemChecked(ByVal nIndex As Long) As Boolean
        '<EhHeader>
        On Error GoTo ItemChecked_Err
        '</EhHeader>
100    ItemChecked = m_tItems(nIndex).bChecked
        '<EhFooter>
        Exit Property

ItemChecked_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemChecked", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let ItemChecked(ByVal nIndex As Long, ByVal value As Boolean)
        '<EhHeader>
        On Error GoTo ItemChecked_Err
        '</EhHeader>
100    m_tItems(nIndex).bChecked = value
        '<EhFooter>
        Exit Property

ItemChecked_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemChecked", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Friend Property Let fItemChecked(ByVal nIndex As Long, ByVal value As Long)
        '<EhHeader>
        On Error GoTo fItemChecked_Err
        '</EhHeader>
100    m_tItems(nIndex).bChecked = value
102    pPaint
        '<EhFooter>
        Exit Property

fItemChecked_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.fItemChecked", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get ItemIcon(ByVal nIndex As Long) As Long
        '<EhHeader>
        On Error GoTo ItemIcon_Err
        '</EhHeader>
100    ItemIcon = m_tItems(nIndex).lIcon
        '<EhFooter>
        Exit Property

ItemIcon_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemIcon", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let ItemIcon(ByVal nIndex As Long, ByVal value As Long)
        '<EhHeader>
        On Error GoTo ItemIcon_Err
        '</EhHeader>
100    m_tItems(nIndex).lIcon = value
        '<EhFooter>
        Exit Property

ItemIcon_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemIcon", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get ItemText(ByVal nIndex As Long) As String
        '<EhHeader>
        On Error GoTo ItemText_Err
        '</EhHeader>
100    ItemText = m_tItems(nIndex).sText
        '<EhFooter>
        Exit Property

ItemText_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemText", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let ItemText(ByVal nIndex As Long, ByVal value As String)
        '<EhHeader>
        On Error GoTo ItemText_Err
        '</EhHeader>
100    m_tItems(nIndex).sText = value
        '<EhFooter>
        Exit Property

ItemText_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemText", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get ItemKey(ByVal nIndex As Long) As String
        '<EhHeader>
        On Error GoTo ItemKey_Err
        '</EhHeader>
100    ItemKey = m_tItems(nIndex).sKey
        '<EhFooter>
        Exit Property

ItemKey_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemKey", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let ItemKey(ByVal nIndex As Long, ByVal value As String)
        '<EhHeader>
        On Error GoTo ItemKey_Err
        '</EhHeader>
100    m_tItems(nIndex).sKey = value
        '<EhFooter>
        Exit Property

ItemKey_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ItemKey", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get IndexForText(ByVal value As String) As Long
        '<EhHeader>
        On Error GoTo IndexForText_Err
        '</EhHeader>
    Dim i As Long
100    For i = 1 To m_iItemCount
102       If (m_tItems(i).sText = value) Then
104          IndexForText = i
             Exit For
          End If
106    Next i
        '<EhFooter>
        Exit Property

IndexForText_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.IndexForText", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get IndexForKey(ByVal value As String) As Long
        '<EhHeader>
        On Error GoTo IndexForKey_Err
        '</EhHeader>
    Dim i As Long
100    For i = 1 To m_iItemCount
102       If (m_tItems(i).sText = value) Then
104          IndexForKey = i
             Exit For
          End If
106    Next i
        '<EhFooter>
        Exit Property

IndexForKey_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.IndexForKey", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Sub AddItem( _
      ByVal sKey As String, _
      Optional ByVal lIcon As Long = -1, _
      Optional ByVal sText As String = "", _
      Optional ByVal bChecked As Boolean = False, _
      Optional ByVal bEnabled As Boolean = True _
   )
        '<EhHeader>
        On Error GoTo AddItem_Err
        '</EhHeader>
100    m_iItemCount = m_iItemCount + 1
102    ReDim Preserve m_tItems(1 To m_iItemCount) As tItem
104    With m_tItems(m_iItemCount)
106       .sKey = sKey
108       .sText = sText
110       .lIcon = lIcon
112       .bChecked = bChecked
114       .bEnabled = bEnabled
       End With
116    EvalSize
        '<EhFooter>
        Exit Sub

AddItem_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.AddItem", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub
Public Sub RemoveItem(ByVal lIndex As Long)
        '<EhHeader>
        On Error GoTo RemoveItem_Err
        '</EhHeader>
    Dim i As Long
100    If (m_iItemCount > 1) Then
102       For i = m_iItemCount - 1 To lIndex Step -1
104          LSet m_tItems(i + 1) = m_tItems(i)
106       Next i
108       m_iItemCount = m_iItemCount + 1
110       ReDim Preserve m_tItems(1 To m_iItemCount) As tItem
       Else
112       m_iItemCount = 0
114       Erase m_tItems
       End If
116    EvalSize
        '<EhFooter>
        Exit Sub

RemoveItem_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.RemoveItem", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub
Public Property Get DropDownShowing() As Boolean
        '<EhHeader>
        On Error GoTo DropDownShowing_Err
        '</EhHeader>
100    DropDownShowing = m_bShowingPopup
        '<EhFooter>
        Exit Property

DropDownShowing_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.DropDownShowing", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Get Font() As IFont
        '<EhHeader>
        On Error GoTo Font_Err
        '</EhHeader>
100    Set Font = UserControl.Font
        '<EhFooter>
        Exit Property

Font_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.Font", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Let Font(iFnt As IFont)
        '<EhHeader>
        On Error GoTo Font_Err
        '</EhHeader>
100    pSetFont iFnt
        '<EhFooter>
        Exit Property

Font_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.Font", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Public Property Set Font(iFnt As IFont)
        '<EhHeader>
        On Error GoTo Font_Err
        '</EhHeader>
100    pSetFont iFnt
        '<EhFooter>
        Exit Property

Font_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.Font", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property
Private Sub pSetFont(iFnt As IFont)
        '<EhHeader>
        On Error GoTo pSetFont_Err
        '</EhHeader>
100    Set UserControl.Font = iFnt
102    PropertyChanged "Font"
        '<EhFooter>
        Exit Sub

pSetFont_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.pSetFont", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub ShowPopup(ByVal hWndRelativeTo As Long, ByVal x As Long, ByVal y As Long)
        '<EhHeader>
        On Error GoTo ShowPopup_Err
        '</EhHeader>
    Dim ctl As ddnMultiSelect
100    If Not (m_bShowingPopup) Then
102       RaiseEvent RequestDropDownInstance(ctl)
104       If Not (ctl Is Nothing) Then
             Dim tP As POINTAPI
106          tP.x = x
108          tP.y = y
110          ScreenToClient hWndRelativeTo, tP
112          ctl.fSetData m_tItems, m_iItemCount
114          ctl.fShowPopup Me, tP.x, tP.y
116          m_bShowingPopup = True
118          m_ptrPopup = ObjPtr(ctl)
120          pPaint
          End If
       End If
        '<EhFooter>
        Exit Sub

ShowPopup_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ShowPopup", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub
Friend Sub fPopupHide()
        '<EhHeader>
        On Error GoTo fPopupHide_Err
        '</EhHeader>
    Dim ctl As ddnMultiSelect
100    If (m_bShowingPopup) Then
102       If Not (m_ptrPopup = 0) Then
104          Set ctl = ObjectFromPtr(m_ptrPopup)
106          ctl.fHidePopup
          End If
       End If
108    m_bShowingPopup = False
        '<EhFooter>
        Exit Sub

fPopupHide_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.fPopupHide", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Friend Sub fSetData(items() As tItem, ByVal lCount As Long)
        '<EhHeader>
        On Error GoTo fSetData_Err
        '</EhHeader>
    Dim i As Long
100    m_iItemCount = lCount
102    ReDim m_tItems(1 To m_iItemCount) As tItem
104    For i = 1 To m_iItemCount
106       LSet m_tItems(i) = items(i)
108    Next i
110    EvalSize
        '<EhFooter>
        Exit Sub

fSetData_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.fSetData", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Friend Sub fShowPopup(ctl As ddnMultiSelect, ByVal x As Long, ByVal y As Long)
        '<EhHeader>
        On Error GoTo fShowPopup_Err
        '</EhHeader>
    Dim rc As RECT
    Dim cM As cMonitor
    Dim rcShow As RECT
   
100    GetWindowRect ctl.hWnd, rc
102    m_lMinWidth = rc.right - rc.left
104    m_hIml = ctl.hIml
106    EvalSize
108    rcShow.left = x
110    rcShow.top = y
112    rcShow.right = rcShow.left + m_lWidth
114    rcShow.bottom = rcShow.top + m_lHeight
   
116    Set cM = New cMonitor
118    cM.CreateFromPoint x, y
120    If (cM.hMonitor = 0) Then
122       If (rcShow.top < 0) Then
124          OffsetRect rcShow, 0, -rcShow.top
          End If
126       If (rcShow.bottom > Screen.Height \ Screen.TwipsPerPixelY) Then
128          OffsetRect rcShow, 0, -(rc.bottom - rc.top) - (rcShow.bottom - rcShow.top)
          End If
130       If (rcShow.left < 0) Then
132          OffsetRect rcShow, -rcShow.left, 0
          End If
134       If (rcShow.right > Screen.Width \ Screen.TwipsPerPixelY) Then
136          OffsetRect rcShow, (Screen.Width \ Screen.TwipsPerPixelY - rcShow.right), 0
          End If
       Else
138       If (rcShow.top < cM.WorkTop) Then
140          OffsetRect rcShow, 0, -rcShow.top
          End If
142       If (rcShow.bottom > cM.WorkTop + cM.WorkHeight) Then
144          OffsetRect rcShow, 0, -(rc.bottom - rc.top) - (rcShow.bottom - rcShow.top)
          End If
146       If (rcShow.left < cM.WorkLeft) Then
148          OffsetRect rcShow, -rcShow.left, 0
          End If
150       If (rcShow.right > cM.WorkLeft + cM.WorkWidth) Then
152          OffsetRect rcShow, (cM.WorkLeft + cM.WorkWidth) - rcShow.right, 0
          End If
       End If
   
154    m_bDropDownMode = True
156    m_ptrOwner = ObjPtr(ctl)
       ' Set the style of the object so it works as a popup:
       Dim lStyle As Long
158    lStyle = GetWindowLong(m_hWnd, GWL_EXSTYLE)
160    lStyle = lStyle Or WS_EX_TOOLWINDOW
162    lStyle = lStyle And Not (WS_EX_APPWINDOW)
164    SetWindowLong m_hWnd, GWL_EXSTYLE, lStyle
166    SetParent m_hWnd, HWND_DESKTOP
168    SetProp m_hWnd, MAGIC_END_EDIT_IGNORE_WINDOW_PROP, 1
170    SetWindowPos m_hWnd, HWND_TOPMOST, rcShow.left, rcShow.top, m_lWidth, m_lHeight, SWP_SHOWWINDOW
172    pPaint
   
        '<EhFooter>
        Exit Sub

fShowPopup_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.fShowPopup", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Friend Sub fHidePopup()
        '<EhHeader>
        On Error GoTo fHidePopup_Err
        '</EhHeader>
100    ShowWindow m_hWnd, SW_HIDE
102    RemoveProp m_hWnd, MAGIC_END_EDIT_IGNORE_WINDOW_PROP
        '<EhFooter>
        Exit Sub

fHidePopup_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.fHidePopup", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
        '<EhHeader>
        On Error GoTo ObjectFromPtr_Err
        '</EhHeader>
    Dim objT As Object
100    If Not (lPtr = 0) Then
          ' Turn the pointer into an illegal, uncounted interface
102       CopyMemory objT, lPtr, 4
          ' Do NOT hit the End button here! You will crash!
          ' Assign to legal reference
104       Set ObjectFromPtr = objT
          ' Still do NOT hit the End button here! You will still crash!
          ' Destroy the illegal reference
106       CopyMemory objT, 0&, 4
       End If
        '<EhFooter>
        Exit Property

ObjectFromPtr_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.ObjectFromPtr", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property

Private Sub VerInitialise()
        '<EhHeader>
        On Error GoTo VerInitialise_Err
        '</EhHeader>
   
       Dim tOSV As OSVERSIONINFO
100    tOSV.dwVersionInfoSize = Len(tOSV)
102    GetVersionEx tOSV
   
104    m_bIsNt = ((tOSV.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
106    If (tOSV.dwMajorVersion > 5) Then
          'm_bHasGradientAndTransparency = True
108       m_bIsXp = True
          'm_bIs2000OrAbove = True
110    ElseIf (tOSV.dwMajorVersion = 5) Then
          'm_bHasGradientAndTransparency = True
          'm_bIs2000OrAbove = True
112       If (tOSV.dwMinorVersion >= 1) Then
114          m_bIsXp = True
          End If
116    ElseIf (tOSV.dwMajorVersion = 4) Then ' NT4 or 9x/ME/SE
          'If (tOSV.dwMinorVersion >= 10) Then
          '   m_bHasGradientAndTransparency = True
          'End If
       Else ' Too old
       End If
   
        '<EhFooter>
        Exit Sub

VerInitialise_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.VerInitialise", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub
Private Sub DrawText( _
      ByVal lhDC As Long, _
      ByVal sText As String, _
      ByVal lLength As Long, _
      tR As RECT, _
      ByVal lFlags As Long _
   )
        '<EhHeader>
        On Error GoTo DrawText_Err
        '</EhHeader>
    Dim lPtr As Long
100    If (m_bIsNt) Then
102       lPtr = StrPtr(sText)
104       If Not (lPtr = 0) Then ' NT4 crashes with ptr = 0
106          DrawTextW lhDC, lPtr, -1, tR, lFlags
          End If
       Else
108       DrawTextA lhDC, sText, -1, tR, lFlags
       End If
        '<EhFooter>
        Exit Sub

DrawText_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.DrawText", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub
Private Sub EvalSize()
        '<EhHeader>
        On Error GoTo EvalSize_Err
        '</EhHeader>
    Dim i As Long
    Dim tR As RECT
    Dim lMaxWidth As Long
    Dim lDefWidth As Long
    Dim lMaxHeight As Long
    Dim lhDC As Long

100    lhDC = UserControl.hdc
      
102    If (m_hIml = 0) Then
104       lMaxHeight = 20
106       lDefWidth = 20
       Else
108       lMaxHeight = m_lIconHeight + 4
110       lDefWidth = (m_lIconWidth + 4) * 2
       End If
112    lMaxWidth = lDefWidth
      
   
114    For i = 1 To m_iItemCount
116       tR.right = 256
118       tR.bottom = 256
120       DrawText lhDC, m_tItems(i).sText, -1, tR, DT_CALCRECT Or DT_SINGLELINE
122       If (tR.bottom - tR.top + 4) > lMaxHeight Then
124          lMaxHeight = tR.bottom - tR.top + 4
          End If
126       If (tR.right - tR.left + 4 + lDefWidth) > lMaxWidth Then
128          lMaxWidth = lDefWidth + (tR.right - tR.left + 4)
          End If
130    Next i
   
132    m_lWidth = lMaxWidth + 8
134    If (m_lWidth < m_lMinWidth) Then
136       m_lWidth = m_lMinWidth
       End If
138    m_lItemHeight = lMaxHeight
140    m_lHeight = lMaxHeight * m_iItemCount + 4
   
        '<EhFooter>
        Exit Sub

EvalSize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.EvalSize", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub pPaint()
        '<EhHeader>
        On Error GoTo pPaint_Err
        '</EhHeader>
100    If (m_bDropDownMode) Then
102       pPaintAsDropDown
       Else
104       pPaintAsSelector
       End If
        '<EhFooter>
        Exit Sub

pPaint_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.pPaint", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub pPaintAsSelector()
        '<EhHeader>
        On Error GoTo pPaintAsSelector_Err
        '</EhHeader>
    Dim lhDC As Long
    Dim rc As RECT
    Dim rcWork As RECT
    Dim hBr As Long
    Dim hBrOutline As Long
    Dim lIconWidth As Long
    Dim lIconHeight As Long
    Dim hTheme As Long

100    lIconWidth = IIf(m_hIml = 0, 16, m_lIconWidth)
102    lIconHeight = IIf(m_hIml = 0, 16, m_lIconHeight)

104    lhDC = UserControl.hdc
106    GetClientRect m_hWnd, rc
   
108    hBr = CreateSolidBrush(TranslateColor(vbWindowBackground))
110    FillRect lhDC, rc, hBr
112    DeleteObject hBr
   
114    hBrOutline = CreateSolidBrush(TranslateColor(vbButtonShadow))
   
116    If (m_bIsXp) Then
          On Error Resume Next
118       hTheme = OpenThemeData(m_hWnd, StrPtr("EDIT"))
          On Error GoTo pPaintAsSelector_Err
       End If
120    If (hTheme = 0) Then
122       FrameRect lhDC, rc, hBrOutline
       Else
124       DrawThemeBackground hTheme, lhDC, 1, IIf(m_bEnabled, 1, 4), rc, rc
126       CloseThemeData hTheme
       End If

128    InflateRect rc, -2, -2
130    If (m_bEnabled And m_bMouseOverButton) Then
132       hBr = CreateSolidBrush(TranslateColor(vbHighlight))
134       FillRect lhDC, rc, hBr
136       DeleteObject hBr
       End If

138    LSet rcWork = rc
140    rcWork.right = rcWork.right - lIconWidth
142    If (m_bEnabled) Then
144       If (m_bMouseOverButton) Then
146          SetTextColor lhDC, TranslateColor(vbHighlightText)
148          SetBkColor lhDC, TranslateColor(vbHighlight)
150          SetBkMode lhDC, OPAQUE
          Else
152          SetTextColor lhDC, TranslateColor(vbWindowText)
154          SetBkMode lhDC, TRANSPARENT
          End If
       Else
156       SetTextColor lhDC, TranslateColor(vbButtonShadow)
158       SetBkMode lhDC, TRANSPARENT
       End If
160    DrawText lhDC, Selection, -1, rcWork, DT_SINGLELINE Or DT_VCENTER Or DT_END_ELLIPSIS
   
162    LSet m_rcButton = rcWork
164    m_rcButton.left = rcWork.right
166    m_rcButton.right = m_rcButton.left + lIconWidth

168    If (m_bIsXp) Then
          On Error Resume Next
170       hTheme = OpenThemeData(m_hWnd, StrPtr("COMBOBOX"))
          On Error GoTo pPaintAsSelector_Err
       End If
   
172    If (hTheme = 0) Then
174       If (m_bShowingPopup) Then
176          hBr = CreateSolidBrush(BlendColor(vbHighlight, vbButtonFace)) ' ,192))
178       ElseIf (m_bMouseOverButton) Then
       '      hBr = CreateSolidBrush(BlendColor(vbHighlight, vbButtonFace))
       '   Else
180          hBr = CreateSolidBrush(TranslateColor(vbButtonFace))
          End If
182       FillRect lhDC, m_rcButton, hBr
184       DeleteObject hBr
   
186       UtilDrawSplitGlyph lhDC, m_rcButton.left, m_rcButton.top, m_rcButton.right - m_rcButton.left, m_rcButton.bottom - m_rcButton.top, m_bEnabled, &H0&
   
188       If (m_bMouseOverButton Or m_bShowingPopup) Then
190          hBr = CreateSolidBrush(TranslateColor(vbHighlight))
192          FrameRect lhDC, m_rcButton, hBr
194          DeleteObject hBr
          Else
196          FrameRect lhDC, m_rcButton, hBrOutline
          End If
   
       Else
198       InflateRect m_rcButton, 1, 1
200       If (m_bEnabled) Then
202          If (m_bShowingPopup) Then
204             DrawThemeBackground hTheme, lhDC, 1, 3, m_rcButton, m_rcButton
             Else
206             DrawThemeBackground hTheme, lhDC, 1, 1, m_rcButton, m_rcButton
             End If
          Else
208          DrawThemeBackground hTheme, lhDC, 1, 4, m_rcButton, m_rcButton
          End If
210       CloseThemeData hTheme
       End If
   
212    DeleteObject hBrOutline
      
214    UserControl.Refresh

        '<EhFooter>
        Exit Sub

pPaintAsSelector_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.pPaintAsSelector", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub pPaintAsDropDown()
        '<EhHeader>
        On Error GoTo pPaintAsDropDown_Err
        '</EhHeader>
    Dim rc As RECT
    Dim i As Long
    Dim rcItem As RECT
    Dim rcWork As RECT
    Dim rcWork2 As RECT
    Dim lhDC As Long
    Dim hBr As Long
    Dim lIconWidth As Long
    Dim lIconHeight As Long
    Dim hTheme As Long

100    lIconWidth = IIf(m_hIml = 0, 16, m_lIconWidth)
102    lIconHeight = IIf(m_hIml = 0, 16, m_lIconHeight)

104    lhDC = UserControl.hdc
106    GetClientRect m_hWnd, rc
   
108    hBr = CreateSolidBrush(TranslateColor(vbWindowBackground))
110    FillRect lhDC, rc, hBr
112    DeleteObject hBr
   
114    hBr = CreateSolidBrush(TranslateColor(vbButtonShadow))
116    FrameRect lhDC, rc, hBr
118    DeleteObject hBr

120    If (m_bIsXp) Then
          On Error Resume Next
122       hTheme = OpenThemeData(m_hWnd, StrPtr("BUTTON"))
          On Error GoTo pPaintAsDropDown_Err
       End If

124    LSet rcItem = rc
126    InflateRect rcItem, -2, 0
128    rcItem.top = 2
130    For i = 1 To m_iItemCount
132       rcItem.bottom = rcItem.top + m_lItemHeight
134       LSet m_tItems(i).rcItem = rcItem
136       If (m_tItems(i).bMouseOver) Then
138          hBr = CreateSolidBrush(BlendColor(vbHighlight, vbWindowBackground))
140          FillRect lhDC, rcItem, hBr
142          DeleteObject hBr
144          hBr = CreateSolidBrush(TranslateColor(vbHighlight))
146          FrameRect lhDC, rcItem, hBr
148          DeleteObject hBr
          End If
         
          ' Check box:
150       LSet rcWork = rcItem
152       InflateRect rcWork, 0, -2
154       rcWork.left = rcWork.left + 2
156       rcWork.right = rcWork.left + lIconWidth
158       LSet rcWork2 = rcWork
160       rcWork2.top = ((rcWork.bottom - rcWork.top) - lIconHeight) \ 2
162       rcWork2.bottom = rcWork2.top + lIconHeight
            
164       If (hTheme = 0) Then
166          hBr = CreateSolidBrush(TranslateColor(vbWindowText))
168          FrameRect lhDC, rcWork, hBr
170          LSet rcWork2 = rcWork
172          InflateRect rcWork2, -1, -1
174          FrameRect lhDC, rcWork2, hBr
176          DeleteObject hBr
178          If (m_tItems(i).bChecked) Then
180             UtilDrawCheckGlyph lhDC, rcWork2.left, rcWork2.top, rcWork2.right - rcWork2.left, rcWork2.bottom - rcWork2.top, True, &H0&
             End If
          Else
182          If (m_tItems(i).bChecked) Then
184             DrawThemeBackground hTheme, lhDC, 3, 5, rcWork, rcWork
             Else
186             DrawThemeBackground hTheme, lhDC, 3, 1, rcWork, rcWork
             End If
          End If
      
188       If Not (m_hIml = 0) Then
             ' Icon
190          rcWork.left = rcWork.left + lIconWidth + 4
192          rcWork.right = rcWork.left + lIconWidth
194          ImageList_Draw m_hIml, m_tItems(i).lIcon, lhDC, rcWork.left, rcWork.top, ILD_TRANSPARENT
          End If
      
          ' Text
196       rcWork.left = rcWork.left + lIconWidth + 4
198       rcWork.right = rcItem.right - 2
200       DrawText lhDC, m_tItems(i).sText, -1, rcWork, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER Or DT_END_ELLIPSIS
      
202       rcItem.top = rcItem.top + m_lItemHeight
204    Next i
   
206    If Not (hTheme = 0) Then
208       CloseThemeData hTheme
       End If
   
210    UserControl.Refresh
   
        '<EhFooter>
        Exit Sub

pPaintAsDropDown_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.pPaintAsDropDown", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Property Get BlendColor( _
      ByVal oColorFrom As OLE_COLOR, _
      ByVal oColorTo As OLE_COLOR, _
      Optional ByVal Alpha As Long = 128 _
   ) As Long
        '<EhHeader>
        On Error GoTo BlendColor_Err
        '</EhHeader>
    Dim lCFrom As Long
    Dim lCTo As Long
100    lCFrom = TranslateColor(oColorFrom)
102    lCTo = TranslateColor(oColorTo)
    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
104    lSrcR = lCFrom And &HFF
106    lSrcG = (lCFrom And &HFF00&) \ &H100&
108    lSrcB = (lCFrom And &HFF0000) \ &H10000
110    lDstR = lCTo And &HFF
112    lDstG = (lCTo And &HFF00&) \ &H100&
114    lDstB = (lCTo And &HFF0000) \ &H10000
     
   
116    BlendColor = RGB( _
          ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
          ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
          ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
          )
      
        '<EhFooter>
        Exit Property

BlendColor_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.BlendColor", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Property



Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
        ' Convert Automation color to Windows color
        '<EhHeader>
        On Error GoTo TranslateColor_Err
        '</EhHeader>
100     If OleTranslateColor(oClr, hPal, TranslateColor) Then
102         TranslateColor = CLR_INVALID
        End If
        '<EhFooter>
        Exit Function

TranslateColor_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.TranslateColor", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Function

Private Sub UtilDrawCheckGlyph( _
      ByVal hdc As Long, _
      ByVal lLeft As Long, _
      ByVal lTop As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long, _
      ByVal bEnabled As Boolean, _
      ByVal color As Long _
   )
        '<EhHeader>
        On Error GoTo UtilDrawCheckGlyph_Err
        '</EhHeader>
    Dim lCentreY As Long
    Dim lCentreX As Long
    Dim tJ As POINTAPI
    Dim hPen As Long
    Dim hPenOld As Long
   
100    lCentreX = lLeft + lWidth \ 2 - 1
102    lCentreY = lTop + lHeight \ 2
   
104    hPen = CreatePen(PS_SOLID, 1, &H0)
106    hPenOld = SelectObject(hdc, hPenOld)
   
108    MoveToEx hdc, lCentreX - 3, lCentreY, tJ
110    LineTo hdc, lCentreX - 1, lCentreY + 2
112    MoveToEx hdc, lCentreX - 3, lCentreY + 1, tJ
114    LineTo hdc, lCentreX - 1, lCentreY + 3
   
116    MoveToEx hdc, lCentreX - 1, lCentreY + 3, tJ
118    LineTo hdc, lCentreX + 5, lCentreY - 3
120    MoveToEx hdc, lCentreX - 1, lCentreY + 2, tJ
122    LineTo hdc, lCentreX + 5, lCentreY - 4
   
124    SelectObject hdc, hPenOld
126    DeleteObject hPen
   
        '<EhFooter>
        Exit Sub

UtilDrawCheckGlyph_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UtilDrawCheckGlyph", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub
Public Sub UtilDrawSplitGlyph( _
      ByVal hdc As Long, _
      ByVal lLeft As Long, _
      ByVal lTop As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long, _
      ByVal bEnabled As Boolean, _
      ByVal color As Long _
   )
        '<EhHeader>
        On Error GoTo UtilDrawSplitGlyph_Err
        '</EhHeader>
    Dim lCentreY As Long
    Dim lCentreX As Long
   
100    lCentreX = lLeft + lWidth \ 2
102    lCentreY = lTop + lHeight \ 2

104    SetPixel hdc, lCentreX - 2, lCentreY - 1, color
106    SetPixel hdc, lCentreX - 1, lCentreY - 1, color
108    SetPixel hdc, lCentreX, lCentreY - 1, color
110    SetPixel hdc, lCentreX + 1, lCentreY - 1, color
112    SetPixel hdc, lCentreX + 2, lCentreY - 1, color
114    SetPixel hdc, lCentreX - 1, lCentreY, color
116    SetPixel hdc, lCentreX, lCentreY, color
118    SetPixel hdc, lCentreX + 1, lCentreY, color
120    SetPixel hdc, lCentreX, lCentreY + 1, color
   
        '<EhFooter>
        Exit Sub

UtilDrawSplitGlyph_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UtilDrawSplitGlyph", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Function plHitTest() As Long
        '<EhHeader>
        On Error GoTo plHitTest_Err
        '</EhHeader>
    Dim tP As POINTAPI
    Dim rc As RECT
    Dim i As Long
100    If (m_bDropDownMode) Then
102       GetCursorPos tP
104       GetWindowRect m_hWnd, rc
106       If Not (PtInRect(rc, tP.x, tP.y) = 0) Then
108          ScreenToClient m_hWnd, tP
110          For i = 1 To m_iItemCount
112             If Not (PtInRect(m_tItems(i).rcItem, tP.x, tP.y) = 0) Then
114                plHitTest = i
                   Exit For
                End If
116          Next i
          End If
       End If
        '<EhFooter>
        Exit Function

plHitTest_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.plHitTest", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Function

Private Sub pInitialise()
        '<EhHeader>
        On Error GoTo pInitialise_Err
        '</EhHeader>
100    m_hWnd = UserControl.hWnd
102    VerInitialise
        '<EhFooter>
        Exit Sub

pInitialise_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.pInitialise", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub selectNext(ByVal iDir As Long)
        '<EhHeader>
        On Error GoTo selectNext_Err
        '</EhHeader>
    Dim i As Long
    Dim iIndexSel As Long
    Dim iNewIndexSel As Long
   
100    If (m_iItemCount = 0) Then
          Exit Sub
       End If
   
102    For i = 1 To m_iItemCount
104       If (m_tItems(i).bMouseOver) Then
106          iIndexSel = i
             Exit For
          End If
108    Next i
   
110    If (iIndexSel = 0) Then
112       If (iDir > 0) Then
114          iNewIndexSel = 1
          Else
116          iNewIndexSel = m_iItemCount
          End If
       Else
118       iNewIndexSel = iIndexSel + iDir
120       If (iNewIndexSel > m_iItemCount) Then
122          iNewIndexSel = m_iItemCount
124       ElseIf (iNewIndexSel < 1) Then
126          iNewIndexSel = 1
          End If
       End If
   
128    If Not (iNewIndexSel = iIndexSel) Then
130       If (iIndexSel > 0) Then
132          m_tItems(iIndexSel).bMouseOver = False
          End If
134       m_tItems(iNewIndexSel).bMouseOver = True
136       pPaint
       End If
   
   
        '<EhFooter>
        Exit Sub

selectNext_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.selectNext", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub checkSelection()
        '<EhHeader>
        On Error GoTo checkSelection_Err
        '</EhHeader>
    Dim i As Long
    Dim iIndexSel As Long
    Dim ctl As ddnMultiSelect
   
100    If (m_iItemCount = 0) Then
          Exit Sub
       End If
   
102    For i = 1 To m_iItemCount
104       If (m_tItems(i).bMouseOver) Then
106          iIndexSel = i
             Exit For
          End If
108    Next i
   
110    If (iIndexSel > 0) Then
112       m_tItems(iIndexSel).bChecked = Not (m_tItems(iIndexSel).bChecked)
114       Set ctl = ObjectFromPtr(m_ptrOwner)
116       ctl.fItemChecked(iIndexSel) = m_tItems(iIndexSel).bChecked
118       pPaint
       End If
   
        '<EhFooter>
        Exit Sub

checkSelection_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.checkSelection", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Friend Sub fPopupKeyDown(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo fPopupKeyDown_Err
        '</EhHeader>
   
100    Select Case KeyCode
       Case vbKeyUp, vbKeyLeft
102       selectNext -1
104    Case vbKeyDown, vbKeyRight
106       selectNext 1
108    Case vbKeyPageUp
110       selectNext -8
112    Case vbKeyPageDown
114       selectNext 8
116    Case vbKeyHome
118       selectNext -10000
120    Case vbKeyEnd
122       selectNext 10000
124    Case vbKeyReturn, vbKeySpace
126       checkSelection
128    Case vbKeyEscape
130       fHidePopup
       End Select
   
        '<EhFooter>
        Exit Sub

fPopupKeyDown_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.fPopupKeyDown", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub


Private Sub tmrLostMouse_Timer()
   '
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
Dim tP As POINTAPI
Dim rc As RECT
Dim i As Long
Dim bChanged As Boolean
Dim bOutsideWin As Boolean

   GetCursorPos tP
   GetWindowRect m_hWnd, rc
   bOutsideWin = (PtInRect(rc, tP.x, tP.y) = 0)

   If (m_bDropDownMode) Then
      If (bOutsideWin) Then
         tmrLostMouse.Enabled = False
         For i = 1 To m_iItemCount
            If (m_tItems(i).bMouseOver) Then
               m_tItems(i).bMouseOver = False
               bChanged = True
            End If
         Next i
         If (bChanged) Then
            pPaint
         End If
      End If
   Else
      If m_bMouseOverButton Then
         If Not (m_bShowingPopup) Then
            If (bOutsideWin) Then
               tmrLostMouse.Enabled = False
               m_bMouseOverButton = False
               pPaint
            End If
         End If
      End If
   End If
   '
End Sub


Private Sub UserControl_Initialize()
        '<EhHeader>
        On Error GoTo UserControl_Initialize_Err
        '</EhHeader>
   
       ' Set defaults
100    m_bEnabled = True
102    m_bCheckBoxes = True
104    m_sDelimiter = ","
   
        '<EhFooter>
        Exit Sub

UserControl_Initialize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_Initialize", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_InitProperties()
        '<EhHeader>
        On Error GoTo UserControl_InitProperties_Err
        '</EhHeader>
100    pInitialise
        '<EhFooter>
        Exit Sub

UserControl_InitProperties_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_InitProperties", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo UserControl_KeyDown_Err
        '</EhHeader>
100    If (m_bDropDownMode) Then
102       fPopupKeyDown KeyCode, Shift
       Else
104       If m_bShowingPopup Then
106          Select Case KeyCode
             Case vbKeyEscape
108             fPopupHide
110          Case Else
                Dim ctl As ddnMultiSelect
112             Set ctl = ObjectFromPtr(m_ptrPopup)
114             ctl.fPopupKeyDown KeyCode, Shift
             End Select
          Else
116          RaiseEvent KeyDown(KeyCode, Shift)
118          Select Case KeyCode
             Case vbKeyDown, vbKeyF4
                Dim rc As RECT
120             GetWindowRect m_hWnd, rc
122             ShowPopup 0, rc.left, rc.bottom
             End Select
          End If
       End If
        '<EhFooter>
        Exit Sub

UserControl_KeyDown_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_KeyDown", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo UserControl_KeyPress_Err
        '</EhHeader>
100    RaiseEvent KeyPress(KeyAscii)
102    If (m_bDropDownMode) Then
       Else
       End If
        '<EhFooter>
        Exit Sub

UserControl_KeyPress_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_KeyPress", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo UserControl_KeyUp_Err
        '</EhHeader>
100    RaiseEvent KeyUp(KeyCode, Shift)
102    If (m_bDropDownMode) Then
       Else
       End If
        '<EhFooter>
        Exit Sub

UserControl_KeyUp_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_KeyUp", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo UserControl_MouseDown_Err
        '</EhHeader>
    Dim ctl As ddnMultiSelect
    Dim lItem As Long
    Dim rc As RECT
       '
100    If (m_bDropDownMode) Then
102       tmrLostMouse.Enabled = False
104       lItem = plHitTest()
106       If (lItem > 0) Then
108          m_tItems(lItem).bChecked = Not (m_tItems(lItem).bChecked)
110          Set ctl = ObjectFromPtr(m_ptrOwner)
112          ctl.fItemChecked(lItem) = m_tItems(lItem).bChecked
114          m_tItems(lItem).bMouseDown = True
116          m_tItems(lItem).bMouseOver = True
118          pPaint
          End If
       Else
120       If (m_bShowingPopup) Then
122          Set ctl = ObjectFromPtr(m_ptrPopup)
124          ctl.fHidePopup
126          m_bShowingPopup = False
128          pPaint
          Else
130          GetWindowRect m_hWnd, rc
132          ShowPopup 0, rc.left, rc.bottom
          End If
       End If
       '
        '<EhFooter>
        Exit Sub

UserControl_MouseDown_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_MouseDown", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo UserControl_MouseMove_Err
        '</EhHeader>
    Dim lItem As Long
    Dim i As Long
    Dim bChanged As Boolean
       '
100    If (m_bDropDownMode) Then
102       lItem = plHitTest()
104       For i = 1 To m_iItemCount
106          If Not (m_tItems(i).bMouseOver) = (i = lItem) Then
108             m_tItems(i).bMouseOver = (i = lItem)
110             bChanged = True
             End If
112       Next i
114       If (bChanged) Then
116          pPaint
          End If
118       If (Button = 0) Then
120          tmrLostMouse.Enabled = True
          End If
       Else
122       If Not (m_bMouseOverButton) Then
124          m_bMouseOverButton = True
126          pPaint
128          If (Button = 0) Then
130             tmrLostMouse.Enabled = True
             End If
          End If
       End If
       '
        '<EhFooter>
        Exit Sub

UserControl_MouseMove_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_MouseMove", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo UserControl_MouseUp_Err
        '</EhHeader>
    Dim bChanged As Boolean
    Dim i As Long
       '
100    If (m_bDropDownMode) Then
102       For i = 1 To m_iItemCount
104          If m_tItems(i).bMouseDown Then
106             m_tItems(i).bMouseDown = False
108             bChanged = True
             End If
110       Next i
112       If (bChanged) Then
114          pPaint
          End If
       Else
   
       End If
       '
        '<EhFooter>
        Exit Sub

UserControl_MouseUp_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_MouseUp", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        '<EhHeader>
        On Error GoTo UserControl_ReadProperties_Err
        '</EhHeader>
100    pInitialise
   
       Dim sFnt As New StdFont
102    sFnt.Name = "Tahoma"
104    sFnt.Size = 8.25
106    Set Font = PropBag.ReadProperty("Font", sFnt)
108    Delimiter = PropBag.ReadProperty("Delimiter", ",")
110    CheckBoxes = PropBag.ReadProperty("CheckBoxes", True)
112    Enabled = PropBag.ReadProperty("Enabled", True)
        '<EhFooter>
        Exit Sub

UserControl_ReadProperties_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_ReadProperties", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
   pPaint
End Sub

Private Sub UserControl_Show()
        '<EhHeader>
        On Error GoTo UserControl_Show_Err
        '</EhHeader>
100    pPaint
        '<EhFooter>
        Exit Sub

UserControl_Show_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_Show", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
   If (m_bDropDownMode) Then
      fHidePopup
   End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        '<EhHeader>
        On Error GoTo UserControl_WriteProperties_Err
        '</EhHeader>
       Dim sFnt As New StdFont
100    sFnt.Name = "Tahoma"
102    sFnt.Size = 8.25
104    PropBag.WriteProperty "Font", Font, sFnt
106    PropBag.WriteProperty "Delimiter", Delimiter, ","
108    PropBag.WriteProperty "CheckBoxes", CheckBoxes, True
110    PropBag.WriteProperty "Enabled", Enabled, True
        '<EhFooter>
        Exit Sub

UserControl_WriteProperties_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ddnMultiSelect.UserControl_WriteProperties", _
                  "ddnMultiSelect component failure"
        '</EhFooter>
End Sub
