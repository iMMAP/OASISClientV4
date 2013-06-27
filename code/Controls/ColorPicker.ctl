VERSION 5.00
Begin VB.UserControl ColorPicker 
   AutoRedraw      =   -1  'True
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   137
   ToolboxBitmap   =   "ColorPicker.ctx":0000
End
Attribute VB_Name = "ColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As ColorPickerCommon.RECT) As Long

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private RClr As ColorPickerCommon.RECT
Private RBut As ColorPickerCommon.RECT

Private IsInFocus As Boolean
Private IsButDown As Boolean

Public Enum cpAppearanceConstants
    Flat
    [3D]
End Enum

Private Const m_def_ShowToolTips = True
Private Const m_def_ShowSysColorButton = True
Private Const m_def_ShowDefault = True
Private Const m_def_ShowCustomColors = True
Private Const m_def_ShowMoreColors = True
Private Const m_def_DefaultCaption = "Default"
Private Const m_def_MoreColorsCaption = "More Colors..."
Private Const m_def_BackColor = &H8000000C
Private Const m_def_Appearance = cpAppearanceConstants.[3D]
Private Const m_def_Color = &HFFFFFF
Private Const m_def_DefaultColor = &HFFFFFF

Private m_ShowToolTips As Boolean
Private m_ShowSysColorButton    As Boolean
Private m_ShowDefault           As Boolean
Private m_ShowCustomColors      As Boolean
Private m_ShowMoreColors        As Boolean
Private m_DefaultCaption        As String
Private m_MoreColorsCaption     As String
Private m_BackColor             As OLE_COLOR
Private m_Appearance            As cpAppearanceConstants
Private m_Color                 As OLE_COLOR
Private m_DefaultColor          As OLE_COLOR

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Resize()
Private Sub UserControl_Click()
        '<EhHeader>
        On Error GoTo UserControl_Click_Err
        '</EhHeader>
100     RaiseEvent Click
        '<EhFooter>
        Exit Sub

UserControl_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_Click", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_GotFocus()
        '<EhHeader>
        On Error GoTo UserControl_GotFocus_Err
        '</EhHeader>
100     IsInFocus = True
102     Call RedrawControl
        '<EhFooter>
        Exit Sub

UserControl_GotFocus_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_GotFocus", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_Initialize()
        '<EhHeader>
        On Error GoTo UserControl_Initialize_Err
        '</EhHeader>
100     ScaleMode = vbPixels
        '<EhFooter>
        Exit Sub

UserControl_Initialize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_Initialize", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_LostFocus()
        '<EhHeader>
        On Error GoTo UserControl_LostFocus_Err
        '</EhHeader>
100     IsInFocus = False
102     Call RedrawControl
        '<EhFooter>
        Exit Sub

UserControl_LostFocus_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_LostFocus", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo UserControl_MouseDown_Err
        '</EhHeader>
100     RaiseEvent MouseDown(Button, Shift, x * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY)
    
102     If Button = 1 Then
104         If (x >= RBut.left And x <= RBut.right) And (y >= RBut.top And y <= RBut.bottom) Then
106             IsButDown = True
108             Call RedrawControl
            End If
        End If
        '<EhFooter>
        Exit Sub

UserControl_MouseDown_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_MouseDown", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo UserControl_MouseMove_Err
        '</EhHeader>
100     RaiseEvent MouseMove(Button, Shift, x * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY)
    
102     If IsButDown Then
104         If Not ((x >= RBut.left And x <= RBut.right) And (y >= RBut.top And y <= RBut.bottom)) Then
106             IsButDown = False
108             Call RedrawControl
            End If
        End If
        '<EhFooter>
        Exit Sub

UserControl_MouseMove_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_MouseMove", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo UserControl_MouseUp_Err
        '</EhHeader>
100     RaiseEvent MouseUp(Button, Shift, x * Screen.TwipsPerPixelX, y * Screen.TwipsPerPixelY)
    
102     If Button = 1 Then
104         If IsButDown Then
106             IsButDown = False
108             Call RedrawControl
            End If
        
110         If ((x >= ScaleLeft And x <= ScaleWidth) And (y >= ScaleTop And y <= ScaleHeight)) Then
112             Call ShowPalette
            End If
        End If
        '<EhFooter>
        Exit Sub

UserControl_MouseUp_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_MouseUp", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    RaiseEvent Resize
    If Height < 285 Then Height = 285
    
    Call RedrawControl
End Sub

Private Sub UserControl_DblClick()
        '<EhHeader>
        On Error GoTo UserControl_DblClick_Err
        '</EhHeader>
100     RaiseEvent DblClick
        '<EhFooter>
        Exit Sub

UserControl_DblClick_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_DblClick", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo UserControl_KeyDown_Err
        '</EhHeader>
100     RaiseEvent KeyDown(KeyCode, Shift)
        '<EhFooter>
        Exit Sub

UserControl_KeyDown_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_KeyDown", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo UserControl_KeyPress_Err
        '</EhHeader>
100     RaiseEvent KeyPress(KeyAscii)
        '<EhFooter>
        Exit Sub

UserControl_KeyPress_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_KeyPress", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo UserControl_KeyUp_Err
        '</EhHeader>
100     RaiseEvent KeyUp(KeyCode, Shift)
        '<EhFooter>
        Exit Sub

UserControl_KeyUp_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_KeyUp", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub RedrawControl()
        '<EhHeader>
        On Error GoTo RedrawControl_Err
        '</EhHeader>
        Dim Rct As ColorPickerCommon.RECT
        Dim Brsh As Long, Clr As Long
    
        Dim lX As Long, ty As Long
        Dim rX As Long, by As Long
    
100     lX = ScaleLeft: ty = ScaleTop
102     rX = ScaleWidth: by = ScaleHeight
    
104     Cls
    
106     Call SetRect(Rct, 0, 0, rX, by)
108     Call OleTranslateColor(m_BackColor, ByVal 0&, Clr)
110     Brsh = CreateSolidBrush(Clr)
112     Call FillRect(hdc, Rct, Brsh)
114     If m_Appearance = [3D] Then
116         Call DrawEdge(hdc, Rct, EDGE_SUNKEN, BF_RECT)
        Else
118         Call DrawEdge(hdc, Rct, BDR_SUNKENOUTER, BF_RECT Or BF_FLAT Or BF_MONO)
        End If
120     Call DeleteObject(Brsh)
122     Call DeleteObject(Clr)
    
        'Draw button
        Dim CurFontName As String
124     CurFontName = Font.Name
126     Font.Name = "Marlett"
128     Call OleTranslateColor(vbButtonFace, ByVal 0&, Clr)
130     Brsh = CreateSolidBrush(Clr)
132     If m_Appearance = [3D] Then
134         If IsButDown Then
136             Call SetRect(RBut, rX - 15, 2, rX - 2, by - 2)
138             Call FillRect(hdc, RBut, Brsh)
140             Call DrawEdge(hdc, RBut, EDGE_RAISED, BF_RECT Or BF_FLAT)
142             Call SetRect(Rct, RBut.left + 2, RBut.top, RBut.right, RBut.bottom)
144             Call DrawText(hdc, "6", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
            Else
146             Call SetRect(RBut, rX - 15, 2, rX - 2, by - 2)
148             Call FillRect(hdc, RBut, Brsh)
150             Call DrawEdge(hdc, RBut, EDGE_RAISED, BF_RECT)
152             Call SetRect(Rct, RBut.left, RBut.top, RBut.right, RBut.bottom - 1)
154             Call DrawText(hdc, "6", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
            End If
        Else
156         Call SetRect(RBut, rX - 15, ty, rX, by)
158         Call FillRect(hdc, RBut, Brsh)
160         Call DrawEdge(hdc, RBut, BDR_SUNKENOUTER, BF_RECT Or BF_FLAT)
162         Call SetRect(Rct, RBut.left + 1, RBut.top, RBut.right, RBut.bottom - 1)
164         Call DrawText(hdc, "6", 1&, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
        End If
166     Font.Name = CurFontName
168     Call DeleteObject(Brsh)
170     Call DeleteObject(Clr)
    
        'Draw Color
172     If m_Appearance = [3D] Then
174         Call SetRect(RClr, 4, 4, rX - 17, by - 4)
        Else
176         Call SetRect(RClr, 3, 3, rX - 17, by - 3)
        End If
178     Call OleTranslateColor(m_Color, ByVal 0&, Clr)
180     Brsh = CreateSolidBrush(Clr)
182     Call FillRect(hdc, RClr, Brsh)
184     Call DeleteObject(Brsh)
186     Call DeleteObject(Clr)
    
        'Draw border to the color
188     Call OleTranslateColor(vbGrayText, ByVal 0&, Clr)
190     Brsh = CreateSolidBrush(Clr)
192     Call FrameRect(hdc, RClr, Brsh)
194     Call DeleteObject(Brsh)
196     Call DeleteObject(Clr)
    
        'Draw focus
198     If m_Appearance = [3D] Then
200         Call SetRect(Rct, 6, 6, rX - 19, by - 6)
        Else
202         Call SetRect(Rct, 5, 5, rX - 19, by - 5)
        End If
204     If IsInFocus Then Call DrawFocusRect(hdc, Rct)
    
206     Refresh
        '<EhFooter>
        Exit Sub

RedrawControl_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.RedrawControl", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

Private Sub ShowPalette()
        '<EhHeader>
        On Error GoTo ShowPalette_Err
        '</EhHeader>
        Dim ClrCtrlPos As RECT
    
100     Call GetWindowRect(hWnd, ClrCtrlPos)
    
102     DefClr = m_DefaultColor
104     CurClr = m_Color
    
106     DefCap = m_DefaultCaption
108     MorCap = m_MoreColorsCaption
    
110     ShwDef = m_ShowDefault
112     ShwMor = m_ShowMoreColors
114     ShwCus = m_ShowCustomColors
116     ShwSys = m_ShowSysColorButton
118     ShwTip = m_ShowToolTips

120     Load frmColorPalette
122     With frmColorPalette
124         .left = ClrCtrlPos.left * Screen.TwipsPerPixelX
126         .top = ClrCtrlPos.bottom * Screen.TwipsPerPixelY
128         If (.top + .Height) > Screen.Height Then
130             .top = ClrCtrlPos.top * Screen.TwipsPerPixelY - .Height
            End If
        
132         .Show vbModal
        
134         If Not .IsCanceled Then m_Color = .SelectedColor
136         Call RedrawControl
        End With
138     Unload frmColorPalette
        '<EhFooter>
        Exit Sub

ShowPalette_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowPalette", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
        '<EhHeader>
        On Error GoTo UserControl_InitProperties_Err
        '</EhHeader>
100     m_DefaultColor = m_def_DefaultColor
102     m_Color = m_def_Color
104     m_Appearance = m_def_Appearance
106     m_BackColor = m_def_BackColor
108     m_ShowDefault = m_def_ShowDefault
110     m_ShowCustomColors = m_def_ShowCustomColors
112     m_ShowMoreColors = m_def_ShowMoreColors
114     m_DefaultCaption = m_def_DefaultCaption
116     m_MoreColorsCaption = m_def_MoreColorsCaption
118     m_ShowSysColorButton = m_def_ShowSysColorButton
120     m_ShowToolTips = m_def_ShowToolTips
    
122     Height = 315
        '<EhFooter>
        Exit Sub

UserControl_InitProperties_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_InitProperties", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        '<EhHeader>
        On Error GoTo UserControl_ReadProperties_Err
        '</EhHeader>
100     m_DefaultColor = PropBag.ReadProperty("DefaultColor", m_def_DefaultColor)
102     m_Color = PropBag.ReadProperty("Value", m_def_Color)
104     m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
106     m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
108     m_ShowDefault = PropBag.ReadProperty("ShowDefault", m_def_ShowDefault)
110     m_ShowCustomColors = PropBag.ReadProperty("ShowCustomColors", m_def_ShowCustomColors)
112     m_ShowMoreColors = PropBag.ReadProperty("ShowMoreColors", m_def_ShowMoreColors)
114     m_DefaultCaption = PropBag.ReadProperty("DefaultCaption", m_def_DefaultCaption)
116     m_MoreColorsCaption = PropBag.ReadProperty("MoreColorsCaption", m_def_MoreColorsCaption)
118     m_ShowSysColorButton = PropBag.ReadProperty("ShowSysColorButton", m_def_ShowSysColorButton)
120     m_ShowToolTips = PropBag.ReadProperty("ShowToolTips", m_def_ShowToolTips)
    
122     Call RedrawControl
        '<EhFooter>
        Exit Sub

UserControl_ReadProperties_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_ReadProperties", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        '<EhHeader>
        On Error GoTo UserControl_WriteProperties_Err
        '</EhHeader>
100     Call PropBag.WriteProperty("DefaultColor", m_DefaultColor, m_def_DefaultColor)
102     Call PropBag.WriteProperty("Value", m_Color, m_def_Color)
104     Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
106     Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
108     Call PropBag.WriteProperty("ShowDefault", m_ShowDefault, m_def_ShowDefault)
110     Call PropBag.WriteProperty("ShowCustomColors", m_ShowCustomColors, m_def_ShowCustomColors)
112     Call PropBag.WriteProperty("ShowMoreColors", m_ShowMoreColors, m_def_ShowMoreColors)
114     Call PropBag.WriteProperty("DefaultCaption", m_DefaultCaption, m_def_DefaultCaption)
116     Call PropBag.WriteProperty("MoreColorsCaption", m_MoreColorsCaption, m_def_MoreColorsCaption)
118     Call PropBag.WriteProperty("ShowSysColorButton", m_ShowSysColorButton, m_def_ShowSysColorButton)
120     Call PropBag.WriteProperty("ShowToolTips", m_ShowToolTips, m_def_ShowToolTips)
        '<EhFooter>
        Exit Sub

UserControl_WriteProperties_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.UserControl_WriteProperties", _
                  "ColorPicker component failure"
        '</EhFooter>
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFFFF&
Public Property Get DefaultColor() As OLE_COLOR
Attribute DefaultColor.VB_Description = "Returns/Sets  the default color"
Attribute DefaultColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
        '<EhHeader>
        On Error GoTo DefaultColor_Err
        '</EhHeader>
100     DefaultColor = m_DefaultColor
        '<EhFooter>
        Exit Property

DefaultColor_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.DefaultColor", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let DefaultColor(ByVal New_DefaultColor As OLE_COLOR)
        '<EhHeader>
        On Error GoTo DefaultColor_Err
        '</EhHeader>
100     m_DefaultColor = New_DefaultColor
102     PropertyChanged "DefaultColor"
        '<EhFooter>
        Exit Property

DefaultColor_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.DefaultColor", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFFFF&
Public Property Get color() As OLE_COLOR
Attribute color.VB_Description = "Returns/Sets the selected color"
Attribute color.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute color.VB_UserMemId = 0
        '<EhHeader>
        On Error GoTo color_Err
        '</EhHeader>
100     color = m_Color
        '<EhFooter>
        Exit Property

color_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.color", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let color(ByVal New_Color As OLE_COLOR)
        '<EhHeader>
        On Error GoTo color_Err
        '</EhHeader>
100     m_Color = New_Color
102     PropertyChanged "Value"
    
104     Call RedrawControl
        '<EhFooter>
        Exit Property

color_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.color", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,cpAppearanceConstants.[3D]
Public Property Get Appearance() As cpAppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
        '<EhHeader>
        On Error GoTo Appearance_Err
        '</EhHeader>
100     Appearance = m_Appearance
        '<EhFooter>
        Exit Property

Appearance_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.Appearance", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let Appearance(ByVal New_Appearance As cpAppearanceConstants)
        '<EhHeader>
        On Error GoTo Appearance_Err
        '</EhHeader>
100     m_Appearance = New_Appearance
102     PropertyChanged "Appearance"
    
104     Call RedrawControl
        '<EhFooter>
        Exit Property

Appearance_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.Appearance", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000C&
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
        '<EhHeader>
        On Error GoTo BackColor_Err
        '</EhHeader>
100     BackColor = m_BackColor
        '<EhFooter>
        Exit Property

BackColor_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.BackColor", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
        '<EhHeader>
        On Error GoTo BackColor_Err
        '</EhHeader>
100     m_BackColor = New_BackColor
102     PropertyChanged "BackColor"
    
104     Call RedrawControl
        '<EhFooter>
        Exit Property

BackColor_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.BackColor", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowDefault() As Boolean
Attribute ShowDefault.VB_Description = "Returns/Sets whether default button will be shown or not"
Attribute ShowDefault.VB_ProcData.VB_Invoke_Property = ";Behavior"
        '<EhHeader>
        On Error GoTo ShowDefault_Err
        '</EhHeader>
100     ShowDefault = m_ShowDefault
        '<EhFooter>
        Exit Property

ShowDefault_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowDefault", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let ShowDefault(ByVal New_ShowDefault As Boolean)
        '<EhHeader>
        On Error GoTo ShowDefault_Err
        '</EhHeader>
100     m_ShowDefault = New_ShowDefault
102     PropertyChanged "ShowDefault"
        '<EhFooter>
        Exit Property

ShowDefault_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowDefault", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowCustomColors() As Boolean
Attribute ShowCustomColors.VB_Description = "Returns/Sets whether custom colors will be shown or not"
Attribute ShowCustomColors.VB_ProcData.VB_Invoke_Property = ";Behavior"
        '<EhHeader>
        On Error GoTo ShowCustomColors_Err
        '</EhHeader>
100     ShowCustomColors = m_ShowCustomColors
        '<EhFooter>
        Exit Property

ShowCustomColors_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowCustomColors", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let ShowCustomColors(ByVal New_ShowCustomColors As Boolean)
        '<EhHeader>
        On Error GoTo ShowCustomColors_Err
        '</EhHeader>
100     m_ShowCustomColors = New_ShowCustomColors
102     PropertyChanged "ShowCustomColors"
        '<EhFooter>
        Exit Property

ShowCustomColors_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowCustomColors", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowMoreColors() As Boolean
Attribute ShowMoreColors.VB_Description = "Returns/Sets whether More Colors button will be shown or not"
Attribute ShowMoreColors.VB_ProcData.VB_Invoke_Property = ";Behavior"
        '<EhHeader>
        On Error GoTo ShowMoreColors_Err
        '</EhHeader>
100     ShowMoreColors = m_ShowMoreColors
        '<EhFooter>
        Exit Property

ShowMoreColors_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowMoreColors", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let ShowMoreColors(ByVal New_ShowMoreColors As Boolean)
        '<EhHeader>
        On Error GoTo ShowMoreColors_Err
        '</EhHeader>
100     m_ShowMoreColors = New_ShowMoreColors
102     PropertyChanged "ShowMoreColors"
        '<EhFooter>
        Exit Property

ShowMoreColors_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowMoreColors", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Default
Public Property Get DefaultCaption() As String
Attribute DefaultCaption.VB_Description = "Returns/Sets the caption in default button"
Attribute DefaultCaption.VB_ProcData.VB_Invoke_Property = ";Appearance"
        '<EhHeader>
        On Error GoTo DefaultCaption_Err
        '</EhHeader>
100     DefaultCaption = m_DefaultCaption
        '<EhFooter>
        Exit Property

DefaultCaption_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.DefaultCaption", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let DefaultCaption(ByVal New_DefaultCaption As String)
        '<EhHeader>
        On Error GoTo DefaultCaption_Err
        '</EhHeader>
100     m_DefaultCaption = New_DefaultCaption
102     PropertyChanged "DefaultCaption"
        '<EhFooter>
        Exit Property

DefaultCaption_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.DefaultCaption", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,More Colors...
Public Property Get MoreColorsCaption() As String
Attribute MoreColorsCaption.VB_Description = "Returns/Sets the caption in the More button"
Attribute MoreColorsCaption.VB_ProcData.VB_Invoke_Property = ";Appearance"
        '<EhHeader>
        On Error GoTo MoreColorsCaption_Err
        '</EhHeader>
100     MoreColorsCaption = m_MoreColorsCaption
        '<EhFooter>
        Exit Property

MoreColorsCaption_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.MoreColorsCaption", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let MoreColorsCaption(ByVal New_MoreColorsCaption As String)
        '<EhHeader>
        On Error GoTo MoreColorsCaption_Err
        '</EhHeader>
100     m_MoreColorsCaption = New_MoreColorsCaption
102     PropertyChanged "MoreColorsCaption"
        '<EhFooter>
        Exit Property

MoreColorsCaption_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.MoreColorsCaption", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowSysColorButton() As Boolean
Attribute ShowSysColorButton.VB_ProcData.VB_Invoke_Property = ";Behavior"
        '<EhHeader>
        On Error GoTo ShowSysColorButton_Err
        '</EhHeader>
100     ShowSysColorButton = m_ShowSysColorButton
        '<EhFooter>
        Exit Property

ShowSysColorButton_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowSysColorButton", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let ShowSysColorButton(ByVal New_ShowSysColorButton As Boolean)
        '<EhHeader>
        On Error GoTo ShowSysColorButton_Err
        '</EhHeader>
100     m_ShowSysColorButton = New_ShowSysColorButton
102     PropertyChanged "ShowSysColorButton"
        '<EhFooter>
        Exit Property

ShowSysColorButton_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowSysColorButton", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowToolTips() As Boolean
Attribute ShowToolTips.VB_ProcData.VB_Invoke_Property = ";Behavior"
        '<EhHeader>
        On Error GoTo ShowToolTips_Err
        '</EhHeader>
100     ShowToolTips = m_ShowToolTips
        '<EhFooter>
        Exit Property

ShowToolTips_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowToolTips", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

Public Property Let ShowToolTips(ByVal New_ShowToolTips As Boolean)
        '<EhHeader>
        On Error GoTo ShowToolTips_Err
        '</EhHeader>
100     m_ShowToolTips = New_ShowToolTips
102     PropertyChanged "ShowToolTips"
        '<EhFooter>
        Exit Property

ShowToolTips_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPicker.ShowToolTips", _
                  "ColorPicker component failure"
        '</EhFooter>
End Property

