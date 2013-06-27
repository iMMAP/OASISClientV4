VERSION 5.00
Begin VB.Form frmColorPalette 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmColorPalette.frx":0000
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias "ChooseColorA" (pChoosecolor As udtCHOOSECOLOR) As Long
Private Type udtCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const CC_FULLOPEN = &H2
Private Const CC_ANYCOLOR = &H100

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_HIDE = 0

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Module specific variable declarations
Private Type cpColorInformation
    clr As OLE_COLOR
    Rct As ColorPickerCommon.RECT
    Tip As String
End Type

Private Clrs(60) As cpColorInformation

Private IsSystemColors As Boolean
Private MouseButId As Integer
Private MouseDownButId As Integer
Private CurClrButId As Integer

Private Const NorClrVal = "&HFFFFFF&HC0C0FF&HC0E0FF&HC0FFFF&HC0FFC0&HFFFFC0&HFFC0C0&HFFC0FF" & _
                          "&HE0E0E0&H8080FF&H80C0FF&H80FFFF&H80FF80&HFFFF80&HFF8080&HFF80FF" & _
                          "&HC0C0C0&H0000FF&H0080FF&H00FFFF&H00FF00&HFFFF00&HFF0000&HFF00FF" & _
                          "&H808080&H0000C0&H0040C0&H00C0C0&H00C000&HC0C000&HC00000&HC000C0" & _
                          "&H404040&H000080&H004080&H008080&H008000&H808000&H800000&H800080" & _
                          "&H000000&H000040&H404080&H004040&H004000&H404000&H400000&H400040"
Private Const SysClrVal = "&H80000000&H80000001&H80000002&H80000003&H80000004&H80000005" & _
                          "&H80000006&H80000007&H80000008&H80000009&H8000000A&H8000000B" & _
                          "&H8000000C&H8000000D&H8000000E&H8000000F&H80000010&H80000011" & _
                          "&H80000012&H80000013&H80000014&H80000015&H80000016&H80000017" & _
                          "&H80000018"
Private Const NorClrTip = ""
Private Const SysClrTip = "Scroll Bars            " & _
                          "Desktop                " & _
                          "Active Title Bar       " & _
                          "Inactive Titl Bar      " & _
                          "Menu Bar               " & _
                          "Window Background      " & _
                          "Window Frame           " & _
                          "Menu Text              " & _
                          "Window Text            " & _
                          "Active Title Bar Text  " & _
                          "Active Border          " & _
                          "Inactive Border        " & _
                          "Application Workspace  " & _
                          "Highlight              " & _
                          "Highlight Text         " & _
                          "Button Face            " & _
                          "Button Shadow          " & _
                          "Disabled Text          " & _
                          "Button Text            " & _
                          "Inactive Title Bar Text" & _
                          "Button Highlight       " & _
                          "Button Dark Shadow     " & _
                          "Button Light Shadow    " & _
                          "ToolTip Text           " & _
                          "ToolTip                "
Private Const OtherTip = "Normal Colors    " & _
                         "System Colors    " & _
                         "Show Color Dialog"

Private pl As Long, Pt12 As Long

Private Const TipTmr1 = 1
Private Const TipTmr2 = 2
Private IsTmr1Active As Boolean
Private IsTmr2Active As Boolean
Private TipButId As Integer

Public SelectedColor As OLE_COLOR
Public IsCanceled As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo Form_KeyDown_Err
        '</EhHeader>
100     If (KeyCode = vbKeyEscape) Then
102         Me.Hide
        End If
        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.Form_KeyDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim R As ColorPickerCommon.RECT
    
100     Me.ScaleMode = vbPixels
102     Me.Font.Name = "Arial"
    
104     Call SetCapture(hwnd)
    
106     IsSystemColors = False
108     MouseButId = -1
110     MouseDownButId = -1
112     IsCanceled = True
    
114     Call Initialize
    
116     Width = (pl + (8 * 16) + 7 + 4) * Screen.TwipsPerPixelX
118     Height = (Pt12 + 4) * Screen.TwipsPerPixelY
    
120     Call SetRect(R, 0, 0, ScaleWidth, ScaleHeight)
122     Call DrawEdge(hdc, R, BDR_RAISEDINNER, BF_RECT)
    
124     Load frmTip
    
126     If Not g_sLanguage = "" Then
128         If Not m_Cnn.State = adStateClosed Then
130             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo Form_MouseDown_Err
        '</EhHeader>
100     If Not (Button = 1) Then Exit Sub
    
102     If Not (MouseButId = -1) Then
104         If (MouseButId = 58) Or (MouseButId = 59) Or (MouseButId = 60) Then
106             Call DrawButton(MouseButId, 1)
            End If
108         Call DrawButEdge(MouseButId, 2)
        
110         MouseDownButId = MouseButId
        
112         Call ShowTip(False)
        End If
        '<EhFooter>
        Exit Sub

Form_MouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.Form_MouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo Form_MouseMove_Err
        '</EhHeader>
        Dim i As Integer
        Dim IsMouseOnBut As Boolean
    
100     If Not (MouseDownButId = -1) Then
            Exit Sub
        End If
    
102     For i = 1 To 60
104         IsMouseOnBut = (x >= Clrs(i).Rct.Left And y >= Clrs(i).Rct.Top) And (x <= Clrs(i).Rct.Right And y <= Clrs(i).Rct.Bottom)
106         If IsMouseOnBut Then
                Exit For
            End If
108     Next i
    
110     If (Not MouseButId = -1) And (Not MouseButId = i) Then
112         Call DrawButEdge(MouseButId, 0)
114         MouseButId = -1
116         Call ShowTip(False)
        End If
    
118     If IsMouseOnBut And (Not MouseButId = i) Then
120         MouseButId = i
122         Call DrawButEdge(MouseButId, 1)
        
124         If ShwTip Then
126             Call SetTimer(Me.hwnd, CLng(TipTmr1), 1000, AddressOf Timer)
128             IsTmr1Active = True
            End If
        End If
    
130     If Not IsMouseOnBut Then
132         If IsTmr1Active Then
134             Call KillTimer(Me.hwnd, CLng(TipTmr1))
136             IsTmr1Active = False
            End If
        End If
    
    '    If (i >= 1) And (i <= 57) Then
    '        If Not Me.MousePointer = vbCustom Then Me.MousePointer = vbCustom
    '    Else
    '        If Not Me.MousePointer = vbDefault Then Me.MousePointer = vbDefault
    '    End If
        '<EhFooter>
        Exit Sub

Form_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.Form_MouseMove " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        '<EhHeader>
        On Error GoTo Form_MouseUp_Err
        '</EhHeader>
        Dim IsMouseOver As Boolean
    
100     If Not (MouseDownButId = -1) Then
102         If (MouseDownButId = 58) Or (MouseDownButId = 59) Or (MouseDownButId = 60) Then
104             Call DrawButton(MouseDownButId, 0)
            End If
106         Call DrawButEdge(MouseDownButId, 1)
        
108         If IsMouseOnBut(MouseDownButId) Then
110             Call DoAction(MouseDownButId)
            End If
        
112         MouseDownButId = -1
        End If
    
114     IsMouseOver = x >= 0 And y >= 0 And x <= ScaleWidth And y <= ScaleHeight
116     If IsMouseOver Then
118         Call SetCapture(Me.hwnd)
        Else
120         Call ReleaseCapture
122         Call Form_KeyDown(vbKeyEscape, 0)
        End If
        '<EhFooter>
        Exit Sub

Form_MouseUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.Form_MouseUp " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub DrawButEdge(ClrId As Integer, EdgeStyle As Integer)
        '<EhHeader>
        On Error GoTo DrawButEdge_Err
        '</EhHeader>
100     Select Case EdgeStyle
            Case 0: Call DrawEdge(hdc, Clrs(ClrId).Rct, BDR_RAISEDINNER, BF_RECT Or BF_FLAT)
102         Case 1: Call DrawEdge(hdc, Clrs(ClrId).Rct, BDR_RAISEDINNER, BF_RECT)
104         Case 2: Call DrawEdge(hdc, Clrs(ClrId).Rct, BDR_SUNKENOUTER, BF_RECT)
        End Select
    
106     Refresh
        '<EhFooter>
        Exit Sub

DrawButEdge_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.DrawButEdge " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Initialize()
        '<EhHeader>
        On Error GoTo Initialize_Err
        '</EhHeader>
        Dim i As Integer
        Dim LPos As Long, TPos As Long
        Dim FrmBkClr As Long
    
100     pl = 4: Pt12 = 0
    
102     If ShwDef Then
104         Call SetRect(Clrs(1).Rct, pl, (Pt12 + 4), pl + 7 + 16 * 8, (Pt12 + 4) + 22)
106         Pt12 = (Pt12 + 4) + 22
        End If
    
108     For i = 2 To 49
110         LPos = (((i - 2) Mod 8) + pl) + (((i - 2) Mod 8) * 16)
112         TPos = (Int((i - 2) / 8) + (Pt12 + 4)) + (Int((i - 2) / 8) * 16)
114         Call SetRect(Clrs(i).Rct, LPos, TPos, LPos + 16, TPos + 16)
116     Next i
118     Pt12 = (Pt12 + 4) + (6 * 16) + 5

120     If ShwCus Then
122         FrmBkClr = Me.ForeColor
124         Me.ForeColor = vb3DShadow
126         CurrentX = 4: CurrentY = Pt12 + 2
128         Line -(16 * 8 + 4 + 7, CurrentY)
130         Me.ForeColor = vb3DHighlight
132         CurrentX = 4: CurrentY = Pt12 + 2 + 1
134         Line -(16 * 8 + 4 + 7, CurrentY)
136         Me.ForeColor = FrmBkClr
        
138         Pt12 = Pt12 + 2 + 1
        
140         For i = 50 To 57
142             LPos = (((i - 50) Mod 8) + 4) + (((i - 50) Mod 8) * 16)
144             TPos = (Int((i - 50) / 8) + (Pt12 + 2)) + (Int((i - 50) / 8) * 16)
146             Call SetRect(Clrs(i).Rct, LPos, TPos, LPos + 16, TPos + 16)
148         Next i
        
150         Pt12 = (Pt12 + 2) + 16
        End If
    
152     If ShwMor Or ShwSys Then
154         FrmBkClr = Me.ForeColor
156         Me.ForeColor = vb3DShadow
158         CurrentX = 4: CurrentY = Pt12 + 2
160         Line -(16 * 8 + 4 + 7, CurrentY)
162         Me.ForeColor = vb3DHighlight
164         CurrentX = 4: CurrentY = Pt12 + 2 + 1
166         Line -(16 * 8 + 4 + 7, CurrentY)
168         Me.ForeColor = FrmBkClr
        
170         Pt12 = Pt12 + 2 + 1
        End If

172     If ShwSys Then
174         For i = 58 To 59
176             LPos = (((i - 58) Mod 2) * 7 + pl) + (((i - 58) Mod 2) * 64)
178             TPos = (Int((i - 58) / 2) + (Pt12 + 2)) + (Int((i - 58) / 2) * 20)
180             Call SetRect(Clrs(i).Rct, LPos, TPos, LPos + 64, TPos + 20)
182         Next i
        
184         Pt12 = (Pt12 + 2) + 20
        End If
    
186     If ShwMor Then
188         Call SetRect(Clrs(60).Rct, pl, (Pt12 + 2), 4 + 7 + 16 * 8, (Pt12 + 2) + 20)
190         Pt12 = (Pt12 + 2) + 20
        End If
    
192     For i = 1 To 60
194         Call DrawButton(i, 0)
196     Next i
        '<EhFooter>
        Exit Sub

Initialize_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.Initialize " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub DrawButton(ButId As Integer, State As Integer)
        '<EhHeader>
        On Error GoTo DrawButton_Err
        '</EhHeader>
        Dim clr As Long, Brsh As Long
        Dim R As ColorPickerCommon.RECT
    
100     Call OleTranslateColor(Me.BackColor, ByVal 0&, clr)
102     Brsh = CreateSolidBrush(clr)
104     Call FillRect(hdc, Clrs(ButId).Rct, Brsh)
106     Call DeleteObject(clr)
108     Call DeleteObject(Brsh)
    
110     Select Case ButId
            Case 1
112             If Not ShwDef Then Exit Sub
            
114             Clrs(1).clr = DefClr
116             Clrs(1).Tip = "Default"
            
118             Call SetRect(R, Clrs(1).Rct.Left + 3, Clrs(1).Rct.Top + 3, Clrs(1).Rct.Right - 3, Clrs(1).Rct.Bottom - 3)
120             Call OleTranslateColor(vbGrayText, ByVal 0&, clr)
122             Brsh = CreateSolidBrush(clr)
124             Call FrameRect(hdc, R, Brsh)
126             Call DeleteObject(Brsh)
128             Call DeleteObject(clr)
            
130             Call SetRect(R, Clrs(1).Rct.Left + 5, Clrs(1).Rct.Top + 5, Clrs(1).Rct.Left + 5 + 12, Clrs(1).Rct.Top + 5 + 12)
132             Call OleTranslateColor(Clrs(1).clr, ByVal 0&, clr)
134             Brsh = CreateSolidBrush(clr)
136             Call FillRect(hdc, R, Brsh)
138             Call DeleteObject(Brsh)
140             Call DeleteObject(clr)
142             Call OleTranslateColor(vbGrayText, ByVal 0&, clr)
144             Brsh = CreateSolidBrush(clr)
146             Call FrameRect(hdc, R, Brsh)
148             Call DeleteObject(Brsh)
150             Call DeleteObject(clr)
            
152             Call SetRect(R, Clrs(1).Rct.Left + 5 + 12, Clrs(1).Rct.Top + 3, Clrs(1).Rct.Right - 2, Clrs(1).Rct.Bottom - 3)
154             Call DrawText(hdc, DefCap, Len(DefCap), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
156         Case 2 To 49
158             If Not IsSystemColors Then
160                 Clrs(ButId).clr = CLng(Mid(NorClrVal, (ButId - 2) * 8 + 1, 8))
162                 Clrs(ButId).Tip = ""
                Else
164                 If (ButId <= 26) Then
166                     Clrs(ButId).clr = CLng(Mid(SysClrVal, (ButId - 2) * 10 + 1, 10))
168                     Clrs(ButId).Tip = Trim(Mid(SysClrTip, (ButId - 2) * 23 + 1, 23))
                    Else
170                     Clrs(ButId).clr = &HFFFFFF
172                     Clrs(ButId).Tip = ""
                    End If
                End If
            
174             Call SetRect(R, Clrs(ButId).Rct.Left + 2, Clrs(ButId).Rct.Top + 2, Clrs(ButId).Rct.Right - 2, Clrs(ButId).Rct.Bottom - 2)
176             Call OleTranslateColor(Clrs(ButId).clr, ByVal 0&, clr)
178             Brsh = CreateSolidBrush(clr)
180             Call FillRect(hdc, R, Brsh)
182             Call DeleteObject(Brsh)
184             Call DeleteObject(clr)
            
186             Call OleTranslateColor(vbGrayText, ByVal 0&, clr)
188             Brsh = CreateSolidBrush(clr)
190             Call FrameRect(hdc, R, Brsh)
192             Call DeleteObject(Brsh)
194             Call DeleteObject(clr)
196         Case 50 To 57
198             If Not ShwCus Then Exit Sub
            
200             Clrs(ButId).clr = &HFFFFFF
202             Clrs(ButId).Tip = "Custom Color " & Trim(str(ButId - 49))
            
204             If Not (LastSavedCustClr = 0) Then
206                 If (UBound(CustClrs) >= (ButId - 49)) Then
208                     Clrs(ButId).clr = CustClrs(ButId - 49)
                    End If
                End If
            
210             Call OleTranslateColor(Clrs(ButId).clr, ByVal 0&, clr)
212             Brsh = CreateSolidBrush(clr)
214             Call SetRect(R, Clrs(ButId).Rct.Left + 2, Clrs(ButId).Rct.Top + 2, Clrs(ButId).Rct.Right - 2, Clrs(ButId).Rct.Bottom - 2)
216             Call FillRect(hdc, R, Brsh)
218             Call DeleteObject(Brsh)
220             Call DeleteObject(clr)
            
222             Call OleTranslateColor(vbGrayText, ByVal 0&, clr)
224             Brsh = CreateSolidBrush(clr)
226             Call FrameRect(hdc, R, Brsh)
228             Call DeleteObject(Brsh)
230             Call DeleteObject(clr)
232         Case 58 To 60
                Dim TmpStr As String
234             Select Case ButId
                    Case 58: TmpStr = "Normal": If Not ShwSys Then Exit Sub
236                 Case 59: TmpStr = "System": If Not ShwSys Then Exit Sub
238                 Case 60: TmpStr = MorCap: If Not ShwMor Then Exit Sub
                End Select
            
240             If State = 0 Then
242                 Call SetRect(R, Clrs(ButId).Rct.Left, Clrs(ButId).Rct.Top, Clrs(ButId).Rct.Right, Clrs(ButId).Rct.Bottom)
                Else
244                 Call SetRect(R, Clrs(ButId).Rct.Left + 1, Clrs(ButId).Rct.Top + 1, Clrs(ButId).Rct.Right + 1, Clrs(ButId).Rct.Bottom + 1)
                End If
246             Call DrawText(hdc, TmpStr, CLng(Len(TmpStr)), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
248             Clrs(ButId).Tip = Trim(Mid(OtherTip, (ButId - 58) * 17 + 1, 17))
        End Select
    
250     Refresh
        '<EhFooter>
        Exit Sub

DrawButton_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.DrawButton " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub DoAction(ButId As Integer)
        '<EhHeader>
        On Error GoTo DoAction_Err
        '</EhHeader>
        Dim i As Integer
    
100     Select Case ButId
            Case 1 To 57
102             SelectedColor = Clrs(ButId).clr
104             IsCanceled = False
106             Call Form_KeyDown(vbKeyEscape, 0)
108         Case 58
110             If IsSystemColors Then
112                 IsSystemColors = False
114                 For i = 2 To 49
116                     Call DrawButton(i, 0)
118                 Next i
                End If
120         Case 59
122             If Not IsSystemColors Then
124                 IsSystemColors = True
126                 For i = 2 To 49
128                     Call DrawButton(i, 0)
130                 Next i
                End If
132         Case 60
134             SelectedColor = ShowColor
136             If Not SelectedColor = -1 Then
138                 Call SaveCustClr(SelectedColor)
140                 IsCanceled = False
                Else
142                 IsCanceled = True
                End If
144             Call Form_KeyDown(vbKeyEscape, 0)
        End Select
        '<EhFooter>
        Exit Sub

DoAction_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.DoAction " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function IsMouseOnBut(ButId As Integer) As Boolean
        '<EhHeader>
        On Error GoTo IsMouseOnBut_Err
        '</EhHeader>
        Dim Ptq As POINTAPI
    
100     Call GetCursorPos(Ptq)
102     Call ScreenToClient(Me.hwnd, Ptq)
104     IsMouseOnBut = (Ptq.x >= Clrs(ButId).Rct.Left And Ptq.x <= Clrs(ButId).Rct.Right) And _
                       (Ptq.y >= Clrs(ButId).Rct.Top And Ptq.y <= Clrs(ButId).Rct.Bottom)
        '<EhFooter>
        Exit Function

IsMouseOnBut_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.IsMouseOnBut " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function ShowColor() As Long
        '<EhHeader>
        On Error GoTo ShowColor_Err
        '</EhHeader>
        Dim ClrInf As udtCHOOSECOLOR
        Static CustomColors(64) As Byte
        Dim i As Integer
    
100     For i = LBound(CustomColors) To UBound(CustomColors)
102         CustomColors(i) = 0
104     Next i
    
106     With ClrInf
108         .lStructSize = Len(ClrInf)              'Size of the structure
110         .hWndOwner = Me.hwnd                    'Handle of owner window
112         .hInstance = App.hInstance              'Instance of application
114         .lpCustColors = StrConv(CustomColors, vbUnicode)       'Array of 16 byte values
116         .Flags = CC_FULLOPEN                    'Flags to open in full mode
        End With
    
118     If Not ChooseColor(ClrInf) = 0 Then
120         ShowColor = ClrInf.rgbResult
        Else
122         ShowColor = -1
        End If
        '<EhFooter>
        Exit Function

ShowColor_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.ShowColor " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub SaveCustClr(ClrVal As OLE_COLOR)
        '<EhHeader>
        On Error GoTo SaveCustClr_Err
        '</EhHeader>
100     If (LastSavedCustClr = 0) Then
102         ReDim Preserve CustClrs(1) As OLE_COLOR
        Else
104         If (UBound(CustClrs) < 8) Then
106             ReDim Preserve CustClrs(UBound(CustClrs) + 1) As OLE_COLOR
            End If
        End If
    
108     LastSavedCustClr = LastSavedCustClr + 1
110     If (LastSavedCustClr > 8) Then LastSavedCustClr = 1
    
112     CustClrs(LastSavedCustClr) = ClrVal
        '<EhFooter>
        Exit Sub

SaveCustClr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.SaveCustClr " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
        Dim i As Integer
    
100     For i = 1 To 60
102         Call SetRectEmpty(Clrs(i).Rct)
104     Next i
    
106     If IsTmr1Active Then
108         Call KillTimer(Me.hwnd, CLng(TipTmr1))
110         IsTmr1Active = False
        End If
    
112     If IsTmr2Active Then
114         Call KillTimer(Me.hwnd, CLng(TipTmr2))
116         IsTmr2Active = False
        End If

118     Unload frmTip
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub TipTimer(hwnd As Long, uMsg As Long, idEvent As Long, dwTime As Long)
        '<EhHeader>
        On Error GoTo TipTimer_Err
        '</EhHeader>
100     Select Case idEvent
            Case 1
102             Call ShowTip(True)
            
104             Call KillTimer(Me.hwnd, CLng(TipTmr1))
106             IsTmr1Active = False
108         Case 2
110             Call ShowTip(False)
        End Select
        '<EhFooter>
        Exit Sub

TipTimer_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.TipTimer " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ShowTip(State As Boolean)
        '<EhHeader>
        On Error GoTo ShowTip_Err
        '</EhHeader>
100     If State Then
            Dim Rct As ColorPickerCommon.RECT
            Dim Pta As POINTAPI
            Dim TipTxt As String
        
            'Store the tip text in a variable
102         TipTxt = Clrs(MouseButId).Tip
104         If TipTxt = "" Then Exit Sub
        
            'Clear Tip Form
106         frmTip.Cls
        
            'Draw Tip text and position the Tip Form
108         Call GetCursorPos(Pta)
110         Call SetRect(Rct, 0, 0, frmTip.ScaleWidth, frmTip.ScaleHeight)
112         Call DrawText(frmTip.hdc, TipTxt, CLng(Len(TipTxt)), Rct, DT_CALCRECT)
114         Call SetRect(Rct, 0, 0, Rct.Right + 8, Rct.Bottom + 6)
116         Call DrawText(frmTip.hdc, TipTxt, CLng(Len(TipTxt)), Rct, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER)
118         Call DrawEdge(frmTip.hdc, Rct, BDR_RAISEDINNER, BF_RECT)
120         frmTip.Move (Pta.x + 2) * Screen.TwipsPerPixelX, (Pta.y + 20) * Screen.TwipsPerPixelY, _
                        Rct.Right * Screen.TwipsPerPixelX, Rct.Bottom * Screen.TwipsPerPixelY
122         frmTip.ZOrder
124         frmTip.Refresh
126         Call ShowWindow(frmTip.hwnd, SW_SHOWNOACTIVATE)
        
            'Set Timer 2 for the duration of tip
128         Call SetTimer(Me.hwnd, CLng(TipTmr2), 4000, AddressOf Timer)
130         IsTmr2Active = True
        Else
            On Error Resume Next
        
            'Hide Tip Form
132         Call ShowWindow(frmTip.hwnd, SW_HIDE)
        
            'Kill Timer 2 if it is active
134         If IsTmr2Active Then
136             Call KillTimer(Me.hwnd, CLng(TipTmr2))
138             IsTmr2Active = False
            End If
        
            'Kill Timer 1 if it is active
140         If IsTmr1Active Then
142             Call KillTimer(Me.hwnd, CLng(TipTmr1))
144             IsTmr1Active = False
            End If
        End If
        '<EhFooter>
        Exit Sub

ShowTip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmColorPalette.ShowTip " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
