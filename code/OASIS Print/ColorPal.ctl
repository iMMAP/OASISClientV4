VERSION 5.00
Begin VB.UserControl ColorPal 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   FillStyle       =   0  'Solid
   MouseIcon       =   "ColorPal.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
End
Attribute VB_Name = "ColorPal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ColorList() As Long
Dim MaxCol      As Integer
Dim TSize       As Integer

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Event ColorSelected(cColor As Long)
Public Event ColorOver(cColor As Long)

Public Sub LoadPalette(Optional PalFile As String)
        '<EhHeader>
        On Error GoTo LoadPalette_Err
        '</EhHeader>
        On Error Resume Next ' GoTo ErrLoad
        Dim FF   As Integer
        Dim tStr As String
        Dim n    As Integer
        Dim cQty As Integer
        Dim Row  As Integer
        Dim Col  As Integer

100     FF = FreeFile

102     If PalFile = "" Or Dir(PalFile) = "" Then PalFile = g_sAppPath & "\Data\User\Default.pal"

104     If Dir(PalFile) <> "" Then
106         Open PalFile For Input As #FF
108         Input #FF, tStr$ 'JASC-PAL

110         If UCase(tStr) <> "JASC-PAL" Then
112             Close #FF
                Exit Sub
            End If

114         Input #FF, tStr$ '0010
116         Input #FF, tStr$ '256 (color qty)
118         cQty = Int(tStr)
120         ReDim ColorList(Int(cQty))
122         n = 0
124         While Not EOF(FF)
126             Input #FF, tStr$
128             ColorList(n) = RGB(Split(tStr, " ")(0), Split(tStr, " ")(1), Split(tStr, " ")(2))
130             n = n + 1
            Wend
132         Close #FF
134         Col = 0
136         Row = 0

138         For n = 0 To cQty - 1
140             UserControl.Line (Col * TSize, Row * TSize)-(Col * TSize + TSize, Row * TSize + TSize), ColorList(n), BF
142             Col = Col + 1

144             If Col = MaxCol Then
146                 Col = 0
148                 Row = Row + 1
                End If

150         Next n

152         UserControl.Width = UserControl.ScaleX((MaxCol * TSize) + 5, vbPixels, vbContainerSize)
154         UserControl.Height = UserControl.ScaleY((cQty / MaxCol * TSize) + TSize + 2, vbPixels, vbContainerSize)
        End If

        Exit Sub
ErrLoad:
156     Close #FF
        '<EhFooter>
        Exit Sub

LoadPalette_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.LoadPalette", _
                  "ColorPal component failure"
        '</EhFooter>
End Sub

Public Property Get ColumnQty() As Integer
        '<EhHeader>
        On Error GoTo ColumnQty_Err
        '</EhHeader>
100     ColumnQty = MaxCol
        '<EhFooter>
        Exit Property

ColumnQty_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.ColumnQty", _
                  "ColorPal component failure"
        '</EhFooter>
End Property

Public Property Let ColumnQty(ByVal iColumnQty As Integer)
        '<EhHeader>
        On Error GoTo ColumnQty_Err
        '</EhHeader>
100     MaxCol = iColumnQty
102     LoadPalette
104     PropertyChanged "ColumnQty"
        '<EhFooter>
        Exit Property

ColumnQty_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.ColumnQty", _
                  "ColorPal component failure"
        '</EhFooter>
End Property

Public Property Get ThumbSize() As Integer
        '<EhHeader>
        On Error GoTo ThumbSize_Err
        '</EhHeader>
100     ThumbSize = TSize
        '<EhFooter>
        Exit Property

ThumbSize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.ThumbSize", _
                  "ColorPal component failure"
        '</EhFooter>
End Property

Public Property Let ThumbSize(ByVal iThumbSize As Integer)
        '<EhHeader>
        On Error GoTo ThumbSize_Err
        '</EhHeader>
100     TSize = iThumbSize
102     LoadPalette
104     PropertyChanged "ThumbSize"
        '<EhFooter>
        Exit Property

ThumbSize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.ThumbSize", _
                  "ColorPal component failure"
        '</EhFooter>
End Property

Private Sub UserControl_Click()
        '<EhHeader>
        On Error GoTo UserControl_Click_Err
        '</EhHeader>
100     RaiseEvent Click
        '<EhFooter>
        Exit Sub

UserControl_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.UserControl_Click", _
                  "ColorPal component failure"
        '</EhFooter>
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
                  "OASISClient.ColorPal.UserControl_DblClick", _
                  "ColorPal component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_InitProperties()
        '<EhHeader>
        On Error GoTo UserControl_InitProperties_Err
        '</EhHeader>
100     TSize = 10
102     MaxCol = 12
        '<EhFooter>
        Exit Sub

UserControl_InitProperties_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.UserControl_InitProperties", _
                  "ColorPal component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
        '<EhHeader>
        On Error GoTo UserControl_MouseDown_Err
        '</EhHeader>
        On Error Resume Next
        Dim tRow   As Integer
        Dim tCol   As Integer
        Dim tColor As Long
        Dim tInd   As Integer

100     If Button = 1 Then
102         tCol = x \ TSize
104         tRow = y \ TSize
106         tInd = tRow * MaxCol + tCol
108         tColor = ColorList(tInd)
110         RaiseEvent ColorSelected(tColor)
        End If

112     RaiseEvent MouseDown(Button, Shift, x, y)
        '<EhFooter>
        Exit Sub

UserControl_MouseDown_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.UserControl_MouseDown", _
                  "ColorPal component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
        '<EhHeader>
        On Error GoTo UserControl_MouseMove_Err
        '</EhHeader>
        On Error Resume Next
        Dim tRow   As Integer
        Dim tCol   As Integer
        Dim tColor As Long
        Dim tInd   As Integer

100     tCol = x \ TSize
102     tRow = y \ TSize
104     tInd = tRow * MaxCol + tCol
106     tColor = ColorList(tInd)
108     RaiseEvent ColorOver(tColor)
110     RaiseEvent MouseMove(Button, Shift, x, y)
        '<EhFooter>
        Exit Sub

UserControl_MouseMove_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.UserControl_MouseMove", _
                  "ColorPal component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
        '<EhHeader>
        On Error GoTo UserControl_MouseUp_Err
        '</EhHeader>
100     RaiseEvent MouseUp(Button, Shift, x, y)
        '<EhFooter>
        Exit Sub

UserControl_MouseUp_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.UserControl_MouseUp", _
                  "ColorPal component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        '<EhHeader>
        On Error GoTo UserControl_ReadProperties_Err
        '</EhHeader>

100     With PropBag
102         MaxCol = .ReadProperty("ColumnQty", 12)
104         TSize = .ReadProperty("Thumbsize", 10)
        End With

        '<EhFooter>
        Exit Sub

UserControl_ReadProperties_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.UserControl_ReadProperties", _
                  "ColorPal component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    LoadPalette
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        '<EhHeader>
        On Error GoTo UserControl_WriteProperties_Err
        '</EhHeader>

100     With PropBag
102         .WriteProperty "ColumnQty", MaxCol, 12
104         .WriteProperty "Thumbsize", TSize, 10
        End With

        '<EhFooter>
        Exit Sub

UserControl_WriteProperties_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.ColorPal.UserControl_WriteProperties", _
                  "ColorPal component failure"
        '</EhFooter>
End Sub

