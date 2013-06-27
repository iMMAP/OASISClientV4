VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{F8F9FBF9-12B5-11D4-8ED3-00E07D815373}#1.0#0"; "MBScroll.ocx"
Begin VB.UserControl OASISDynamicRSEditor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   ScaleHeight     =   4770
   ScaleWidth      =   8250
   Begin XpressEditorsLibCtl.dxMemoEdit dxMemoEdit0 
      Height          =   510
      Index           =   0
      Left            =   2370
      OleObjectBlob   =   "OASISDynamicRSEditor.ctx":0000
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   1635
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   4770
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8250
      _cx             =   14552
      _cy             =   8414
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   0
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   6
      GridCols        =   3
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"OASISDynamicRSEditor.ctx":00FC
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   4770
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   8250
         _cx             =   14552
         _cy             =   8414
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483639
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Picture         =   "OASISDynamicRSEditor.ctx":0189
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   0
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"OASISDynamicRSEditor.ctx":3874E
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin MBScroller.Scroller Scroller1 
            Height          =   4770
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   8250
            _ExtentX        =   14552
            _ExtentY        =   8414
            Appearance      =   0
            BorderStyle     =   0
         End
         Begin CONTROLSLibCtl.dxProgressBar dxProgressBar1 
            Height          =   4770
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   8250
            _Version        =   65536
            _cx             =   14552
            _cy             =   8414
            ForeColor       =   0
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MinPos          =   0
            MaxPos          =   100
            Pos             =   50
            Step            =   10
            ShowText        =   -1  'True
            Orientation     =   0
            StartColor      =   16711680
            EndColor        =   16777215
            DrawBorderStyle =   1
            ShowTextStyle   =   0
            DrawBarStyle    =   2
            DrawBarBorderStyle=   2
         End
      End
      Begin XpressEditorsLibCtl.dxTextEdit dxTextEdit0 
         Height          =   315
         Index           =   0
         Left            =   5535
         OleObjectBlob   =   "OASISDynamicRSEditor.ctx":38788
         TabIndex        =   1
         Top             =   0
         Width           =   2715
      End
      Begin XpressEditorsLibCtl.dxTextEdit dxIntegerEdit0 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   5535
         OleObjectBlob   =   "OASISDynamicRSEditor.ctx":387FE
         TabIndex        =   2
         Top             =   810
         Width           =   2715
      End
      Begin XpressEditorsLibCtl.dxTextEdit dxDoubleEdit0 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   5535
         OleObjectBlob   =   "OASISDynamicRSEditor.ctx":38874
         TabIndex        =   3
         Top             =   1605
         Width           =   2715
      End
      Begin XpressEditorsLibCtl.dxDateEdit dxDateEdit0 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MMM-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   5535
         OleObjectBlob   =   "OASISDynamicRSEditor.ctx":388EA
         TabIndex        =   4
         Top             =   3225
         Width           =   2715
      End
      Begin XpressEditorsLibCtl.dxCheckEdit dxCheckEdit0 
         Height          =   315
         Index           =   0
         Left            =   5535
         OleObjectBlob   =   "OASISDynamicRSEditor.ctx":3898A
         TabIndex        =   5
         Top             =   4020
         Width           =   2715
      End
      Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpEdit0 
         Height          =   315
         Index           =   0
         Left            =   5535
         OleObjectBlob   =   "OASISDynamicRSEditor.ctx":38A68
         TabIndex        =   10
         Top             =   2415
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Label Label0 
         Caption         =   "Label"
         Height          =   735
         Index           =   0
         Left            =   2775
         TabIndex        =   8
         Top             =   810
         Width           =   2700
      End
   End
End
Attribute VB_Name = "OASISDynamicRSEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Enum SIMPLE_FIELD_TYPES
    FIELD_TEXT = &H1
    FIELD_INT = &H2
    FIELD_DOUBLE = &H4
    FIELD_MEMO = &H6 '
    FIELD_PICK = &H10
    FIELD_DATE = &H20
    FIELD_CHECKBOX = &H80
End Enum

Private Type FIELD_SPEC
    Name As String
    Caption As String
    FieldType As DataTypeEnum
    SimpleFieldType As SIMPLE_FIELD_TYPES
    DefinedSize As Long
End Type

Private mRS As ADODB.Recordset
Private mConn As ADODB.Connection
Private Fields() As FIELD_SPEC
Private cControl() As Control

Private lCountOfFields As Long
Private lCountOfRecords As Long
Private lCountOfCharsInLabel As Long
Private iCountOfGUIDs As Integer

Private sGUID1Value As String
Private sGUID2Value As String
Private sTableNamePrefix As String
Private sExcludedFields As String

Private bLoaded As Boolean
Private bLocked As Boolean
Private bIsGeoTable As Boolean

Public bChangeMade As Boolean
Public Event SoundChangeMade()

Public Function GetIsLocked() As Boolean
        '<EhHeader>
        On Error GoTo GetIsLocked_Err
        '</EhHeader>
100     GetIsLocked = bLocked
        '<EhFooter>
        Exit Function

GetIsLocked_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.GetIsLocked", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Function

Private Sub C1Elastic1_ResizeChildren()
        '<EhHeader>
        On Error GoTo C1Elastic1_ResizeChildren_Err
        '</EhHeader>

100     ResizeFields

        '<EhFooter>
        Exit Sub

C1Elastic1_ResizeChildren_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.C1Elastic1_ResizeChildren", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Public Sub Reset()
        '<EhHeader>
        On Error GoTo Reset_Err
        '</EhHeader>

100     Set mRS = Nothing
102     Set mConn = Nothing
104     DeleteFields
106     bLoaded = False
108     bChangeMade = False
110     lCountOfFields = 0
112     lCountOfRecords = 0
114     dxProgressBar1.Visible = False
116     Scroller1.Visible = False
118     C1Elastic2.Refresh
        '<EhFooter>
        Exit Sub

Reset_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.Reset", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Public Sub Init(oRS As ADODB.Recordset, _
                oConn As ADODB.Connection, _
                sTablePrefix As String, _
                Optional lBackColour As Long = 666666, _
                Optional lTextColour As Long = 666666, _
                Optional iNumOfGUIDs As Integer = 1, _
                Optional bIsAGeoTable As Boolean = False, Optional sExcludes As String = "")
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>

100     Set mConn = oConn
102     Set mRS = oRS

104     bIsGeoTable = bIsAGeoTable

106     If bLoaded Then Call DeleteFields
108     iCountOfGUIDs = iNumOfGUIDs
110     Call GetRSInfo

112     sTableNamePrefix = sTablePrefix
114     dxProgressBar1.Visible = False
116     Scroller1.Visible = False

118     If Not lBackColour = 666666 Then Scroller1.BackColor = lBackColour
120     If Not lTextColour = 666666 Then Label0(0).ForeColor = lTextColour
        
122     sExcludedFields = sExcludes
124     DisplayFields
126     bLoaded = True
128     bChangeMade = False

        '<EhFooter>
        Exit Sub

Init_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.Init", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub GetRSInfo()
        '<EhHeader>
        On Error GoTo GetRSInfo_Err
        '</EhHeader>

        Dim i As Long
        Dim lNumOfCharsInLabel As Long
        
100     lCountOfCharsInLabel = 0
102     lCountOfFields = mRS.Fields.Count
104     lCountOfRecords = mRS.RecordCount
106     dxProgressBar1.Visible = True
108     dxProgressBar1.Pos = 1
110     dxProgressBar1.MaxPos = lCountOfFields
112     Scroller1.HValue = 1
114     Scroller1.Visible = False
116     C1Elastic2.Refresh

118     ReDim Fields(0)
120     ReDim Fields(lCountOfFields - 1)
    
122     i = 0

124     Do Until i = lCountOfFields

126         Debug.Print "GetRSInfo: " & mRS.Fields(i).Name
    
128         Fields(i).Name = mRS.Fields(i).Name
130         Fields(i).Caption = getFieldCaption(mConn, mRS, Fields(i).Name)
132         Fields(i).FieldType = mRS.Fields(i).Type
134         Fields(i).DefinedSize = mRS.Fields(i).DefinedSize
            
136         Select Case Fields(i).FieldType
                
                Case adLongVarWChar
138                 Fields(i).SimpleFieldType = FIELD_MEMO
            
140             Case adVarWChar
142                 Fields(i).SimpleFieldType = FIELD_TEXT

144             Case adVarChar
146                 Fields(i).SimpleFieldType = FIELD_TEXT

148             Case adWChar
150                 Fields(i).SimpleFieldType = FIELD_TEXT

152             Case adLongVarChar
154                 Fields(i).SimpleFieldType = FIELD_TEXT

156             Case adChar
158                 Fields(i).SimpleFieldType = FIELD_TEXT
        
160             Case adInteger
162                 Fields(i).SimpleFieldType = FIELD_INT

164             Case adTinyInt
166                 Fields(i).SimpleFieldType = FIELD_INT

168             Case adSmallInt
170                 Fields(i).SimpleFieldType = FIELD_INT
        
172             Case adDouble
174                 Fields(i).SimpleFieldType = FIELD_DOUBLE

176             Case adDecimal
178                 Fields(i).SimpleFieldType = FIELD_DOUBLE

180             Case adSingle
182                 Fields(i).SimpleFieldType = FIELD_DOUBLE

184             Case adNumeric
186                 Fields(i).SimpleFieldType = FIELD_DOUBLE
        
188             Case adDate
190                 Fields(i).SimpleFieldType = FIELD_DATE

192             Case adBoolean
194                 Fields(i).SimpleFieldType = FIELD_CHECKBOX
        
            End Select
        
196         If Left$(Fields(i).Name, 2) = "dd" Then Fields(i).SimpleFieldType = FIELD_PICK
198         If lCountOfCharsInLabel < Len(Fields(i).Caption) Then lCountOfCharsInLabel = Len(Fields(i).Caption)
200         i = i + 1
202         dxProgressBar1.DoStepBy 1
204         C1Elastic2.Refresh
        Loop
        
206     dxProgressBar1.Visible = False

        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Not accounted for
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' adArray
        ' adBigInt
        ' adBinary
        ' adBSTR
        ' adChapter
        ' adCurrency
        ' adDBDate
        ' adDBTime
        ' adDBTimeStamp
        ' adEmpty
        ' adError
        ' adFileTime
        ' adGUID
        ' adIDispatch
        ' adIUnknown
        ' adLongVarBinary
        ' adPropVariant
        ' adUnsignedBigInt
        ' adUnsignedInt
        ' adUnsignedSmallInt
        ' adUnsignedTinyInt
        ' adUserDefined
        ' adVarBinary
        ' adVariant
        ' adVarNumeric
        ' adVarWChar
        ''''''''''''''''''''''''''''''''''''''''''''''''''''

        '<EhFooter>
        Exit Sub

GetRSInfo_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.GetRSInfo", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub DisplayFields()
        '<EhHeader>
        On Error GoTo DisplayFields_Err
        '</EhHeader>

        Dim i As Long
        Dim iFieldHeight As Long
        Dim iFieldWidth As Long
        Dim lLabelWidth As Long
        Dim lCurrentTop As Long
        
100     bChangeMade = False
102     dxProgressBar1.Visible = True

104     ReDim cControl(lCountOfFields)
    
106     Scroller1.Visible = False

108     iFieldHeight = 300
110     lLabelWidth = lCountOfCharsInLabel * 100

112     iFieldWidth = C1Elastic2.Width - 1.5 * (lLabelWidth)
114     lCurrentTop = iFieldHeight
    
116     Do Until i = lCountOfFields

118         Load Label0(i + 1)
120         Set Label0(i + 1).Container = Scroller1
122         Label0(i + 1).Move 0, lCurrentTop + 50
124         Label0(i + 1).Height = iFieldHeight
126         Label0(i + 1).Width = lLabelWidth
128         Label0(i + 1).Caption = Fields(i).Caption
130         Label0(i + 1).Alignment = vbRightJustify
132         Label0(i + 1).BackColor = Scroller1.BackColor
134         Label0(i + 1).Visible = True

136         Debug.Print "DisplayFields: " & Fields(i).Name

138         Select Case Fields(i).SimpleFieldType

                Case FIELD_TEXT
140                 Load dxTextEdit0(i + 1)
142                 Set dxTextEdit0(i + 1).Container = Scroller1
144                 dxTextEdit0(i + 1).Move (lLabelWidth + 300), lCurrentTop
146                 dxTextEdit0(i + 1).Height = iFieldHeight
148                 dxTextEdit0(i + 1).Width = iFieldWidth
150                 dxTextEdit0(i + 1).Visible = True
152                 Set dxTextEdit0(i + 1).DataSource = mRS
154                 dxTextEdit0(i + 1).DataField = Fields(i).Name
156                 dxTextEdit0(i + 1).EditMaxLength = Fields(i).DefinedSize
158                 Set cControl(i + 1) = dxTextEdit0(i + 1)

160             Case FIELD_MEMO
162                 Load dxMemoEdit0(i + 1)
164                 Set dxMemoEdit0(i + 1).Container = Scroller1
166                 dxMemoEdit0(i + 1).Move (lLabelWidth + 300), lCurrentTop
168                 dxMemoEdit0(i + 1).Height = iFieldHeight * 3
170                 dxMemoEdit0(i + 1).Width = iFieldWidth
172                 dxMemoEdit0(i + 1).Visible = True
174                 Set dxMemoEdit0(i + 1).DataSource = mRS
176                 dxMemoEdit0(i + 1).DataField = Fields(i).Name
                    'dxMemoEdit0(i + 1).EditMaxLength = Fields(i).DefinedSize
178                 Set cControl(i + 1) = dxMemoEdit0(i + 1)
180                 lCurrentTop = lCurrentTop + iFieldHeight + (iFieldHeight * 1.25)

182             Case FIELD_INT
184                 Load dxIntegerEdit0(i + 1)
186                 Set dxIntegerEdit0(i + 1).Container = Scroller1
188                 dxIntegerEdit0(i + 1).Move (lLabelWidth + 300), lCurrentTop
190                 dxIntegerEdit0(i + 1).Height = iFieldHeight
192                 dxIntegerEdit0(i + 1).Width = iFieldWidth
194                 dxIntegerEdit0(i + 1).Visible = True
196                 Set dxIntegerEdit0(i + 1).DataSource = mRS
198                 dxIntegerEdit0(i + 1).DataField = Fields(i).Name
200                 dxIntegerEdit0(i + 1).EditMaxLength = 9 'Fields(I).DefinedSize
202                 Set cControl(i + 1) = dxIntegerEdit0(i + 1)

204             Case FIELD_DOUBLE
206                 Load dxDoubleEdit0(i + 1)
208                 Set dxDoubleEdit0(i + 1).Container = Scroller1
210                 dxDoubleEdit0(i + 1).Move (lLabelWidth + 300), lCurrentTop
212                 dxDoubleEdit0(i + 1).Height = iFieldHeight
214                 dxDoubleEdit0(i + 1).Width = iFieldWidth
216                 dxDoubleEdit0(i + 1).Visible = True
218                 Set dxDoubleEdit0(i + 1).DataSource = mRS
220                 dxDoubleEdit0(i + 1).DataField = Fields(i).Name
222                 dxDoubleEdit0(i + 1).EditMaxLength = 16 'Fields(i).DefinedSize
224                 Set cControl(i + 1) = dxDoubleEdit0(i + 1)

226             Case FIELD_DATE
228                 Load dxDateEdit0(i + 1)
230                 Set dxDateEdit0(i + 1).Container = Scroller1
232                 dxDateEdit0(i + 1).Move (lLabelWidth + 300), lCurrentTop
234                 dxDateEdit0(i + 1).Height = iFieldHeight
236                 dxDateEdit0(i + 1).Width = iFieldWidth
238                 dxDateEdit0(i + 1).Visible = True
240                 Set dxDateEdit0(i + 1).DataSource = mRS
242                 dxDateEdit0(i + 1).DataField = Fields(i).Name
244                 Set cControl(i + 1) = dxDateEdit0(i + 1)

246             Case FIELD_CHECKBOX
248                 Load dxCheckEdit0(i + 1)
250                 Set dxCheckEdit0(i + 1).Container = Scroller1
252                 dxCheckEdit0(i + 1).Move (lLabelWidth + 300), lCurrentTop
254                 dxCheckEdit0(i + 1).Height = iFieldHeight
256                 dxCheckEdit0(i + 1).Width = 300 'iFieldWidth
258                 dxCheckEdit0(i + 1).Visible = True
260                 dxCheckEdit0(i + 1).Tag = "CheckBox"
262                 Set dxCheckEdit0(i + 1).DataSource = mRS
264                 dxCheckEdit0(i + 1).DataField = Fields(i).Name
266                 Set cControl(i + 1) = dxCheckEdit0(i + 1)

268             Case FIELD_PICK
270                 Load dxLookUpEdit0(i + 1)
272                 Set dxLookUpEdit0(i + 1).Container = Scroller1
274                 dxLookUpEdit0(i + 1).Move (lLabelWidth + 300), lCurrentTop
276                 dxLookUpEdit0(i + 1).Height = iFieldHeight
278                 dxLookUpEdit0(i + 1).Width = iFieldWidth
280                 dxLookUpEdit0(i + 1).Visible = True
282                 Set dxLookUpEdit0(i + 1).DataSource = mRS
284                 dxLookUpEdit0(i + 1).DataField = Fields(i).Name
286                 Set cControl(i + 1) = dxLookUpEdit0(i + 1)
288                 PopulateCombo dxLookUpEdit0(i + 1), Fields(i).Name
                
290             Case Else
292                 Load dxTextEdit0(i + 1)
294                 Set dxTextEdit0(i + 1).Container = Scroller1
296                 dxTextEdit0(i + 1).Move (lLabelWidth + 300), lCurrentTop
298                 dxTextEdit0(i + 1).Height = iFieldHeight
300                 dxTextEdit0(i + 1).Width = iFieldWidth
302                 dxTextEdit0(i + 1).Visible = True
                    'Set dxTextEdit0(i + 1).DataSource = mRS
                    'dxTextEdit0(i + 1).DataField = Fields(i).Name
                    'dxTextEdit0(i + 1).EditMaxLength = Fields(i).DefinedSize
304                 dxTextEdit0(i + 1) = "N/A"
306                 Set cControl(i + 1) = dxTextEdit0(i + 1)
                
            End Select
            
308         If bIsGeoTable Then
310             cControl(i + 1).Enabled = False
                'Label0(i + 1).Enabled = False
312         ElseIf Fields(i).Name = "UID" Then
314             cControl(i + 1).Enabled = False
316         ElseIf (i = 0) And (iCountOfGUIDs = 1 Or iCountOfGUIDs = 2) Then
318             cControl(i + 1).Enabled = False
320         ElseIf (i = 1) And (iCountOfGUIDs = 2) Then
322             cControl(i + 1).Enabled = False
            End If
            
324         If InStr(1, sExcludedFields, Fields(i).Name, vbTextCompare) <> 0 Then
326             cControl(i + 1).Enabled = False
            End If
        
328         lCurrentTop = lCurrentTop + (iFieldHeight * 1.25)
330         i = i + 1
        Loop
    
332     If lCountOfFields > 0 Then
334         Load Label0(i + 1)
336         Set Label0(i + 1).Container = Scroller1
338         Label0(i + 1).Move 0, lCurrentTop
340         Label0(i + 1).Height = iFieldHeight / 2
342         Label0(i + 1).Visible = False
        End If

344     Scroller1.Refresh
346     Scroller1.Visible = True
348     bChangeMade = False
350     RaiseEvent SoundChangeMade
        '<EhFooter>
        Exit Sub

DisplayFields_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.DisplayFields", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub DeleteFields()
        '<EhHeader>
        On Error GoTo DeleteFields_Err
        '</EhHeader>

        Dim i As Integer
100     i = 0
102     Scroller1.Visible = False

104     If lCountOfFields > 0 Then

106         Unload Label0(lCountOfFields + 1)

108         Do Until i = lCountOfFields
    
110             Unload cControl(i + 1)
112             Unload Label0(i + 1)
114             i = i + 1
            Loop
        
        End If
    
116     Scroller1.VValue = 1
118     Scroller1.Refresh
120     Scroller1.Visible = True

        '<EhFooter>
        Exit Sub

DeleteFields_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.DeleteFields", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub ResizeFields()
        '<EhHeader>
        On Error GoTo ResizeFields_Err
        '</EhHeader>

        Dim i As Long
        Dim lLabelWidth As Long
        Dim iFieldWidth As Long
    
100     i = 0
        'Scroller1.Visible = False
    
102     lLabelWidth = lCountOfCharsInLabel * 100

104     iFieldWidth = C1Elastic2.Width - 1.5 * (lLabelWidth)

106     If iFieldWidth > 0 And lCountOfFields > 0 Then
        
108         Do Until i = lCountOfFields

110             If Not cControl(i + 1).Tag = "CheckBox" Then
112                 cControl(i + 1).Width = iFieldWidth

114                 If bIsGeoTable Then cControl(i + 1).Width = cControl(i + 1).Width * 0.9
                End If

116             i = i + 1
            Loop
        
        End If
    
        'Scroller1.Refresh
        'Scroller1.Visible = True

        '<EhFooter>
        Exit Sub

ResizeFields_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.ResizeFields", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub dxCheckEdit0_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo dxCheckEdit0_Change_Err
        '</EhHeader>
    
100     bChangeMade = True
102     RaiseEvent SoundChangeMade
        '<EhFooter>
        Exit Sub

dxCheckEdit0_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.dxCheckEdit0_Change", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub dxDateEdit0_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo dxDateEdit0_Change_Err
        '</EhHeader>
100     bChangeMade = True
102     RaiseEvent SoundChangeMade
        '<EhFooter>
        Exit Sub

dxDateEdit0_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.dxDateEdit0_Change", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub dxDoubleEdit0_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo dxDoubleEdit0_Change_Err
        '</EhHeader>
100     bChangeMade = True
102     RaiseEvent SoundChangeMade
        '<EhFooter>
        Exit Sub

dxDoubleEdit0_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.dxDoubleEdit0_Change", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub dxDoubleEdit0_KeyPress(Index As Integer, _
                                   KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo dxDoubleEdit0_KeyPress_Err
        '</EhHeader>

100     If InStr(dxDoubleEdit0(Index), "-") > 0 And Chr(KeyAscii) = "-" Then
102         KeyAscii = 0
104     ElseIf Chr(KeyAscii) = "-" Then

106     ElseIf InStr(dxDoubleEdit0(Index), ".") > 0 And Chr(KeyAscii) = "." Then
108         KeyAscii = 0
110     ElseIf Not Chr(KeyAscii) = "." And Not IsNumeric(Chr(KeyAscii)) Then
112         KeyAscii = 0
        End If

        '<EhFooter>
        Exit Sub

dxDoubleEdit0_KeyPress_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.dxDoubleEdit0_KeyPress", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub dxIntegerEdit0_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo dxIntegerEdit0_Change_Err
        '</EhHeader>
100     bChangeMade = True
102     RaiseEvent SoundChangeMade
        '<EhFooter>
        Exit Sub

dxIntegerEdit0_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.dxIntegerEdit0_Change", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub dxIntegerEdit0_KeyPress(Index As Integer, _
                                    KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo dxIntegerEdit0_KeyPress_Err
        '</EhHeader>
    
100     If InStr(dxIntegerEdit0(Index), "-") > 0 And Chr(KeyAscii) = "-" Then
102         KeyAscii = 0
104     ElseIf Chr(KeyAscii) = "-" Then

106     ElseIf Not IsNumeric(Chr(KeyAscii)) Then
108         KeyAscii = 0
        End If
 
        '<EhFooter>
        Exit Sub

dxIntegerEdit0_KeyPress_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.dxIntegerEdit0_KeyPress", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub dxLookUpEdit0_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo dxLookUpEdit0_Change_Err
        '</EhHeader>
100     bChangeMade = True
102     RaiseEvent SoundChangeMade
        '<EhFooter>
        Exit Sub

dxLookUpEdit0_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.dxLookUpEdit0_Change", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub dxMemoEdit0_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo dxMemoEdit0_Change_Err
        '</EhHeader>
100     bChangeMade = True
102     RaiseEvent SoundChangeMade
        '<EhFooter>
        Exit Sub

dxMemoEdit0_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.dxMemoEdit0_Change", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub dxTextEdit0_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo dxTextEdit0_Change_Err
        '</EhHeader>
100     bChangeMade = True
102     RaiseEvent SoundChangeMade
        '<EhFooter>
        Exit Sub

dxTextEdit0_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.dxTextEdit0_Change", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_Initialize()
        '<EhHeader>
        On Error GoTo UserControl_Initialize_Err
        '</EhHeader>
100     Set mRS = New ADODB.Recordset
102     Set mConn = New ADODB.Connection

        '<EhFooter>
        Exit Sub

UserControl_Initialize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.UserControl_Initialize", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    DeleteFields
    Set mRS = Nothing
    Set mConn = Nothing
End Sub

Private Function getFieldCaption(oConn As ADODB.Connection, _
                                 RSLocalRecordset As ADODB.Recordset, _
                                 sFieldName As String)
        '<EhHeader>
        On Error GoTo getFieldCaption_Err
        '</EhHeader>
 
        Dim oDB As ADOx.Catalog
        Dim itbl As ADOx.Table
        Dim fld As ADOx.Column
 
100     Set oDB = New ADOx.Catalog
102     Set itbl = New ADOx.Table
104     Set oDB.ActiveConnection = oConn
106     getFieldCaption = "desc not defined"

108     For Each itbl In oDB.Tables
            
110         If itbl.Name = RSLocalRecordset.Fields(0).Properties(1) Then
                'UNCOMMENT TO AUTOCHANGE CAPTIONS
112             getFieldCaption = itbl.Columns(sFieldName).Properties(2).Value
                ' Columns(iFieldIndex)
            End If
        
        Next
        
114     Set oDB = Nothing
116     Set itbl = Nothing
118     Set fld = Nothing

        '<EhFooter>
        Exit Function

getFieldCaption_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.getFieldCaption", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Function

Private Sub PopulateCombo(dxCombo As dxLookUpEdit, _
                          sFieldName As String)
        '<EhHeader>
        On Error GoTo PopulateCombo_Err
        '</EhHeader>

        Dim oRS As ADODB.Recordset
        Dim strQuery As String
        Dim i As Long
        Dim sFieldsInCombo
        
        If Not Right$(sFieldName, 4) = "_FEA" And Not Right$(sFieldName, 4) = "_GEO" Then
        
            If InStr(1, sFieldName, "_", vbTextCompare) > 0 Then
        
                sFieldName = Left$(sFieldName, InStr(1, sFieldName, "_", vbTextCompare) - 1)
        
            End If
        
        End If

100     strQuery = "SELECT * FROM [" & sTableNamePrefix & sFieldName & "] ORDER BY option"
102     Set oRS = New ADODB.Recordset

104     With oRS
106         Set .ActiveConnection = mConn
108         .CursorType = adOpenDynamic  'adOpenKeyset
110         .LockType = adLockOptimistic
112         .Source = strQuery
114         .Open
        End With
        
116     i = 1
118     sFieldsInCombo = ""

120     Do Until i = oRS.Fields.Count
        
122         If Not Left$(oRS.Fields(i).Name, 2) = "dd" Then
124             sFieldsInCombo = sFieldsInCombo & ";" & oRS.Fields(i).Name
            End If

126         i = i + 1
        Loop
        
128     If Len(sFieldsInCombo) > 2 Then
130         sFieldsInCombo = Right$(sFieldsInCombo, Len(sFieldsInCombo) - 1)
        End If
        
132     If Not oRS.EOF Or Not oRS.BOF Then

134         Set dxCombo.LookUpRecordset = oRS
136         dxCombo.LookUpKeyFieldName = "GUID1"
138         dxCombo.LookUpDisplayFieldName = "option"
140         dxCombo.ListFieldName = sFieldsInCombo
142         dxCombo.ListColumns = "*"
        
        End If

        '<EhFooter>
        Exit Sub

PopulateCombo_Err:
        Err.Raise vbObjectError + 100, "OASISDynamRSEditor.OASISDynamicRSEditor.PopulateCombo", "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Public Sub SaveRecord()
        '<EhHeader>
        On Error GoTo SaveRecord_Err
        '</EhHeader>

        Dim i As Long
        Dim lLabelWidth As Long
        Dim iFieldWidth As Long

100     i = 0

102     Do Until i = lCountOfFields

104         If cControl(i + 1) = "" Then cControl(i + 1) = Null
106         i = i + 1
        Loop

108     mRS.UpdateBatch adAffectCurrent
       
        '<EhFooter>
        Exit Sub

SaveRecord_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.SaveRecord", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Public Sub SetGUID1(sGUID As String)
        '<EhHeader>
        On Error GoTo SetGUID1_Err
        '</EhHeader>

100     If Not sGUID = "" And Not IsNull(sGUID) Then
    
102         mRS.Fields(0).Value = sGUID
104         cControl(1).Enabled = False
    
        End If

        '<EhFooter>
        Exit Sub

SetGUID1_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.SetGUID1", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Public Sub SetUID(lUID As Long)
        '<EhHeader>
        On Error GoTo SetUID_Err
        '</EhHeader>

100     If Not lUID = 0 And Not IsNull(lUID) Then
    
102         mRS.Fields("UID").Value = lUID
    
        End If

        '<EhFooter>
        Exit Sub

SetUID_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.SetUID", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Public Sub SetGUID2(sGUID As String)
        '<EhHeader>
        On Error GoTo SetGUID2_Err
        '</EhHeader>

100     If Not sGUID = "" And Not IsNull(sGUID) Then
    
102         mRS.Fields(1).Value = sGUID
104         cControl(2).Enabled = False
    
        End If

        '<EhFooter>
        Exit Sub

SetGUID2_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.SetGUID2", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

Public Sub LockUnlockAllFields(bLock As Boolean)
        '<EhHeader>
        On Error GoTo LockUnlockAllFields_Err
        '</EhHeader>
    
        Dim i As Integer
100     i = 0

102     If lCountOfFields > 0 Then
        
104         Do Until i = lCountOfFields
    
106             If bLock Then
108                 cControl(i + 1).ReadOnly = True
110                 bLocked = True
                Else
112                 cControl(i + 1).ReadOnly = False
114                 bLocked = False
                End If
            
116             i = i + 1
            Loop
        
        End If

        '<EhFooter>
        Exit Sub

LockUnlockAllFields_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISDynamRSEditor.OASISDynamicRSEditor.LockUnlockAllFields", _
                  "OASISDynamicRSEditor component failure"
        '</EhFooter>
End Sub

