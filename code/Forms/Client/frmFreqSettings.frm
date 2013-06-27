VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFreqSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Analysis Settings"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5010
   Icon            =   "frmFreqSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   8205
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5010
      _cx             =   8837
      _cy             =   14473
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
      BorderWidth     =   2
      ChildSpacing    =   2
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
      GridRows        =   2
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmFreqSettings.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic elControls 
         Height          =   480
         Left            =   30
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   7695
         Width           =   4950
         _cx             =   8731
         _cy             =   847
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
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
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
            Height          =   375
            Left            =   3915
            TabIndex        =   12
            Top             =   45
            Width           =   960
         End
      End
      Begin C1SizerLibCtl.C1Tab c1TabOptions 
         Height          =   7635
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   4950
         _cx             =   8731
         _cy             =   13467
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
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "General|Analysis|Advanced"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic elAdvanced 
            Height          =   7260
            Left            =   5895
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   330
            Width           =   4860
            _cx             =   8573
            _cy             =   12806
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
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   6
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
            GridRows        =   0
            GridCols        =   0
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
         Begin C1SizerLibCtl.C1Elastic elAnalysis 
            Height          =   7260
            Left            =   5595
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   330
            Width           =   4860
            _cx             =   8573
            _cy             =   12806
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
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   6
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
            GridRows        =   0
            GridCols        =   0
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame FraAnalysisSettings 
               Caption         =   "Analysis Settings:"
               Height          =   6720
               Left            =   90
               TabIndex        =   13
               Top             =   45
               Width           =   4695
               Begin VB.Frame FraAnalysisType 
                  Caption         =   "Type:"
                  Height          =   960
                  Left            =   2430
                  TabIndex        =   16
                  Top             =   225
                  Width           =   2130
                  Begin VB.CommandButton cmdIndividual 
                     Caption         =   "Individual"
                     Height          =   600
                     Left            =   1080
                     Picture         =   "frmFreqSettings.frx":6895
                     Style           =   1  'Graphical
                     TabIndex        =   18
                     Top             =   225
                     Width           =   915
                  End
                  Begin VB.CommandButton cmdRanges 
                     Caption         =   "Ranges"
                     Height          =   600
                     Left            =   135
                     Picture         =   "frmFreqSettings.frx":719B
                     Style           =   1  'Graphical
                     TabIndex        =   17
                     Top             =   225
                     Width           =   915
                  End
               End
               Begin VB.Frame FraAnalysisDetails 
                  Caption         =   "Analysis Details:"
                  Height          =   5415
                  Left            =   135
                  TabIndex        =   14
                  Top             =   1215
                  Width           =   4470
                  Begin OASISClient.ColorPicker ColorPicker1 
                     Height          =   240
                     Index           =   0
                     Left            =   2025
                     TabIndex        =   22
                     Top             =   1125
                     Width           =   1995
                     _ExtentX        =   3519
                     _ExtentY        =   503
                  End
                  Begin VB.ComboBox ComAnalysisField 
                     Height          =   315
                     Left            =   90
                     Style           =   2  'Dropdown List
                     TabIndex        =   19
                     Top             =   450
                     Width           =   4245
                  End
                  Begin DXDBGRIDLibCtl.dxDBGrid dxDataGrid 
                     Height          =   4440
                     Left            =   135
                     OleObjectBlob   =   "frmFreqSettings.frx":7A5D
                     TabIndex        =   21
                     Top             =   855
                     Width           =   4170
                  End
                  Begin MSComctlLib.ListView lvUniqueValues 
                     Height          =   2625
                     Left            =   360
                     TabIndex        =   20
                     Top             =   855
                     Width           =   3435
                     _ExtentX        =   6059
                     _ExtentY        =   4630
                     View            =   3
                     LabelWrap       =   -1  'True
                     HideSelection   =   -1  'True
                     _Version        =   393217
                     ForeColor       =   -2147483640
                     BackColor       =   -2147483643
                     BorderStyle     =   1
                     Appearance      =   1
                     NumItems        =   0
                  End
                  Begin VB.Label lblAnalysisField 
                     AutoSize        =   -1  'True
                     Caption         =   "Field:"
                     Height          =   195
                     Left            =   135
                     TabIndex        =   15
                     Top             =   225
                     Width           =   375
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elgeneral 
            Height          =   7260
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   4860
            _cx             =   8573
            _cy             =   12806
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
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   6
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
            GridRows        =   0
            GridCols        =   0
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame FraLegend 
               Caption         =   "Legend:"
               Height          =   1995
               Left            =   45
               TabIndex        =   28
               Top             =   2745
               Width           =   4740
               Begin VB.CommandButton cmdView 
                  Caption         =   "View"
                  Height          =   285
                  Left            =   3690
                  TabIndex        =   32
                  Top             =   1395
                  Width           =   825
               End
               Begin VB.TextBox txtLegendTitle 
                  Height          =   285
                  Left            =   180
                  TabIndex        =   30
                  Top             =   990
                  Width           =   4335
               End
               Begin VB.CheckBox chkShowLegend 
                  Caption         =   "Show Legend"
                  Height          =   240
                  Left            =   180
                  TabIndex        =   29
                  Top             =   315
                  Width           =   4155
               End
               Begin VB.Label lblLegendTitle 
                  AutoSize        =   -1  'True
                  Caption         =   "Legend Title:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   31
                  Top             =   720
                  Width           =   930
               End
            End
            Begin VB.Frame FraZeroNull 
               Caption         =   "Zero/Null Values"
               Height          =   2310
               Left            =   45
               TabIndex        =   23
               Top             =   4725
               Visible         =   0   'False
               Width           =   4740
               Begin VB.Frame FraZerosOr 
                  Caption         =   "Zeros or Blanks Color:"
                  Height          =   780
                  Left            =   135
                  TabIndex        =   26
                  Top             =   1260
                  Width           =   4470
                  Begin OASISClient.ColorPicker ColorPickerZeros 
                     Height          =   375
                     Left            =   135
                     TabIndex        =   27
                     Top             =   270
                     Width           =   4155
                     _ExtentX        =   7329
                     _ExtentY        =   661
                  End
               End
               Begin VB.CheckBox chkUseColor 
                  Caption         =   "Use color for Zeros or Blanks (Default Transparent)"
                  Height          =   420
                  Left            =   180
                  TabIndex        =   25
                  Top             =   720
                  Width           =   3390
               End
               Begin VB.CheckBox chkIgnoreZeros 
                  Caption         =   "Ignore Zeros or Blanks (Null)"
                  Height          =   420
                  Left            =   180
                  TabIndex        =   24
                  Top             =   270
                  Width           =   2445
               End
            End
            Begin VB.Frame FraOverlayLayer 
               Caption         =   "Overlay layer:"
               Height          =   1140
               Left            =   45
               TabIndex        =   8
               Top             =   1485
               Width           =   4740
               Begin VB.ComboBox ComOverlayLayer 
                  Height          =   315
                  Left            =   180
                  Style           =   2  'Dropdown List
                  TabIndex        =   9
                  Top             =   585
                  Width           =   4290
               End
               Begin VB.Label lblOverlayLayer 
                  Caption         =   "Name:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   10
                  Top             =   315
                  Width           =   2580
               End
            End
            Begin VB.Frame FraDataSettings 
               Caption         =   "Analysis layer:"
               Height          =   1230
               Left            =   45
               TabIndex        =   5
               Top             =   180
               Width           =   4740
               Begin VB.ComboBox ComLayerName 
                  Height          =   315
                  Left            =   180
                  Style           =   2  'Dropdown List
                  TabIndex        =   6
                  Top             =   585
                  Width           =   4380
               End
               Begin VB.Label lblDataTo 
                  Caption         =   "Name:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   7
                  Top             =   315
                  Width           =   2580
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmFreqSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_sLastFieldName As String
Private m_sLastLayer As String
Private m_sAnalysisLayer As String
Public Event GetFields(sLayer As String)
Public Event GetUniqueValues(sLayer As String, sField As String)
Public Event ChangeScope(sScope As String, sLayer As String)
Public Event ShowSpatialAnalysisLegend()
Public Event ApplyAnalysis()
Private M_sVals As Variant
Private m_iCurListItem As Integer
Private m_ColColors As New Collection

Public Property Get AnalysisLayer() As String
        '<EhHeader>
        On Error GoTo AnalysisLayer_Err
        '</EhHeader>
100     AnalysisLayer = m_sAnalysisLayer
        '<EhFooter>
        Exit Property

AnalysisLayer_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.AnalysisLayer " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Property

Public Property Get OverLayField() As String
        '<EhHeader>
        On Error GoTo OverLayField_Err
        '</EhHeader>
100     OverLayField = m_sLastFieldName
        '<EhFooter>
        Exit Property

OverLayField_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.OverLayField " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Property

Public Property Get OverLayLayer() As String
        '<EhHeader>
        On Error GoTo OverLayLayer_Err
        '</EhHeader>
100     OverLayLayer = m_sLastLayer
        '<EhFooter>
        Exit Property

OverLayLayer_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.OverLayLayer " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Property

Public Sub Init()
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        Dim i As Integer

100     SafeMoveFirst g_RSGISGridTableSettings
    
102     ComLayerName.Clear
104     ComOverlayLayer.Clear
    
106     Do While Not g_RSGISGridTableSettings.EOF
108         ComLayerName.AddItem g_RSGISGridTableSettings.Fields.Item("alias").Value
110         ComOverlayLayer.AddItem g_RSGISGridTableSettings.Fields.Item("alias").Value
112         g_RSGISGridTableSettings.MoveNext
        Loop
 
114     If m_oColUserLayers.Count > 0 Then

116         For i = 1 To m_oColUserLayers.Count - 1
118             ComLayerName.AddItem m_oColUserLayers.Item(i)
120             ComOverlayLayer.AddItem m_oColUserLayers.Item(i)
            Next

        End If

122     If ComLayerName.ListCount > 0 Then
124         ComLayerName.ListIndex = 0
126         ComOverlayLayer.ListIndex = 0
        End If
    
128     ColorPicker1(0).ShowDefault = True
130     ColorPicker1(0).ShowCustomColors = True
132     ColorPicker1(0).ShowSysColorButton = True
134     ColorPicker1(0).ShowMoreColors = True
136     ColorPicker1(0).ShowToolTips = True
    
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub cmdApply_Click()
        '<EhHeader>
        On Error GoTo cmdApply_Click_Err
        '</EhHeader>
100     RaiseEvent ApplyAnalysis
        '<EhFooter>
        Exit Sub

cmdApply_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.cmdApply_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdView_Click()
        '<EhHeader>
        On Error GoTo cmdView_Click_Err
        '</EhHeader>
100     RaiseEvent ShowSpatialAnalysisLegend
        '<EhFooter>
        Exit Sub

cmdView_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.cmdView_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComLayerName_Click()
        '<EhHeader>
        On Error GoTo ComLayerName_Click_Err
        '</EhHeader>
100     SafeMoveFirst g_RSGISGridTableSettings
102     g_RSGISGridTableSettings.Find "alias = '" & ComLayerName.List(ComLayerName.ListIndex) & "'"

104     If Not g_RSGISGridTableSettings.EOF Then
106         m_sAnalysisLayer = g_RSGISGridTableSettings.Fields.Item("name").Value
        Else
108         m_sAnalysisLayer = ""
        End If

        '<EhFooter>
        Exit Sub

ComLayerName_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.ComLayerName_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComOverlayLayer_Click()
        '<EhHeader>
        On Error GoTo ComOverlayLayer_Click_Err
        '</EhHeader>
100    SafeMoveFirst g_RSGISGridTableSettings
102     g_RSGISGridTableSettings.Find "alias = '" & ComOverlayLayer.List(ComOverlayLayer.ListIndex) & "'"

104     If Not g_RSGISGridTableSettings.EOF Then
106         m_sLastLayer = g_RSGISGridTableSettings.Fields.Item("name").Value
108         RaiseEvent GetFields(m_sLastLayer)
        Else
110         m_sLastLayer = ""
        End If

        '<EhFooter>
        Exit Sub

ComOverlayLayer_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.ComOverlayLayer_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComAnalysisField_Click()
        '<EhHeader>
        On Error GoTo ComAnalysisField_Click_Err
        '</EhHeader>
    Dim sLyr As String

100     SafeMoveFirst g_RSGISGridTableSettings
102     g_RSGISGridTableSettings.Find "alias = '" & ComOverlayLayer.List(ComOverlayLayer.ListIndex) & "'"
    
104     If Not g_RSGISGridTableSettings.EOF Then
106         sLyr = g_RSGISGridTableSettings.Fields.Item("name").Value
108         RaiseEvent GetUniqueValues(sLyr, ComAnalysisField.List(ComAnalysisField.ListIndex))
        End If
    
110     m_sLastFieldName = sLyr
    
        '<EhFooter>
        Exit Sub

ComAnalysisField_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.ComAnalysisField_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub RemoveDuplicates(lst As ListView)
        '<EhHeader>
        On Error GoTo RemoveDuplicates_Err
        '</EhHeader>

        Dim lRet As ListItem
        Dim strTemp As String
        Dim intCnt As Integer
100     intCnt = 0

102     Do While intCnt <= lst.ListItems.Count - 1
        
104         intCnt = intCnt + 1
            'Save the text that was in the listvew i
            '     ndex
106         strTemp = lst.ListItems.Item(intCnt).Text

            On Error Resume Next

            Do
108             lst.ListItems.Item(intCnt).Text = "" 'Remove the text inside the specific index
                'Use the FindItem() call to search for t
                '     he specific item
110             Set lRet = lst.FindItem(strTemp, lvwText, lvwPartial)
                'If the item is found, then it is a dupl
                '     icate and is removed

112             If Not lRet Is Nothing Then
114                 lst.ListItems.Remove (lRet.Index)
                End If

116         Loop While Not lRet Is Nothing 'If no item is found the loop is exited
        
118         lst.ListItems.Item(intCnt).Text = strTemp 'reset the listitem index text back To what it was, and Then continue
120         DebugPrint intCnt

122         DoEvents 'Added To ensure that the application does Not lock up when doing large amounts of data.
            
        Loop

124     LoadLayerAttrDataToGrid

        '<EhFooter>
        Exit Sub

RemoveDuplicates_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.RemoveDuplicates " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadLayerAttrDataToGrid()
        '<EhHeader>
        On Error GoTo LoadLayerAttrDataToGrid_Err
        '</EhHeader>
        Dim arVals As Variant
        Dim arVal As Variant
        Dim varVal As Variant
        Dim i As Integer
        Dim lngFldType As Long
        Dim lngDataLength As Long
        Dim flds() As String
        Dim col As DXDBGRIDLibCtl.dxGridColumn ' Variant
        Dim iRed As Integer
        Dim iGreen As Integer
        Dim iBlue As Integer
    
100     Set m_ColColors = New Collection
    
102     With dxDataGrid
104         .Dataset.Close
 
106         If .Dataset.FieldCount > 0 Then

108             For i = 0 To .Dataset.FieldCount - 1
110                 .Dataset.MemoryDataset.DeleteField .Dataset.FieldByNo(i)
                Next

            End If
            
112         .Dataset.Refresh
        
114         .Columns.DestroyColumns
116         .Options.Unset (egoShowGroupPanel)
118         .Options.Set (egoAutoWidth)
            '   .Options.Set (egoShowGroupPanel)
            '   .Options.Set (egoBandMoving)
            '   .Options.Set (egoColumnMoving)
            '    .Options.Set (egoMultiSort)
            '    .Options.Set (egoShowFooter)
120         .Options.Set (egoAutoSort)
122         .Options.Set (egoShowButtons)
            '    .Options.Set (egoShowRowFooter)
            '    .Options.Set (egoAutoSearch)
            '    .Options.Set (egoAutoExpandOnSearch)
124         .Options.Set (egoAnsiSort)
126         .Options.Set (egoLoadAllRecords)
128         .Options.Set (egoAutoSearch)
        
            
130         .Options.Unset (egoCanNavigation)
    
132         .DatasetType = dtMemoryDataset
134         .Dataset.MemoryDataset.ClearData
136         .Filter.FilterActive = True
            '.Filter.AutoDataSetFilter = True
            
138         .Filter.FilterStatus = fsAlways
    
140         ReDim flds(0 To lvUniqueValues.ListItems.Count - 1)
            
142         .Dataset.MemoryDataset.AddField "Unique Value", xftString, 254
144         Set col = .Columns.Add(gedTextEdit)
146         col.caption = "Unique Value"
148         col.FieldName = "Unique Value"
150         col.Visible = True
            
152         flds(0) = "Unique Value"
        
154         .KeyField = "Unique Value"
156         .Dataset.Open
        
158         For i = 1 To lvUniqueValues.ListItems.Count
160             ReDim arVal(0)
162             arVal(0) = lvUniqueValues.ListItems.Item(i).Text
164             GetRandomColor iRed, iGreen, iBlue
166             m_ColColors.Add RGB(iRed, iGreen, iBlue), lvUniqueValues.ListItems.Item(i).Text
168             arVals = arVal
170             .Dataset.AppendRecord arVals
            Next
        
        End With
    
        '<EhFooter>
        Exit Sub

LoadLayerAttrDataToGrid_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.LoadLayerAttrDataToGrid " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GetRandomColor(iR As Integer, _
                           iG As Integer, _
                           ib As Integer)
        '<EhHeader>
        On Error GoTo GetRandomColor_Err
        '</EhHeader>
100     Randomize
102     iR = Int(Rnd * 255)

104     Randomize
106     iG = Int(Rnd * 255)

108     Randomize
110     ib = Int(Rnd * 255)

        '<EhFooter>
        Exit Sub

GetRandomColor_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.GetRandomColor " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDataGrid_OnCustomDrawCell(ByVal HDC As Long, _
                                        ByVal Left As Single, _
                                        ByVal Top As Single, _
                                        ByVal Right As Single, _
                                        ByVal Bottom As Single, _
                                        ByVal Node As DXDBGRIDLibCtl.IdxGridNode, _
                                        ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, _
                                        ByVal Selected As Boolean, _
                                        ByVal Focused As Boolean, _
                                        ByVal NewItemRow As Boolean, _
                                        Text As String, _
                                        Color As Long, _
                                        ByVal Font As stdole.IFontDisp, _
                                        FontColor As Long, _
                                        Alignment As DXDBGRIDLibCtl.ExAlignment, _
                                        Done As Boolean)
        '<EhHeader>
        On Error GoTo dxDataGrid_OnCustomDrawCell_Err
        '</EhHeader>
        Dim s As String
        Dim q As Integer
        Dim iNumLevel1 As Integer
 
        On Error Resume Next
 
100     s = Node.values(dxDataGrid.Columns.ColumnByFieldName("Unique Value").Index)
  
102     Color = m_ColColors.Item(s)
    
        'Load ColorPicker1(ColorPicker1.UBound + 1)
        'ColorPicker1(ColorPicker1.UBound).Top = ColorPicker1(ColorPicker1.UBound).Top + ColorPicker1(ColorPicker1.UBound).Height * ColorPicker1.UBound
        'Set ColorPicker1(ColorPicker1.UBound).Parent = ColorPicker1(0).Parent
        'ColorPicker1(ColorPicker1.UBound).Visible = True
        'ColorPicker1(ColorPicker1.UBound).ZOrder 0
        'ColorPicker1(ColorPicker1.UBound).DefaultColor = m_ColColors.Item(s)
    
104     Select Case s

            Case "High"
                'Color = vbRed
                'Color = LColor1(0)

106         Case "Medium"
                'Color = vbBlue
                'Color = LColor1(3)

108         Case "Low"
                'Color = vbCyan
                'Color = LColor1(1)

110         Case "None"
                'Color = vbGreen
                'Color = LColor1(2)
        End Select

        '<EhFooter>
        Exit Sub

dxDataGrid_OnCustomDrawCell_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.dxDataGrid_OnCustomDrawCell " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDataGrid_OnMouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal x As Single, ByVal y As Single)
        '<EhHeader>
        On Error GoTo dxDataGrid_OnMouseUp_Err
        '</EhHeader>
    Dim ptp As POINTAPI

100 If Button = vbRightButton Then
    
        'DebugPrint elAppHolder.left & " : elAppHolder.Left"
        'DebugPrint tbAppHolder.left & " : tbAppHolder.Left"
        'DebugPrint Me.left & " : Me.Left"
        'DebugPrint X & " : X"
    
        'DebugPrint abGridPop.Bands("popGrid").left & " : Pop Left"
        'abGridPop.Bands("popGrid").PopupMenu , X, 0   '+ ScaleX(Me.Left, vbTwips, vbPixels), 0   'ScaleX(X, vbPixels, vbTwips), 0 'ScaleY(Y, vbPixels, vbTwips)
        ''abGridPop.Bands("popGrid").Left = X  '& " : Pop Left1"
        'DebugPrint abGridPop.Bands("popGrid").left & " : Pop Left 1" & vbCrLf
    
            ' Get the position of the cursor
102     GetCursorPos ptp
        'abGridPop.Bands("popGrid").PopupMenu , ScaleX(Point.X, vbPixels, vbTwips) - Me.left, ScaleY(Point.Y, vbPixels, vbTwips) - Me.top
104     If dxDataGrid.ex.SelectedCount > 0 Then
            'PopupMenu mnuGridAction, X:=ScaleX(ptPopUpPos.X, vbPixels, vbTwips) - Me.Left, Y:=ScaleY(ptPopUpPos.Y, vbPixels, vbTwips) - (Me.Top + 250)
        End If
    End If

        '<EhFooter>
        Exit Sub

dxDataGrid_OnMouseUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.dxDataGrid_OnMouseUp " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     If Not g_sLanguage = "" Then
102         If Not m_Cnn.State = adStateClosed Then
104             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        '<EhHeader>
        On Error GoTo Form_QueryUnload_Err
        '</EhHeader>
100     Me.Hide
102     Cancel = 1
        '<EhFooter>
        Exit Sub

Form_QueryUnload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmFreqSettings.Form_QueryUnload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
