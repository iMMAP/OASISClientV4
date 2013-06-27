VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmAddPointWZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Operations - OASIS feature"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmAddPointWZ.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elAddPt 
      Height          =   4725
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5745
      _cx             =   10134
      _cy             =   8334
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
      _GridInfo       =   $"frmAddPointWZ.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic elNav 
         Height          =   660
         Left            =   30
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4035
         Width           =   5685
         _cx             =   10028
         _cy             =   1164
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
         Begin VB.CommandButton cmdAddShape 
            Caption         =   "Ok"
            Height          =   420
            Left            =   4545
            TabIndex        =   26
            Top             =   135
            Width           =   960
         End
      End
      Begin C1SizerLibCtl.C1Tab c1TabAddPt 
         Height          =   3975
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   5685
         _cx             =   10028
         _cy             =   7011
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
         Caption         =   "Start|General|Location|Attachments|Summary"
         Align           =   0
         CurrTab         =   2
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
         Flags(0)        =   2
         Flags(1)        =   2
         Flags(3)        =   2
         Flags(4)        =   2
         Begin C1SizerLibCtl.C1Elastic elStep 
            Height          =   3600
            Index           =   0
            Left            =   -6240
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   330
            Width           =   5595
            _cx             =   9869
            _cy             =   6350
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
            Begin VB.TextBox En 
               Height          =   285
               Left            =   225
               TabIndex        =   6
               Text            =   "info"
               Top             =   180
               Width           =   1365
            End
         End
         Begin C1SizerLibCtl.C1Elastic elIntro 
            Height          =   3600
            Left            =   -6540
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   5595
            _cx             =   9869
            _cy             =   6350
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
            Begin VB.Label lblStart 
               Caption         =   "start"
               Height          =   1230
               Left            =   1080
               TabIndex        =   5
               Top             =   720
               Width           =   2490
            End
         End
         Begin C1SizerLibCtl.C1Elastic elStep 
            Height          =   3600
            Index           =   1
            Left            =   45
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   330
            Width           =   5595
            _cx             =   9869
            _cy             =   6350
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
            Begin VB.Frame FraCoordinate 
               Caption         =   "Coordinate"
               Height          =   1680
               Left            =   1890
               TabIndex        =   17
               Top             =   630
               Width           =   3660
               Begin VB.Frame FraXY 
                  Caption         =   "X:Y"
                  Height          =   1005
                  Left            =   135
                  TabIndex        =   19
                  Top             =   270
                  Width           =   3300
                  Begin VB.TextBox txtY 
                     Height          =   285
                     Left            =   1260
                     TabIndex        =   21
                     Top             =   630
                     Width           =   1905
                  End
                  Begin VB.TextBox txtX 
                     Height          =   285
                     Left            =   1260
                     TabIndex        =   20
                     Top             =   270
                     Width           =   1905
                  End
                  Begin VB.Label lblY 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Y"
                     Height          =   195
                     Left            =   1035
                     TabIndex        =   23
                     Top             =   630
                     Width           =   105
                  End
                  Begin VB.Label lblX 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "X"
                     Height          =   195
                     Left            =   1035
                     TabIndex        =   22
                     Top             =   270
                     Width           =   105
                  End
               End
               Begin VB.TextBox txtMGRS 
                  Height          =   285
                  Left            =   720
                  TabIndex        =   18
                  Top             =   1305
                  Width           =   2715
               End
               Begin VB.Label lblMGRS 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "MGRS:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   24
                  Top             =   1350
                  Width           =   555
               End
            End
            Begin VB.Frame FraType1 
               Caption         =   "Type"
               Height          =   555
               Left            =   90
               TabIndex        =   13
               Top             =   45
               Width           =   3660
               Begin VB.OptionButton OptFeatureType 
                  Caption         =   "Dynamic"
                  Height          =   195
                  Index           =   2
                  Left            =   2385
                  TabIndex        =   16
                  Top             =   225
                  Width           =   960
               End
               Begin VB.OptionButton OptFeatureType 
                  Caption         =   "Temporary"
                  Height          =   195
                  Index           =   1
                  Left            =   1305
                  TabIndex        =   15
                  Top             =   225
                  Width           =   1320
               End
               Begin VB.OptionButton OptFeatureType 
                  Caption         =   "Permanent"
                  Height          =   195
                  Index           =   0
                  Left            =   135
                  TabIndex        =   14
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   1320
               End
            End
            Begin VB.Frame FraShapeType 
               Caption         =   "Shape"
               Height          =   1365
               Left            =   90
               TabIndex        =   8
               Top             =   630
               Width           =   1725
               Begin VB.OptionButton OptShpType 
                  Caption         =   "Multi Point"
                  Height          =   195
                  Index           =   3
                  Left            =   180
                  TabIndex        =   12
                  Top             =   1035
                  Width           =   1230
               End
               Begin VB.OptionButton OptShpType 
                  Caption         =   "Polygon"
                  Height          =   195
                  Index           =   2
                  Left            =   180
                  TabIndex        =   11
                  Top             =   780
                  Width           =   1230
               End
               Begin VB.OptionButton OptShpType 
                  Caption         =   "Line"
                  Height          =   195
                  Index           =   1
                  Left            =   180
                  TabIndex        =   10
                  Top             =   525
                  Width           =   1230
               End
               Begin VB.OptionButton OptShpType 
                  Caption         =   "Point"
                  Height          =   195
                  Index           =   0
                  Left            =   180
                  TabIndex        =   9
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   1230
               End
            End
            Begin VB.TextBox txtTxtLocation 
               Height          =   1230
               Left            =   45
               TabIndex        =   7
               Top             =   2340
               Width           =   5505
            End
         End
      End
   End
End
Attribute VB_Name = "frmAddPointWZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lCurShapeType As TatukGIS_XDK9.XGIS_ShapeType
Private lCurLocationType As OASISLocationType


Public Event AddShape(dtShapeType As TatukGIS_XDK9.XGIS_ShapeType, dtFeatureType As OASISLocationType)

Public Sub Init()
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
100   lCurShapeType = TatukGIS_XDK9.XgisShapeTypePoint
102   lCurLocationType = Permanent
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddPointWZ.init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAddShape_Click()
        '<EhHeader>
        On Error GoTo cmdAddShape_Click_Err
        '</EhHeader>
100     RaiseEvent AddShape(lCurShapeType, lCurLocationType)
        '<EhFooter>
        Exit Sub

cmdAddShape_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddPointWZ.cmdAddShape_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetCoordinateValue(x As String, xLabel As String, y As String, yLabel As String, sCoordSysName As String, sMGRS As String)
        '<EhHeader>
        On Error GoTo SetCoordinateValue_Err
        '</EhHeader>
100     txtX.Text = x
102     lblX.caption = xLabel
104     txtY.Text = y
106     lblY.caption = yLabel
108     txtMGRS.Text = sMGRS
110     FraXY.caption = sCoordSysName
        '<EhFooter>
        Exit Sub

SetCoordinateValue_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddPointWZ.SetCoordinateValue " & _
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
               "in OASISClient.frmAddPointWZ.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub OptFeatureType_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo OptFeatureType_Click_Err
        '</EhHeader>
    
100     Select Case Index
    
            Case 0
102             lCurLocationType = Permanent
104         Case 1
106             lCurLocationType = Temporary
108         Case 2
110             lCurLocationType = Dynamic
        End Select
    
        '<EhFooter>
        Exit Sub

OptFeatureType_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddPointWZ.OptFeatureType_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub OptShpType_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo OptShpType_Click_Err
        '</EhHeader>

100     Select Case Index
    
            Case 0
102             lCurShapeType = TatukGIS_XDK9.XgisShapeTypePoint

104         Case 1
106             lCurShapeType = TatukGIS_XDK9.XgisShapeTypeArc

108         Case 2
110             lCurShapeType = TatukGIS_XDK9.XgisShapeTypePolygon

112         Case 3
114             lCurShapeType = TatukGIS_XDK9.XgisShapeTypeMultiPoint
        End Select

        '<EhFooter>
        Exit Sub

OptShpType_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddPointWZ.OptShpType_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
