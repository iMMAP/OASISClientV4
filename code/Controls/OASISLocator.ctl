VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Begin VB.UserControl OASISLocator 
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   Picture         =   "OASISLocator.ctx":0000
   ScaleHeight     =   3555
   ScaleWidth      =   3825
   ToolboxBitmap   =   "OASISLocator.ctx":6852
   Begin C1SizerLibCtl.C1Elastic elHolder 
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3825
      _cx             =   6747
      _cy             =   6271
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
      Begin VB.CommandButton cmdGetCoordinate 
         Height          =   510
         Left            =   1560
         Picture         =   "OASISLocator.ctx":6B64
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Get Coordinate from Map"
         Top             =   3570
         Width           =   600
      End
      Begin VB.CommandButton cmdCheckIn 
         Height          =   510
         Left            =   60
         Picture         =   "OASISLocator.ctx":D3B6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Check In Map"
         Top             =   3600
         Width           =   600
      End
      Begin C1SizerLibCtl.C1Elastic elHolder1 
         Height          =   3570
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   3855
         _cx             =   6800
         _cy             =   6297
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
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"OASISLocator.ctx":13C08
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab tabPCode 
            Height          =   3510
            Left            =   30
            TabIndex        =   7
            Top             =   30
            Width           =   3795
            _cx             =   6694
            _cy             =   6191
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
            Caption         =   "Admin Locator|Geo Locator"
            Align           =   0
            CurrTab         =   1
            FirstTab        =   0
            Style           =   3
            Position        =   0
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   -1  'True
            TabsPerPage     =   2
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
            Begin C1SizerLibCtl.C1Elastic elGeo 
               Height          =   3135
               Left            =   45
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   330
               Width           =   3705
               _cx             =   6535
               _cy             =   5530
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
               Begin VB.CommandButton cmdRadius 
                  Height          =   330
                  Left            =   60
                  Picture         =   "OASISLocator.ctx":13C3E
                  Style           =   1  'Graphical
                  TabIndex        =   42
                  Top             =   2760
                  Width           =   330
               End
               Begin VB.CommandButton cmdCheckInMap 
                  Height          =   330
                  Left            =   750
                  Picture         =   "OASISLocator.ctx":14067
                  Style           =   1  'Graphical
                  TabIndex        =   41
                  ToolTipText     =   "Check Coordinate in Map"
                  Top             =   2760
                  Width           =   330
               End
               Begin VB.CommandButton cmdGetFromMap 
                  Height          =   330
                  Left            =   420
                  Picture         =   "OASISLocator.ctx":144B1
                  Style           =   1  'Graphical
                  TabIndex        =   40
                  ToolTipText     =   "Get Coordinates From Map"
                  Top             =   2760
                  Width           =   330
               End
               Begin VB.TextBox txtMGRS 
                  Height          =   315
                  Left            =   1050
                  TabIndex        =   39
                  Top             =   810
                  Width           =   2610
               End
               Begin VB.Frame FraAutoLocator 
                  Caption         =   "Auto Locator"
                  Height          =   1530
                  Left            =   30
                  TabIndex        =   35
                  Top             =   1170
                  Width           =   3660
                  Begin VB.Label lblProvince_ 
                     AutoSize        =   -1  'True
                     Caption         =   "Province:____________________"
                     Height          =   195
                     Left            =   75
                     TabIndex        =   38
                     Top             =   375
                     Width           =   3465
                  End
                  Begin VB.Label lblDistrict_ 
                     AutoSize        =   -1  'True
                     Caption         =   "District:______________________"
                     Height          =   195
                     Left            =   75
                     TabIndex        =   37
                     Top             =   810
                     Width           =   3465
                  End
                  Begin VB.Label lblNearestTown 
                     AutoSize        =   -1  'True
                     Caption         =   "Nearest Town:________________"
                     Height          =   195
                     Left            =   75
                     TabIndex        =   36
                     Top             =   1200
                     Width           =   3420
                  End
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxEdtLat 
                  Height          =   315
                  Left            =   1035
                  OleObjectBlob   =   "OASISLocator.ctx":148FD
                  TabIndex        =   9
                  Top             =   180
                  Width           =   2610
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxedtLong 
                  Height          =   315
                  Left            =   1035
                  OleObjectBlob   =   "OASISLocator.ctx":14955
                  TabIndex        =   10
                  Top             =   495
                  Width           =   2610
               End
               Begin VB.Label lblLatitudeX 
                  AutoSize        =   -1  'True
                  Caption         =   "Latitude Y:"
                  Height          =   195
                  Left            =   0
                  TabIndex        =   13
                  Top             =   180
                  Width           =   765
               End
               Begin VB.Label lblLongitudeX 
                  AutoSize        =   -1  'True
                  Caption         =   "Longitude X:"
                  Height          =   195
                  Left            =   0
                  TabIndex        =   12
                  Top             =   540
                  Width           =   900
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "MGRS:"
                  Height          =   195
                  Left            =   0
                  TabIndex        =   11
                  Top             =   900
                  Width           =   525
               End
            End
            Begin C1SizerLibCtl.C1Elastic elAdmin 
               Height          =   3135
               Left            =   -4350
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   330
               Width           =   3705
               _cx             =   6535
               _cy             =   5530
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
               Begin VB.ComboBox ComPlace 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   810
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   660
                  Width           =   2865
               End
               Begin VB.ComboBox ComDistrict 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   810
                  Style           =   2  'Dropdown List
                  TabIndex        =   27
                  Top             =   360
                  Width           =   2865
               End
               Begin VB.ComboBox ComProvince 
                  Height          =   315
                  Left            =   810
                  Style           =   2  'Dropdown List
                  TabIndex        =   26
                  Top             =   30
                  Width           =   2865
               End
               Begin VB.Frame FraPcodes 
                  Caption         =   "Admin Location Search"
                  Height          =   2025
                  Left            =   30
                  TabIndex        =   15
                  Top             =   1020
                  Width           =   3540
                  Begin VB.CheckBox chkUseFuzzy 
                     Caption         =   "Use Fuzzy Search (May Be Slow) "
                     Height          =   315
                     Left            =   60
                     TabIndex        =   43
                     Top             =   630
                     Width           =   2895
                  End
                  Begin VB.TextBox txtPcode 
                     Height          =   330
                     Left            =   900
                     TabIndex        =   23
                     Top             =   270
                     Width           =   2085
                  End
                  Begin VB.CommandButton cmdSearch 
                     Height          =   315
                     Left            =   3000
                     MaskColor       =   &H00FFFFFF&
                     Picture         =   "OASISLocator.ctx":149CB
                     Style           =   1  'Graphical
                     TabIndex        =   22
                     ToolTipText     =   "Search"
                     Top             =   270
                     UseMaskColor    =   -1  'True
                     Width           =   450
                  End
                  Begin VB.ListBox lstResult 
                     Height          =   840
                     ItemData        =   "OASISLocator.ctx":14E7E
                     Left            =   630
                     List            =   "OASISLocator.ctx":14E80
                     TabIndex        =   21
                     Top             =   930
                     Width           =   2835
                  End
                  Begin VB.OptionButton OptSearchType 
                     Caption         =   "PCode"
                     Height          =   195
                     Index           =   0
                     Left            =   45
                     TabIndex        =   20
                     Top             =   210
                     Value           =   -1  'True
                     Width           =   780
                  End
                  Begin VB.OptionButton OptSearchType 
                     Caption         =   "Village"
                     Height          =   195
                     Index           =   1
                     Left            =   45
                     TabIndex        =   19
                     Top             =   420
                     Width           =   780
                  End
                  Begin VB.Frame FraSource 
                     Caption         =   "Source:"
                     BeginProperty Font 
                        Name            =   "MS Serif"
                        Size            =   6.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   420
                     Left            =   1080
                     TabIndex        =   16
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   2250
                     Begin VB.OptionButton OptDatasource 
                        Caption         =   "OCHA/HIC"
                        BeginProperty Font 
                           Name            =   "MS Serif"
                           Size            =   6.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   195
                        Index           =   0
                        Left            =   90
                        TabIndex        =   18
                        Top             =   180
                        Value           =   -1  'True
                        Width           =   1275
                     End
                     Begin VB.OptionButton OptDatasource 
                        Caption         =   "LIS"
                        BeginProperty Font 
                           Name            =   "MS Serif"
                           Size            =   6.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   195
                        Index           =   1
                        Left            =   1410
                        TabIndex        =   17
                        Top             =   150
                        Width           =   570
                     End
                  End
                  Begin VB.Label lblResult 
                     AutoSize        =   -1  'True
                     Caption         =   "Result:"
                     Height          =   195
                     Left            =   90
                     TabIndex        =   25
                     Top             =   1230
                     Width           =   495
                  End
                  Begin VB.Label lblRecords 
                     AutoSize        =   -1  'True
                     Caption         =   "records"
                     Height          =   195
                     Left            =   45
                     TabIndex        =   24
                     Top             =   1725
                     Width           =   525
                  End
               End
               Begin VB.Label lblProvince 
                  AutoSize        =   -1  'True
                  Caption         =   "Province:"
                  Height          =   195
                  Left            =   15
                  TabIndex        =   34
                  Top             =   90
                  Width           =   675
               End
               Begin VB.Label lblCommunity 
                  AutoSize        =   -1  'True
                  Caption         =   "Place:"
                  Height          =   195
                  Left            =   15
                  TabIndex        =   33
                  Top             =   750
                  Width           =   450
               End
               Begin VB.Label lblDistrict 
                  AutoSize        =   -1  'True
                  Caption         =   "District:"
                  Height          =   195
                  Left            =   15
                  TabIndex        =   32
                  Top             =   450
                  Width           =   525
               End
               Begin VB.Label lblProvincePcode 
                  AutoSize        =   -1  'True
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   7875
                  TabIndex        =   31
                  Top             =   360
                  Width           =   75
               End
               Begin VB.Label lblDistrictPCode 
                  AutoSize        =   -1  'True
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   7875
                  TabIndex        =   30
                  Top             =   900
                  Width           =   75
               End
               Begin VB.Label lblPlaceCode 
                  AutoSize        =   -1  'True
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   7875
                  TabIndex        =   29
                  Top             =   1440
                  Width           =   75
               End
            End
         End
      End
      Begin VB.Label lblCheckYour 
         Caption         =   "Check Your position"
         Height          =   510
         Left            =   660
         TabIndex        =   5
         Top             =   3570
         Width           =   1185
      End
      Begin VB.Label lblGetCoordinates 
         Caption         =   "Get Position From Map"
         Height          =   510
         Left            =   2160
         TabIndex        =   4
         Top             =   3600
         Width           =   1185
      End
      Begin VB.Label lblLocationDescription 
         Caption         =   "Location Description"
         Height          =   285
         Left            =   60
         TabIndex        =   3
         Top             =   1950
         Width           =   1710
      End
   End
End
Attribute VB_Name = "OASISLocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event CheckCoordInMap(x As Double, y As Double, sMGRS As String)
Public Event PcodePicked()
Public Event FocusRecieved(Name As String)
Public Event LocRadiusSelectTool()
Public Event GetCoordFromMapTool()
Public Event hglAdminLevel0()
Public Event hglAdminLevel1()
Public Event hglAdminLevel2()
Public Event hglAdminLevel3()
Public Event hglAdminLocation()
Public Event LocationFound(sName As String)

Private m_cn As adodb.Connection
Private m_bGetCoordTool As Boolean
Private m_oShpPt As TatukGIS_XDK9.XGIS_ShapePoint
Private m_oDrawLyr As TatukGIS_XDK9.XGIS_LayerVector
Dim lL As TatukGIS_XDK9.XGIS_LayerVector
Dim ll2 As TatukGIS_XDK9.XGIS_LayerVector

Private M_RS As adodb.Recordset
Public DetailClose As Boolean
Dim DBName As String
Dim ilName As String
Dim i As Long
Dim RRequery As Boolean

Private m_OGIS As TatukGIS_XDK9.XGIS_Viewer
Private m_lngLastAddedLocationID As Long
'Private Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long
Private m_bShowGeoSelector As Boolean
Private m_bShowAdminSelector As Boolean
Private m_intStartTab As Integer
Private m_DouLatitude As Double
Private m_DouLongitude As Double
Private m_StrMGRS As String
Private m_StrProvince As String
Private m_StrDistrict As String
Private m_StrLocation As String


Public Property Get Location() As String
    Location = m_StrLocation
End Property

Public Property Let Location(ByVal StrValue As String)
    m_StrLocation = StrValue
End Property

Public Property Get District() As String
    District = m_StrDistrict
End Property

Public Property Let District(ByVal StrValue As String)
    m_StrDistrict = StrValue
End Property

Public Property Get Province() As String
    Province = m_StrProvince
End Property

Public Property Let Province(ByVal StrValue As String)
    m_StrProvince = StrValue
End Property

Public Property Get MGRS() As String
    MGRS = txtMGRS.Text
End Property

Public Property Let MGRS(ByVal StrValue As String)
    m_StrMGRS = StrValue
End Property

Public Property Get Longitude() As Double
    If IsNumeric(dxedtLong.EditValue) Then
        Longitude = dxedtLong.EditValue
    Else
        Longitude = 0
    End If
End Property

Public Property Let Longitude(ByVal DouValue As Double)
    m_DouLongitude = DouValue
End Property

Public Property Get Latitude() As Double
    If IsNumeric(dxEdtLat.EditValue) Then
        Latitude = dxEdtLat.EditValue
    Else
        Latitude = 0
    End If
End Property

Public Property Let Latitude(ByVal DouValue As Double)
    m_DouLatitude = DouValue
End Property

Public Property Get StartTab() As Integer
    m_intStartTab = tabPCode.CurrTab
    StartTab = m_intStartTab
End Property

Public Property Let StartTab(ByVal intValue As Integer)
    m_intStartTab = intValue
    tabPCode.CurrTab = intValue
End Property

Public Property Get ShowAdminSelector() As Boolean
    ShowAdminSelector = m_bShowAdminSelector
End Property

Public Property Let ShowAdminSelector(ByVal bValue As Boolean)
    m_bShowAdminSelector = bValue
    tabPCode.TabVisible(0) = bValue
End Property

Public Property Get ShowGeoSelector() As Boolean
    ShowGeoSelector = m_bShowGeoSelector
End Property

Public Property Let ShowGeoSelector(ByVal bValue As Boolean)
    m_bShowGeoSelector = bValue
    tabPCode.TabVisible(1) = bValue
End Property
 
Public Sub LoadAdmin0(v As Variant)
        '*************PROVINCE******************************
        '<EhHeader>
        On Error GoTo LoadAdmin0_Err
        '</EhHeader>

        Dim oLyr As TatukGIS_XDK9.XGIS_LayerVector

100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel0'"

104     Set oLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)

106     If oLyr Is Nothing Then Exit Sub

108     oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True

110     Do While Not oLyr.EOF
112         ComProvince.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
116         oLyr.MoveNext
        Loop

        '*****************************************

        '<EhFooter>
        Exit Sub

LoadAdmin0_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLocator.LoadAdmin0 " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadAdmin1(v As Variant)
               '*************DISTRICT******************************
        '<EhHeader>
        On Error GoTo LoadAdmin1_Err
        '</EhHeader>
        
                Dim oLyr As TatukGIS_XDK9.XGIS_LayerVector
                Dim oProvLayer As TatukGIS_XDK9.XGIS_LayerVector
                Dim aShape As TatukGIS_XDK9.XGIS_Shape
                Dim shp As TatukGIS_XDK9.XGIS_Shape
        
100             SafeMoveFirst g_RSAppSettings
102             g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel1'"
        
104             Set oLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
        
                If oLyr Is Nothing Then Exit Sub
        
106             If ComProvince.List(ComProvince.ListIndex) = "--ALL--" Then
108                 oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
        
110                 Do While Not oLyr.EOF
112                     ComDistrict.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
116                     oLyr.MoveNext
                    Loop
        
118                 ComDistrict.ListIndex = 0
                Else
        
120                 SafeMoveFirst g_RSAppSettings
122                 g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel0'"
        
124                 Set oProvLayer = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
                    
                    If oProvLayer Is Nothing Then Exit Sub
                    
126                 Set shp = oProvLayer.FindFirst(m_OGIS.Extent, g_RSAppSettings.Fields.Item("SettingValue2").Value & " = '" & ComProvince.List(ComProvince.ListIndex) & "'", Nothing, "", True)
                
                    If shp Is Nothing Then Exit Sub
                    
130                 m_OGIS.VisibleExtent = shp.Extent
132                 shp.Flash
        
134                 ComDistrict.Clear
136                 ComDistrict.AddItem "--ALL--"
        
138                 SafeMoveFirst g_RSAppSettings
140                 g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel1'"
        
142                 Set aShape = oLyr.FindFirst(m_OGIS.Extent, "", shp, GisUtils.GIS_RELATE_CONTAINS, True)
                    
                    If aShape Is Nothing Then Exit Sub
                    
144                 Do While Not oLyr.EOF
146                     ComDistrict.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
        
148                     oLyr.MoveNext
                    Loop
        
150                 ComDistrict.ListIndex = 0
        
                End If
        
                '*****************************************

        '<EhFooter>
        Exit Sub

LoadAdmin1_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLocator.LoadAdmin1 " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadAdmin2(v As Variant)

End Sub


Public Sub LoadAdminLocation(v As Variant)
        '<EhHeader>
        On Error GoTo LoadAdminLocation_Err
        '</EhHeader>
    Dim oLyr As TatukGIS_XDK9.XGIS_LayerVector

            '*************Place******************************
    
    
100         SafeMoveFirst g_RSAppSettings
102         g_RSAppSettings.Find "SettingName = 'PCodeAdminLocation'"
    
104         Set oLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            
            If oLyr Is Nothing Then Exit Sub
            
106         oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
        
108         ComPlace.Clear
        
110         Do While Not oLyr.EOF
112             ComPlace.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
114             m_frmDebug.DebugPrint oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
116             oLyr.MoveNext
            Loop
    
118         ComPlace.Enabled = True
            '*****************************************

        '<EhFooter>
        Exit Sub

LoadAdminLocation_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLocator.LoadAdminLocation " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

  
Public Function GetLastAddedLocationID() As Long
        '<EhHeader>
        On Error GoTo GetLastAddedLocationID_Err
        '</EhHeader>
100     GetLastAddedLocationID = m_lngLastAddedLocationID
        '<EhFooter>
        Exit Function

GetLastAddedLocationID_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISLocator.GetLastAddedLocationID", _
                  "OASISLocator component failure"
        '</EhFooter>
End Function

Private Sub cmdCheckIn_Click()
        '<EhHeader>
        On Error GoTo cmdCheckIn_Click_Err
        '</EhHeader>
        Dim ptg As New TatukGIS_XDK9.XGIS_Point
            
100     ptg.Prepare CDbl(dxedtLong.EditValue), CDbl(dxEdtLat.EditValue)

102     m_OGIS.CenterViewport ptg
        
104     m_OGIS.UpDate
        '<EhFooter>
        Exit Sub

cmdCheckIn_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISLocator.cmdCheckIn_Click", _
                  "OASISLocator component failure"
        '</EhFooter>
End Sub

Public Sub Init(GIS As TatukGIS_XDK9.XGIS_Viewer)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        
100             SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel0'"
104     lblProvince.caption = g_RSAppSettings.Fields.Item("SettingValue1").Value & ":"

106              SafeMoveFirst g_RSAppSettings
108     g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel1'"
110     lblDistrict.caption = g_RSAppSettings.Fields.Item("SettingValue1").Value & ":"
        
112             SafeMoveFirst g_RSAppSettings
114     g_RSAppSettings.Find "SettingName = 'PCodeAdminLocation'"
116     lblCommunity.caption = g_RSAppSettings.Fields.Item("SettingValue1").Value & ":"


                        
118     Set m_cn = m_Cnn
    
120     Set m_OGIS = GIS
             
122     ComDistrict.Clear
124     ComProvince.Clear
       
126     ComDistrict.AddItem "--ALL--"
128     ComProvince.AddItem "--ALL--"
130     LoadAdmin0 1
132     txtMGRS.Text = ""

134     If Not ComProvince.ListCount > 0 Then ComProvince.ListIndex = 0
    
136     CreateLyrs
        
        '<EhFooter>
        Exit Sub

Init_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISLocator.Init", _
                  "OASISLocator component failure"
        '</EhFooter>
End Sub

Public Function IsInitialized() As Boolean
        '<EhHeader>
        On Error GoTo IsInitialized_Err
        '</EhHeader>

100     If m_cn Is Nothing Then
102         IsInitialized = False
        Else
104         IsInitialized = True
        End If

        '<EhFooter>
        Exit Function

IsInitialized_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISLocator.IsInitialized", _
                  "OASISLocator component failure"
        '</EhFooter>
End Function

Private Sub cmdGetFrom_Click()
        '<EhHeader>
        On Error GoTo cmdGetFrom_Click_Err
        '</EhHeader>
100     m_bGetCoordTool = True
        '<EhFooter>
        Exit Sub

cmdGetFrom_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISLocator.cmdGetFrom_Click", _
                  "OASISLocator component failure"
        '</EhFooter>
End Sub


Private Sub cmdCheckInMap_Click()
    RaiseEvent CheckCoordInMap(CDbl(dxedtLong.EditValue), CDbl(dxEdtLat.EditValue), txtMGRS.Text)
End Sub

Private Sub cmdGetFromMap_Click()
    RaiseEvent GetCoordFromMapTool
End Sub

Private Sub cmdRadius_Click()
        '<EhHeader>
        On Error GoTo cmdRadius_Click_Err
        '</EhHeader>
        RaiseEvent LocRadiusSelectTool
        
        '<EhFooter>
        Exit Sub

cmdRadius_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISLocator.cmdRadius_Click", _
                  "OASISLocator component failure"
        '</EhFooter>
End Sub

Public Sub SetAdmValues(sAdmVal1 As String, sAdmVal2 As String, sAdmloc As String)
        
    If sAdmVal1 = "" Then
        lblProvince_.caption = "Province: N/A"
    Else
        lblProvince_.caption = "Province: " & sAdmVal1
    End If

    If sAdmVal2 = "" Then
        lblDistrict_.caption = "District: N/A"
    Else
        lblDistrict_.caption = "District: " & sAdmVal2
    End If

    If sAdmloc = "" Then
        lblNearestTown.caption = "Nearest Town: N/A"
    Else
        lblNearestTown.caption = "Nearest Town: " & sAdmloc
    End If

End Sub

Public Sub SetCoordinateValue(x As String, xLabel As String, y As String, yLabel As String, sCoordSysName As String, sMGRS As String)
    dxedtLong.EditValue = x
    'lblX.caption = xLabel
    dxEdtLat.EditValue = y
    'lblY.caption = yLabel
    txtMGRS.Text = sMGRS
    'FraXY.caption = sCoordSysName
End Sub

Public Sub SetXY(x As String, y As String, bZoomTo As Boolean)
    
    dxedtLong.EditValue = x
    dxEdtLat.EditValue = y
    
    If bZoomTo Then Call cmdCheckInMap_Click

End Sub

Private Sub cmdSearch_Click()
        '<EhHeader>
        On Error GoTo cmdSearch_Click_Err
        '</EhHeader>
        Dim RS As New adodb.Recordset
        Dim sSearchfield As String
        Dim sHelperSearchField As String
        Dim sTable2Search As String
        Dim oSearchLyr As TatukGIS_XDK9.XGIS_LayerVector
        Dim oshp As TatukGIS_XDK9.XGIS_Shape
    
100     lblRecords.caption = ""
102     SafeMoveFirst g_RSAppSettings
    
106     Set RS = New adodb.Recordset
    
        RaiseEvent FocusRecieved("Search")
    
110     g_RSAppSettings.Find "SettingName = 'HICPcodeLayer'"
112     sTable2Search = g_RSAppSettings.Fields.Item("SettingValue1").Value
        
        Set oSearchLyr = m_OGIS.get("af_nga_place_names") '(sTable2Search) '
        
        lstResult.Clear
        
        If oSearchLyr Is Nothing Then Exit Sub
        
114     If OptSearchType(0).Value = True Then
116         sSearchfield = g_RSAppSettings.Fields.Item("SettingValue3").Value
118         sHelperSearchField = g_RSAppSettings.Fields.Item("SettingValue4").Value
        Else
120         sSearchfield = g_RSAppSettings.Fields.Item("SettingValue4").Value
122         sHelperSearchField = g_RSAppSettings.Fields.Item("SettingValue3").Value
        End If
            
        oSearchLyr.Lock
        
        If chkUseFuzzy.Value = vbChecked Then
            Set oshp = oSearchLyr.FindFirst(oSearchLyr.Extent, "SORT_NAME LIKE '%" & txtPcode.Text & "%'", Nothing, "", True)
        Else
            Set oshp = oSearchLyr.FindFirst(oSearchLyr.Extent, "SORT_NAME = '" & txtPcode.Text & "'", Nothing, "", True)
        End If
    
        Do While Not oshp Is Nothing
            'lstResult.AddItem sSearchfield & " = " & .Item(sSearchfield).Value & " ; " & sHelperSearchField & " = " & .Item(sHelperSearchField).Value
            lstResult.AddItem oshp.GetField("SORT_NAME") & " Code; " & oshp.GetField("GIS_UID")
            lstResult.ItemData(lstResult.ListCount - 1) = CLng(oshp.GetField("GIS_UID"))
            Set oshp = oSearchLyr.FindNext
        Loop
        
        oSearchLyr.Unlock
    
        '132     RS.Open "SELECT * FROM " & sTable2Search & " WHERE " & sSearchfield & " LIKE '%" & Replace(txtPcode.Text, "'", "''") & "%'", cn
    
        '134     If RS.EOF And RS.BOF Then
        '136         MsgBox "Your search did not get any results... Refine your search and try again.", vbInformation, "OASIS Search Utility"
        '            Exit Sub
        '        End If
        '
        '138     RS.MoveFirst
        '140     lstResult.Clear
        '
        '142     With RS
        '
        '144         If sHelperSearchField <> "" Then
        '
        '146             Do While Not .EOF
        '
        '148                 With .Fields
        '150                     lstResult.AddItem sSearchfield & " = " & .Item(sSearchfield).Value & " ; " & sHelperSearchField & " = " & .Item(sHelperSearchField).Value
        '                    End With
        '
        '152                 .MoveNext
        '
        '                Loop
        '
        '            Else
        '
        '154             Do While Not .EOF
        '156                 lstResult.AddItem sSearchfield & " = " & .Fields.Item(sSearchfield).Value & " ; " & "PCode = N/A"
        '158                 .MoveNext
        '                Loop
        '
        '            End If
        '
        '        End With
        '
160     lblRecords.caption = "found"
        
        '<EhFooter>
        Exit Sub

cmdSearch_Click_Err:
        MsgBox Err.Description ' vbObjectError + 100, "OASISClient.OASISLocator.cmdSearch_Click", "OASISLocator component failure"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComDistrict_Click()
        '<EhHeader>
        On Error GoTo ComDistrict_Click_Err
        '</EhHeader>
100     'If m_bInitReady Then
            '*************Place******************************
    
            Dim oLyr As TatukGIS_XDK9.XGIS_LayerVector
            Dim oDistLayer As TatukGIS_XDK9.XGIS_LayerVector
            Dim aShape As TatukGIS_XDK9.XGIS_Shape
            Dim shp As TatukGIS_XDK9.XGIS_Shape
    
102         ComPlace.Enabled = False
        
104         SafeMoveFirst g_RSAppSettings
106         g_RSAppSettings.Find "SettingName = 'PCodeAdminLocation'"
    
108         Set oLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
    
            If oLyr Is Nothing Then Exit Sub
    
110         If ComDistrict.List(ComDistrict.ListIndex) = "--ALL--" Then
112             oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    
114             ComPlace.Clear
116             ComPlace.AddItem "--ALL--"
    
118             Do While Not oLyr.EOF
120                 ComPlace.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
124                 oLyr.MoveNext
                Loop
    
126             ComPlace.ListIndex = 0
            Else
    
128             'If chkAutoHighlight.Value = vbChecked Then
130                 RaiseEvent hglAdminLevel1
                'End If
    
132             SafeMoveFirst g_RSAppSettings
134             g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel1'"
    
136             Set oDistLayer = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
    
                If oDistLayer Is Nothing Then Exit Sub
                
138             Set shp = oDistLayer.FindFirst(m_OGIS.Extent, g_RSAppSettings.Fields.Item("SettingValue2").Value & " = '" & ComDistrict.List(ComDistrict.ListIndex) & "'", Nothing, "", True)
    
140             If shp Is Nothing Then
142                 MsgBox "The following Item could not be found on the map: "
                    Exit Sub
                End If
    
144             m_OGIS.VisibleExtent = shp.Extent
146              shp.Flash
    
148             SafeMoveFirst g_RSAppSettings
150             g_RSAppSettings.Find "SettingName = 'PCodeAdminLocation'"
    
152             ComPlace.Clear
154             ComPlace.AddItem "--ALL--"
    
156             Set aShape = oLyr.FindFirst(m_OGIS.Extent, "", shp, GisUtils.GIS_RELATE_CONTAINS, True)
                If aShape Is Nothing Then Exit Sub
                
158             Do While Not oLyr.EOF
160                 ComPlace.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
162                 oLyr.MoveNext
                Loop
    
164             ComPlace.ListIndex = 0
    
                'oLyr.FindFirst
                'oLyr.f
            End If

        'End If
    
166     ComPlace.Enabled = True

        '<EhFooter>
        Exit Sub

ComDistrict_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLocator.ComDistrict_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComPlace_Click()
        '<EhHeader>
        On Error GoTo ComPlace_Click_Err
        '</EhHeader>
       
        '*************DISTRICT******************************
    
        Dim oVillageLayer As TatukGIS_XDK9.XGIS_LayerVector
        Dim aShape As TatukGIS_XDK9.XGIS_Shape
        Dim shp As TatukGIS_XDK9.XGIS_Shape
    
100     If Not ComPlace.List(ComPlace.ListIndex) = "--ALL--" Then
102         SafeMoveFirst g_RSAppSettings
104         g_RSAppSettings.Find "SettingName = 'HICPcodeLayer'"

106         Set oVillageLayer = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
        
108         Set shp = oVillageLayer.FindFirst(m_OGIS.Extent, g_RSAppSettings.Fields.Item("SettingValue4").Value & " = '" & Replace(ComPlace.List(ComPlace.ListIndex), "'", "''") & "'", Nothing, "", True)
                
110         If Not shp Is Nothing Then
112             m_OGIS.VisibleExtent = shp.Extent
114             lblProvincePcode.caption = "Province Code: " & shp.GetField("ADM2CODE")
116             lblDistrictPCode.caption = "District Code: " & shp.GetField("ADM3CODE")
118             lblPlaceCode.caption = "Place Code: " & shp.GetField("NEW_PCODE")
        
                'GIS.Zoom = m_OGIS.Zoom / 2 'CInt(g_RSAppSettings.Fields.Item("SettingValue5").value)
120             Set shp = shp.MakeEditable
122             shp.Flash
                'shp.IsSelected = True
            End If
        End If

        '*****************************************

        '<EhFooter>
        Exit Sub

ComPlace_Click_Err:
        Err.Raise vbObjectError + 100, "OASISClient.OASISLocator.ComPlace_Click", "OASISLocator component failure"
        '</EhFooter>
End Sub

Private Sub ComProvince_Click()
        '<EhHeader>
        On Error GoTo ComProvince_Click_Err
        '</EhHeader>
100     'If m_bInitReady Then
102         ComDistrict.Enabled = True
104         'If chkAutoHighlight.Value = vbChecked Then
106             RaiseEvent hglAdminLevel0
            'End If
108    '     m_bInitReady = False
110         LoadAdmin1 "s"
            DoEvents
112    '     m_bInitReady = True
        'End If
        '<EhFooter>
        Exit Sub

ComProvince_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmLocator.ComProvince_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateLyrs()
        '<EhHeader>
        On Error GoTo CreateLyrs_Err
        '</EhHeader>

100     Set m_oDrawLyr = m_OGIS.get("Draw Layer")
112     Set ll2 = m_OGIS.get("Buffers")

        '<EhFooter>
        Exit Sub

CreateLyrs_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISLocator.CreateLyrs", _
                  "OASISLocator component failure"
        '</EhFooter>
End Sub

Private Sub GetAdmCode(ptg As TatukGIS_XDK9.XGIS_Point, _
                       sAdm0 As String, _
                       sAdm1 As String, _
                       sAdm2 As String, _
                       sAdm3 As String, _
                       sAdm4 As String, _
                       sAdm5 As String, _
                       sAdmloc As String)
        '<EhHeader>
        On Error GoTo GetAdmCode_Err
        '</EhHeader>

        Dim oVecLyr As TatukGIS_XDK9.XGIS_LayerVector
        Dim shp As TatukGIS_XDK9.XGIS_Shape
        Dim oshp As TatukGIS_XDK9.XGIS_Shape
        Dim j As Integer
    
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'AdminLevel0'"

104     Set oVecLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
                        
106     If Not oVecLyr Is Nothing Then
        
108         Set oshp = oVecLyr.Locate(ptg, 5 / m_OGIS.zoom, True)
                        
110         If Not oshp Is Nothing Then
        
112             SafeMoveFirst g_RSAppSettings
114             g_RSAppSettings.Find "SettingName = 'AdminLevel0'"
116             sAdm0 = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
118             sAdm0 = sAdm0 & " PCODE:" & oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
 
            End If
        End If
    
120     SafeMoveFirst g_RSAppSettings
122     g_RSAppSettings.Find "SettingName = 'AdminLevel1'"

124     If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
126         If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
128             Set oVecLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
130             Set oVecLyr = Nothing
            End If

        Else
132         Set oVecLyr = Nothing
        End If
    
134     If Not oVecLyr Is Nothing Then
        
136         Set oshp = oVecLyr.Locate(ptg, 5 / m_OGIS.zoom, True)
                        
138         If Not oshp Is Nothing Then
140             If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
                
142                 If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
144                     sAdm1 = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
146                     sAdm1 = sAdm1 & " PCODE:" & oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
                    End If
                End If
            End If
        End If
            
148     SafeMoveFirst g_RSAppSettings
150     g_RSAppSettings.Find "SettingName = 'AdminLevel2'"
            
152     If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
154         If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
156             Set oVecLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
158             Set oVecLyr = Nothing
            End If

        Else
160         Set oVecLyr = Nothing
        End If
                        
162     If Not oVecLyr Is Nothing Then
        
164         Set oshp = oVecLyr.Locate(ptg, 5 / m_OGIS.zoom, True)
                        
166         If Not oshp Is Nothing Then
            
168             If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
170                 If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
172                     sAdm2 = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
174                     sAdm2 = sAdm2 & oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
                    End If
                End If
            
            End If
        End If
            
176    SafeMoveFirst g_RSAppSettings
178     g_RSAppSettings.Find "SettingName = 'AdminLevel3'"
            
180     If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
182         If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
184             Set oVecLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
186             Set oVecLyr = Nothing
            End If

        Else
188         Set oVecLyr = Nothing
        End If
                        
190     If Not oVecLyr Is Nothing Then
        
192         Set oshp = oVecLyr.Locate(ptg, 5 / m_OGIS.zoom, True)
                        
194         If Not oshp Is Nothing Then
            
196             If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
198                 If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
200                     sAdm3 = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
202                     sAdm3 = sAdm3 & oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
                    End If
                End If
            
            End If
        End If
            
204     SafeMoveFirst g_RSAppSettings
206     g_RSAppSettings.Find "SettingName = 'AdminLevel4'"
            
208     If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
210         If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
212             Set oVecLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
214             Set oVecLyr = Nothing
            End If

        Else
216         Set oVecLyr = Nothing
        End If
                        
218     If Not oVecLyr Is Nothing Then
        
220         Set oshp = oVecLyr.Locate(ptg, 5 / m_OGIS.zoom, True)
                        
222         If Not oshp Is Nothing Then
        
224             If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
226                 If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
228                     sAdm4 = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
230                     sAdm4 = sAdm4 & oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
                    End If
                End If
    
            End If
        End If
    
232     SafeMoveFirst g_RSAppSettings
234     g_RSAppSettings.Find "SettingName = 'AdminLevel5'"
            
236     If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
238         If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
240             Set oVecLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
242             Set oVecLyr = Nothing
            End If

        Else
244         Set oVecLyr = Nothing
        End If
                        
246     If Not oVecLyr Is Nothing Then
        
248         Set oshp = oVecLyr.Locate(ptg, 5 / m_OGIS.zoom, True)
                        
250         If Not oshp Is Nothing Then
        
252             If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
254                 If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
256                     sAdm5 = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
258                     sAdm5 = sAdm5 & oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
                    End If
                End If
            End If
        End If
    
260     SafeMoveFirst g_RSAppSettings
262     g_RSAppSettings.Find "SettingName = 'AdminLocation'"
            
264     If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
266         If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
268             Set oVecLyr = m_OGIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
270             Set oVecLyr = Nothing
            End If

        Else
272         Set oVecLyr = Nothing
        End If
                        
274     If Not oVecLyr Is Nothing Then
            Dim iIncremental As Integer
        
276         Set oshp = Nothing
        
278         Do While oshp Is Nothing
280             iIncremental = iIncremental + 5
282             Set oshp = oVecLyr.Locate(ptg, iIncremental / m_OGIS.zoom, True)

284             If iIncremental > 500 Then
286                 GoTo ExitLoop
                End If

            Loop
        
ExitLoop:
        
288         If Not oshp Is Nothing Then
        
290             If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
292                 If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
294                     sAdmloc = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
296                     sAdmloc = sAdmloc & oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
298                     sAdmloc = sAdmloc & " Distance to community:" & oshp.distance(ptg, 4) '& oVecLyr.Units.Units
                    End If
                End If
    
            End If
        End If
    
        '<EhFooter>
        Exit Sub

GetAdmCode_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISLocator.GetAdmCode", _
                  "OASISLocator component failure"
        '</EhFooter>
End Sub

Private Sub lstResult_Click()
        '<EhHeader>
        On Error GoTo lstResult_Click_Err
        '</EhHeader>
        Dim oHicLayer As TatukGIS_XDK9.XGIS_LayerVector
        Dim shp As TatukGIS_XDK9.XGIS_Shape
        Dim sVals() As String
    
100     m_frmDebug.DebugPrint lstResult.List(lstResult.ListIndex)
    
102     SafeMoveFirst g_RSAppSettings
    
106     g_RSAppSettings.Find "SettingName = 'HICPcodeLayer'"

110    ' sVals = Split(lstResult.List(lstResult.ListIndex), " Code; ")
    
        'PETRI: We need to change this!!!!
112     Set oHicLayer = m_OGIS.get("af_nga_place_names") 'g_RSAppSettings.Fields.Item("SettingValue1").Value)
        
        If oHicLayer Is Nothing Then Exit Sub
        
        RaiseEvent FocusRecieved("PlaceFound")
        
114     Set shp = oHicLayer.FindFirst(m_OGIS.Extent, "GIS_UID = " & lstResult.ItemData(lstResult.ListIndex), Nothing, "", True)
        
116     If shp Is Nothing Then
118         MsgBox "The following Item could not be found on the map: "
            Exit Sub
        End If
        
120     m_OGIS.VisibleExtent = shp.Extent

        'GIS.Zoom = m_OGIS.Zoom / 2 'CInt(g_RSAppSettings.Fields.Item("SettingValue5").value)
122     Set shp = shp.MakeEditable
124     shp.Flash
        
        dxedtLong.EditValue = shp.Centroid.x
        dxEdtLat.EditValue = shp.Centroid.y
       
        '<EhFooter>
        Exit Sub

lstResult_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISLocator.lstResult_Click", _
                  "OASISLocator component failure"
        '</EhFooter>
End Sub

Private Sub OptDatasource_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo OptDatasource_Click_Err
        '</EhHeader>
    
100     lstResult.Clear
       ' lblLblRecords.caption = ""
102     lblRecords.caption = ""
    
104     If OptDatasource(0).Value = True Then
106         OptSearchType(0).Enabled = True
108         OptSearchType(1).Enabled = True
        Else
110         OptSearchType(0).Enabled = False
112         OptSearchType(1).Enabled = False
        End If
        '<EhFooter>
        Exit Sub

OptDatasource_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISLocator.OptDatasource_Click", _
                  "OASISLocator component failure"
        '</EhFooter>
End Sub

Private Sub UserControl_GotFocus()
    RaiseEvent FocusRecieved("")
End Sub

Private Sub UserControl_Initialize()
    m_bShowGeoSelector = True
    m_bShowAdminSelector = True
    m_intStartTab = 0
    m_DouLatitude = 0
    m_DouLongitude = 0
End Sub
