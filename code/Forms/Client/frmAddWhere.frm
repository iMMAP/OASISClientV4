VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Begin VB.Form frmAddWhere 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Where?"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   7170
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4380
      _cx             =   7726
      _cy             =   12647
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
      Begin C1SizerLibCtl.C1Elastic elNavigator 
         Height          =   420
         Left            =   180
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   6750
         Width           =   3750
         _cx             =   6615
         _cy             =   741
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
         Begin VB.CommandButton cmdOK 
            Caption         =   "OK"
            Height          =   285
            Left            =   1980
            TabIndex        =   18
            Top             =   90
            Width           =   645
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            Height          =   285
            Left            =   1350
            TabIndex        =   17
            Top             =   90
            Width           =   645
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit"
            Height          =   285
            Left            =   765
            TabIndex        =   16
            Top             =   90
            Width           =   600
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   285
            Left            =   135
            TabIndex        =   15
            Top             =   90
            Width           =   645
         End
      End
      Begin C1SizerLibCtl.C1Tab tabWhere 
         Height          =   6585
         Left            =   135
         TabIndex        =   1
         Top             =   135
         Width           =   4020
         _cx             =   7091
         _cy             =   11615
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
         Caption         =   "General|Location|Privacy"
         Align           =   0
         CurrTab         =   1
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
         Begin C1SizerLibCtl.C1Elastic elLocation 
            Height          =   6210
            Left            =   45
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   330
            Width           =   3930
            _cx             =   6932
            _cy             =   10954
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
            Begin TatukGIS_DK.XGIS_ViewerWnd GIS 
               Height          =   2580
               Left            =   90
               TabIndex        =   10
               Top             =   90
               Width           =   3165
               BigExtentMargin =   -10
               RestrictedDrag  =   -1  'True
               CachedPaint     =   -1  'True
               IncrementalPaint=   -1  'True
               FullPaint       =   -1  'True
               CodePage        =   0
               OutCodePage     =   0
               CharSet         =   0
               UseRTree        =   0   'False
               PrinterTileSize =   512
               PrintTitle      =   ""
               PrintSubtitle   =   ""
               PrintFooter     =   ""
               BeginProperty PrintTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty PrintSubtitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty PrintFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PrintTitleFontColor=   -16777208
               PrintSubtitleFontColor=   -16777208
               PrintFooterFontColor=   -16777208
               SelectionColor  =   16777215
               SelectionPattern=   "frmAddWhere.frx":0000
               SelectionTransparency=   100
               SelectionWidth  =   100
               SelectionOutlineOnly=   0   'False
               OldCachedPaint  =   0   'False
               PrinterModeDraft=   0   'False
               PrinterModeForceBitmap=   0   'False
               Mode            =   0
               BorderStyle     =   1
               CursorForDrag   =   0
               CursorForSelect =   0
               CursorForZoom   =   0
               CursorForEdit   =   0
               MinZoomSize     =   -5
               ScrollBars      =   0
               AutoCenter      =   0   'False
               Align           =   0
               Ctl3D           =   -1  'True
               ParentColor     =   0   'False
               ParentCtl3D     =   0   'False
               Object.Visible         =   -1  'True
               Cursor          =   16
               DoubleBuffered  =   0   'False
               ModeMouseButton =   0
               CursorForUserDefined=   0
            End
            Begin C1SizerLibCtl.C1Tab tabPCode 
               Height          =   2580
               Left            =   90
               TabIndex        =   19
               Top             =   3555
               Width           =   3525
               _cx             =   6218
               _cy             =   4551
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
               Caption         =   "Admin Locator|Geo Locator|PCode Locator"
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
               Begin C1SizerLibCtl.C1Elastic elPCodeLoc 
                  Height          =   2205
                  Left            =   4470
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   3435
                  _cx             =   6059
                  _cy             =   3889
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
               End
               Begin C1SizerLibCtl.C1Elastic elGeo 
                  Height          =   2205
                  Left            =   4170
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   3435
                  _cx             =   6059
                  _cy             =   3889
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
                  Begin VB.TextBox txtLatitude 
                     Height          =   285
                     Left            =   1035
                     TabIndex        =   27
                     Text            =   "Latitude"
                     Top             =   270
                     Width           =   1950
                  End
                  Begin VB.TextBox txtLongitude 
                     Height          =   285
                     Left            =   1035
                     TabIndex        =   26
                     Text            =   "Longitude"
                     Top             =   585
                     Width           =   1950
                  End
                  Begin VB.TextBox txtMGRS 
                     Height          =   285
                     Left            =   1035
                     TabIndex        =   25
                     Text            =   "MGRS"
                     Top             =   900
                     Width           =   1950
                  End
                  Begin VB.CommandButton cmdGetFrom 
                     Height          =   330
                     Left            =   3060
                     Picture         =   "frmAddWhere.frx":5D7E
                     Style           =   1  'Graphical
                     TabIndex        =   24
                     ToolTipText     =   "Get Coordinates From Map"
                     Top             =   495
                     Width           =   330
                  End
                  Begin VB.CommandButton cmdCheckIn 
                     Height          =   330
                     Left            =   3060
                     Picture         =   "frmAddWhere.frx":61CA
                     Style           =   1  'Graphical
                     TabIndex        =   23
                     ToolTipText     =   "Check Coordinate in Map"
                     Top             =   855
                     Width           =   330
                  End
                  Begin VB.CommandButton cmdRadius 
                     Height          =   330
                     Left            =   3060
                     Picture         =   "frmAddWhere.frx":6614
                     Style           =   1  'Graphical
                     TabIndex        =   22
                     Top             =   135
                     Width           =   330
                  End
                  Begin VB.Label lblLatitudeX 
                     AutoSize        =   -1  'True
                     Caption         =   "Latitude Y:"
                     Height          =   195
                     Left            =   0
                     TabIndex        =   30
                     Top             =   360
                     Width           =   765
                  End
                  Begin VB.Label lblLongitudeX 
                     AutoSize        =   -1  'True
                     Caption         =   "Longitude X:"
                     Height          =   195
                     Left            =   0
                     TabIndex        =   29
                     Top             =   630
                     Width           =   900
                  End
                  Begin VB.Label lblMGRS 
                     AutoSize        =   -1  'True
                     Caption         =   "MGRS:"
                     Height          =   195
                     Left            =   0
                     TabIndex        =   28
                     Top             =   900
                     Width           =   525
                  End
               End
               Begin C1SizerLibCtl.C1Elastic elAdmin 
                  Height          =   2205
                  Left            =   45
                  TabIndex        =   20
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   3435
                  _cx             =   6059
                  _cy             =   3889
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
                  Begin VB.ComboBox ComPlace 
                     Height          =   315
                     Left            =   45
                     Style           =   2  'Dropdown List
                     TabIndex        =   39
                     Top             =   1845
                     Width           =   2265
                  End
                  Begin VB.ComboBox ComDistrict 
                     Height          =   315
                     Left            =   45
                     Style           =   2  'Dropdown List
                     TabIndex        =   33
                     Top             =   1305
                     Width           =   2265
                  End
                  Begin VB.ComboBox ComProvince 
                     Height          =   315
                     Left            =   45
                     Style           =   2  'Dropdown List
                     TabIndex        =   32
                     Top             =   765
                     Width           =   2265
                  End
                  Begin VB.ComboBox ComCountry 
                     BackColor       =   &H80000011&
                     Enabled         =   0   'False
                     Height          =   315
                     Left            =   45
                     Style           =   2  'Dropdown List
                     TabIndex        =   31
                     Top             =   225
                     Width           =   2310
                  End
                  Begin VB.Label lblCommunity 
                     AutoSize        =   -1  'True
                     Caption         =   "Place:"
                     Height          =   195
                     Left            =   45
                     TabIndex        =   38
                     Top             =   1620
                     Width           =   450
                  End
                  Begin VB.Label lblDistrict 
                     AutoSize        =   -1  'True
                     Caption         =   "District:"
                     Height          =   195
                     Left            =   45
                     TabIndex        =   36
                     Top             =   1080
                     Width           =   525
                  End
                  Begin VB.Label lblProvince 
                     AutoSize        =   -1  'True
                     Caption         =   "Province:"
                     Height          =   195
                     Left            =   45
                     TabIndex        =   35
                     Top             =   540
                     Width           =   675
                  End
                  Begin VB.Label lblCountry 
                     AutoSize        =   -1  'True
                     Caption         =   "Country:"
                     Height          =   195
                     Left            =   45
                     TabIndex        =   34
                     Top             =   0
                     Width           =   585
                  End
               End
            End
            Begin ActiveBar3LibraryCtl.ActiveBar3 ActiveBar31 
               Height          =   1410
               Left            =   3375
               TabIndex        =   11
               Top             =   180
               Width           =   420
               _LayoutVersion  =   2
               _ExtentX        =   741
               _ExtentY        =   2487
               _DataPath       =   ""
               Bands           =   "frmAddWhere.frx":6A3D
            End
            Begin VB.Label lblAdminlevel 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   90
               TabIndex        =   14
               Top             =   2925
               Width           =   3165
            End
            Begin VB.Label lblCoords 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "X: n/a Y: n/a"
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   90
               TabIndex        =   12
               Top             =   2700
               Width           =   3165
            End
         End
         Begin C1SizerLibCtl.C1Elastic elGeneral 
            Height          =   6210
            Left            =   -4575
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   3930
            _cx             =   6932
            _cy             =   10954
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
            Begin VB.TextBox txtDescription 
               Height          =   2445
               Left            =   90
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   9
               Text            =   "frmAddWhere.frx":10BD5
               Top             =   1260
               Width           =   3525
            End
            Begin VB.ComboBox ComPlaceType 
               Height          =   315
               Left            =   1125
               TabIndex        =   8
               Text            =   "Place Type"
               Top             =   495
               Width           =   2265
            End
            Begin VB.TextBox txtPlaceName 
               Height          =   315
               Left            =   1125
               TabIndex        =   7
               Text            =   "PlaceName"
               Top             =   135
               Width           =   2265
            End
            Begin VB.Label lblDescription 
               AutoSize        =   -1  'True
               Caption         =   "Description:"
               Height          =   195
               Left            =   90
               TabIndex        =   6
               Top             =   990
               Width           =   840
            End
            Begin VB.Label lblPlaceType 
               AutoSize        =   -1  'True
               Caption         =   "Place Type:"
               Height          =   195
               Left            =   90
               TabIndex        =   5
               Top             =   585
               Width           =   855
            End
            Begin VB.Label lblPlaceName 
               AutoSize        =   -1  'True
               Caption         =   "Place Name:"
               Height          =   195
               Left            =   90
               TabIndex        =   4
               Top             =   135
               Width           =   915
            End
         End
         Begin C1SizerLibCtl.C1Elastic elSettings 
            Height          =   6210
            Left            =   4665
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   330
            Width           =   3930
            _cx             =   6932
            _cy             =   10954
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
            Begin VB.Frame FraPrivacy 
               Caption         =   "Privacy:"
               Height          =   735
               Left            =   270
               TabIndex        =   48
               Top             =   270
               Width           =   3390
               Begin VB.OptionButton OptPrivacy 
                  Caption         =   "Public"
                  Height          =   330
                  Index           =   2
                  Left            =   2250
                  TabIndex        =   49
                  Top             =   225
                  Width           =   1050
               End
               Begin VB.OptionButton OptPrivacy 
                  Caption         =   "User group"
                  Height          =   330
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   50
                  Top             =   225
                  Width           =   1230
               End
               Begin VB.OptionButton OptPrivacy 
                  Caption         =   "Private"
                  Height          =   330
                  Index           =   0
                  Left            =   135
                  TabIndex        =   51
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   1050
               End
            End
            Begin VB.Frame FraLocationVisibility 
               Caption         =   "Location visibility level:"
               Height          =   2670
               Left            =   270
               TabIndex        =   41
               Top             =   1215
               Width           =   3345
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "National"
                  Height          =   375
                  Index           =   0
                  Left            =   270
                  TabIndex        =   47
                  Top             =   225
                  Value           =   1  'Checked
                  Width           =   1680
               End
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "Province"
                  Height          =   375
                  Index           =   1
                  Left            =   270
                  TabIndex        =   46
                  Top             =   615
                  Value           =   1  'Checked
                  Width           =   1680
               End
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "District"
                  Height          =   375
                  Index           =   2
                  Left            =   270
                  TabIndex        =   45
                  Top             =   1005
                  Width           =   1680
               End
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "Sub District"
                  Height          =   375
                  Index           =   3
                  Left            =   270
                  TabIndex        =   44
                  Top             =   1380
                  Width           =   1680
               End
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "Town"
                  Height          =   375
                  Index           =   4
                  Left            =   270
                  TabIndex        =   43
                  Top             =   1770
                  Width           =   1680
               End
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "Actual Location"
                  Height          =   375
                  Index           =   5
                  Left            =   270
                  TabIndex        =   42
                  Top             =   2160
                  Width           =   1680
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmAddWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event PcodePicked()

Private m_cn As adodb.Connection
Private m_bGetCoordTool As Boolean
Private m_bSelectRadiusTool As Boolean
Private m_oShpPt As TatukGIS_XDK9.XGIS_ShapePoint
Private m_oDrawLyr As TatukGIS_XDK9.XGIS_LayerVector
Private m_PTUID As Long
Const RELATE_INTERSECT = "T"

Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Const R2_XORPEN = 7
Dim oldPos As New xPoint
Dim oldRadius As Integer
Dim lL As TatukGIS_XDK9.XGIS_LayerVector
Dim ll2 As TatukGIS_XDK9.XGIS_LayerVector

Public Sub Init(CN As Connection)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        Dim RS As New adodb.Recordset

100     Set m_cn = CN
    
102     SafeMoveFirst g_RSAppSettings
104     g_RSAppSettings.Find "SettingName = 'PCODEPickerMap'"
106     GIS.Open g_sAppPath & g_RSAppSettings.Fields.Item("SettingValue1").Value
             
        'GIS.FullExtent

108     RS.Open "SELECT name, id FROM 1placeType", m_cn, adOpenForwardOnly, adLockReadOnly
    
110     ComPlaceType.Clear
112     ComDistrict.Clear
114     ComProvince.Clear
116     ComCountry.Clear
    
        'ComPlaceType.AddItem "--ALL--"
118     ComDistrict.AddItem "--ALL--"
120     ComProvince.AddItem "--ALL--"
122     ComCountry.AddItem "--ALL--"
    
124     SafeMoveFirst RS
    
126     Do While Not RS.EOF

128         With RS.Fields
130             ComPlaceType.AddItem .Item("name").Value
            End With

132         RS.MoveNext
        Loop
    
134     Set RS = New adodb.Recordset

136     RS.Open "SELECT name, id FROM 1country", m_cn, adOpenForwardOnly, adLockReadOnly

138     Do While Not RS.EOF

140         With RS.Fields
142             ComCountry.AddItem .Item("name").Value
            End With

144         RS.MoveNext
        Loop
    

    
    '    Set RS = New ADODB.Recordset
    '
    '    RS.Open "SELECT name, id FROM 1admin2Names", m_cn, adOpenForwardOnly, adLockReadOnly
    '
    '    Do While Not RS.EOF
    '
    '        With RS.Fields
    '            ComDistrict.AddItem .Item("name").value
    '        End With
    '
    '        RS.MoveNext
    '    Loop
    
        '*************PROVINCE******************************
    
        Dim oLyr As TatukGIS_XDK9.XGIS_LayerVector
    
146     SafeMoveFirst g_RSAppSettings
148     g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel0'"
    
150     Set oLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
    
152     oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    
154     Do While Not oLyr.EOF
156         ComProvince.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
158         m_frmDebug.DebugPrint oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
160         oLyr.MoveNext
        Loop
    
        '*****************************************
    
    '    '*************Place******************************
    '
    '
    '    g_RSAppSettings.MoveFirst
    '    g_RSAppSettings.Find "SettingName = 'PCodeAdminLocation'"
    '
    '    Set oLyr = GIS.Get(g_RSAppSettings.Fields.Item("SettingValue1").value)
    '
    '    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    '
    '    Do While Not oLyr.EOF
    '        ComPlace.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value)
    '        m_frmDebug.debugprint oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue3").value)
    '        oLyr.MoveNext
    '    Loop
    '
    '    '*****************************************
    
        '    Set RS = New ADODB.Recordset
        '
        '    RS.Open "SELECT name, id FROM 1admin1Names", m_cn, adOpenForwardOnly, adLockReadOnly
        '
        '    Do While Not RS.EOF
        '        With RS.Fields
        '            ComProvince.AddItem .Item("name").value
        '        End With
        '
        '        RS.MoveNext
        '    Loop
    
162     txtDescription.Text = ""
164     txtLatitude.Text = ""
166     txtLongitude.Text = ""
168     txtMGRS.Text = ""
170     txtPlaceName.Text = ""
    
        On Error Resume Next
172     ComPlaceType.ListIndex = 0
        'ComDistrict.ListIndex = 0
174     ComProvince.ListIndex = 0
176     ComCountry.ListIndex = 0
    
178     tabWhere.TabVisible(0) = True
180     tabWhere.CurrTab = 0
    
182     SafeMoveFirst g_RSAppSettings
184     g_RSAppSettings.Find "SettingName = 'w3WHERELocator'"
186     tabWhere.TabVisible(1) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)

188     SafeMoveFirst g_RSAppSettings
190     g_RSAppSettings.Find "SettingName = 'w3WHERELocation'"
192     tabWhere.TabVisible(2) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)
    
194     SafeMoveFirst g_RSAppSettings
196     g_RSAppSettings.Find "SettingName = 'w3WHEREPrivacy'"
198     tabWhere.TabVisible(3) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = "1", True, False)

200     CreateLyrs
    
202     ItemInBox "Iraq", ComCountry
    
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.init " & _
               "at line " & Erl
        Resume Next
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.IsInitialized " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdAdd_Click()
        '<EhHeader>
        On Error GoTo cmdAdd_Click_Err
        '</EhHeader>
100     txtDescription.Text = ""
102     txtLatitude.Text = ""
104     txtLongitude.Text = ""
106     txtMGRS.Text = ""
108     txtPlaceName.Text = ""
    
        On Error Resume Next
110     ComPlaceType.ListIndex = 0
112     ComDistrict.ListIndex = 0
114     ComProvince.ListIndex = 0
116     ComCountry.ListIndex = 0
    
118     tabWhere.CurrTab = 0
120     cmdAdd.Enabled = False
        '<EhFooter>
        Exit Sub

cmdAdd_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.cmdAdd_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdGetFrom_Click()
        '<EhHeader>
        On Error GoTo cmdGetFrom_Click_Err
        '</EhHeader>
100     m_bGetCoordTool = True
        '<EhFooter>
        Exit Sub

cmdGetFrom_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.cmdGetFrom_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOk_Click()
        '<EhHeader>
        On Error GoTo cmdOk_Click_Err
        '</EhHeader>
    Dim sVals As String
    Dim sID As String
    Dim RS As adodb.Recordset

100     If cmdAdd.Enabled Then
            'An Edit Update
        Else
102         sID = GUIDGen
104         sVals = "'" & sID & "'"
106         sVals = sVals & ", '" & txtPlaceName.Text & "'"
        
108         Set RS = New adodb.Recordset
110         RS.Open "SELECT id,name FROM 1placeType WHERE name = '" & ComPlaceType.List(ComPlaceType.ListIndex) & "'"
112         sVals = sVals & ", '" & RS.Fields.Item("id").Value & "'"
        
114         sVals = sVals & ", " & Replace(txtLatitude.Text, ",", ".")
116         sVals = sVals & ", " & Replace(txtLongitude.Text, ",", ".")
        
118         sVals = sVals & ", '0'" 'ADm1
120         sVals = sVals & ", '0'" 'ADm2
122         sVals = sVals & ", '0'" 'ADm3
124         sVals = sVals & ", '0'" 'ADm4
        
126         m_cn.Execute "INSERT INTO 1placeName (id, name, placeTypeId, latX, longY, admin1Id, admin2Id, admin3Id, admin4Id) VALUES (" & sVals & ")"
        
        
128         m_cn.Execute "INSERT INTO visibility (ExternalID, ExternalTable, Privacy, visNational, visProvince, visDistrict, visSubDistrict, visCommunity, visActualLocation) VALUES (" & sVals & ")"
    
        End If
        '<EhFooter>
        Exit Sub

cmdOk_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.cmdOK_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdRadius_Click()
        '<EhHeader>
        On Error GoTo cmdRadius_Click_Err
        '</EhHeader>
100     GIS.Mode = TatukGIS_XDK9.XgisSelect
102     m_bSelectRadiusTool = True
        '<EhFooter>
        Exit Sub

cmdRadius_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.cmdRadius_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComDistrict_Click()
        '*************Place******************************
        '<EhHeader>
        On Error GoTo ComDistrict_Click_Err
        '</EhHeader>
    
        Dim oLyr As TatukGIS_XDK9.XGIS_LayerVector
        Dim oDistLayer As TatukGIS_XDK9.XGIS_LayerVector
        Dim aShape As TatukGIS_XDK9.XGIS_Shape
        Dim shp As TatukGIS_XDK9.XGIS_Shape
    
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'PCodeAdminLocation'"
    
104     Set oLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
    
106     If ComDistrict.List(ComDistrict.ListIndex) = "--ALL--" Then
108         oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
        
110         ComPlace.Clear
112         ComPlace.AddItem "--ALL--"
        
114         Do While Not oLyr.EOF
116             ComPlace.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
118             m_frmDebug.DebugPrint oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
120             oLyr.MoveNext
            Loop
        
122         ComPlace.ListIndex = 0
        Else
        
124         SafeMoveFirst g_RSAppSettings
126         g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel1'"

128         Set oDistLayer = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
        
130         Set shp = oDistLayer.FindFirst(GIS.Extent, g_RSAppSettings.Fields.Item("SettingValue2").Value & " = '" & ComDistrict.List(ComDistrict.ListIndex) & "'", Nothing, "", True)
        
132         If shp Is Nothing Then
134             MsgBox "The following Item could not be found on the map: "
                Exit Sub
            End If
        
136         GIS.VisibleExtent = shp.Extent
138         GIS.zoom = GIS.zoom / CInt(g_RSAppSettings.Fields.Item("SettingValue5").Value)
140         shp.Flash

142         SafeMoveFirst g_RSAppSettings
144         g_RSAppSettings.Find "SettingName = 'PCodeAdminLocation'"
        
146         ComPlace.Clear
148         ComPlace.AddItem "--ALL--"
        
150         Set aShape = oLyr.FindFirst(GIS.Extent, "", shp, GisUtils.GIS_RELATE_CONTAINS, True)
152         Do While Not oLyr.EOF
154             ComPlace.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
156             oLyr.MoveNext
            Loop
        
158         ComPlace.ListIndex = 0
        
        End If
        '<EhFooter>
        Exit Sub

ComDistrict_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.ComDistrict_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComPlace_Click()
        '*************DISTRICT******************************
        '<EhHeader>
        On Error GoTo ComPlace_Click_Err
        '</EhHeader>
    
        Dim oVillageLayer As TatukGIS_XDK9.XGIS_LayerVector
        Dim aShape As TatukGIS_XDK9.XGIS_Shape
        Dim shp As TatukGIS_XDK9.XGIS_Shape
    
    
100     If Not ComPlace.List(ComPlace.ListIndex) = "--ALL--" Then
102         SafeMoveFirst g_RSAppSettings
104         g_RSAppSettings.Find "SettingName = 'PCodeAdminLocation'"

106         Set oVillageLayer = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
        
108         Set shp = oVillageLayer.FindFirst(GIS.Extent, g_RSAppSettings.Fields.Item("SettingValue2").Value & " = '" & ComPlace.List(ComPlace.ListIndex) & "'", Nothing, "", True)
                
110         GIS.VisibleExtent = shp.Extent
112         GIS.zoom = GIS.zoom / CInt(g_RSAppSettings.Fields.Item("SettingValue5").Value)
114         shp.Flash
                
        End If

        '*****************************************

        '<EhFooter>
        Exit Sub

ComPlace_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.ComPlace_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComProvince_Click()
        '<EhHeader>
        On Error GoTo ComProvince_Click_Err
        '</EhHeader>
        
        '*************DISTRICT******************************
    
        Dim oLyr As TatukGIS_XDK9.XGIS_LayerVector
        Dim oProvLayer As TatukGIS_XDK9.XGIS_LayerVector
        Dim aShape As TatukGIS_XDK9.XGIS_Shape
        Dim shp As TatukGIS_XDK9.XGIS_Shape
    
100    SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel1'"
    
104     Set oLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
    
106     If ComProvince.List(ComProvince.ListIndex) = "--ALL--" Then
108         oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    
110         Do While Not oLyr.EOF
112             ComDistrict.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
114             m_frmDebug.DebugPrint oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
116             oLyr.MoveNext
            Loop
        
118         ComDistrict.ListIndex = 0
        Else
        
120         SafeMoveFirst g_RSAppSettings
122         g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel0'"

124         Set oProvLayer = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
        
126         Set shp = oProvLayer.FindFirst(GIS.Extent, g_RSAppSettings.Fields.Item("SettingValue2").Value & " = '" & ComProvince.List(ComProvince.ListIndex) & "'", Nothing, "", True)
                
128         m_frmDebug.DebugPrint shp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
                
130         GIS.VisibleExtent = shp.Extent
132         GIS.zoom = GIS.zoom / CInt(g_RSAppSettings.Fields.Item("SettingValue5").Value)
134         shp.Flash
                
136         ComDistrict.Clear
138         ComDistrict.AddItem "--ALL--"
        
140         SafeMoveFirst g_RSAppSettings
142         g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel1'"
        
144         Set aShape = oLyr.FindFirst(GIS.Extent, "", shp, GisUtils.GIS_RELATE_CONTAINS, True)
146         Do While Not oLyr.EOF
148             ComDistrict.AddItem oLyr.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
            
150             oLyr.MoveNext
            Loop
        
152         ComDistrict.ListIndex = 0
        
            'oLyr.FindFirst
            'oLyr.f
        End If

        '*****************************************

        '<EhFooter>
        Exit Sub

ComProvince_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.ComProvince_Click " & _
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
               "in OASISClient.frmAddWhere.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS_OnMouseDown(translated As Boolean, ByVal Button As TatukGIS_XDK9.XMouseButton, ByVal Shift As TatukGIS_XDK9.XShiftState, ByVal x As Long, ByVal y As Long)
        '<EhHeader>
        On Error GoTo GIS_OnMouseDown_Err
        '</EhHeader>
100   Set oldPos = GisUtils.Point(x, y)
102   oldRadius = 0
        '<EhFooter>
        Exit Sub

GIS_OnMouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.GIS_OnMouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS_OnMouseMove(translated As Boolean, _
                            ByVal Shift As TatukGIS_XDK9.XShiftState, _
                            ByVal x As Long, _
                            ByVal y As Long)
        '<EhHeader>
        On Error GoTo GIS_OnMouseMove_Err
        '</EhHeader>
        Dim ptg As TatukGIS_XDK9.XGIS_Point
        Dim shp As TatukGIS_XDK9.XGIS_Shape
        Dim i As Integer
  
100     Set ptg = GIS.ScreenToMap(GisUtils.Point(x, y))
  
102     Set ptg = GIS.ScreenToMap(GisUtils.Point(x, y))
104     Set shp = GIS.Locate(ptg, 5 / GIS.zoom) ' 5 pixels precision

106     lblCoords.caption = "X:" & Round(ptg.x, 6) & " Y:" & Round(ptg.y, 6)

108     If Not shp Is Nothing Then
110         Me.caption = shp.GetField("Name")
        End If
        
112     If GIS.Mode <> TatukGIS_XDK9.XgisSelect Then Exit Sub
114     If Not (Shift = XssLeft) Then Exit Sub

116     SetROP2 GIS.hdc, R2_XORPEN
   
118     If oldRadius <> 0 Then
120         Ellipse GIS.hdc, oldPos.x - oldRadius, oldPos.y - oldRadius, oldPos.x + oldRadius, oldPos.y + oldRadius
        End If

122     oldRadius = Round(Sqr(((oldPos.x - x) * (oldPos.x - x)) + ((oldPos.y - y) * (oldPos.y - y))))

124     Ellipse GIS.hdc, oldPos.x - oldRadius, oldPos.y - oldRadius, oldPos.x + oldRadius, oldPos.y + oldRadius
    
126     translated = True
        '<EhFooter>
        Exit Sub

GIS_OnMouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.GIS_OnMouseMove " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ActiveBar31_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)
        '<EhHeader>
        On Error GoTo ActiveBar31_ToolClick_Err
        '</EhHeader>

100     m_bGetCoordTool = False

102     Select Case Tool.Name

            Case Is = "btnZoomin"
104             GIS.zoom = GIS.zoom * 2

106         Case Is = "btnZoomout"
108             GIS.zoom = GIS.zoom / 2

110         Case Is = "btnZoom"
112             GIS.Mode = TatukGIS_XDK9.XgisZoomEx

114         Case Is = "btnPan"
116             GIS.Mode = TatukGIS_XDK9.XgisDrag

118         Case Is = "btnSelect"
120             GIS.Mode = TatukGIS_XDK9.XgisSelect

122         Case "btnZoom"

124         Case "btnSelect"

126         Case "btnDeselect"

128         Case "btnInfo"
130             GIS.Mode = TatukGIS_XDK9.XgisSelect

132         Case "btnFullExtent"
134             GIS.FullExtent

136         Case "btnSelectionExtent"

138         Case "btnPrevExtent"

140         Case "btnRefreshMap"
142             GIS.UpDate

        End Select

        '<EhFooter>
        Exit Sub

ActiveBar31_ToolClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.ActiveBar31_ToolClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateLyrs()
        '<EhHeader>
        On Error GoTo CreateLyrs_Err
        '</EhHeader>

100     Set m_oDrawLyr = New TatukGIS_XDK9.XGIS_LayerVector
102     m_oDrawLyr.Params.area.Color = RGB(0, 0, 255)
104     m_oDrawLyr.Transparency = 50
106     m_oDrawLyr.Name = "Draw Layer"
108     m_oDrawLyr.HideFromLegend = True
  
110     GIS.Add m_oDrawLyr
  
112     Set ll2 = New TatukGIS_XDK9.XGIS_LayerVector
114     ll2.Params.area.Color = RGB(0, 0, 255)
116     ll2.Params.area.OutlineColor = RGB(0, 0, 255)
118     ll2.Transparency = 60
120     ll2.Name = "Buffers"
122     GIS.Add ll2

        '<EhFooter>
        Exit Sub

CreateLyrs_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.CreateLyrs " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS_OnMouseUp(translated As Boolean, _
                          ByVal Button As TatukGIS_XDK9.XMouseButton, _
                          ByVal Shift As TatukGIS_XDK9.XShiftState, _
                          ByVal x As Long, _
                          ByVal y As Long)
        '<EhHeader>
        On Error GoTo GIS_OnMouseUp_Err
        '</EhHeader>
                          
        Dim tpl As TatukGIS_XDK9.IXGIS_Topology
        Dim lL As TatukGIS_XDK9.IXGIS_LayerVector
        Dim tmp As TatukGIS_XDK9.IXGIS_Shape
        Dim buf1 As TatukGIS_XDK9.IXGIS_Shape
        Dim buf2 As TatukGIS_XDK9.IXGIS_Shape
        Dim ptg As TatukGIS_XDK9.IXGIS_Point
        Dim ptg1 As TatukGIS_XDK9.IXGIS_Point
        Dim distance As Double
        Dim sVals As String
    
100     If m_bGetCoordTool Then
102         translated = True

            Dim lUID As Long
            Dim i As Integer
            Dim j As Integer
            Dim fdesc As String
            Dim oVecLyr As TatukGIS_XDK9.XGIS_LayerVector
        
            Dim shp As TatukGIS_XDK9.XGIS_Shape
            Dim oshp As TatukGIS_XDK9.XGIS_Shape
  
104         Set ptg = GIS.ScreenToMap(GisUtils.Point(x, y))
106         Set shp = GIS.Locate(ptg, 5 / GIS.zoom) ' 5 pixels precision
            
108         If Not m_oShpPt Is Nothing Then
110             m_oDrawLyr.Delete m_oShpPt.uID
            End If
            
112         Set m_oShpPt = m_oDrawLyr.CreateShape(XgisShapeTypePoint)
        
            'm_PTUID = m_oShpPt.Uid
        
114         m_oShpPt.Lock TatukGIS_XDK9.XgisLockExtent
116         m_oShpPt.AddPart

118         m_oShpPt.AddPoint ptg
            'm_oShpPt.SetField "ID", Rnd(300000)
            'm_oShpPt.SetField "Name", m_fmrAddIncident.txtEnteredBy.Text
            'm_oShpPt.SetField "Type", m_fmrAddIncident.ComIncType.List(m_fmrAddIncident.ComIncType.ListIndex)
            'm_oShpPt.SetField "Target", m_fmrAddIncident.ComIncTarget.List(m_fmrAddIncident.ComIncTarget.ListIndex)
            ' m_oIncShpPt.SetField "Time", Now
            Dim sAdm0 As String
            Dim sAdm1 As String
            Dim sAdm2 As String
            Dim sAdm3 As String
            Dim sAdm4 As String
            Dim sAdm5 As String
            Dim sAdmloc As String
        
120         txtLatitude.Text = ptg.y
122         txtLongitude.Text = ptg.x
        
124         m_oShpPt.Unlock

126         GetAdmCode ptg, sAdm0, sAdm1, sAdm2, sAdm3, sAdm4, sAdm5, sAdmloc
128         lblAdminlevel.caption = "Province: " & sAdm0 & " District: " & sAdm1 & " Community:" & sAdmloc
130         GIS.UpDate
132         m_bGetCoordTool = False
134     ElseIf m_bSelectRadiusTool Then
136         If oldRadius = 0 Then
              Exit Sub
            End If

138         Set ptg = GIS.ScreenToMap(oldPos)
            'GIS.Get("Points").Lock
140         m_oDrawLyr.Lock
142         Set tmp = m_oDrawLyr.CreateShape(XgisShapeTypePoint)
144         tmp.Params.Marker.Size = 0
146         tmp.Lock TatukGIS_XDK9.XgisLockExtent
148         tmp.AddPart
150         tmp.AddPoint ptg
152         tmp.Unlock
154         GIS.get("Buffers").RevertAll

156         m_oDrawLyr.Unlock
158         Set tpl = New TatukGIS_XDK9.XGIS_Topology
            'distance recalc
160         Set ptg1 = GIS.ScreenToMap(GisUtils.Point(oldPos.x + oldRadius, y))
162         distance = ptg1.x - ptg.x

164         Set buf1 = tpl.MakeBuffer(tmp, distance, 36, True)
166         Set buf2 = GIS.get("Buffers").AddShape(buf1)
168         Set buf1 = Nothing
170         Set tpl = Nothing

172         SafeMoveFirst g_RSAppSettings
174         g_RSAppSettings.Find "SettingName = 'PCodeAdminLocation'"
        
176         Set lL = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value) '("completegazetteer") '("Districts_region")
178         If lL Is Nothing Then
180             GIS.UpDate
                Exit Sub
            End If

182         lL.DeselectAll

            'lstItems.Clear
        
            ' check all shapes
184         Set tmp = lL.FindFirst(buf2.Extent, "", buf2, RELATE_INTERSECT, True)
186         While Not tmp Is Nothing
                ' if any has a common point with buffer mark it
188             Set tmp = tmp.MakeEditable
190             sVals = sVals & vbCrLf & tmp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value) '("CITY")
192             tmp.IsSelected = True
194             Set tmp = lL.FindNext
            Wend
196         GIS.UpDate
198         MsgBox sVals
        End If

        '<EhFooter>
        Exit Sub

GIS_OnMouseUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.GIS_OnMouseUp " & _
               "at line " & Erl
        Resume Next
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

104     Set oVecLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
                        
106     If Not oVecLyr Is Nothing Then
        
108         Set oshp = oVecLyr.Locate(ptg, 5 / GIS.zoom, True)
                        
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
128             Set oVecLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
130             Set oVecLyr = Nothing
            End If

        Else
132         Set oVecLyr = Nothing
        End If
    
134     If Not oVecLyr Is Nothing Then
        
136         Set oshp = oVecLyr.Locate(ptg, 5 / GIS.zoom, True)
                        
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
156             Set oVecLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
158             Set oVecLyr = Nothing
            End If

        Else
160         Set oVecLyr = Nothing
        End If
                        
162     If Not oVecLyr Is Nothing Then
        
164         Set oshp = oVecLyr.Locate(ptg, 5 / GIS.zoom, True)
                        
166         If Not oshp Is Nothing Then
            
168             If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
170                 If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
172                     sAdm2 = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
174                     sAdm2 = sAdm2 & oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue3").Value)
                    End If
                End If
            
            End If
        End If
            
176     SafeMoveFirst g_RSAppSettings
178     g_RSAppSettings.Find "SettingName = 'AdminLevel3'"
            
180     If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
182         If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
184             Set oVecLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
186             Set oVecLyr = Nothing
            End If

        Else
188         Set oVecLyr = Nothing
        End If
                        
190     If Not oVecLyr Is Nothing Then
        
192         Set oshp = oVecLyr.Locate(ptg, 5 / GIS.zoom, True)
                        
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
212             Set oVecLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
214             Set oVecLyr = Nothing
            End If

        Else
216         Set oVecLyr = Nothing
        End If
                        
218     If Not oVecLyr Is Nothing Then
        
220         Set oshp = oVecLyr.Locate(ptg, 5 / GIS.zoom, True)
                        
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
240             Set oVecLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
            Else
242             Set oVecLyr = Nothing
            End If

        Else
244         Set oVecLyr = Nothing
        End If
                        
246     If Not oVecLyr Is Nothing Then
        
248         Set oshp = oVecLyr.Locate(ptg, 5 / GIS.zoom, True)
                        
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
268             Set oVecLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
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
282             Set oshp = oVecLyr.Locate(ptg, iIncremental / GIS.zoom, True)
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
    
        'For j = 0 To oVecLyr.Fields.Count - 1
        '    m_frmDebug.debugprint oVecLyr.Name & "_____" & oVecLyr.Fields.Item(j).Name & ":" & oShp.GetField(oVecLyr.Fields.Item(j).Name)
        'Next

        '<EhFooter>
        Exit Sub

GetAdmCode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddWhere.GetAdmCode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

