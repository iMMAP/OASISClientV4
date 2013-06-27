VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Begin VB.Form frmSelections 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Selections"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   300
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elAttr 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5805
      _cx             =   10239
      _cy             =   12356
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
      AutoSizeChildren=   8
      BorderWidth     =   0
      ChildSpacing    =   0
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
      _GridInfo       =   $"frmSelections.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic elTop 
         Height          =   555
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   5805
         _cx             =   10239
         _cy             =   979
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
         BorderWidth     =   1
         ChildSpacing    =   1
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
         _GridInfo       =   $"frmSelections.frx":003E
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.ComboBox ComSelLayer 
            Height          =   315
            Left            =   15
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   240
            Width           =   5775
         End
         Begin VB.Label lblActiveInfo 
            AutoSize        =   -1  'True
            Caption         =   "Selection/Reporting Layer:"
            Height          =   210
            Left            =   15
            TabIndex        =   2
            Top             =   15
            Width           =   5775
         End
      End
      Begin C1SizerLibCtl.C1Tab tbShapAttr 
         Height          =   6450
         Left            =   0
         TabIndex        =   3
         Top             =   555
         Width           =   5805
         _cx             =   10239
         _cy             =   11377
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
         Caption         =   "Attributes|Geo Summary|Reports|Settings"
         Align           =   0
         CurrTab         =   3
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   4
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
         Begin TatukGIS_DK.XGIS_ControlAttributes GEO1 
            Height          =   6075
            Left            =   -6660
            TabIndex        =   4
            Top             =   330
            Width           =   5715
            ReadOnly        =   -1  'True
            AllowRestructure=   -1  'True
            ColorHeader     =   -16777201
            ColorGrid       =   -16777211
            Align           =   0
            BevelInner      =   0
            BevelOuter      =   0
            Ctl3D           =   -1  'True
            BorderStyle     =   0
            Color           =   -16777201
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ParentColor     =   0   'False
            ParentCtl3D     =   -1  'True
            Object.Visible         =   -1  'True
            ParentBackground=   0   'False
            DoubleBuffered  =   0   'False
            AllowNull       =   0   'False
            BevelWidth      =   1
            BorderWidth     =   0
            HelpContextId   =   0
            ParentFont      =   -1  'True
            TabOrder        =   -1
            TabStop         =   0   'False
         End
         Begin C1SizerLibCtl.C1Elastic elSelSettings 
            Height          =   6075
            Left            =   45
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   330
            Width           =   5715
            _cx             =   10081
            _cy             =   10716
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
            Begin VB.Frame FraGeneral 
               Caption         =   "General:"
               Height          =   1575
               Left            =   120
               TabIndex        =   32
               Top             =   1200
               Width           =   2295
               Begin VB.CheckBox chkAutoZoom 
                  Caption         =   "Auto Zoom"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   37
                  Top             =   1140
                  Width           =   1155
               End
               Begin VB.CheckBox chkAutoSelect 
                  Caption         =   "Auto Select"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   36
                  Top             =   780
                  Width           =   1215
               End
               Begin VB.CheckBox chkAutoFlash 
                  Caption         =   "Auto Flash"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   35
                  Top             =   540
                  Width           =   1335
               End
               Begin VB.CheckBox chkAutomaticClear 
                  Caption         =   "Automatic Clear"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   33
                  Top             =   240
                  Value           =   1  'Checked
                  Width           =   1515
               End
            End
            Begin VB.Frame FraSelectionSettings 
               Caption         =   "Selection Settings:"
               Height          =   1035
               Left            =   0
               TabIndex        =   22
               Top             =   0
               Width           =   2445
               Begin VB.ComboBox ComBuffLevel 
                  Height          =   315
                  ItemData        =   "frmSelections.frx":007A
                  Left            =   150
                  List            =   "frmSelections.frx":0096
                  Style           =   2  'Dropdown List
                  TabIndex        =   24
                  Top             =   210
                  Width           =   1335
               End
               Begin VB.ComboBox txtSpatialOperation 
                  Height          =   315
                  ItemData        =   "frmSelections.frx":00C2
                  Left            =   150
                  List            =   "frmSelections.frx":00C9
                  Style           =   2  'Dropdown List
                  TabIndex        =   23
                  Top             =   540
                  Width           =   2175
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   6075
            Left            =   -6960
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   330
            Width           =   5715
            _cx             =   10081
            _cy             =   10716
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
            GridRows        =   2
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmSelections.frx":00D8
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame Frame2 
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   30
               TabIndex        =   7
               Top             =   5685
               Width           =   5655
               Begin VB.ComboBox ComFeatureLayer 
                  Height          =   315
                  Left            =   3060
                  Style           =   2  'Dropdown List
                  TabIndex        =   42
                  Top             =   60
                  Width           =   2535
               End
               Begin OASISClient.OASISButton OASSelect 
                  Height          =   315
                  Left            =   60
                  TabIndex        =   41
                  Top             =   60
                  Width           =   1875
                  _ExtentX        =   3307
                  _ExtentY        =   556
                  ButtonStyle     =   1
                  Text            =   "Select"
                  TextColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin VB.CommandButton cmdClear 
                  Caption         =   "Clear All"
                  Height          =   285
                  Left            =   4980
                  TabIndex        =   8
                  Top             =   60
                  Width           =   975
               End
               Begin VB.Label lblFeatureLayer 
                  AutoSize        =   -1  'True
                  Caption         =   "Feature Layer:"
                  Height          =   195
                  Left            =   1980
                  TabIndex        =   43
                  Top             =   120
                  Width           =   1020
               End
            End
            Begin C1SizerLibCtl.C1Tab C1TTab1Tab2 
               Height          =   5625
               Left            =   30
               TabIndex        =   6
               Top             =   30
               Width           =   5655
               _cx             =   9975
               _cy             =   9922
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
               FrontTabColor   =   -2147483633
               BackTabColor    =   -2147483633
               TabOutlineColor =   -2147483633
               FrontTabForeColor=   -2147483630
               Caption         =   "Tab&1"
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
               TabHeight       =   1
               TabCaptionPos   =   4
               TabPicturePos   =   0
               CaptionEmpty    =   ""
               Separators      =   0   'False
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   37
               Begin C1SizerLibCtl.C1Elastic AttributeHolder 
                  Height          =   5595
                  Left            =   15
                  TabIndex        =   25
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   5625
                  _cx             =   9922
                  _cy             =   9869
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
                  AutoSizeChildren=   7
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
                  Begin VB.Frame frAttrProperties 
                     BorderStyle     =   0  'None
                     Height          =   810
                     Left            =   135
                     TabIndex        =   27
                     Top             =   4680
                     Width           =   5355
                     Begin VB.CheckBox chkSelectIn 
                        Caption         =   "Select In Map"
                        Height          =   255
                        Left            =   60
                        TabIndex        =   40
                        Top             =   300
                        Width           =   1395
                     End
                     Begin VB.CommandButton cmdZoomTo 
                        Caption         =   "Zoom To"
                        Height          =   315
                        Left            =   4020
                        TabIndex        =   39
                        Top             =   360
                        Width           =   1215
                     End
                     Begin VB.CommandButton cmdFlash 
                        Caption         =   "Flash"
                        Height          =   315
                        Left            =   4020
                        TabIndex        =   38
                        Top             =   60
                        Width           =   1215
                     End
                     Begin VB.CheckBox chkEdit 
                        Caption         =   "Edit"
                        Height          =   255
                        Left            =   60
                        TabIndex        =   34
                        Top             =   540
                        Width           =   1275
                     End
                     Begin VB.CommandButton cmdRemove 
                        Caption         =   "Remove"
                        Height          =   435
                        Left            =   5040
                        TabIndex        =   29
                        Top             =   420
                        Visible         =   0   'False
                        Width           =   795
                     End
                     Begin VB.CheckBox chkUseInRprt 
                        Caption         =   "Use In Report"
                        Height          =   375
                        Left            =   60
                        TabIndex        =   28
                        Top             =   0
                        Width           =   1635
                     End
                  End
                  Begin C1SizerLibCtl.C1Tab AttribTabs 
                     Height          =   4485
                     Left            =   135
                     TabIndex        =   26
                     Top             =   105
                     Width           =   5355
                     _cx             =   9446
                     _cy             =   7911
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
                     Appearance      =   3
                     MousePointer    =   0
                     Version         =   801
                     BackColor       =   -2147483633
                     ForeColor       =   -2147483630
                     FrontTabColor   =   -2147483633
                     BackTabColor    =   -2147483633
                     TabOutlineColor =   -2147483632
                     FrontTabForeColor=   -2147483630
                     Caption         =   "no"
                     Align           =   0
                     CurrTab         =   0
                     FirstTab        =   0
                     Style           =   0
                     Position        =   1
                     AutoSwitch      =   -1  'True
                     AutoScroll      =   -1  'True
                     TabPreview      =   -1  'True
                     ShowFocusRect   =   -1  'True
                     TabsPerPage     =   3
                     BorderWidth     =   0
                     BoldCurrent     =   -1  'True
                     DogEars         =   -1  'True
                     MultiRow        =   -1  'True
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
                     Begin C1SizerLibCtl.C1Elastic elDynHolder 
                        Height          =   4110
                        Index           =   0
                        Left            =   45
                        TabIndex        =   30
                        TabStop         =   0   'False
                        Top             =   45
                        Width           =   5265
                        _cx             =   9287
                        _cy             =   7250
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
                        GridRows        =   2
                        GridCols        =   1
                        Frame           =   3
                        FrameStyle      =   0
                        FrameWidth      =   1
                        FrameColor      =   -2147483628
                        FrameShadow     =   -2147483632
                        FloodStyle      =   1
                        _GridInfo       =   $"frmSelections.frx":011B
                        AccessibleName  =   ""
                        AccessibleDescription=   ""
                        AccessibleValue =   ""
                        AccessibleRole  =   9
                        Begin TatukGIS_DK.XGIS_ControlAttributes Attributes1 
                           Height          =   3930
                           Index           =   0
                           Left            =   90
                           TabIndex        =   31
                           Top             =   90
                           Width           =   5085
                           ReadOnly        =   -1  'True
                           AllowRestructure=   -1  'True
                           ColorHeader     =   -16777201
                           ColorGrid       =   -16777211
                           Align           =   0
                           BevelInner      =   0
                           BevelOuter      =   0
                           Ctl3D           =   -1  'True
                           BorderStyle     =   0
                           Color           =   -16777201
                           Enabled         =   -1  'True
                           BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                              Name            =   "MS Sans Serif"
                              Size            =   8.25
                              Charset         =   0
                              Weight          =   400
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           ParentColor     =   0   'False
                           ParentCtl3D     =   -1  'True
                           Object.Visible         =   -1  'True
                           ParentBackground=   0   'False
                           DoubleBuffered  =   0   'False
                           AllowNull       =   0   'False
                           BevelWidth      =   1
                           BorderWidth     =   0
                           HelpContextId   =   0
                           ParentFont      =   -1  'True
                           TabOrder        =   -1
                           TabStop         =   0   'False
                        End
                     End
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   6075
            Left            =   -6360
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   330
            Width           =   5715
            _cx             =   10081
            _cy             =   10716
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
            Begin VB.CheckBox chkIncludeGeo 
               Caption         =   "Include Geo ID"
               Height          =   285
               Left            =   120
               TabIndex        =   16
               Top             =   990
               Width           =   2235
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "Print"
               Height          =   285
               Left            =   120
               TabIndex        =   15
               Top             =   2160
               Width           =   1005
            End
            Begin VB.TextBox txtTitle 
               Height          =   255
               Left            =   1080
               TabIndex        =   14
               Top             =   60
               Width           =   2655
            End
            Begin VB.CheckBox chkIncludeArea 
               Caption         =   "Include Area"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   1260
               Width           =   1725
            End
            Begin VB.CheckBox chkIncludeLength 
               Caption         =   "Include Length"
               Height          =   225
               Left            =   120
               TabIndex        =   12
               Top             =   1530
               Width           =   1605
            End
            Begin VB.CheckBox chkIncludeCentroid 
               Caption         =   "Include Centroid"
               Height          =   285
               Left            =   120
               TabIndex        =   11
               Top             =   1770
               Width           =   1485
            End
            Begin VB.TextBox txtMapTitle 
               Height          =   315
               Left            =   1020
               TabIndex        =   10
               Top             =   720
               Width           =   2625
            End
            Begin VB.CheckBox chkIncludeMap 
               Caption         =   "Include Map"
               Height          =   285
               Left            =   120
               TabIndex        =   17
               Top             =   450
               Width           =   1245
            End
            Begin VB.Label lblTitle 
               AutoSize        =   -1  'True
               Caption         =   "Report Title:"
               Height          =   195
               Left            =   60
               TabIndex        =   19
               Top             =   90
               Width           =   870
            End
            Begin VB.Label lblMapTitle 
               AutoSize        =   -1  'True
               Caption         =   "Map Title:"
               Height          =   195
               Left            =   90
               TabIndex        =   18
               Top             =   750
               Width           =   705
            End
         End
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPop"
      Visible         =   0   'False
      Begin VB.Menu mnuLineSelect 
         Caption         =   "Line Select"
      End
      Begin VB.Menu mnuAreaSelect 
         Caption         =   "Area Select"
      End
      Begin VB.Menu mnuCircleSelect 
         Caption         =   "Circle Select"
      End
      Begin VB.Menu mnuFeatureSelect 
         Caption         =   "Feature Select"
      End
   End
End
Attribute VB_Name = "frmSelections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event RemoveSelItem(uID As Long)
Public Event DoneAll()
Public Event ShowInMap(uID As Long, sLayerName As String)
Public Event FlashInMap(uID As Long, sLayerName As String)
Public Event SelectInMap(uID As Long, sLayerName As String, bSelected As Boolean)
Public Event zoomTo(uID As Long, sLayerName As String)
Public Event UpdateGeoSum(sLayerName As String)
Public Event GetSHP(uID As Long, sLayerName As String, oshp As TatukGIS_XDK9.XGIS_Shape)

Public Event SelectionTypeChanged(enmtool As OASIS_TOOLS)

Public m_LyrCol As Collection
Private m_shpProps() As ShpProps
Private m_bProgress As Boolean

Public Sub ClearAll()
        '<EhHeader>
        On Error GoTo ClearAll_Err
        '</EhHeader>
        Dim i As Integer

        m_bProgress = True

100     ReDim m_shpProps(0)
102     GEO1.Clear
104     chkEdit.Value = vbUnchecked
106     chkUseInRprt.Value = vbUnchecked
        chkSelectIn.Value = vbUnchecked
        
108     If AttribTabs.NumTabs < 1 Then Exit Sub
    
110     For i = AttribTabs.NumTabs To 1 Step -1

112         If Not i - 1 = 0 Then
114             Unload Attributes1(i - 1)
116             Unload elDynHolder(i - 1)
            Else
118             Attributes1(0).Clear
            End If
        
120         AttribTabs.RemoveTab i - 1
        Next
    
         m_bProgress = False
   
        '<EhFooter>
        Exit Sub

ClearAll_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.ClearAll " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub AddSelection(shp As Object, _
                        Optional bShowGeo As Boolean = True)
        '<EhHeader>
        On Error GoTo AddSelection_Err
        '</EhHeader>

100     If shp Is Nothing Then Exit Sub
        
102     m_shpProps(UBound(m_shpProps)).Editable = False
104     m_shpProps(UBound(m_shpProps)).sLayerName = shp.layer.Name
        m_shpProps(UBound(m_shpProps)).sLayerCaption = shp.layer.caption
106     m_shpProps(UBound(m_shpProps)).sTabCaption = shp.layer.caption & " GEO ID:" & shp.uID
108     m_shpProps(UBound(m_shpProps)).uID = shp.uID
110     m_shpProps(UBound(m_shpProps)).UseInReport = True
        m_shpProps(UBound(m_shpProps)).bSelect = True
            
112     AddTabs shp.layer.caption & " GEO ID:" & shp.uID, shp, bShowGeo
114     ReDim Preserve m_shpProps(UBound(m_shpProps) + 1)

        '<EhFooter>
        Exit Sub

AddSelection_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.AddSelection " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Init(GIS As Object)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        Dim i As Integer
        Dim sCurrItem As String
        Dim sCurFeatItm As String
        Dim keyArray() As Variant
        Dim element As Variant

100     Set m_LyrCol = New Collection

102     If chkAutomaticClear.Value = vbChecked Then ClearAll

104     ComBuffLevel.ListIndex = 2
        
106     If ComSelLayer.ListCount > 0 Then
        
108         sCurrItem = ComSelLayer.List(ComSelLayer.ListIndex)
        
        End If
        
        If ComFeatureLayer.ListCount > 0 Then
            sCurFeatItm = ComFeatureLayer.List(ComFeatureLayer.ListIndex)
        End If
        
110     txtSpatialOperation.Clear
112     keyArray = DE9IM.Keys

114     For Each element In keyArray
116         txtSpatialOperation.AddItem element
        Next

118     txtSpatialOperation.ListIndex = 5
        
120     ComSelLayer.Clear
        ComFeatureLayer.Clear
        
122     For i = 0 To GIS.Items.Count - 1

124         If GisUtils.IsInherited(GIS.Items.Item(i), "XGIS_LayerVector") Then
126             m_LyrCol.Add GIS.Items.Item(i).Name, GIS.Items.Item(i).caption
128             ComSelLayer.AddItem GIS.Items.Item(i).caption 'Name
                ComFeatureLayer.AddItem GIS.Items.Item(i).caption
            End If
        
        Next
       
130     If ComSelLayer.ListCount > 0 Then
            
132         If Len(sCurrItem) > 0 Then
134             FindIndexStrEx ComSelLayer, sCurrItem
            Else
136             ComSelLayer.ListIndex = 0 'FindIndexStrEx ComSelLayer, "--All--"
            End If
            
        End If

        If ComFeatureLayer.ListCount > 0 Then
            If Len(sCurFeatItm) > 0 Then
                FindIndexStrEx ComFeatureLayer, sCurFeatItm
            Else
                ComFeatureLayer.ListIndex = 0
            End If
        End If

        mnuLineSelect_Click

        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.Init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub AddTabs(sCaption As String, _
                    Optional oshp As TatukGIS_XDK9.XGIS_Shape, _
                    Optional bShowGeo As Boolean = True)
        '<EhHeader>
        On Error GoTo AddTabs_Err
        '</EhHeader>
        
100     Set Attributes1(Attributes1.UBound).Container = elDynHolder(elDynHolder.UBound)
102     Attributes1(Attributes1.UBound).Visible = True
    
104     If Not oshp Is Nothing Then
        
106         Attributes1(Attributes1.UBound).AllowRestructure = True
108         Attributes1(Attributes1.UBound).ShowShape oshp
110         GEO1.ShowSelected oshp.layer
    
        End If

112     If Attributes1.UBound > 0 Then
114         AttribTabs.AddTab sCaption, AttribTabs.NumTabs
116         AttribTabs.Refresh
    
118         elDynHolder(elDynHolder.UBound).Visible = True
            'm_shpProps(UBound(m_shpProps)).lCurrTab = i - 1

            'On Error Resume Next
120         Call AttribTabs.AttachPageToTab(elDynHolder(elDynHolder.UBound).hWnd, AttribTabs.NumTabs - 1)
        Else

122         If AttribTabs.NumTabs = 0 Then
124             AttribTabs.AddTab sCaption, AttribTabs.NumTabs
126             AttribTabs.Refresh
    
128             elDynHolder(elDynHolder.UBound).Visible = True
130             Call AttribTabs.AttachPageToTab(elDynHolder(elDynHolder.UBound).hWnd, AttribTabs.NumTabs - 1)
        
            Else
132             AttribTabs.TabCaption(0) = sCaption
            End If
       
        End If
    
134     Load elDynHolder(elDynHolder.UBound + 1)
136     Load Attributes1(Attributes1.UBound + 1)
    
        '<EhFooter>
        Exit Sub

AddTabs_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.AddTabs " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub RemoveTabs()
        '<EhHeader>
        On Error GoTo RemoveTabs_Err
        '</EhHeader>
        Dim i As Integer
    
100     If AttribTabs.CurrTab < 0 Then Exit Sub
102     i = AttribTabs.CurrTab
104     AttribTabs.RemoveTab AttribTabs.CurrTab
106     AttribTabs.CurrTab = i - 1
        '<EhFooter>
        Exit Sub

RemoveTabs_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.RemoveTabs " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub AttribTabs_Switch(OldTab As Integer, _
                              NewTab As Integer, _
                              Cancel As Integer)
        '<EhHeader>
        On Error GoTo AttribTabs_Switch_Err
        '</EhHeader>
100     m_bProgress = True

102     If m_shpProps(NewTab).Editable Then
104         Attributes1(NewTab).ReadOnly = False
106         chkEdit.Value = vbChecked
        Else
108         Attributes1(NewTab).ReadOnly = True
110         chkEdit.Value = vbUnchecked
        End If

112     If m_shpProps(NewTab).bSelect Then
114         chkSelectIn.Value = vbChecked
        Else
116         chkSelectIn.Value = vbUnchecked
        End If

        If m_shpProps(NewTab).UseInReport Then
            chkUseInRprt.Value = vbChecked
        Else
            chkUseInRprt.Value = vbUnchecked
        End If

118     If chkAutozoom.Value = vbChecked Then
120         RaiseEvent zoomTo(m_shpProps(NewTab).uID, m_shpProps(NewTab).sLayerName)
        End If

122     If chkAutoSelect.Value = vbChecked Then
            RaiseEvent SelectInMap(m_shpProps(NewTab).uID, m_shpProps(NewTab).sLayerName, True)
        End If

124     If chkAutoFlash.Value = vbChecked Then
126         RaiseEvent FlashInMap(m_shpProps(NewTab).uID, m_shpProps(NewTab).sLayerName)
        End If
    
128     m_bProgress = False
        '<EhFooter>
        Exit Sub

AttribTabs_Switch_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.AttribTabs_Switch " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkEdit_Click()
        '<EhHeader>
        On Error GoTo chkEdit_Click_Err
        '</EhHeader>

100     If m_bProgress Then Exit Sub
102     m_shpProps(AttribTabs.CurrTab).Editable = IIf(chkEdit.Value = vbChecked, True, False)
104     Attributes1(AttribTabs.CurrTab).ReadOnly = Not m_shpProps(AttribTabs.CurrTab).Editable
106     Attributes1(AttribTabs.CurrTab).Paint
        '<EhFooter>
        Exit Sub

chkEdit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.chkEdit_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub chkSelectIn_Click()
        '<EhHeader>
        On Error GoTo chkSelectIn_Click_Err
        '</EhHeader>

100     If m_bProgress Then Exit Sub
102     m_shpProps(AttribTabs.CurrTab).bSelect = IIf(chkSelectIn.Value = vbChecked, True, False)
    
104     If m_shpProps(AttribTabs.CurrTab).bSelect Then
106         RaiseEvent SelectInMap(m_shpProps(AttribTabs.CurrTab).uID, m_shpProps(AttribTabs.CurrTab).sLayerName, True)
        Else
108         RaiseEvent SelectInMap(m_shpProps(AttribTabs.CurrTab).uID, m_shpProps(AttribTabs.CurrTab).sLayerName, False)
        End If
    
        RaiseEvent UpdateGeoSum(m_shpProps(AttribTabs.CurrTab).sLayerName)
    
        '<EhFooter>
        Exit Sub

chkSelectIn_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.chkSelectIn_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub




Public Sub SetGeoSum(oLyr As TatukGIS_XDK9.XGIS_LayerVector)
    GEO1.ShowSelected oLyr
End Sub

Private Sub chkUseInRprt_Click()
    '<EhHeader>
    
    '</EhHeader>

    If m_bProgress Then Exit Sub

    m_shpProps(AttribTabs.CurrTab).UseInReport = IIf(chkUseInRprt.Value = vbChecked, True, False)
    Exit Sub

End Sub


Private Sub cmdClear_Click()
        '<EhHeader>
        On Error GoTo cmdClear_Click_Err
        '</EhHeader>
100     ClearAll
        '<EhFooter>
        Exit Sub

cmdClear_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.cmdClear_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdFlash_Click()
        '<EhHeader>
        On Error GoTo cmdFlash_Click_Err
        '</EhHeader>

100     If m_bProgress Then Exit Sub
    
102     RaiseEvent FlashInMap(m_shpProps(AttribTabs.CurrTab).uID, m_shpProps(AttribTabs.CurrTab).sLayerName)
    
        '<EhFooter>
        Exit Sub

cmdFlash_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.cmdFlash_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdPrint_Click()

        Dim oLyr As TatukGIS_XDK9.XGIS_LayerVector
        Dim i As Integer
        Dim oRS As New adodb.Recordset
        Dim iType As Integer
        Dim iFldCount As Integer
        Dim k As Integer
        Dim oshp As TatukGIS_XDK9.XGIS_Shape
        Dim bTableTemplateCreated As Boolean
    
108     If AttribTabs.NumTabs < 1 Then Exit Sub
    
110     For k = 0 To AttribTabs.NumTabs - 1

            If m_shpProps(k).UseInReport And (m_shpProps(k).sLayerCaption = ComSelLayer.List(ComSelLayer.ListIndex)) Then
                RaiseEvent GetSHP(m_shpProps(k).uID, m_shpProps(k).sLayerName, oshp)

                If Not oshp Is Nothing Then
                    Set oLyr = oshp.layer
         
                    If Not oLyr Is Nothing Then
                        If Not bTableTemplateCreated Then

                            For i = 0 To oLyr.Fields.Count - 1
                                m_frmDebug.DebugPrint oLyr.FieldInfo(i).Name
                            
                                Select Case oLyr.FieldInfo(i).FieldType
           
                                    Case TatukGIS_XDK9.XgisFieldTypeBoolean
                                        iType = adBoolean

                                    Case TatukGIS_XDK9.XgisFieldTypeDate
                                        iType = adDate

                                    Case TatukGIS_XDK9.XgisFieldTypeFloat
                                        iType = adDouble

                                    Case TatukGIS_XDK9.XgisFieldTypeNumber
                                        iType = adDouble

                                    Case TatukGIS_XDK9.XgisFieldTypeString
                                        iType = adVarChar
                                End Select

                                oRS.Fields.Append oLyr.Fields.Item(i).Name, iType, oLyr.FieldInfo(i).Width ', , oSHP.GetField(oLyr.Fields.Item(i).Name)
                            Next
                    
                            iFldCount = oRS.Fields.Count - 1

                            If chkIncludeArea.Value = vbChecked Then
                                oRS.Fields.Append "Geo_AREA", adDouble
                            End If
            
                            If chkIncludeGeo.Value = vbChecked Then
                                oRS.Fields.Append "GEO_ID", adDouble
                            End If
            
                            If chkIncludeLength.Value = vbChecked Then
                                oRS.Fields.Append "GEO_Length", adDouble
                            End If
            
                            If chkIncludeCentroid.Value = vbChecked Then
                                oRS.Fields.Append "GEO_CenterX", adDouble
                                oRS.Fields.Append "GEO_CenterY", adDouble
                            End If
            
                            oRS.Open , , adOpenDynamic, adLockBatchOptimistic
                    
                            bTableTemplateCreated = True
                        End If

                        oRS.AddNew
            
                        For i = 0 To oRS.Fields.Count - 1
                            m_frmDebug.DebugPrint oRS.Fields.Item(i).Type
                            m_frmDebug.DebugPrint oRS.Fields.Item(i).Name
                            oRS.Fields.Item(i).Value = oshp.GetField(oRS.Fields.Item(i).Name)
                        Next
                
                        If chkIncludeArea.Value = vbChecked Then
                            oRS.Fields.Item("Geo_AREA").Value = oshp.area
                        End If
                
                        If chkIncludeGeo.Value = vbChecked Then
                            oRS.Fields.Item("GEO_ID").Value = oshp.uID
                        End If
            
                        If chkIncludeLength.Value = vbChecked Then
                            oRS.Fields.Item("GEO_Length").Value = oshp.Length
                        End If
            
                        If chkIncludeCentroid.Value = vbChecked Then
                            oRS.Fields.Item("GEO_CenterX").Value = oshp.Centroid.x
                            oRS.Fields.Item("GEO_CenterY").Value = oshp.Centroid.y
                        End If
            
                    End If
                End If
            End If

        Next

        If oRS.RecordCount = 0 Then Exit Sub

        If chkIncludeMap.Value = vbChecked Then
            Clipboard.Clear
            oLyr.viewer.PrintClipboard
            frmReportsFromRS.SetReportRS txtTitle.Text, oRS, "", Clipboard.GetData(vbCFEMetafile), txtMapTitle.Text, ""
        Else
            Clipboard.Clear
            frmReportsFromRS.SetReportRS txtTitle.Text, oRS, ""
        End If
            
        frmReportsFromRS.ShowReport
        frmReportsFromRS.Show vbModal, Me
End Sub

Private Sub cmdRemove_Click()
        '<EhHeader>
        On Error GoTo cmdRemove_Click_Err
        '</EhHeader>
100     RemoveTabs
        '<EhFooter>
        Exit Sub

cmdRemove_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.cmdRemove_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdZoomTo_Click()
        '<EhHeader>
        On Error GoTo cmdZoomTo_Click_Err
        '</EhHeader>
100     RaiseEvent zoomTo(m_shpProps(AttribTabs.CurrTab).uID, m_shpProps(AttribTabs.CurrTab).sLayerName)
        '<EhFooter>
        Exit Sub

cmdZoomTo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.cmdZoomTo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     ReDim m_shpProps(0)
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSelections.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuAreaSelect_Click()
    RaiseEvent SelectionTypeChanged(oAreaSelect)
    ComFeatureLayer.Enabled = False
    OASSelect.Text = "Area Select"
    OASSelect.toolTipText = "Current Tool Is Area Select tool"
End Sub

Private Sub mnuCircleSelect_Click()
    RaiseEvent SelectionTypeChanged(oCircleSelect)
    ComFeatureLayer.Enabled = False
    OASSelect.Text = "Circle Select"
    OASSelect.toolTipText = "Current Tool Is Circle Select tool"
End Sub

Private Sub mnuFeatureSelect_Click()
    RaiseEvent SelectionTypeChanged(oFeatureSelect)
    ComFeatureLayer.Enabled = True
    OASSelect.Text = "Feature Select"
    OASSelect.toolTipText = "Current Tool Is Feature Select tool"
End Sub

Private Sub mnuLineSelect_Click()
    RaiseEvent SelectionTypeChanged(oLineSelect)
    ComFeatureLayer.Enabled = False
    OASSelect.Text = "Line Select"
    OASSelect.toolTipText = "Current Tool Is Line Select tool"
End Sub

Private Sub OASSelect_MouseDownOnDropdown()
    PopupMenu mnuPopUp
End Sub

