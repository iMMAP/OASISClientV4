VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{605925BE-4799-4093-A2E7-39323147E70E}#1.0#0"; "C1Query8.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmChartQueries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chart Query Wizard"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1EMain 
      Height          =   4125
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   9825
      _cx             =   17330
      _cy             =   7276
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
      Caption         =   "Main"
      Align           =   0
      AutoSizeChildren=   6
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
      Begin C1SizerLibCtl.C1Tab C1TTab1Tab2 
         Height          =   3945
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   9645
         _cx             =   17013
         _cy             =   6959
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
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "Settings|Data Sources|Chart Data Preview"
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
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   3915
            Left            =   15
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   15
            Width           =   9615
            _cx             =   16960
            _cy             =   6906
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
            Begin VB.Frame Frame3 
               Caption         =   "Data Source"
               Height          =   3585
               Left            =   30
               TabIndex        =   51
               Top             =   0
               Width           =   9495
               Begin VB.CommandButton cmdSwapTable 
                  Caption         =   "..."
                  Height          =   225
                  Left            =   8940
                  TabIndex        =   52
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   465
               End
               Begin MSDataGridLib.DataGrid DataGrid1 
                  Height          =   2985
                  Left            =   4830
                  TabIndex        =   53
                  Top             =   510
                  Width           =   4575
                  _ExtentX        =   8070
                  _ExtentY        =   5265
                  _Version        =   393216
                  HeadLines       =   1
                  RowHeight       =   15
                  BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ColumnCount     =   2
                  BeginProperty Column00 
                     DataField       =   ""
                     Caption         =   ""
                     BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                        Type            =   0
                        Format          =   ""
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   1033
                        SubFormatType   =   0
                     EndProperty
                  EndProperty
                  BeginProperty Column01 
                     DataField       =   ""
                     Caption         =   ""
                     BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                        Type            =   0
                        Format          =   ""
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   1033
                        SubFormatType   =   0
                     EndProperty
                  EndProperty
                  SplitCount      =   1
                  BeginProperty Split0 
                     BeginProperty Column00 
                     EndProperty
                     BeginProperty Column01 
                     EndProperty
                  EndProperty
               End
               Begin C1SizerLibCtl.C1Tab C1TTabDS 
                  Height          =   3315
                  Left            =   150
                  TabIndex        =   54
                  Top             =   210
                  Width           =   4575
                  _cx             =   8070
                  _cy             =   5847
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
                  Caption         =   "Conditions|Fields|SQL"
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
                  Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                     Height          =   2940
                     Index           =   0
                     Left            =   45
                     TabIndex        =   55
                     TabStop         =   0   'False
                     Top             =   330
                     Width           =   4485
                     _cx             =   7911
                     _cy             =   5186
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
                     Begin C1Query80Ctl.C1QueryFrame C1QueryFrame2 
                        Height          =   2805
                        Left            =   90
                        TabIndex        =   56
                        Top             =   90
                        Width           =   4275
                        _cx             =   7541
                        _cy             =   4948
                        DesignTemplates =   ""
                        ManualRender    =   0   'False
                        Enabled         =   -1  'True
                        DebugContextMenu=   0   'False
                        Border          =   -1  'True
                        TabInQuery      =   0   'False
                        FullFieldNames  =   0   'False
                        SchemaControl   =   "C1Query1"
                        ContentsType    =   1
                        BackColor       =   -2147483633
                        ForeColor       =   -2147483630
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DesignTimeTemplates=   -1  'True
                        TypedEditing    =   -1  'True
                        FormatDate      =   2
                        CheckBoxes      =   0   'False
                        CheckValues     =   -1  'True
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                     Height          =   2940
                     Index           =   1
                     Left            =   5220
                     TabIndex        =   57
                     TabStop         =   0   'False
                     Top             =   330
                     Width           =   4485
                     _cx             =   7911
                     _cy             =   5186
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
                     Begin C1Query80Ctl.C1QueryFrame C1QueryFrame1 
                        Height          =   2865
                        Left            =   0
                        TabIndex        =   58
                        Top             =   0
                        Width           =   4455
                        _cx             =   7858
                        _cy             =   5054
                        DesignTemplates =   ""
                        ManualRender    =   0   'False
                        Enabled         =   -1  'True
                        DebugContextMenu=   0   'False
                        Border          =   -1  'True
                        TabInQuery      =   0   'False
                        FullFieldNames  =   0   'False
                        SchemaControl   =   "C1Query1"
                        ContentsType    =   2
                        BackColor       =   -2147483633
                        ForeColor       =   -2147483630
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   204
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        DesignTimeTemplates=   -1  'True
                        TypedEditing    =   -1  'True
                        FormatDate      =   2
                        CheckBoxes      =   0   'False
                        CheckValues     =   -1  'True
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                     Height          =   2940
                     Index           =   2
                     Left            =   5520
                     TabIndex        =   59
                     TabStop         =   0   'False
                     Top             =   330
                     Width           =   4485
                     _cx             =   7911
                     _cy             =   5186
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
                     Begin VB.TextBox txtSQL 
                        Height          =   2865
                        Left            =   60
                        MultiLine       =   -1  'True
                        ScrollBars      =   2  'Vertical
                        TabIndex        =   60
                        Top             =   30
                        Width           =   4365
                     End
                  End
               End
            End
            Begin VB.Frame FraChartSettings 
               Caption         =   "Chart Settings:"
               Height          =   3300
               Left            =   30
               TabIndex        =   6
               Top             =   60
               Width           =   9495
               Begin VB.Frame FraMiscellenous 
                  Caption         =   "Miscellenous:"
                  Height          =   675
                  Left            =   90
                  TabIndex        =   40
                  Top             =   2520
                  Width           =   9315
                  Begin VB.CheckBox chkPointLabelsGen 
                     Caption         =   "Point Labels"
                     Height          =   255
                     Left            =   2430
                     TabIndex        =   46
                     Top             =   300
                     Width           =   1245
                  End
                  Begin VB.CheckBox chkMultipleColors 
                     Caption         =   "Multiple Colors"
                     Height          =   255
                     Left            =   5460
                     TabIndex        =   45
                     Top             =   300
                     Width           =   1335
                  End
                  Begin VB.CheckBox chkCrossHairs 
                     Caption         =   "Cross Hairs"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   44
                     Top             =   300
                     Width           =   1155
                  End
                  Begin VB.CheckBox chkShowTips 
                     Caption         =   "Show tips"
                     Height          =   255
                     Left            =   1260
                     TabIndex        =   43
                     Top             =   300
                     Width           =   1155
                  End
                  Begin VB.CheckBox chkCluster 
                     Caption         =   "Cluster"
                     Height          =   255
                     Left            =   4590
                     TabIndex        =   42
                     Top             =   300
                     Width           =   825
                  End
                  Begin VB.CheckBox chkBorder 
                     Caption         =   "Border"
                     Height          =   255
                     Left            =   3720
                     TabIndex        =   41
                     Top             =   300
                     Width           =   825
                  End
               End
               Begin VB.TextBox txtYAxis 
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   39
                  Top             =   600
                  Width           =   3255
               End
               Begin VB.TextBox txtXAxis 
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   38
                  Top             =   240
                  Width           =   3255
               End
               Begin VB.TextBox txtNotes 
                  Height          =   285
                  Left            =   900
                  TabIndex        =   37
                  Top             =   570
                  Width           =   4035
               End
               Begin VB.TextBox txtTitle 
                  Height          =   285
                  Left            =   570
                  TabIndex        =   36
                  Top             =   240
                  Width           =   4365
               End
               Begin VB.Frame FraDataHighlighting 
                  Caption         =   "Data Highlighting:"
                  Height          =   945
                  Left            =   5520
                  TabIndex        =   32
                  Top             =   1590
                  Width           =   2385
                  Begin VB.CheckBox chkPointLabels 
                     Caption         =   "Point Labels"
                     Height          =   255
                     Left            =   1110
                     TabIndex        =   35
                     Top             =   210
                     Width           =   1215
                  End
                  Begin VB.CheckBox chkDimmed 
                     Caption         =   "Dimmed"
                     Height          =   255
                     Left            =   90
                     TabIndex        =   34
                     Top             =   510
                     Width           =   945
                  End
                  Begin VB.CheckBox chkAllowData 
                     Caption         =   "Allow"
                     Height          =   255
                     Left            =   90
                     TabIndex        =   33
                     Top             =   210
                     Width           =   765
                  End
               End
               Begin VB.Frame FraDataEditor 
                  Caption         =   "Data Editor:"
                  Height          =   945
                  Left            =   3180
                  TabIndex        =   28
                  Top             =   1590
                  Width           =   2325
                  Begin VB.ComboBox ComAlignment 
                     Height          =   315
                     Index           =   2
                     ItemData        =   "frmChartQueries.frx":0000
                     Left            =   930
                     List            =   "frmChartQueries.frx":0010
                     Style           =   2  'Dropdown List
                     TabIndex        =   30
                     Top             =   510
                     Width           =   1305
                  End
                  Begin VB.CheckBox chkDataEditor 
                     Caption         =   "Show"
                     Height          =   255
                     Left            =   90
                     TabIndex        =   29
                     Top             =   240
                     Width           =   765
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Alignement:"
                     Height          =   195
                     Index           =   4
                     Left            =   60
                     TabIndex        =   31
                     Top             =   570
                     Width           =   825
                  End
               End
               Begin VB.Frame FraLegend 
                  Caption         =   "Legend:"
                  Height          =   945
                  Left            =   90
                  TabIndex        =   21
                  Top             =   1590
                  Width           =   3075
                  Begin VB.ComboBox ComAlignment 
                     Height          =   315
                     Index           =   1
                     ItemData        =   "frmChartQueries.frx":002E
                     Left            =   1710
                     List            =   "frmChartQueries.frx":003E
                     Style           =   2  'Dropdown List
                     TabIndex        =   25
                     Top             =   510
                     Width           =   1305
                  End
                  Begin VB.ComboBox ComAlignment 
                     Height          =   315
                     Index           =   0
                     ItemData        =   "frmChartQueries.frx":005C
                     Left            =   1710
                     List            =   "frmChartQueries.frx":006C
                     Style           =   2  'Dropdown List
                     TabIndex        =   24
                     Top             =   180
                     Width           =   1305
                  End
                  Begin VB.CheckBox chkPointLegend 
                     Caption         =   "Value"
                     Height          =   285
                     Left            =   30
                     TabIndex        =   23
                     Top             =   570
                     Width           =   735
                  End
                  Begin VB.CheckBox chkSeriesLegend 
                     Caption         =   "Series"
                     Height          =   285
                     Left            =   30
                     TabIndex        =   22
                     Top             =   240
                     Width           =   765
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Alignement:"
                     Height          =   195
                     Index           =   3
                     Left            =   810
                     TabIndex        =   27
                     Top             =   600
                     Width           =   825
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Alignement:"
                     Height          =   195
                     Index           =   2
                     Left            =   810
                     TabIndex        =   26
                     Top             =   270
                     Width           =   825
                  End
               End
               Begin VB.Frame FraChartTools 
                  Caption         =   "Chart:"
                  Height          =   645
                  Left            =   90
                  TabIndex        =   15
                  Top             =   930
                  Width           =   9315
                  Begin VB.CheckBox chkEnableChartTBR 
                     Caption         =   "Enable Tools"
                     Height          =   285
                     Left            =   2400
                     TabIndex        =   19
                     Top             =   240
                     Width           =   1275
                  End
                  Begin VB.CheckBox chk3D 
                     Caption         =   "Use 3D"
                     Height          =   255
                     Left            =   5490
                     TabIndex        =   18
                     Top             =   240
                     Width           =   945
                  End
                  Begin VB.CheckBox chkEnable 
                     Caption         =   "Enable Annotations"
                     Height          =   285
                     Left            =   3720
                     TabIndex        =   17
                     Top             =   240
                     Width           =   1695
                  End
                  Begin VB.ComboBox ComType 
                     Height          =   315
                     ItemData        =   "frmChartQueries.frx":008A
                     Left            =   480
                     List            =   "frmChartQueries.frx":00D5
                     Style           =   2  'Dropdown List
                     TabIndex        =   16
                     Top             =   180
                     Width           =   1815
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Type:"
                     Height          =   195
                     Index           =   9
                     Left            =   30
                     TabIndex        =   20
                     Top             =   270
                     Width           =   405
                  End
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   0
                  Left            =   4980
                  TabIndex        =   14
                  Top             =   210
                  Width           =   255
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   1
                  Left            =   4980
                  TabIndex        =   13
                  Top             =   540
                  Width           =   255
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   2
                  Left            =   9150
                  TabIndex        =   12
                  Top             =   240
                  Width           =   255
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   3
                  Left            =   9150
                  TabIndex        =   11
                  Top             =   570
                  Width           =   255
               End
               Begin VB.Frame FraXAxis 
                  Caption         =   "X Axis:"
                  Height          =   945
                  Left            =   7920
                  TabIndex        =   7
                  Top             =   1590
                  Width           =   1485
                  Begin VB.TextBox txtXAngle 
                     Height          =   285
                     Left            =   630
                     TabIndex        =   9
                     Text            =   "45"
                     Top             =   210
                     Width           =   795
                  End
                  Begin VB.CheckBox chkStaggered 
                     Caption         =   "Staggered"
                     Height          =   285
                     Left            =   60
                     TabIndex        =   8
                     Top             =   570
                     Width           =   1395
                  End
                  Begin VB.Label lblAngle 
                     AutoSize        =   -1  'True
                     Caption         =   "Angle:"
                     Height          =   195
                     Left            =   60
                     TabIndex        =   10
                     Top             =   210
                     Width           =   450
                  End
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Y Axis:"
                  Height          =   195
                  Index           =   8
                  Left            =   5370
                  TabIndex        =   50
                  Top             =   600
                  Width           =   480
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "X Axis:"
                  Height          =   195
                  Index           =   7
                  Left            =   5340
                  TabIndex        =   49
                  Top             =   270
                  Width           =   480
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Notes:"
                  Height          =   195
                  Index           =   6
                  Left            =   180
                  TabIndex        =   48
                  Top             =   660
                  Width           =   465
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Title:"
                  Height          =   195
                  Index           =   5
                  Left            =   150
                  TabIndex        =   47
                  Top             =   300
                  Width           =   345
               End
            End
            Begin VB.ComboBox ComChartSQL 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   3600
               Width           =   1995
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "Save"
               Height          =   315
               Left            =   2940
               TabIndex        =   4
               Top             =   3600
               Width           =   855
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "Delete"
               Height          =   315
               Left            =   2100
               TabIndex        =   3
               Top             =   3600
               Width           =   855
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3915
            Left            =   10260
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   15
            Width           =   9615
            _cx             =   16960
            _cy             =   6906
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
            AutoSizeChildren=   3
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   3915
            Left            =   10560
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   15
            Width           =   9615
            _cx             =   16960
            _cy             =   6906
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
            Begin VB.CommandButton cmd 
               Caption         =   "..."
               Height          =   375
               Left            =   8850
               TabIndex        =   68
               Top             =   990
               Width           =   405
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Load database schema from this database:"
               Height          =   405
               Left            =   30
               TabIndex        =   67
               Top             =   960
               Width           =   4335
            End
            Begin VB.ListBox List1 
               Height          =   840
               Left            =   30
               TabIndex        =   66
               Top             =   1650
               Width           =   4335
            End
            Begin VB.ListBox List2 
               Height          =   840
               Left            =   4470
               TabIndex        =   65
               Top             =   1650
               Width           =   4335
            End
            Begin VB.TextBox Text1 
               Height          =   375
               Left            =   4470
               TabIndex        =   64
               Top             =   990
               Width           =   4335
            End
            Begin C1Query80Ctl.C1Query C1Query1 
               Height          =   540
               Left            =   10
               TabIndex        =   63
               Top             =   1
               Width           =   540
               _cx             =   952
               _cy             =   952
               DesignTemplates =   ""
               MainViewName    =   ""
               DataMember      =   ""
               FilterMode      =   0   'False
               ApplyExtensions =   3
               NameSubstitute  =   ""
               SaveSchemaAsString=   0   'False
               PathSeparator   =   "."
               SchemaData      =   "frmChartQueries.frx":0183
            End
            Begin VB.Label Label1 
               Caption         =   "Create Schemas and Chart Queries."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   30
               TabIndex        =   72
               Top             =   90
               Width           =   3135
            End
            Begin VB.Label Label2 
               Caption         =   $"frmChartQueries.frx":01A3
               Height          =   735
               Left            =   270
               TabIndex        =   71
               Top             =   330
               Width           =   8055
            End
            Begin VB.Label Label3 
               Caption         =   "Main table:"
               Height          =   255
               Left            =   30
               TabIndex        =   70
               Top             =   1410
               Width           =   1575
            End
            Begin VB.Label Label4 
               Caption         =   "has natural joins with:"
               Height          =   375
               Left            =   4470
               TabIndex        =   69
               Top             =   1410
               Width           =   4575
            End
         End
      End
   End
End
Attribute VB_Name = "frmChartQueries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim doc As DOMDocument
'Dim viewsNode As MSXML2.IXMLDOMElement
'Dim relationsNode As MSXML2.IXMLDOMElement
'Dim catalogsNode As IXMLDOMElement
'Dim catalogNode As IXMLDOMElement
'Dim RSLocalUserGroups As New ADODB.Recordset
'' this array will store all the queries in the XML data file
'Private XMLValues() As String
'Private XMLValuesQRY() As String
'Private XMLValuesC As Integer
'Private XMLValuesQC As Integer
'
'Private Sub XMLValuesQRYClear()
'        '<EhHeader>
'        On Error GoTo XMLValuesQRYClear_Err
'        '</EhHeader>
'100     XMLValuesQC = 0
'102     ReDim XMLValuesQRY(0) As String
'        '<EhFooter>
'        Exit Sub
'
'XMLValuesQRYClear_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.XMLValuesQRYClear " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub XMLValuesAddQRY(XMLValue As String)
'        '<EhHeader>
'        On Error GoTo XMLValuesAddQRY_Err
'        '</EhHeader>
'100     XMLValuesQRY(XMLValuesQC) = XMLValue
'102     XMLValuesQC = XMLValuesQC + 1
'104     ReDim Preserve XMLValuesQRY(0 To XMLValuesQC) As String
'        '<EhFooter>
'        Exit Sub
'
'XMLValuesAddQRY_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.XMLValuesAddQRY " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'
'Private Sub XMLValuesClear()
'        '<EhHeader>
'        On Error GoTo XMLValuesClear_Err
'        '</EhHeader>
'100     XMLValuesC = 0
'102     ReDim XMLValues(0) As String
'        '<EhFooter>
'        Exit Sub
'
'XMLValuesClear_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.XMLValuesClear " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub XMLValuesAdd(XMLValue As String)
'        '<EhHeader>
'        On Error GoTo XMLValuesAdd_Err
'        '</EhHeader>
'100     XMLValues(XMLValuesC) = XMLValue
'102     XMLValuesC = XMLValuesC + 1
'104     ReDim Preserve XMLValues(0 To XMLValuesC) As String
'        '<EhFooter>
'        Exit Sub
'
'XMLValuesAdd_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.XMLValuesAdd " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub LoadQueries(sFIle As String)
'        '<EhHeader>
'        On Error GoTo LoadQueries_Err
'        '</EhHeader>
'        Dim xmlDoc, Node
'        Dim i As Integer
'        Dim fs As New FileSystemObject
'
'100     ComChartSQL.Clear
'
'102     If Not fs.FileExists(g_sAppPath & "\data\templates\" & sFIle) Then
'104         ComChartSQL.AddItem "---- N/A ----"
'106         ComChartSQL.ListIndex = 0
'            Exit Sub
'        End If
'
'108     ComChartSQL.AddItem "--- Clear All ---"
'
'110     XMLValuesClear
'
'112     Set xmlDoc = CreateObject("Msxml.DOMDocument")
'114     xmlDoc.async = False
'
'116     xmlDoc.Load (g_sAppPath & "\data\templates\" & sFIle)
'
'118     If xmlDoc.documentElement Is Nothing Then
'            Exit Sub
'        End If
'
'        ' go through every first-level child node
'120     For i = 0 To xmlDoc.documentElement.childNodes.Length - 1
'122         Set Node = xmlDoc.documentElement.childNodes.Item(i)
'            ' the first second-level child node is the description
'124         ComChartSQL.AddItem Node.childNodes.Item(0).Text
'            ' the first second-level child node is the value
'126         XMLValuesAdd (Node.childNodes.Item(1).Text)
'        Next
'
'
'128     XMLValuesQRYClear
'
'130     Set xmlDoc = CreateObject("Msxml.DOMDocument")
'132     xmlDoc.async = False
'
'134     xmlDoc.Load (g_sAppPath & "\data\templates\constraints.xml")
'
'136     If xmlDoc.documentElement Is Nothing Then
'            Exit Sub
'        End If
'
'        ' go through every first-level child node
'138     For i = 0 To xmlDoc.documentElement.childNodes.Length - 1
'140         Set Node = xmlDoc.documentElement.childNodes.Item(i)
'            ' the first second-level child node is the value
'142         XMLValuesAddQRY (Node.childNodes.Item(1).Text)
'        Next
'
'144     ComChartSQL.ListIndex = 0
'
'        '<EhFooter>
'        Exit Sub
'
'LoadQueries_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.LoadQueries " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub AddChildTextNode(parent As MSXML2.IXMLDOMNode, _
'                             nodeName As String, _
'                             Data As String)
'        '<EhHeader>
'        On Error GoTo AddChildTextNode_Err
'        '</EhHeader>
'        Dim node1 As IXMLDOMElement
'100     Set node1 = doc.createElement(nodeName)
'102     parent.appendChild node1
'        Dim node2 As IXMLDOMNode
'104     Set node2 = doc.createTextNode(Data)
'106     node1.appendChild node2
'        '<EhFooter>
'        Exit Sub
'
'AddChildTextNode_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.AddChildTextNode " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub C1Query1_Error(ByVal ErrorNumber As Long, _
'                           Description As String, _
'                           CancelDisplay As Boolean)
'        '<EhHeader>
'        On Error GoTo C1Query1_Error_Err
'        '</EhHeader>
'100     CancelDisplay = True
'        '<EhFooter>
'        Exit Sub
'
'C1Query1_Error_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.C1Query1_Error " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub C1QueryFrame1_Change(ByVal ItemID As Long, _
'                                 ByVal ItemElement As C1Query80Ctl.ElementTypeEnum)
'        '<EhHeader>
'        On Error GoTo C1QueryFrame1_Change_Err
'        '</EhHeader>
'
'        On Error Resume Next
'
'100     C1Query1.BuildSQL
'
'102     If Len(C1Query1.SQL) > 0 Then
'104         Adodc1.ConnectionString = m_Cnn.ConnectionString
'106         Adodc1.CommandType = adCmdText
'108         Adodc1.RecordSource = C1Query1.SQL
'110         txtSQL.Text = C1Query1.SQL
'112         Adodc1.Refresh
'114         Set DataGrid1.DataSource = Adodc1
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'C1QueryFrame1_Change_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.C1QueryFrame1_Change " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub C1QueryFrame1_Error(ByVal ErrorNumber As Long, _
'                                Description As String, _
'                                CancelDisplay As Boolean)
'        '<EhHeader>
'        On Error GoTo C1QueryFrame1_Error_Err
'        '</EhHeader>
'100     CancelDisplay = True
'        '<EhFooter>
'        Exit Sub
'
'C1QueryFrame1_Error_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.C1QueryFrame1_Error " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub C1QueryFrame2_Change(ByVal ItemID As Long, _
'                                 ByVal ItemElement As C1Query80Ctl.ElementTypeEnum)
'        '<EhHeader>
'        On Error GoTo C1QueryFrame2_Change_Err
'        '</EhHeader>
'
'        On Error Resume Next
'
'100     C1Query1.BuildSQL
'
'102     If Len(C1Query1.SQL) > 0 Then
'104         Adodc1.ConnectionString = m_Cnn.ConnectionString
'106         Adodc1.CommandType = adCmdText
'108         Adodc1.RecordSource = C1Query1.SQL
'110         txtSQL.Text = C1Query1.SQL
'112         Adodc1.Refresh
'114         Set DataGrid1.DataSource = Adodc1
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'C1QueryFrame2_Change_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.C1QueryFrame2_Change " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub C1QueryFrame2_Error(ByVal ErrorNumber As Long, _
'                                Description As String, _
'                                CancelDisplay As Boolean)
'        '<EhHeader>
'        On Error GoTo C1QueryFrame2_Error_Err
'        '</EhHeader>
'100     CancelDisplay = True
'        '<EhFooter>
'        Exit Sub
'
'C1QueryFrame2_Error_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.C1QueryFrame2_Error " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdAssign_Click()
'        '<EhHeader>
'        On Error GoTo cmdAssign_Click_Err
'        '</EhHeader>
'
'        On Error Resume Next
'
'100     Adodc1.ConnectionString = m_Cnn.ConnectionString
'102     Adodc1.CommandType = adCmdText
'
'104     If Len(txtSQL.Text) > 0 Then
'106         Adodc1.RecordSource = txtSQL.Text
'108         Adodc1.Refresh
'110         Set DataGrid1.DataSource = Adodc1
'        End If
'
'112     CreateChartSettings , , Adodc1.Recordset
'        '<EhFooter>
'        Exit Sub
'
'cmdAssign_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.cmdAssign_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub CreateChartSettings(Optional sPath As String, _
'                                Optional sName As String, _
'                                Optional oRS As ADODB.Recordset)
'        '<EhHeader>
'        On Error GoTo CreateChartSettings_Err
'        '</EhHeader>
'        Dim sConnStr As String
'        Dim i As Integer
'        Dim udtOASISChartO As OASISChartObj
'
'        'sConnStr = Adodc1.ConnectionString
'
'
'100     With udtOASISChartO
'
'102         If chkEnableChartTBR.Value = vbChecked Then
'104             .bChartTBR = True
'106             ReDim .sChartTools(2)
'108             .sChartTools(0) = 0
'110             .sChartTools(1) = 20
'            End If
'
'112         If chkEnable.Value = vbChecked Then
'114             .bAnnoTBR = True
'116             ReDim .sAnnoTools(2)
'            End If
'
'118         If chkDataEditor.Value = vbChecked Then .bDataEdtr = True
'
'120         If chkSeriesLegend.Value = vbChecked Then .bSeriesLGD = True
'
'122         If chkPointLegend.Value = vbChecked Then .bValueLGD = True
'
'124         If chkAllowData.Value = vbChecked Then .bDataHigls = True
'
'126         .bContextMenu = False
'
'128         .bMenuBar = False
'
'130         .bDataEdtrAllowEdit = False
'
'132         .bDataEdtrAllowDrag = False
'
'134         If chkMultipleColors.Value = vbChecked Then .bMultipleColors = True
'
'136         If chkPointLabelsGen.Value = vbChecked Then .bPointLabelsGen = True
'
'138         .bScrollable = False
'
'140         If chkPointLabels.Value = vbChecked Then .bHGlsPointLabel = True
'
'142         If chkDimmed.Value = vbChecked Then .bHGlsDimmed = True
'
'144         .sConnStr = m_Cnn.ConnectionString
'146         .sSQL = txtSQL.Text
'
'148         .iSerLgdAlign = GetAlignment(ComAlignment(0).ListIndex)
'150         .iValLgdAlign = GetAlignment(ComAlignment(1).ListIndex)
'152         .iDataEdtrAlign = GetAlignment(ComAlignment(2).ListIndex)
'
'154         .sXAxis = txtXAxis.Text
'156         .XAxisAngle = IIf(IsNumeric(txtXAngle.Text), txtXAngle.Text, 0)
'158         .XAxisStaggered = IIf(chkStaggered.Value = vbChecked, True, False)
'160         .sYAxis = txtYAxis.Text
'
'162         .iHeight = 5000
'164         .iWidth = 6000
'166         .iParentHeight = 6500
'168         .iParentWidth = 7060
'
'170         With .udtTitle
'172             .CT_Text = txtTitle.Text
'174             .CT_BackColor = 2147483647
'176             .CT_DrawingArea = True
'178             .CT_Alignment = CA_StringAlignment_Center
'180             .CT_DockArea = CDA_Area_Top
'182             .CT_Flags = CF_TitleFlag_DrawingArea
'184             .CT_LineAlignment = CA_StringAlignment_Far
'186             .CT_LineGap = 2
'
'188             With .CT_Font
'190                 .CF_Size = lblPrev(0).Font.Size
'192                 .CF_Bold = lblPrev(0).Font.Bold
'194                 .CF_Name = lblPrev(0).Font.Name
'196                 .CF_Italic = lblPrev(0).Font.Italic
'198                 .CF_Strikethrough = lblPrev(0).Font.Strikethrough
'200                 .CF_Underline = lblPrev(0).Font.Underline
'202                 .CF_Weight = lblPrev(0).Font.Weight ' 500
'                End With
'            End With
'
'204         With .udtNotes
'206             .CT_Text = txtNotes.Text
'208             .CT_BackColor = 2147483647
'210             .CT_DrawingArea = True
'212             .CT_Alignment = CA_StringAlignment_Far
'214             .CT_DockArea = CDA_Area_Bottom
'216             .CT_Flags = CF_TitleFlag_DrawingArea
'218             .CT_LineAlignment = CA_StringAlignment_Far
'220             .CT_LineGap = 2
'
'222             With .CT_Font
'224                 .CF_Size = lblPrev(1).Font.Size
'226                 .CF_Bold = lblPrev(1).Font.Bold
'228                 .CF_Name = lblPrev(1).Font.Name
'230                 .CF_Italic = lblPrev(1).Font.Italic
'232                 .CF_Strikethrough = lblPrev(1).Font.Strikethrough
'234                 .CF_Underline = lblPrev(1).Font.Underline
'236                 .CF_Weight = lblPrev(1).Font.Weight ' 500
'                End With
'            End With
'
'238         .enmChartType = ComType.ListIndex + 1 ' Chart_Bubble
'
'240         .enmScheme = CS_Solid
'242         .iAngleX = 30
'244         .iAngleY = 30
'246         .enmAxesStyle = CA_FlatFrame
'248         .lngBackColor = 14935011
'
'250         If chkBorder.Value = vbChecked Then .bBorder = True
'252         .lngBorderColor = 11053224
'254         .enmBorderEffect = CBE_Dark
'
'256         If chkCluster.Value = vbChecked Then .bCluster = False
'258         If chk3D.Value = vbChecked Then .bChart3D = True
'260         If chkCrossHairs.Value = vbChecked Then .bCrossHairs = False
'262         .sngCylSides = 0
'264         .enmGrid = CG_None
'266         .lngInsideColor = 16777215
'268         .enmMarkerShape = CMS_Many
'270         .iMarkerSize = 12
'272         .sngMarkerStep = 3
'274         .sngPerspective = 1
'
'276         If chkShowTips.Value = vbChecked Then .bShowTips = True
'278         .enmSmoothFlags = CSF_Fill
'280         .enmStacked = CS_No
'282         .iWallWidth = 4
'284         .bZoom = False
'
'            '        If Len(sPath) > 0 Then
'            '        'Binary Template
'            '            With .udtExports(2)
'            '                .bForceKill = True
'            '                .enmExportFormat = tplBin
'            '                .sFileName = sName
'            '                .sPath = sPath
'            '            End With
'            '        End If
'
'        End With
'
'286     If oRS Is Nothing Then
'288         frmOASISCharts.SetChart udtOASISChartO
'        Else
'290         frmOASISCharts.SetChartWithRS udtOASISChartO, oRS
'        End If
'
'292     frmOASISCharts.Show vbModeless, Me
'
'        '<EhFooter>
'        Exit Sub
'
'CreateChartSettings_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.CreateChartSettings " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Function GetAlignment(iIndex As Integer) As Integer
'        '<EhHeader>
'        On Error GoTo GetAlignment_Err
'        '</EhHeader>
'
'100     GetAlignment = 513
'
'102     Select Case iIndex
'
'            Case 0
'104             GetAlignment = 513
'
'106         Case 1
'108             GetAlignment = 515
'
'110         Case 2
'112             GetAlignment = 256
'
'114         Case 3
'116             GetAlignment = 258
'        End Select
'
'        '<EhFooter>
'        Exit Function
'
'GetAlignment_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.GetAlignment " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Function
'
'Private Sub cmdDelete_Click()
'        '<EhHeader>
'        On Error GoTo cmdDelete_Click_Err
'        '</EhHeader>
'        Dim res As VbMsgBoxResult
'
'100     res = MsgBox("This will permanently erase all stored settings." & vbCrLf & "Are you sure?", vbApplicationModal + vbYesNo + vbQuestion)
'102     If res = vbYes Then
'            ' delete the file then refresh
'            ' this will create an empty array in our memory
'            On Error Resume Next
'104         Kill (g_sAppPath & "\data\templates\queries.xml")
'106         Kill (g_sAppPath & "\data\templates\constraints.xml")
'108         LoadQueries "queries.xml"
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'cmdDelete_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.cmdDelete_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdFontTitle_Click(Index As Integer)
'        '<EhHeader>
'        On Error GoTo cmdFontTitle_Click_Err
'        '</EhHeader>
'        Dim c As New cCommonDialog
'
'100     c.Font = lblPrev(Index).Font
'102     c.ShowFont
'
'104     With lblPrev(Index)
'106         .FontBold = c.FontBold
'108         .FontItalic = c.FontItalic
'110         .FontName = c.FontName
'112         .FontSize = c.FontSize
'        End With
'
'        '<EhFooter>
'        Exit Sub
'
'cmdFontTitle_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.cmdFontTitle_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub SaveSQLSettings(sFIle As String, desc As String, xml As String)
'        '<EhHeader>
'        On Error GoTo SaveSQLSettings_Err
'        '</EhHeader>
'        Dim i As Integer
'        Dim found As Boolean
'        Dim xmlDoc, Node
'        Dim res As VbMsgBoxResult
'
'100     If desc = "" Then
'            Exit Sub
'        End If
'
'        ' create an empty parser object
'102     Set xmlDoc = CreateObject("Msxml.DOMDocument")
'104     xmlDoc.async = False
'
'        ' load the XML storage file
'        ' if the file does not exist this statement will do nothing
'106     xmlDoc.Load (g_sAppPath & "\data\templates\" & sFIle)
'
'108     If xmlDoc.documentElement Is Nothing Then
'            ' if the file was empty create a parent node and an empty child node
'110         xmlDoc.documentElement = xmlDoc.createElement("storage")
'112         Set Node = xmlDoc.documentElement.appendChild(xmlDoc.createElement("query"))
'114         Call Node.setAttribute("id", xmlDoc.documentElement.childNodes.Length)
'        Else
'            ' if the file is not empty check that there is no query with the
'            ' same description as the one we are adding
'116         found = False
'
'118         For i = 0 To xmlDoc.documentElement.childNodes.Length - 1
'120             Set Node = xmlDoc.documentElement.childNodes.Item(i)
'
'122             If Node.childNodes.Item(0).Text = desc Then
'                    ' there is a duplicate query - the user must choose a
'                    ' different description
'124                 found = True
'126                 res = MsgBox("A setting by this name already exists in the repository." & vbCrLf & "Would you like to replace it?", vbYesNoCancel + vbQuestion + vbApplicationModal)
'
'128                 If res = vbYes Then
'                        ' user wishes to overwrite the existing query
'                        ' exit this loop with "node" set to the existing query
'                        Exit For
'130                 ElseIf res = vbNo Then
'                        ' user wants to choose another name
'132                     cmdSave_Click
'                        Exit Sub
'                    Else
'                        ' user wishes to abort the saving
'                        Exit Sub
'                    End If
'                End If
'
'            Next
'
'134         If Not found Then
'                ' no existing query found, add a new first-level child node
'136             Set Node = xmlDoc.documentElement.appendChild(xmlDoc.createElement("query"))
'138             Call Node.setAttribute("id", xmlDoc.documentElement.childNodes.Length)
'            End If
'        End If
'
'        ' clear the existing query
'140     While (Node.childNodes.Length > 0)
'142         Node.removeChild (Node.childNodes.Item(0))
'        Wend
'
'        ' set "node" to the current query
'        ' first second-level child node is the description
'144     Node.appendChild xmlDoc.createElement("description")
'146     Node.lastChild.Text = desc
'        ' second second-level child node is the value
'148     Node.appendChild xmlDoc.createElement("value")
'150     Node.lastChild.Text = xml
'
'        ' update the XML storage file
'152     xmlDoc.Save (g_sAppPath & "\data\templates\" & sFIle)
'
'        '<EhFooter>
'        Exit Sub
'
'SaveSQLSettings_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.SaveSQLSettings " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdSave_Click()
'        '<EhHeader>
'        On Error GoTo cmdSave_Click_Err
'        '</EhHeader>
'        Dim xml As String
'        Dim desc As String
'
'100     xml = C1QueryFrame1.SaveToXML
'102     desc = InputBox("Enter description for this query", "OASIS Charting")
'
'104     SaveSQLSettings "queries.xml", desc, xml
'
'106     xml = C1QueryFrame2.SaveToXML
'108     SaveSQLSettings "constraints.xml", desc, xml
'
'        ' refresh our memory
'110     LoadQueries "queries.xml"
'
'        '<EhFooter>
'        Exit Sub
'
'cmdSave_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.cmdSave_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdSwapTable_Click()
'        '<EhHeader>
'        On Error GoTo cmdSwapTable_Click_Err
'        '</EhHeader>
'100     frmTableChooser.Show vbModal, Me
'        '<EhFooter>
'        Exit Sub
'
'cmdSwapTable_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.cmdSwapTable_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub ComChartSQL_Click()
'        '<EhHeader>
'        On Error GoTo ComChartSQL_Click_Err
'        '</EhHeader>
'
'100     If Not ComChartSQL.List(ComChartSQL.ListIndex) = "---- N/A ----" Then
'102         If Not ComChartSQL.List(ComChartSQL.ListIndex) = "--- Clear All ---" Then
'
'104             C1QueryFrame1.LoadFromXML (XMLValues(ComChartSQL.ListIndex - 1))
'                ' batched post-load query formatting goes here
'106             C1QueryFrame1.Render
'108             C1QueryFrame2.LoadFromXML (XMLValuesQRY(ComChartSQL.ListIndex - 1))
'110             C1QueryFrame2.Render
'
'                On Error Resume Next
'
'112             C1Query1.BuildSQL
'
'114             If Len(C1Query1.SQL) > 0 Then
'116                 Adodc1.ConnectionString = m_Cnn.ConnectionString
'118                 Adodc1.CommandType = adCmdText
'120                 Adodc1.RecordSource = C1Query1.SQL
'122                 txtSQL.Text = C1Query1.SQL
'124                 Adodc1.Refresh
'126                 Set DataGrid1.DataSource = Adodc1
'                End If
'
'            Else
'128             C1QueryFrame1.Clear
'                C1QueryFrame1.Render
'130             C1QueryFrame2.Clear
'                C1QueryFrame2.Render
'                txtSQL.Text = ""
''132             C1Query1.BuildSQL
'134             Set DataGrid1.DataSource = Nothing
'            End If
'
'        Else
'136         C1QueryFrame1.Clear
'138         C1QueryFrame2.Clear
'            C1QueryFrame1.Render
'            C1QueryFrame2.Render
'            txtSQL.Text = ""
''140         C1Query1.BuildSQL
'142         Set DataGrid1.DataSource = Nothing
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'ComChartSQL_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.ComChartSQL_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub Form_Load()
'        '<EhHeader>
'        On Error GoTo Form_Load_Err
'        '</EhHeader>
'100     ComAlignment(0).ListIndex = 0
'102     ComAlignment(1).ListIndex = 0
'104     ComAlignment(2).ListIndex = 0
'106     ComType.ListIndex = 1
'
'        'C1Query2.
'
'108     LoadSchema
'110     LoadTable "oincidents_FEA"
'
'112     XMLValuesClear
'114     LoadQueries "queries.xml"
'        '<EhFooter>
'        Exit Sub
'
'Form_Load_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.Form_Load " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub LoadSchema()
'        '<EhHeader>
'        On Error GoTo LoadSchema_Err
'        '</EhHeader>
'
'        Dim Cnxn As ADODB.Connection
'        Dim rstSchema As ADODB.Recordset
'        Dim strCnxn As String
'        Dim schemaNode As IXMLDOMElement
'        Dim rootFolderNode As IXMLDOMElement
'        Dim foldersNode As IXMLDOMElement
'        Dim viewNode As IXMLDOMElement
'        Dim folderNode As IXMLDOMElement
'        Dim TableName As String
'        Dim viewFieldNode As IXMLDOMElement
'        Dim tableInfosNode As IXMLDOMElement
'        Dim tableInfoNode As IXMLDOMElement
'        Dim viewFieldsNode As IXMLDOMElement
'        Dim folderFieldsNode As IXMLDOMElement
'        Dim rstColumns As ADODB.Recordset
'        Dim folderFieldNode As IXMLDOMElement
'        Dim leng As String
'
'100     Set Cnxn = New ADODB.Connection
'102     strCnxn = m_Cnn.ConnectionString
'104     Cnxn.Open strCnxn
'
'        ' Open database schema
'106     Set rstSchema = Cnxn.OpenSchema(adSchemaTables)
'
'108     Set doc = New DOMDocument
'
'110     Set schemaNode = doc.createElement("SCHEMA")
'
'112     doc.appendChild schemaNode
'
'114     Set rootFolderNode = doc.createElement("FOLDER")
'116     schemaNode.appendChild rootFolderNode
'
'118     Set foldersNode = doc.createElement("SUBFOLDERS")
'120     rootFolderNode.appendChild foldersNode
'
'122     Set viewsNode = doc.createElement("VIEWS")
'124     schemaNode.appendChild viewsNode
'
'126     Set relationsNode = doc.createElement("JOINRELATIONS")
'128     schemaNode.appendChild relationsNode
'
'130     Set catalogsNode = doc.createElement("CATALOGS")
'132     schemaNode.appendChild catalogsNode
'
'134     AddChildTextNode schemaNode, "IMPORTEDFROM", m_Cnn.ConnectionString
'136     AddChildTextNode schemaNode, "CP", "0"
'
'        ' creating views and folders
'
'138     frmTableChooser.List1.Clear
'
'140     rstSchema.Filter = "TABLE_NAME = 'oincidents_FEA'"
'
'142     Do Until rstSchema.EOF
'
'144         If (rstSchema!TABLE_TYPE = "TABLE") Then
'146             Set folderNode = doc.createElement("FOLDER")
'148             foldersNode.appendChild folderNode
'150             Set viewNode = doc.createElement("VIEW")
'152             TableName = rstSchema!TABLE_NAME
'154             AddChildTextNode viewNode, "VIEWNAME", TableName
'156             AddChildTextNode folderNode, "FOLDERNAME", TableName
'158             viewsNode.appendChild viewNode
'160             List1.AddItem TableName
'162             frmTableChooser.List1.AddItem TableName
'                ' table info in a view
'164             Set tableInfosNode = doc.createElement("VIEWTABLES")
'166             viewNode.appendChild tableInfosNode
'168             Set tableInfoNode = doc.createElement("VIEWTABLE")
'170             tableInfosNode.appendChild tableInfoNode
'172             AddChildTextNode tableInfoNode, "TABLENAME", TableName
'174             AddChildTextNode tableInfoNode, "REQUIRED", "1"
'                ' for each table column create a view field and a folder field
'176             Set viewFieldsNode = doc.createElement("VIEWFIELDS")
'178             viewNode.appendChild viewFieldsNode
'180             Set folderFieldsNode = doc.createElement("FOLDERFIELDS")
'182             folderNode.appendChild folderFieldsNode
'
'184             Set rstColumns = Cnxn.OpenSchema(adSchemaColumns, Array(Empty, Empty, TableName))
'
'186             Do Until rstColumns.EOF
'188                 Set viewFieldNode = doc.createElement("VIEWFIELD")
'190                 viewFieldsNode.appendChild viewFieldNode
'192                 Set folderFieldNode = doc.createElement("FOLDERFIELD")
'194                 folderFieldsNode.appendChild folderFieldNode
'196                 AddChildTextNode viewFieldNode, "FIELDNAME", rstColumns!COLUMN_NAME
'198                 AddChildTextNode viewFieldNode, "FIELDTABLENAME", TableName
'200                 AddChildTextNode viewFieldNode, "FIELDTYPE", rstColumns!DATA_TYPE
'202                 AddChildTextNode folderFieldNode, "NAME", rstColumns!COLUMN_NAME
'204                 AddChildTextNode folderFieldNode, "VIEW", TableName
'206                 AddChildTextNode folderFieldNode, "TABLENAME", TableName
'208                 leng = "0"
'
'210                 If rstColumns!CHARACTER_MAXIMUM_LENGTH <> Empty Then
'212                     leng = rstColumns!CHARACTER_MAXIMUM_LENGTH
'                    End If
'
'214                 AddChildTextNode viewFieldNode, "FIELDSIZE", leng
'216                 leng = "0"
'
'218                 If rstColumns!NUMERIC_PRECISION <> Empty Then
'220                     leng = rstColumns!NUMERIC_PRECISION
'                    End If
'
'222                 AddChildTextNode viewFieldNode, "FIELDPREC", leng
'224                 leng = "0"
'
'226                 If rstColumns!NUMERIC_SCALE <> Empty Then
'228                     leng = rstColumns!NUMERIC_SCALE
'                    End If
'
'230                 AddChildTextNode viewFieldNode, "FIELDSCALE", leng
'232                 rstColumns.MoveNext
'                Loop
'
'            End If
'
'234         rstSchema.MoveNext
'        Loop
'
'        '<EhFooter>
'        Exit Sub
'
'LoadSchema_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.LoadSchema " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub LoadSchemaEX()
'        '<EhHeader>
'        On Error GoTo LoadSchemaEX_Err
'        '</EhHeader>
'
'        Dim Cnxn As ADODB.Connection
'        Dim rstSchema As ADODB.Recordset
'        Dim strCnxn As String
'        Dim schemaNode As IXMLDOMElement
'        Dim rootFolderNode As IXMLDOMElement
'        Dim foldersNode As IXMLDOMElement
'        Dim viewNode As IXMLDOMElement
'        Dim folderNode As IXMLDOMElement
'        Dim TableName As String
'        Dim viewFieldNode As IXMLDOMElement
'        Dim tableInfosNode As IXMLDOMElement
'        Dim tableInfoNode As IXMLDOMElement
'        Dim viewFieldsNode As IXMLDOMElement
'        Dim folderFieldsNode As IXMLDOMElement
'        Dim rstColumns As ADODB.Recordset
'        Dim folderFieldNode As IXMLDOMElement
'        Dim leng As String
'
'100     Set Cnxn = New ADODB.Connection
'102     strCnxn = m_Cnn.ConnectionString
'104     Cnxn.Open strCnxn
'
'        ' Open database schema
'106     Set rstSchema = Cnxn.OpenSchema(adSchemaTables)
'
'108     Set doc = New DOMDocument
'
'110     Set schemaNode = doc.createElement("SCHEMA")
'
'112     doc.appendChild schemaNode
'
'114     Set rootFolderNode = doc.createElement("FOLDER")
'116     schemaNode.appendChild rootFolderNode
'
'118     Set foldersNode = doc.createElement("SUBFOLDERS")
'120     rootFolderNode.appendChild foldersNode
'
'122     Set viewsNode = doc.createElement("VIEWS")
'124     schemaNode.appendChild viewsNode
'
'126     Set relationsNode = doc.createElement("JOINRELATIONS")
'128     schemaNode.appendChild relationsNode
'
'130     Set catalogsNode = doc.createElement("CATALOGS")
'132     schemaNode.appendChild catalogsNode
'
'134     AddChildTextNode schemaNode, "IMPORTEDFROM", m_Cnn.ConnectionString
'136     AddChildTextNode schemaNode, "CP", "0"
'
'        ' creating views and folders
'
'138     frmTableChooser.List1.Clear
'
'140     Do Until rstSchema.EOF
'
'142         If (rstSchema!TABLE_TYPE = "TABLE") Then
'144             Set folderNode = doc.createElement("FOLDER")
'146             foldersNode.appendChild folderNode
'148             Set viewNode = doc.createElement("VIEW")
'150             TableName = rstSchema!TABLE_NAME
'152             AddChildTextNode viewNode, "VIEWNAME", TableName
'154             AddChildTextNode folderNode, "FOLDERNAME", TableName
'156             viewsNode.appendChild viewNode
'158             List1.AddItem TableName
'160             frmTableChooser.List1.AddItem TableName
'                ' table info in a view
'162             Set tableInfosNode = doc.createElement("VIEWTABLES")
'164             viewNode.appendChild tableInfosNode
'166             Set tableInfoNode = doc.createElement("VIEWTABLE")
'168             tableInfosNode.appendChild tableInfoNode
'170             AddChildTextNode tableInfoNode, "TABLENAME", TableName
'172             AddChildTextNode tableInfoNode, "REQUIRED", "1"
'                ' for each table column create a view field and a folder field
'174             Set viewFieldsNode = doc.createElement("VIEWFIELDS")
'176             viewNode.appendChild viewFieldsNode
'178             Set folderFieldsNode = doc.createElement("FOLDERFIELDS")
'180             folderNode.appendChild folderFieldsNode
'
'182             Set rstColumns = Cnxn.OpenSchema(adSchemaColumns, Array(Empty, Empty, TableName))
'
'184             Do Until rstColumns.EOF
'186                 Set viewFieldNode = doc.createElement("VIEWFIELD")
'188                 viewFieldsNode.appendChild viewFieldNode
'190                 Set folderFieldNode = doc.createElement("FOLDERFIELD")
'192                 folderFieldsNode.appendChild folderFieldNode
'194                 AddChildTextNode viewFieldNode, "FIELDNAME", rstColumns!COLUMN_NAME
'196                 AddChildTextNode viewFieldNode, "FIELDTABLENAME", TableName
'198                 AddChildTextNode viewFieldNode, "FIELDTYPE", rstColumns!DATA_TYPE
'200                 AddChildTextNode folderFieldNode, "NAME", rstColumns!COLUMN_NAME
'202                 AddChildTextNode folderFieldNode, "VIEW", TableName
'204                 AddChildTextNode folderFieldNode, "TABLENAME", TableName
'206                 leng = "0"
'
'208                 If rstColumns!CHARACTER_MAXIMUM_LENGTH <> Empty Then
'210                     leng = rstColumns!CHARACTER_MAXIMUM_LENGTH
'                    End If
'
'212                 AddChildTextNode viewFieldNode, "FIELDSIZE", leng
'214                 leng = "0"
'
'216                 If rstColumns!NUMERIC_PRECISION <> Empty Then
'218                     leng = rstColumns!NUMERIC_PRECISION
'                    End If
'
'220                 AddChildTextNode viewFieldNode, "FIELDPREC", leng
'222                 leng = "0"
'
'224                 If rstColumns!NUMERIC_SCALE <> Empty Then
'226                     leng = rstColumns!NUMERIC_SCALE
'                    End If
'
'228                 AddChildTextNode viewFieldNode, "FIELDSCALE", leng
'230                 rstColumns.MoveNext
'                Loop
'
'            End If
'
'232         rstSchema.MoveNext
'        Loop
'
'        '<EhFooter>
'        Exit Sub
'
'LoadSchemaEX_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.LoadSchemaEX " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub LoadTable(sTable As String)
'        '<EhHeader>
'        On Error GoTo LoadTable_Err
'        '</EhHeader>
'        Dim view1 As String
'        Dim viewFieldsNode As IXMLDOMElement
'        Dim tableCandidate As String
'        Dim fieldsNode As IXMLDOMNode
'        Dim added As Boolean
'        Dim fieldCandidateName As String
'        Dim relationNode As IXMLDOMElement
'        Dim viewLinksNode As IXMLDOMElement
'        Dim relateViewsNode As IXMLDOMElement
'        Dim viewLinkNode As IXMLDOMNode
'        Dim joinsNode As IXMLDOMElement
'        Dim joinNode As IXMLDOMElement
'        Dim candidate As String
'        Dim found As Boolean
'        Dim newCatNode As IXMLDOMNode
'        Dim i As Integer
'        Dim j As Integer
'        Dim K As Integer
'
'100     List2.Clear
'102     While relationsNode.childNodes.Length > 0
'104         relationsNode.removeChild relationsNode.childNodes(0)
'        Wend
'106     While catalogsNode.childNodes.Length > 0
'108         catalogsNode.removeChild catalogsNode.childNodes(0)
'        Wend
'        ' Add 1st view group and main view name (first item in the view group)
'110     Set catalogNode = doc.createElement("CATALOG")
'112     catalogsNode.appendChild catalogNode
'114     AddChildTextNode catalogNode, "VIEW", sTable
'
'        ' Add join relationships and catalogs
'116     For i = 0 To viewsNode.childNodes.Length - 1
'
'118         If viewsNode.childNodes(i).childNodes(0).nodeTypedValue = sTable Then
'120             view1 = sTable
'122             Set viewFieldsNode = viewsNode.childNodes(i).childNodes(2)
'                Exit For
'            End If
'
'        Next
'
'124     For i = 0 To viewsNode.childNodes.Length - 1
'
'126         tableCandidate = viewsNode.childNodes(i).childNodes(0).nodeTypedValue
'
'128         If tableCandidate <> sTable Then
'130             Set fieldsNode = viewsNode.childNodes(i).childNodes(2) ' 2nd node is VIEWFIELDS
'
'132             For j = 0 To fieldsNode.childNodes.Length - 1
'134                 added = False
'136                 fieldCandidateName = fieldsNode.childNodes(j).childNodes(0).nodeTypedValue
'
'138                 For K = 0 To viewFieldsNode.childNodes.Length - 1
'
'140                     If viewFieldsNode.childNodes(K).childNodes(0).nodeTypedValue = fieldCandidateName Then
'142                         List2.AddItem tableCandidate & " (using " & fieldCandidateName & " field)"
'                            ' Add relationship
'144                         Set relationNode = doc.createElement("JOINRELATION")
'146                         relationsNode.appendChild relationNode
'148                         Set relateViewsNode = doc.createElement("RELATEVIEWS")
'150                         AddChildTextNode relateViewsNode, "VIEWNAME", sTable
'152                         AddChildTextNode relateViewsNode, "VIEWNAME", tableCandidate
'154                         relationNode.appendChild relateViewsNode
'156                         Set viewLinksNode = doc.createElement("VIEWLINKS")
'158                         relationNode.appendChild viewLinksNode
'160                         Set viewLinkNode = doc.createElement("VIEWLINK")
'162                         viewLinksNode.appendChild viewLinkNode
'164                         AddChildTextNode viewLinkNode, "VIEW1", sTable
'166                         AddChildTextNode viewLinkNode, "JOINTYPE", "0"
'168                         AddChildTextNode viewLinkNode, "VIEW2", tableCandidate
'                            ' Joins
'170                         Set joinsNode = doc.createElement("JOINS")
'172                         relationNode.appendChild joinsNode
'174                         Set joinNode = doc.createElement("JOIN")
'176                         joinsNode.appendChild joinNode
'178                         AddChildTextNode joinNode, "FIELD1", fieldCandidateName
'180                         AddChildTextNode joinNode, "TABLE1", sTable
'182                         AddChildTextNode joinNode, "VIEW1", sTable
'184                         AddChildTextNode joinNode, "OPERATOR", "0"  '' is equal to
'186                         AddChildTextNode joinNode, "FIELD2", fieldCandidateName
'188                         AddChildTextNode joinNode, "TABLE2", tableCandidate
'190                         AddChildTextNode joinNode, "VIEW2", tableCandidate
'                            ' Add this table to the view group
'192                         AddChildTextNode catalogNode, "VIEW", tableCandidate
'194                         added = True
'                            Exit For
'                        End If
'
'                    Next
'
'196                 If added = True Then
'                        Exit For
'                    End If
'
'                Next
'
'            End If
'
'        Next
'
'        ' Create the remaining (2nd, etc) view groups (catalogs).
'198     For i = 0 To viewsNode.childNodes.Length - 1
'200         candidate = viewsNode.childNodes(i).childNodes(0).nodeTypedValue
'202         found = False
'
'204         For j = 0 To catalogNode.childNodes.Length - 1
'
'206             If catalogNode.childNodes(j).nodeTypedValue = candidate Then
'208                 found = True
'                    Exit For
'                End If
'
'            Next
'
'210         If Not found Then
'212             Set newCatNode = doc.createElement("CATALOG")
'214             catalogsNode.appendChild newCatNode
'216             AddChildTextNode newCatNode, "VIEW", candidate
'            End If
'
'        Next
'
'        ' Assign the schema to the control
'218     C1Query1.schema = doc.xml
'
'        ' Initialize the C1QueryFrame controls
'220     C1QueryFrame1.Clear
'222     C1QueryFrame1.CurrentItemID = 0
'224     C1QueryFrame1.Render
'226     C1QueryFrame2.Clear
'228     C1QueryFrame2.CurrentItemID = 0
'230     C1QueryFrame2.Render
'
'        '<EhFooter>
'        Exit Sub
'
'LoadTable_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmChartProviderSettings.LoadTable " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'
'
