VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChartSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "frmChartSettings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Tab C1TbChart 
      Height          =   6315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _cx             =   12726
      _cy             =   11139
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
      Caption         =   "Standard|Detailed|Combined|Word Export"
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
      Flags(3)        =   2
      Begin C1SizerLibCtl.C1Elastic elWordExport 
         Height          =   5940
         Left            =   8460
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   330
         Width           =   7125
         _cx             =   12568
         _cy             =   10478
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
         Begin VB.Frame FraChooseExport 
            Caption         =   "Choose Export Areas:"
            Height          =   4425
            Left            =   3825
            TabIndex        =   145
            Top             =   90
            Width           =   3210
            Begin MSComctlLib.ListView lvBookMarks 
               Height          =   3930
               Left            =   135
               TabIndex        =   146
               Top             =   360
               Width           =   2940
               _ExtentX        =   5186
               _ExtentY        =   6932
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Export Section"
                  Object.Width           =   4304
               EndProperty
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic elCombined 
         Height          =   5940
         Left            =   8160
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   330
         Width           =   7125
         _cx             =   12568
         _cy             =   10478
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
         Begin VB.Frame FraAnalysisFilter2 
            Caption         =   "Analysis Filter"
            Height          =   4500
            Left            =   3510
            TabIndex        =   128
            Top             =   0
            Width           =   3615
            Begin VB.CheckBox chkIgnoreZero1 
               Caption         =   "Ignore Zero values"
               Height          =   195
               Left            =   90
               TabIndex        =   143
               Top             =   4230
               Width           =   2445
            End
            Begin VB.Frame FraQuery2 
               Caption         =   "Query"
               Height          =   1230
               Left            =   45
               TabIndex        =   140
               Top             =   2970
               Width           =   3525
               Begin VB.TextBox txtQRy2 
                  Height          =   600
                  Left            =   45
                  TabIndex        =   142
                  Top             =   540
                  Width           =   3435
               End
               Begin VB.CheckBox chkUseQuery2 
                  Caption         =   "Use Query Expression"
                  Height          =   240
                  Left            =   45
                  TabIndex        =   141
                  Top             =   225
                  Width           =   1995
               End
            End
            Begin VB.Frame FraDateIntervalls3 
               Caption         =   "Date Intervalls:"
               Enabled         =   0   'False
               Height          =   2400
               Left            =   135
               TabIndex        =   130
               Top             =   495
               Width           =   3255
               Begin VB.Frame FraIntervall2 
                  Caption         =   "Intervall 2:"
                  Height          =   1230
                  Left            =   90
                  TabIndex        =   131
                  Top             =   1035
                  Width           =   3075
                  Begin XpressEditorsLibCtl.dxDateEdit dxDateEditInt2 
                     Height          =   315
                     Left            =   990
                     OleObjectBlob   =   "frmChartSettings.frx":6852
                     TabIndex        =   132
                     Top             =   675
                     Width           =   1995
                  End
                  Begin XpressEditorsLibCtl.dxDateEdit dxDateEditInt22 
                     Height          =   315
                     Left            =   990
                     OleObjectBlob   =   "frmChartSettings.frx":68F2
                     TabIndex        =   133
                     Top             =   315
                     Width           =   1995
                  End
                  Begin VB.Label lblToDate75 
                     AutoSize        =   -1  'True
                     Caption         =   "To Date:"
                     Height          =   195
                     Left            =   135
                     TabIndex        =   135
                     Top             =   720
                     Width           =   630
                  End
                  Begin VB.Label lblFromDate967 
                     AutoSize        =   -1  'True
                     Caption         =   "From Date:"
                     Height          =   195
                     Left            =   90
                     TabIndex        =   134
                     Top             =   360
                     Width           =   780
                  End
               End
               Begin XpressEditorsLibCtl.dxDateEdit dxDateEdit2 
                  Height          =   315
                  Left            =   1035
                  OleObjectBlob   =   "frmChartSettings.frx":6992
                  TabIndex        =   136
                  Top             =   630
                  Width           =   1995
               End
               Begin XpressEditorsLibCtl.dxDateEdit dxDateEdit3 
                  Height          =   315
                  Left            =   1035
                  OleObjectBlob   =   "frmChartSettings.frx":6A32
                  TabIndex        =   137
                  Top             =   270
                  Width           =   1995
               End
               Begin VB.Label lblToDate2 
                  AutoSize        =   -1  'True
                  Caption         =   "To Date:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   139
                  Top             =   675
                  Width           =   630
               End
               Begin VB.Label lblFromDate2 
                  AutoSize        =   -1  'True
                  Caption         =   "From Date:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   138
                  Top             =   315
                  Width           =   780
               End
            End
            Begin VB.CheckBox chkUseDate2 
               Caption         =   "Use Date Intervalls"
               Height          =   195
               Left            =   180
               TabIndex        =   129
               Top             =   270
               Width           =   1725
            End
         End
         Begin VB.Frame FraSettings7 
            Caption         =   "Settings"
            Height          =   2895
            Left            =   0
            TabIndex        =   106
            Top             =   0
            Width           =   3480
            Begin VB.Frame FraChartSettings7 
               Caption         =   "Chart Settings:"
               Height          =   2520
               Left            =   90
               TabIndex        =   107
               Top             =   270
               Width           =   3345
               Begin VB.TextBox txtMWidth 
                  Height          =   285
                  Left            =   1710
                  TabIndex        =   117
                  Text            =   "400"
                  Top             =   1440
                  Width           =   690
               End
               Begin VB.TextBox txtMultLgdTop 
                  Height          =   285
                  Left            =   2520
                  TabIndex        =   116
                  Text            =   "100"
                  Top             =   1440
                  Width           =   690
               End
               Begin VB.TextBox txtMultLGDLeft 
                  Height          =   285
                  Left            =   1755
                  TabIndex        =   115
                  Text            =   "100"
                  Top             =   2070
                  Width           =   690
               End
               Begin VB.TextBox txtMultCHTop 
                  Height          =   285
                  Left            =   2565
                  TabIndex        =   114
                  Text            =   "40"
                  Top             =   2070
                  Width           =   690
               End
               Begin VB.TextBox txtMultYCaption 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   113
                  Text            =   "# Of Incidents"
                  Top             =   945
                  Width           =   3120
               End
               Begin VB.TextBox txtMultCaption 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   112
                  Text            =   "Incidents Distribution Over Day By Province"
                  Top             =   450
                  Width           =   3165
               End
               Begin VB.TextBox txtMultBGWidth 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   111
                  Text            =   "500"
                  Top             =   1440
                  Width           =   645
               End
               Begin VB.TextBox txtMultBGHeight 
                  Height          =   285
                  Left            =   855
                  TabIndex        =   110
                  Text            =   "320"
                  Top             =   1440
                  Width           =   735
               End
               Begin VB.TextBox txtMultCHWidth 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   109
                  Text            =   "280"
                  Top             =   2070
                  Width           =   690
               End
               Begin VB.TextBox txtMultCHHeight 
                  Height          =   285
                  Left            =   900
                  TabIndex        =   108
                  Text            =   "240"
                  Top             =   2070
                  Width           =   690
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  Caption         =   "Lgd Left"
                  Height          =   195
                  Left            =   1710
                  TabIndex        =   127
                  Top             =   1215
                  Width           =   585
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Lgd Top"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   126
                  Top             =   1215
                  Width           =   600
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Chart Left"
                  Height          =   195
                  Left            =   1755
                  TabIndex        =   125
                  Top             =   1845
                  Width           =   690
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Chart Top"
                  Height          =   195
                  Left            =   2565
                  TabIndex        =   124
                  Top             =   1845
                  Width           =   705
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Chart Caption:"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   123
                  Top             =   225
                  Width           =   1005
               End
               Begin VB.Label Label14 
                  AutoSize        =   -1  'True
                  Caption         =   "BGWidth"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   122
                  Top             =   1215
                  Width           =   645
               End
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  Caption         =   "BGHeight"
                  Height          =   195
                  Left            =   855
                  TabIndex        =   121
                  Top             =   1215
                  Width           =   690
               End
               Begin VB.Label Label16 
                  AutoSize        =   -1  'True
                  Caption         =   "Chart Wdt"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   120
                  Top             =   1845
                  Width           =   720
               End
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
                  Caption         =   "Chart Hgt"
                  Height          =   195
                  Left            =   900
                  TabIndex        =   119
                  Top             =   1845
                  Width           =   675
               End
               Begin VB.Label Label18 
                  Caption         =   "Y Axis Caption"
                  Height          =   330
                  Left            =   90
                  TabIndex        =   118
                  Top             =   720
                  Width           =   1185
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic elChartDetailed 
         Height          =   5940
         Left            =   7860
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   330
         Width           =   7125
         _cx             =   12568
         _cy             =   10478
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
         Begin VB.CommandButton cmdApplyDetailed 
            Caption         =   "Apply"
            Height          =   375
            Left            =   5940
            TabIndex        =   104
            Top             =   5535
            Width           =   1140
         End
         Begin VB.Frame FraStyleSettings 
            Caption         =   "Style Settings:"
            Height          =   1410
            Index           =   1
            Left            =   0
            TabIndex        =   97
            Top             =   4095
            Width           =   3525
            Begin VB.TextBox txtXValRotation 
               Height          =   285
               Index           =   1
               Left            =   1485
               MaxLength       =   3
               TabIndex        =   100
               Text            =   "0"
               Top             =   180
               Width           =   420
            End
            Begin VB.TextBox txtLargeValueStep 
               Height          =   285
               Index           =   1
               Left            =   1485
               MaxLength       =   3
               TabIndex        =   99
               Text            =   "1"
               Top             =   540
               Width           =   420
            End
            Begin VB.TextBox txtSmallValueStep 
               Height          =   285
               Index           =   1
               Left            =   1485
               MaxLength       =   3
               TabIndex        =   98
               Text            =   "0"
               Top             =   945
               Width           =   420
            End
            Begin VB.Label lblXValRotation 
               AutoSize        =   -1  'True
               Caption         =   "X Val Rotation:"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   103
               Top             =   225
               Width           =   1065
            End
            Begin VB.Label lblLargeValueStep 
               AutoSize        =   -1  'True
               Caption         =   "Large Value Step:"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   102
               Top             =   585
               Width           =   1275
            End
            Begin VB.Label lblSmallValueStep 
               AutoSize        =   -1  'True
               Caption         =   "Small Value Step:"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   101
               Top             =   990
               Width           =   1245
            End
         End
         Begin VB.Frame FraAvailableFields1 
            Caption         =   "Analysis Filter"
            Height          =   3195
            Left            =   3645
            TabIndex        =   86
            Top             =   0
            Width           =   3435
            Begin VB.CheckBox chkUseDate1 
               Caption         =   "Use Date Intervalls"
               Height          =   195
               Left            =   180
               TabIndex        =   96
               Top             =   270
               Width           =   1725
            End
            Begin VB.Frame FraDateIntervalls1 
               Caption         =   "Date Intervalls:"
               Enabled         =   0   'False
               Height          =   1140
               Left            =   135
               TabIndex        =   91
               Top             =   495
               Width           =   3255
               Begin XpressEditorsLibCtl.dxDateEdit dxDtTo1 
                  Height          =   315
                  Left            =   1035
                  OleObjectBlob   =   "frmChartSettings.frx":6AD2
                  TabIndex        =   92
                  Top             =   630
                  Width           =   1995
               End
               Begin XpressEditorsLibCtl.dxDateEdit dxDtFROM1 
                  Height          =   315
                  Left            =   1035
                  OleObjectBlob   =   "frmChartSettings.frx":6B72
                  TabIndex        =   93
                  Top             =   270
                  Width           =   1995
               End
               Begin VB.Label lblFromDate1 
                  AutoSize        =   -1  'True
                  Caption         =   "From Date:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   95
                  Top             =   315
                  Width           =   780
               End
               Begin VB.Label lblToDate1 
                  AutoSize        =   -1  'True
                  Caption         =   "To Date:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   94
                  Top             =   675
                  Width           =   630
               End
            End
            Begin VB.Frame FraQuery1 
               Caption         =   "Query"
               Height          =   1230
               Left            =   135
               TabIndex        =   88
               Top             =   1620
               Width           =   3255
               Begin VB.CheckBox chkUseQuery1 
                  Caption         =   "Use Query Expression"
                  Height          =   240
                  Left            =   45
                  TabIndex        =   90
                  Top             =   225
                  Width           =   1995
               End
               Begin VB.TextBox txtQry1 
                  Height          =   600
                  Left            =   45
                  TabIndex        =   89
                  Top             =   540
                  Width           =   3165
               End
            End
            Begin VB.CheckBox chkIgnoreZero 
               Caption         =   "Ignore Zero values"
               Height          =   195
               Left            =   180
               TabIndex        =   87
               Top             =   2880
               Width           =   2445
            End
         End
         Begin VB.Frame FraSettings1 
            Caption         =   "Settings"
            Height          =   4065
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   3525
            Begin VB.Frame FraChartSettings13 
               Caption         =   "Chart Settings:"
               Height          =   2520
               Left            =   90
               TabIndex        =   65
               Top             =   270
               Width           =   3345
               Begin VB.TextBox txtDeCHHgt 
                  Height          =   285
                  Left            =   900
                  TabIndex        =   75
                  Text            =   "240"
                  Top             =   2070
                  Width           =   690
               End
               Begin VB.TextBox txtDeCHWidth 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   74
                  Text            =   "280"
                  Top             =   2070
                  Width           =   690
               End
               Begin VB.TextBox txtDeBGHeight 
                  Height          =   285
                  Left            =   855
                  TabIndex        =   73
                  Text            =   "320"
                  Top             =   1440
                  Width           =   735
               End
               Begin VB.TextBox txtDeBGWidth 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   72
                  Text            =   "500"
                  Top             =   1440
                  Width           =   645
               End
               Begin VB.TextBox txtHeadCaption 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   71
                  Text            =   "Incidents Distribution Over Day By Province"
                  Top             =   450
                  Width           =   3165
               End
               Begin VB.TextBox Text1 
                  Height          =   285
                  Left            =   90
                  TabIndex        =   70
                  Text            =   "# Of Incidents"
                  Top             =   945
                  Width           =   3120
               End
               Begin VB.TextBox txtDeCHTop 
                  Height          =   285
                  Left            =   2565
                  TabIndex        =   69
                  Text            =   "40"
                  Top             =   2070
                  Width           =   690
               End
               Begin VB.TextBox txtDeCHLeft 
                  Height          =   285
                  Left            =   1755
                  TabIndex        =   68
                  Text            =   "100"
                  Top             =   2070
                  Width           =   690
               End
               Begin VB.TextBox txtDeLGDTop 
                  Height          =   285
                  Left            =   2520
                  TabIndex        =   67
                  Text            =   "100"
                  Top             =   1440
                  Width           =   690
               End
               Begin VB.TextBox txtDeLGDLeft 
                  Height          =   285
                  Left            =   1710
                  TabIndex        =   66
                  Text            =   "400"
                  Top             =   1440
                  Width           =   690
               End
               Begin VB.Label lblYAxisCap 
                  Caption         =   "Y Axis Caption"
                  Height          =   330
                  Left            =   90
                  TabIndex        =   85
                  Top             =   720
                  Width           =   1185
               End
               Begin VB.Label lblChartHeight1 
                  AutoSize        =   -1  'True
                  Caption         =   "Chart Hgt"
                  Height          =   195
                  Left            =   900
                  TabIndex        =   84
                  Top             =   1845
                  Width           =   675
               End
               Begin VB.Label lblChartWidth1 
                  AutoSize        =   -1  'True
                  Caption         =   "Chart Wdt"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   83
                  Top             =   1845
                  Width           =   720
               End
               Begin VB.Label lblBGHeight1 
                  AutoSize        =   -1  'True
                  Caption         =   "BGHeight"
                  Height          =   195
                  Left            =   855
                  TabIndex        =   82
                  Top             =   1215
                  Width           =   690
               End
               Begin VB.Label lblBGWidth1 
                  AutoSize        =   -1  'True
                  Caption         =   "BGWidth"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   81
                  Top             =   1215
                  Width           =   645
               End
               Begin VB.Label lblHeaderCaption 
                  AutoSize        =   -1  'True
                  Caption         =   "Chart Caption:"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   80
                  Top             =   225
                  Width           =   1005
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Chart Top"
                  Height          =   195
                  Left            =   2565
                  TabIndex        =   79
                  Top             =   1845
                  Width           =   705
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Chart Left"
                  Height          =   195
                  Left            =   1755
                  TabIndex        =   78
                  Top             =   1845
                  Width           =   690
               End
               Begin VB.Label Label7 
                  AutoSize        =   -1  'True
                  Caption         =   "Lgd Top"
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   77
                  Top             =   1215
                  Width           =   600
               End
               Begin VB.Label Label8 
                  AutoSize        =   -1  'True
                  Caption         =   "Lgd Left"
                  Height          =   195
                  Left            =   1710
                  TabIndex        =   76
                  Top             =   1215
                  Width           =   585
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic elCHMain 
         Height          =   5940
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   7125
         _cx             =   12568
         _cy             =   10478
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
         Begin VB.CommandButton cmdBrowseWord 
            Caption         =   "..."
            Height          =   240
            Left            =   5040
            TabIndex        =   148
            Top             =   5535
            Width           =   420
         End
         Begin VB.CheckBox chkExportTo 
            Caption         =   "Export to word"
            Height          =   330
            Left            =   3600
            TabIndex        =   147
            Top             =   5490
            Width           =   1500
         End
         Begin VB.Frame FraStyleSettings 
            Caption         =   "Style Settings:"
            Height          =   1410
            Index           =   0
            Left            =   0
            TabIndex        =   56
            Top             =   4455
            Width           =   3525
            Begin VB.TextBox txtSmallValueStep 
               Height          =   285
               Index           =   0
               Left            =   1485
               MaxLength       =   3
               TabIndex        =   62
               Text            =   "0"
               Top             =   945
               Width           =   420
            End
            Begin VB.TextBox txtLargeValueStep 
               Height          =   285
               Index           =   0
               Left            =   1485
               MaxLength       =   3
               TabIndex        =   61
               Text            =   "1"
               Top             =   540
               Width           =   420
            End
            Begin VB.TextBox txtXValRotation 
               Height          =   285
               Index           =   0
               Left            =   1485
               MaxLength       =   3
               TabIndex        =   58
               Text            =   "0"
               Top             =   180
               Width           =   420
            End
            Begin VB.Label lblSmallValueStep 
               AutoSize        =   -1  'True
               Caption         =   "Small Value Step:"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   60
               Top             =   990
               Width           =   1245
            End
            Begin VB.Label lblLargeValueStep 
               AutoSize        =   -1  'True
               Caption         =   "Large Value Step:"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   59
               Top             =   585
               Width           =   1275
            End
            Begin VB.Label lblXValRotation 
               AutoSize        =   -1  'True
               Caption         =   "X Val Rotation:"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   57
               Top             =   225
               Width           =   1065
            End
         End
         Begin VB.Frame FraTrendSettings 
            Caption         =   "Trend Settings:"
            Height          =   645
            Left            =   0
            TabIndex        =   50
            Top             =   3780
            Width           =   3525
            Begin VB.OptionButton OptFreqSetting 
               Caption         =   "Days"
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   53
               Top             =   270
               Value           =   -1  'True
               Width           =   915
            End
            Begin VB.OptionButton OptFreqSetting 
               Caption         =   "Weeks"
               Height          =   195
               Index           =   1
               Left            =   1125
               TabIndex        =   52
               Top             =   270
               Width           =   915
            End
            Begin VB.CheckBox chkTrendLine 
               Caption         =   "Use 95% Trend Prediction"
               Height          =   330
               Left            =   2025
               TabIndex        =   51
               Top             =   180
               Width           =   1410
            End
         End
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
            Height          =   375
            Left            =   5940
            TabIndex        =   49
            Top             =   5535
            Width           =   1140
         End
         Begin VB.Frame FraGeneral 
            Caption         =   "General settings:"
            Height          =   1395
            Left            =   0
            TabIndex        =   40
            Top             =   0
            Width           =   3525
            Begin VB.Frame FraChartType 
               Caption         =   "Chart Type"
               Height          =   1140
               Left            =   1620
               TabIndex        =   45
               Top             =   180
               Width           =   1320
               Begin VB.OptionButton OptChartType 
                  Caption         =   "Linear"
                  Height          =   240
                  Index           =   2
                  Left            =   90
                  TabIndex        =   48
                  Top             =   675
                  Width           =   825
               End
               Begin VB.OptionButton OptChartType 
                  Caption         =   "Pie"
                  Height          =   240
                  Index           =   1
                  Left            =   90
                  TabIndex        =   47
                  Top             =   450
                  Width           =   690
               End
               Begin VB.OptionButton OptChartType 
                  Caption         =   "Bar"
                  Height          =   240
                  Index           =   0
                  Left            =   90
                  TabIndex        =   46
                  Top             =   225
                  Value           =   -1  'True
                  Width           =   735
               End
            End
            Begin VB.Frame FraLevel 
               Caption         =   "Analysis Level"
               Height          =   1140
               Left            =   90
               TabIndex        =   41
               Top             =   180
               Width           =   1410
               Begin VB.OptionButton OptAdminLevel 
                  Caption         =   "Town"
                  Height          =   240
                  Index           =   2
                  Left            =   135
                  TabIndex        =   44
                  Top             =   810
                  Width           =   915
               End
               Begin VB.OptionButton OptAdminLevel 
                  Caption         =   "District"
                  Height          =   240
                  Index           =   1
                  Left            =   135
                  TabIndex        =   43
                  Top             =   540
                  Width           =   960
               End
               Begin VB.OptionButton OptAdminLevel 
                  Caption         =   "Province"
                  Height          =   240
                  Index           =   0
                  Left            =   135
                  TabIndex        =   42
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   1005
               End
            End
         End
         Begin VB.Frame FraChartSettings 
            Caption         =   "Chart Settings:"
            Height          =   1800
            Left            =   0
            TabIndex        =   27
            Top             =   1395
            Width           =   3525
            Begin VB.TextBox txtYCaption 
               Height          =   285
               Left            =   90
               TabIndex        =   33
               Text            =   "# Of Incidents"
               Top             =   945
               Width           =   3120
            End
            Begin VB.TextBox txtCaption 
               Height          =   285
               Left            =   90
               TabIndex        =   32
               Text            =   "Incidents By Province"
               Top             =   450
               Width           =   3165
            End
            Begin VB.TextBox txtWidth 
               Height          =   285
               Index           =   0
               Left            =   90
               TabIndex        =   31
               Text            =   "600"
               Top             =   1440
               Width           =   645
            End
            Begin VB.TextBox txtHeight 
               Height          =   285
               Index           =   0
               Left            =   855
               TabIndex        =   30
               Text            =   "360"
               Top             =   1440
               Width           =   735
            End
            Begin VB.TextBox txtWidth 
               Height          =   285
               Index           =   1
               Left            =   1665
               TabIndex        =   29
               Text            =   "500"
               Top             =   1440
               Width           =   690
            End
            Begin VB.TextBox txtHeight 
               Height          =   285
               Index           =   1
               Left            =   2475
               TabIndex        =   28
               Text            =   "280"
               Top             =   1440
               Width           =   690
            End
            Begin VB.Label lblCaption 
               AutoSize        =   -1  'True
               Caption         =   "X Axis Caption:"
               Height          =   195
               Left            =   90
               TabIndex        =   39
               Top             =   225
               Width           =   1065
            End
            Begin VB.Label lblBGWidth 
               AutoSize        =   -1  'True
               Caption         =   "Width"
               Height          =   195
               Left            =   90
               TabIndex        =   38
               Top             =   1215
               Width           =   420
            End
            Begin VB.Label lblBGHeight 
               AutoSize        =   -1  'True
               Caption         =   "Height"
               Height          =   195
               Left            =   855
               TabIndex        =   37
               Top             =   1215
               Width           =   465
            End
            Begin VB.Label lblChartWidth 
               AutoSize        =   -1  'True
               Caption         =   "Chart Width"
               Height          =   195
               Left            =   1620
               TabIndex        =   36
               Top             =   1215
               Width           =   840
            End
            Begin VB.Label lblChartHeight 
               AutoSize        =   -1  'True
               Caption         =   "Chart Height"
               Height          =   195
               Left            =   2520
               TabIndex        =   35
               Top             =   1215
               Width           =   885
            End
            Begin VB.Label lblYAxis 
               Caption         =   "Y Axis Caption"
               Height          =   330
               Left            =   90
               TabIndex        =   34
               Top             =   720
               Width           =   1185
            End
         End
         Begin VB.Frame FraAvailableFields 
            Caption         =   "Analysis Filter"
            Height          =   4095
            Left            =   3555
            TabIndex        =   12
            Top             =   0
            Width           =   3525
            Begin VB.Frame FraTimeFrames 
               Caption         =   "Time Frames:"
               Enabled         =   0   'False
               Height          =   1230
               Left            =   135
               TabIndex        =   22
               Top             =   2745
               Width           =   3210
               Begin VB.CheckBox chkTimeframe 
                  Caption         =   "Night (22-05)"
                  Height          =   285
                  Index           =   3
                  Left            =   180
                  TabIndex        =   26
                  Top             =   900
                  Width           =   1725
               End
               Begin VB.CheckBox chkTimeframe 
                  Caption         =   "Evening (18-22)"
                  Height          =   285
                  Index           =   2
                  Left            =   180
                  TabIndex        =   25
                  Top             =   675
                  Width           =   1725
               End
               Begin VB.CheckBox chkTimeframe 
                  Caption         =   "Afternoon (12-18)"
                  Height          =   285
                  Index           =   1
                  Left            =   180
                  TabIndex        =   24
                  Top             =   450
                  Width           =   1725
               End
               Begin VB.CheckBox chkTimeframe 
                  Caption         =   "Morning (05-12)"
                  Height          =   285
                  Index           =   0
                  Left            =   180
                  TabIndex        =   23
                  Top             =   225
                  Width           =   1725
               End
            End
            Begin VB.CheckBox chkUseTime 
               Caption         =   "Use Time Frame:"
               Height          =   375
               Left            =   90
               TabIndex        =   21
               Top             =   2340
               Width           =   2805
            End
            Begin VB.Frame FraDateIntervalls 
               Caption         =   "Date Intervalls:"
               Enabled         =   0   'False
               Height          =   1140
               Left            =   135
               TabIndex        =   16
               Top             =   1215
               Width           =   3255
               Begin XpressEditorsLibCtl.dxDateEdit dxDtTo 
                  Height          =   315
                  Left            =   1035
                  OleObjectBlob   =   "frmChartSettings.frx":6C12
                  TabIndex        =   17
                  Top             =   630
                  Width           =   1995
               End
               Begin XpressEditorsLibCtl.dxDateEdit dxDtFROM 
                  Height          =   315
                  Left            =   1035
                  OleObjectBlob   =   "frmChartSettings.frx":6CB2
                  TabIndex        =   18
                  Top             =   270
                  Width           =   1995
               End
               Begin VB.Label lblToDate 
                  AutoSize        =   -1  'True
                  Caption         =   "To Date:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   20
                  Top             =   675
                  Width           =   630
               End
               Begin VB.Label lblFromDate 
                  AutoSize        =   -1  'True
                  Caption         =   "From Date:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   19
                  Top             =   315
                  Width           =   780
               End
            End
            Begin VB.CheckBox chkUseDate 
               Caption         =   "Use Date Intervalls"
               Height          =   195
               Left            =   180
               TabIndex        =   15
               Top             =   945
               Width           =   1725
            End
            Begin VB.Frame FraNumberOf 
               Caption         =   "Number Of Areas To Analyse:"
               Height          =   645
               Left            =   135
               TabIndex        =   13
               Top             =   225
               Width           =   3300
               Begin VB.ComboBox ComNumOfAreasToAnalyse 
                  Height          =   315
                  ItemData        =   "frmChartSettings.frx":6D52
                  Left            =   135
                  List            =   "frmChartSettings.frx":6D95
                  Style           =   2  'Dropdown List
                  TabIndex        =   14
                  Top             =   225
                  Width           =   2850
               End
            End
         End
         Begin VB.Frame FraFrequencyField 
            Caption         =   "Analysis Field"
            Height          =   600
            Left            =   0
            TabIndex        =   10
            Top             =   3195
            Width           =   3525
            Begin VB.ComboBox ComFlds 
               Height          =   315
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   225
               Width           =   3300
            End
         End
         Begin VB.PictureBox NavigatePad 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   1365
            Left            =   0
            ScaleHeight     =   1335
            ScaleWidth      =   1515
            TabIndex        =   8
            Top             =   4455
            Width           =   1545
            Begin VB.Label NavigateWindow 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   795
               Left            =   150
               TabIndex        =   9
               Top             =   210
               Width           =   1125
            End
         End
         Begin VB.Frame FraQuery 
            Caption         =   "Query"
            Height          =   1230
            Left            =   3555
            TabIndex        =   2
            Top             =   4140
            Width           =   3525
            Begin VB.ComboBox ComSQLPre 
               Height          =   315
               ItemData        =   "frmChartSettings.frx":6DE7
               Left            =   1845
               List            =   "frmChartSettings.frx":6E00
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   225
               Visible         =   0   'False
               Width           =   1635
            End
            Begin VB.ComboBox ComQueryFields 
               Height          =   315
               Left            =   45
               TabIndex        =   6
               Text            =   "QueryFields"
               Top             =   225
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.ComboBox ComQryValue 
               Height          =   315
               Left            =   45
               TabIndex        =   5
               Text            =   "QryValue"
               Top             =   540
               Visible         =   0   'False
               Width           =   3435
            End
            Begin VB.TextBox txtQry 
               Height          =   600
               Left            =   45
               TabIndex        =   4
               Top             =   540
               Width           =   3435
            End
            Begin VB.CheckBox chkUseQuery 
               Caption         =   "Use Query Expression"
               Height          =   240
               Left            =   45
               TabIndex        =   3
               Top             =   225
               Width           =   1995
            End
         End
         Begin MSComctlLib.Slider ZoomBar 
            Height          =   615
            Left            =   1755
            TabIndex        =   54
            Top             =   5085
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   1085
            _Version        =   393216
            LargeChange     =   10
            Min             =   1
            Max             =   100
            SelStart        =   1
            TickStyle       =   2
            TickFrequency   =   10
            Value           =   1
         End
         Begin VB.Label Label2 
            Caption         =   "Chart Zoom Level"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1755
            TabIndex        =   55
            Top             =   4500
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "frmChartSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Apply(sSQL As String)
Public Event DoDetailedChart()
Public Event DoTrendAnalysis(sSQL() As String, labels As Variant)
Public Event ZoomInMode()
Public Event ZoomOutMode()
Public Event ArrowMode()
Public Event ZoomBar()
Public Event NavWindowMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event NavWindowMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event NavWindowMouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public c As New cCommonDialog


Public Sub SetIncidentFields(sArFields As Variant)
        '<EhHeader>
        On Error GoTo SetIncidentFields_Err
        '</EhHeader>
    Dim i As Integer
100     ComFlds.Clear
    
102     For i = LBound(sArFields) + 1 To UBound(sArFields)
104         ComFlds.AddItem sArFields(i)
        Next

106     ComFlds.ListIndex = 0
    
        '<EhFooter>
        Exit Sub

SetIncidentFields_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartSettings.SetIncidentFields " & _
               "at line " & Erl
        '</EhFooter>
End Sub

Private Sub chkUseDate_Click()
        '<EhHeader>
        On Error GoTo chkUseDate_Click_Err
        '</EhHeader>
100     FraDateIntervalls.Enabled = IIf(chkUseDate.Value = vbChecked, True, False)
        '<EhFooter>
        Exit Sub

chkUseDate_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartSettings.chkUseDate_Click " & _
               "at line " & Erl
        '</EhFooter>
End Sub

Private Sub chkUseDate1_Click()
        FraDateIntervalls1.Enabled = IIf(chkUseDate1.Value = vbChecked, True, False)
End Sub

Private Sub chkUseTime_Click()
        '<EhHeader>
        On Error GoTo chkUseTime_Click_Err
        '</EhHeader>
100     FraTimeFrames.Enabled = IIf(chkUseTime.Value = vbChecked, True, False)
        '<EhFooter>
        Exit Sub

chkUseTime_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartSettings.chkUseTime_Click " & _
               "at line " & Erl
        '</EhFooter>
End Sub

Private Sub cmdApply_Click()
        '<EhHeader>
        On Error GoTo cmdApply_Click_Err
        '</EhHeader>
        Dim lDays As Long
        Dim lWeeks As Long
        Dim i As Long
        Dim labels As Variant
        Dim RSadmin As ADODB.Recordset
        Dim sadminUnits As String
        Dim sRestrictedWhere As String
        Dim sCurAdmUnit As String
        Dim dateFrom As Date
        Dim dateTo As Date

        Dim sAddQry As String
        Dim sWhereSQl As String
        Dim sWhereSQls() As String
        Dim sTimeWhere As String
        
100     If Not OptChartType(2).Value Then
102         If chkUseDate.Value = vbChecked Then
                '"INSERT INTO oincidentstrans SELECT * From oincidents_FEA WHERE ((Incident_DATE Between '07/17/2007' And '07/19/2007'));
104             sWhereSQl = " WHERE ((Incident_DATE BETWEEN #" & dxDtFROM.EditValue - 1 & "# AND #" & dxDtTo.EditValue + 1 & "#))"
                '        DebugPrint "SELECT * FROM oincidents_FEA WHERE Incident_DATE BETWEEN '" & dxDateFrom.EditValue & "' AND '" & dxDateTo.EditValue & "'"
                'sWhereSQl = Replace(sWhereSQl, "-", "/")
            End If
    
106         If OptAdminLevel(0).Value Then
108             sCurAdmUnit = "Province"
110         ElseIf OptAdminLevel(1).Value Then
112             sCurAdmUnit = "District"
            Else
114             sCurAdmUnit = "Town"
            End If
        
            '"--All--"
116         If Not ComNumOfAreasToAnalyse.ListIndex = 0 Then
118             Set RSadmin = New ADODB.Recordset
120             RSadmin.Open "SELECT oincidents_FEA." & sCurAdmUnit & ", Count(oincidents_FEA.UID) AS Incidents FROM oincidents_FEA GROUP BY oincidents_FEA." & sCurAdmUnit & " ORDER BY Count(oincidents_FEA.UID) DESC;", m_Cnn
            
122             If Not RSadmin.EOF And Not RSadmin.Bof Then
124                 sRestrictedWhere = ""
                
126                 Do While Not RSadmin.EOF
128                     sRestrictedWhere = sRestrictedWhere & " " & sCurAdmUnit & " = '" & RSadmin.Fields(sCurAdmUnit).Value & "'"
130                     RSadmin.MoveNext
132                     i = i + 1

134                     If i = CInt(ComNumOfAreasToAnalyse.List(ComNumOfAreasToAnalyse.ListIndex)) Or i = RSadmin.RecordCount Then Exit Do
136                     sRestrictedWhere = sRestrictedWhere & " OR "
                    Loop

                End If

            End If
    
138         If chkUseTime.Value = vbChecked Then
    
140             If Len(sWhereSQl) > 0 Then
142                 sTimeWhere = " AND ("
                Else
144                 sTimeWhere = " WHERE ("
                End If
    
146             If chkTimeframe(0).Value = vbChecked Then
148                 sTimeWhere = sTimeWhere & "TIME00 = 'Morning (05-12)'"
                End If
        
150             If chkTimeframe(1).Value = vbChecked Then
152                 If Len(sTimeWhere) > 10 Then sTimeWhere = sTimeWhere & " OR "
154                 sTimeWhere = sTimeWhere & "TIME00 = 'Afternoon (12-18)'"
                End If
        
156             If chkTimeframe(2).Value = vbChecked Then
158                 If Len(sTimeWhere) > 10 Then sTimeWhere = sTimeWhere & " OR "
160                 sTimeWhere = sTimeWhere & "TIME00 = 'Evening (18-22)'"
                End If
        
162             If chkTimeframe(3).Value = vbChecked Then
164                 If Len(sTimeWhere) > 10 Then sTimeWhere = sTimeWhere & " OR "
166                 sTimeWhere = sTimeWhere & "TIME00 = 'Night (22-05)'"
                End If
    
168             sTimeWhere = sTimeWhere & ")"
        
170             If Len(sTimeWhere) < 10 Then sTimeWhere = ""
            End If
    
172         If Len(sWhereSQl) < 1 And Len(sTimeWhere) < 1 Then
174             If Not Len(sRestrictedWhere) < 1 Then sRestrictedWhere = " WHERE " & sRestrictedWhere
            Else
176             If Not Len(sRestrictedWhere) < 1 Then sRestrictedWhere = " AND (" & sRestrictedWhere & ")"
            End If
        
178         If chkUseQuery.Value = vbChecked Then
180                     If Len(sWhereSQl) < 1 And Len(sTimeWhere) < 1 And Len(sRestrictedWhere) < 1 Then
182                         sAddQry = " WHERE " & txtQry.Text
                        Else
184                         sAddQry = " AND " & "(" & txtQry.Text & ")"
                        End If

186             sWhereSQl = "INSERT INTO oincidentstrans SELECT * From oincidents_FEA" & sWhereSQl & " " & sTimeWhere & sRestrictedWhere & sAddQry
            
            
188             If sWhereSQl = "" Then
190             ReDim Preserve sWhereSQls(UBound(sWhereSQls) + 1)
192             sWhereSQls(UBound(sWhereSQls)) = sAddQry
                End If
            Else
194             sWhereSQl = "INSERT INTO oincidentstrans SELECT * From oincidents_FEA" & sWhereSQl & " " & sTimeWhere & sRestrictedWhere
            End If
        
196         DebugPrint sWhereSQl
        
198         RaiseEvent Apply(sWhereSQl)

        Else
        
200         If OptAdminLevel(0).Value Then
202             sCurAdmUnit = "Province"
204         ElseIf OptAdminLevel(1).Value Then
206             sCurAdmUnit = "District"
            Else
208             sCurAdmUnit = "Town"
            End If
        
210         If chkUseDate.Value = vbChecked Then
212             dateFrom = dxDtFROM.EditValue
214             dateTo = dxDtTo.EditValue
            Else
216             Set RSadmin = New ADODB.Recordset
218             RSadmin.Open "SELECT MAX(Incident_DATE) FROM oincidents_FEA", m_Cnn
        
220             If Not RSadmin.EOF And Not RSadmin.Bof Then
222                 dateTo = RSadmin.Fields.Item(0).Value
224                 Set RSadmin = New ADODB.Recordset
226                 RSadmin.Open "SELECT MIN(Incident_DATE) FROM oincidents_FEA", m_Cnn
228                 dateFrom = RSadmin.Fields.Item(0).Value
230                 dxDtFROM.EditValue = dateFrom
232                 dxDtTo.EditValue = dateTo
                
                End If
            End If

            '"--All--"
234         If Not ComNumOfAreasToAnalyse.ListIndex = 0 Then
236             Set RSadmin = New ADODB.Recordset
238             RSadmin.Open "SELECT oincidents_FEA." & sCurAdmUnit & ", Count(oincidents_FEA.UID) AS Incidents FROM oincidents_FEA GROUP BY oincidents_FEA." & sCurAdmUnit & " ORDER BY Count(oincidents_FEA.UID) DESC;", m_Cnn
            
240             If Not RSadmin.EOF And Not RSadmin.Bof Then
242                 sRestrictedWhere = " WHERE "
                
244                 Do While Not RSadmin.EOF
246                     sRestrictedWhere = sRestrictedWhere & " " & sCurAdmUnit & " = '" & RSadmin.Fields(sCurAdmUnit).Value & "'"
248                     RSadmin.MoveNext
250                     i = i + 1

252                     If i = CInt(ComNumOfAreasToAnalyse.List(ComNumOfAreasToAnalyse.ListIndex)) Or i = RSadmin.RecordCount Then Exit Do
254                     sRestrictedWhere = sRestrictedWhere & " OR "
                    Loop

                End If

256             sRestrictedWhere = "Select DISTINCT " & sCurAdmUnit & " From oincidents_FEA" & sRestrictedWhere
            Else
258             sRestrictedWhere = "Select Distinct " & sCurAdmUnit & " From oincidents_FEA"
            End If
    
260         Set RSadmin = New ADODB.Recordset
262         RSadmin.Open sRestrictedWhere, m_Cnn
264         ReDim sWhereSQls(0)
        
266         Do While Not RSadmin.EOF
            
268             sWhereSQl = " WHERE ((Incident_DATE BETWEEN #" & dateFrom - 1 & "# AND #" & dateTo + 1 & "#)) AND (" & sCurAdmUnit & " = '" & Replace(RSadmin.Fields(sCurAdmUnit).Value, "'", "''") & "')"
            
270             sadminUnits = sadminUnits & RSadmin.Fields(sCurAdmUnit).Value & ","
    
272             If chkUseTime.Value = vbChecked Then
    
274                 If Len(sWhereSQl) > 0 Then
276                     sTimeWhere = " AND ("
                    Else
278                     sTimeWhere = " WHERE ("
                    End If
    
280                 If chkTimeframe(0).Value = vbChecked Then
282                     sTimeWhere = sTimeWhere & "TIME00 = 'Morning (05-12)'"
                    End If
        
284                 If chkTimeframe(1).Value = vbChecked Then
286                     If Len(sTimeWhere) > 10 Then sTimeWhere = sTimeWhere & " OR "
288                     sTimeWhere = sTimeWhere & "TIME00 = 'Afternoon (12-18)'"
                    End If
        
290                 If chkTimeframe(2).Value = vbChecked Then
292                     If Len(sTimeWhere) > 10 Then sTimeWhere = sTimeWhere & " OR "
294                     sTimeWhere = sTimeWhere & "TIME00 = 'Evening (18-22)'"
                    End If
        
296                 If chkTimeframe(3).Value = vbChecked Then
298                     If Len(sTimeWhere) > 10 Then sTimeWhere = sTimeWhere & " OR "
300                     sTimeWhere = sTimeWhere & "TIME00 = 'Night (22-05)'"
                    End If
    
302                 sTimeWhere = sTimeWhere & ")"
        
304                 If Len(sTimeWhere) < 10 Then sTimeWhere = ""
                End If
            
306             If chkUseQuery.Value = vbChecked Then
308                 If Len(sWhereSQl) < 1 And Len(sTimeWhere) < 1 Then
310                     sAddQry = " WHERE (" & txtQry.Text & ") "
                    Else
312                     sAddQry = " AND (" & txtQry.Text & ") "
                    End If
                End If
    
314             sWhereSQl = "INSERT INTO oincidentstrans SELECT * From oincidents_FEA" & sWhereSQl & " " & sTimeWhere & sAddQry
316             DebugPrint sWhereSQl
318             sWhereSQls(UBound(sWhereSQls)) = sWhereSQl
320             RSadmin.MoveNext
322             If Not RSadmin.EOF Then
324             ReDim Preserve sWhereSQls(UBound(sWhereSQls) + 1)
                End If
            Loop
    
326         If chkUseQuery.BackColor = vbGreen Then
328             sAddQry = " AND (" & txtQry.Text & ") "
330             ReDim Preserve sWhereSQls(UBound(sWhereSQls) + 1)
332             sWhereSQls(UBound(sWhereSQls)) = sAddQry
            End If
        
334         labels = Split(sadminUnits, ",")
    
336         RaiseEvent DoTrendAnalysis(sWhereSQls, labels)

        End If

        '<EhFooter>
        Exit Sub

cmdApply_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartSettings.cmdApply_Click " & _
               "at line " & Erl
       ' Resume Next
        '</EhFooter>
End Sub

Private Sub cmdApplyDetailed_Click()
    RaiseEvent DoDetailedChart
End Sub

Private Sub cmdBrowseWord_Click()
    c.Filter = "Microsoft Word Documents (*.doc)|*.doc|Microsoft Word Documents (*.rtf)|*.rtf"
    c.ShowOpen
End Sub

Private Sub ComFlds_Click()
        '<EhHeader>
        On Error GoTo ComFlds_Click_Err
        '</EhHeader>
100     txtCaption.Text = ComFlds.List(ComFlds.ListIndex) & " By " & IIf(OptAdminLevel(0).Value, "Province", IIf(OptAdminLevel(1).Value, "District", "Town"))
102     txtYCaption.Text = "# Of " & ComFlds.List(ComFlds.ListIndex)
        '<EhFooter>
        Exit Sub

ComFlds_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartSettings.ComFlds_Click " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
100     ComNumOfAreasToAnalyse.Clear
    
102     ComNumOfAreasToAnalyse.AddItem "--All--"
104     ComNumOfAreasToAnalyse.AddItem "1"
106     ComNumOfAreasToAnalyse.AddItem "2"
108     ComNumOfAreasToAnalyse.AddItem "3"
110     ComNumOfAreasToAnalyse.AddItem "4"
112     ComNumOfAreasToAnalyse.AddItem "5"
114     ComNumOfAreasToAnalyse.AddItem "6"
116     ComNumOfAreasToAnalyse.AddItem "7"
118     ComNumOfAreasToAnalyse.AddItem "8"
120     ComNumOfAreasToAnalyse.AddItem "9"
122     ComNumOfAreasToAnalyse.AddItem "10"
124     ComNumOfAreasToAnalyse.AddItem "11"
126     ComNumOfAreasToAnalyse.AddItem "12"
128     ComNumOfAreasToAnalyse.AddItem "13"
130     ComNumOfAreasToAnalyse.AddItem "14"
132     ComNumOfAreasToAnalyse.AddItem "15"
134     ComNumOfAreasToAnalyse.AddItem "16"
136     ComNumOfAreasToAnalyse.AddItem "17"
138     ComNumOfAreasToAnalyse.AddItem "18"
140     ComNumOfAreasToAnalyse.AddItem "19"
142     ComNumOfAreasToAnalyse.AddItem "20"
    
144     ComNumOfAreasToAnalyse.ListIndex = 0
146     ComSQLPre.ListIndex = 0
    
        If Not g_sLanguage = "" Then
            If Not m_Cnn.State = adStateClosed Then
                LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmChartSettings.Form_Load " & "at line " & Erl
        'Resume Next
        '</EhFooter>
End Sub

Private Sub PointerPB_Click()
    RaiseEvent ArrowMode
End Sub

Private Sub ZoomInPB_Click()
    RaiseEvent ZoomInMode
End Sub

Private Sub ZoomOutPB_Click()
    RaiseEvent ZoomOutMode
End Sub

Private Sub NavigateWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent NavWindowMouseDown(Button, Shift, x, y)
End Sub

Private Sub NavigateWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent NavWindowMouseMove(Button, Shift, x, y)
End Sub

Private Sub NavigateWindow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent NavWindowMouseUp(Button, Shift, x, y)
End Sub

Private Sub ZoomBar_Scroll()
    RaiseEvent ZoomBar
End Sub
