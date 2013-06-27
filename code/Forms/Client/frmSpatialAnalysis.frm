VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Begin VB.Form frmSpatialAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Spatial Analyser"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10425
   Icon            =   "frmSpatialAnalysis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExportToHtml 
      Caption         =   "Export to HTML"
      Height          =   255
      Left            =   1950
      TabIndex        =   150
      Top             =   8250
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdExportToXLS 
      Caption         =   "Export to XLS"
      Height          =   255
      Left            =   60
      TabIndex        =   149
      Top             =   8250
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdExportToXML 
      Caption         =   "Export to XML"
      Height          =   255
      Left            =   3840
      TabIndex        =   148
      Top             =   8250
      Visible         =   0   'False
      Width           =   1815
   End
   Begin C1SizerLibCtl.C1Elastic C1EMain 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10545
      _cx             =   18600
      _cy             =   14420
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
         Height          =   7995
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   10365
         _cx             =   18283
         _cy             =   14102
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
         Caption         =   "Templates|Spatial Template Designer|Chart Template Designer|Please wait....."
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
         Begin VB.PictureBox Picture3 
            Height          =   7620
            Left            =   11010
            ScaleHeight     =   7560
            ScaleWidth      =   10215
            TabIndex        =   151
            Top             =   330
            Width           =   10275
            Begin VB.Label lblPleaseWait 
               BackStyle       =   0  'Transparent
               Caption         =   "Please wait......."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   855
               Left            =   4680
               TabIndex        =   152
               Top             =   6360
               Width           =   5415
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   7620
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   10275
            _cx             =   18124
            _cy             =   13441
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
            Begin VB.CommandButton cmdSave 
               Caption         =   "Show Chart"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   7470
               TabIndex        =   131
               Top             =   7080
               Width           =   2625
            End
            Begin VB.Frame FraChartSettings 
               Caption         =   "Chart Settings:"
               Height          =   6315
               Left            =   90
               TabIndex        =   8
               Top             =   1380
               Width           =   10095
               Begin VB.CommandButton cmdSaveTemplate 
                  BackColor       =   &H80000000&
                  Caption         =   "Save Template"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   465
                  Left            =   4680
                  TabIndex        =   157
                  Top             =   5700
                  Width           =   2625
               End
               Begin VB.Frame FraXAxis 
                  Caption         =   "X Axis:"
                  Height          =   675
                  Left            =   5430
                  TabIndex        =   153
                  Top             =   2940
                  Width           =   4575
                  Begin VB.CheckBox chkStaggered 
                     Caption         =   "Staggered"
                     Height          =   285
                     Left            =   1860
                     TabIndex        =   155
                     Top             =   240
                     Width           =   1185
                  End
                  Begin VB.TextBox txtXAngle 
                     Height          =   285
                     Left            =   780
                     TabIndex        =   154
                     Text            =   "45"
                     Top             =   240
                     Width           =   795
                  End
                  Begin VB.Label lblAngle 
                     AutoSize        =   -1  'True
                     Caption         =   "Angle:"
                     Height          =   195
                     Left            =   180
                     TabIndex        =   156
                     Top             =   300
                     Width           =   450
                  End
               End
               Begin VB.Frame FraDChart 
                  Caption         =   "3D Chart"
                  Height          =   1035
                  Left            =   9030
                  TabIndex        =   107
                  Top             =   900
                  Width           =   945
                  Begin VB.CheckBox chk3D 
                     Caption         =   "3D"
                     Height          =   285
                     Left            =   120
                     TabIndex        =   108
                     Top             =   240
                     Width           =   585
                  End
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   3
                  Left            =   9690
                  TabIndex        =   96
                  Top             =   570
                  Width           =   255
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   2
                  Left            =   9690
                  TabIndex        =   95
                  Top             =   210
                  Width           =   255
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   1
                  Left            =   4980
                  TabIndex        =   94
                  Top             =   540
                  Width           =   255
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   0
                  Left            =   4980
                  TabIndex        =   93
                  Top             =   210
                  Width           =   255
               End
               Begin VB.Frame FraChartTools 
                  Caption         =   "Chart Tools:"
                  Height          =   1035
                  Left            =   90
                  TabIndex        =   65
                  Top             =   900
                  Width           =   7785
                  Begin VB.Frame Frame1 
                     BorderStyle     =   0  'None
                     Enabled         =   0   'False
                     Height          =   525
                     Left            =   840
                     TabIndex        =   84
                     Top             =   180
                     Width           =   6855
                     Begin VB.PictureBox Picture2 
                        Appearance      =   0  'Flat
                        AutoSize        =   -1  'True
                        BackColor       =   &H80000005&
                        ForeColor       =   &H80000008&
                        Height          =   495
                        Left            =   60
                        Picture         =   "frmSpatialAnalysis.frx":6852
                        ScaleHeight     =   465
                        ScaleWidth      =   6705
                        TabIndex        =   91
                        Top             =   30
                        Width           =   6735
                     End
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   0
                     Left            =   930
                     TabIndex        =   83
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   1
                     Left            =   1290
                     TabIndex        =   82
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   2
                     Left            =   1650
                     TabIndex        =   81
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   3
                     Left            =   2130
                     TabIndex        =   80
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   4
                     Left            =   2580
                     TabIndex        =   79
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   5
                     Left            =   3030
                     TabIndex        =   78
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   6
                     Left            =   3540
                     TabIndex        =   77
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   7
                     Left            =   3870
                     TabIndex        =   76
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   8
                     Left            =   4230
                     TabIndex        =   75
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   9
                     Left            =   4620
                     TabIndex        =   74
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   10
                     Left            =   4950
                     TabIndex        =   73
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   11
                     Left            =   5310
                     TabIndex        =   72
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   12
                     Left            =   5670
                     TabIndex        =   71
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   13
                     Left            =   6000
                     TabIndex        =   70
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   14
                     Left            =   6360
                     TabIndex        =   69
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   15
                     Left            =   6690
                     TabIndex        =   68
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkMainTool 
                     Height          =   255
                     Index           =   16
                     Left            =   7140
                     TabIndex        =   67
                     Top             =   690
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkEnableChartTBR 
                     Caption         =   "Enable"
                     Height          =   285
                     Left            =   60
                     TabIndex        =   66
                     Top             =   270
                     Width           =   795
                  End
               End
               Begin VB.Frame FraLegend 
                  Caption         =   "Legend:"
                  Height          =   945
                  Left            =   90
                  TabIndex        =   56
                  Top             =   1980
                  Width           =   5295
                  Begin VB.Frame Frame2 
                     BorderStyle     =   0  'None
                     Caption         =   "Frame2"
                     Height          =   285
                     Left            =   1650
                     TabIndex        =   102
                     Top             =   540
                     Width           =   3465
                     Begin VB.OptionButton OptValPlacement 
                        Caption         =   "Top"
                        Height          =   255
                        Index           =   2
                        Left            =   1740
                        TabIndex        =   105
                        Top             =   0
                        Width           =   615
                     End
                     Begin VB.OptionButton OptValPlacement 
                        Caption         =   "Bottom"
                        Height          =   255
                        Index           =   3
                        Left            =   2610
                        TabIndex        =   106
                        Top             =   0
                        Width           =   855
                     End
                     Begin VB.OptionButton OptValPlacement 
                        Caption         =   "Right"
                        Height          =   255
                        Index           =   1
                        Left            =   870
                        TabIndex        =   104
                        Top             =   0
                        Width           =   705
                     End
                     Begin VB.OptionButton OptValPlacement 
                        Caption         =   "Left"
                        Height          =   255
                        Index           =   0
                        Left            =   120
                        TabIndex        =   103
                        Top             =   0
                        Value           =   -1  'True
                        Width           =   615
                     End
                  End
                  Begin VB.CheckBox chkSeriesLegend 
                     Caption         =   "Series"
                     Height          =   285
                     Left            =   180
                     TabIndex        =   62
                     Top             =   240
                     Width           =   735
                  End
                  Begin VB.CheckBox chkPointLegend 
                     Caption         =   "Value"
                     Height          =   285
                     Left            =   180
                     TabIndex        =   61
                     Top             =   510
                     Width           =   735
                  End
                  Begin VB.OptionButton OptLGDPlacement 
                     Caption         =   "Left"
                     Height          =   255
                     Index           =   0
                     Left            =   1770
                     TabIndex        =   60
                     Top             =   270
                     Value           =   -1  'True
                     Width           =   615
                  End
                  Begin VB.OptionButton OptLGDPlacement 
                     Caption         =   "Right"
                     Height          =   255
                     Index           =   1
                     Left            =   2520
                     TabIndex        =   59
                     Top             =   270
                     Width           =   705
                  End
                  Begin VB.OptionButton OptLGDPlacement 
                     Caption         =   "Top"
                     Height          =   255
                     Index           =   2
                     Left            =   3390
                     TabIndex        =   58
                     Top             =   270
                     Width           =   615
                  End
                  Begin VB.OptionButton OptLGDPlacement 
                     Caption         =   "Bottom"
                     Height          =   255
                     Index           =   3
                     Left            =   4260
                     TabIndex        =   57
                     Top             =   270
                     Width           =   795
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Alignment:"
                     Height          =   195
                     Index           =   2
                     Left            =   930
                     TabIndex        =   64
                     Top             =   270
                     Width           =   825
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Alignment:"
                     Height          =   195
                     Index           =   3
                     Left            =   930
                     TabIndex        =   63
                     Top             =   540
                     Width           =   825
                  End
               End
               Begin VB.Frame FraDataEditor 
                  Caption         =   "Data Editor:"
                  Height          =   945
                  Left            =   5430
                  TabIndex        =   47
                  Top             =   1980
                  Width           =   4575
                  Begin VB.CheckBox chkDataEditor 
                     Caption         =   "Show"
                     Height          =   255
                     Left            =   180
                     TabIndex        =   54
                     Top             =   240
                     Width           =   1125
                  End
                  Begin VB.OptionButton OptDEPlacement 
                     Caption         =   "Left"
                     Height          =   255
                     Index           =   0
                     Left            =   1080
                     TabIndex        =   53
                     Top             =   570
                     Width           =   615
                  End
                  Begin VB.OptionButton OptDEPlacement 
                     Caption         =   "Right"
                     Height          =   255
                     Index           =   1
                     Left            =   1800
                     TabIndex        =   52
                     Top             =   570
                     Width           =   705
                  End
                  Begin VB.OptionButton OptDEPlacement 
                     Caption         =   "Top"
                     Height          =   255
                     Index           =   2
                     Left            =   2640
                     TabIndex        =   51
                     Top             =   570
                     Width           =   615
                  End
                  Begin VB.OptionButton OptDEPlacement 
                     Caption         =   "Bottom"
                     Height          =   255
                     Index           =   3
                     Left            =   3420
                     TabIndex        =   50
                     Top             =   570
                     Value           =   -1  'True
                     Width           =   795
                  End
                  Begin VB.CheckBox chkAllowEdit 
                     Caption         =   "Allow Edit"
                     Height          =   255
                     Left            =   1590
                     TabIndex        =   49
                     Top             =   240
                     Width           =   1425
                  End
                  Begin VB.CheckBox chkAllowDrag 
                     Caption         =   "Allow Drag"
                     Height          =   285
                     Left            =   3210
                     TabIndex        =   48
                     Top             =   210
                     Width           =   1245
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Alignment:"
                     Height          =   195
                     Index           =   4
                     Left            =   180
                     TabIndex        =   55
                     Top             =   600
                     Width           =   735
                  End
               End
               Begin VB.Frame FraDataHighlighting 
                  Caption         =   "Data Highlighting:"
                  Height          =   675
                  Left            =   90
                  TabIndex        =   43
                  Top             =   2940
                  Width           =   5295
                  Begin VB.CheckBox chkAllowData 
                     Caption         =   "Allow"
                     Height          =   255
                     Left            =   180
                     TabIndex        =   46
                     Top             =   270
                     Width           =   765
                  End
                  Begin VB.CheckBox chkDimmed 
                     Caption         =   "Dimmed"
                     Height          =   255
                     Left            =   1320
                     TabIndex        =   45
                     Top             =   270
                     Width           =   945
                  End
                  Begin VB.CheckBox chkPointLabels 
                     Caption         =   "Point Labels"
                     Height          =   255
                     Left            =   2700
                     TabIndex        =   44
                     Top             =   270
                     Width           =   1215
                  End
               End
               Begin VB.Frame FraAnnotations 
                  Caption         =   "Annotations:"
                  Height          =   1035
                  Left            =   90
                  TabIndex        =   19
                  Top             =   3630
                  Width           =   9915
                  Begin VB.CheckBox chkEnable 
                     Caption         =   "Enable"
                     Height          =   285
                     Left            =   60
                     TabIndex        =   42
                     Top             =   240
                     Width           =   795
                  End
                  Begin VB.PictureBox Picture1 
                     BorderStyle     =   0  'None
                     Height          =   465
                     Left            =   870
                     Picture         =   "frmSpatialAnalysis.frx":10B54
                     ScaleHeight     =   465
                     ScaleWidth      =   8055
                     TabIndex        =   41
                     Top             =   210
                     Width           =   8055
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   0
                     Left            =   990
                     TabIndex        =   40
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   1
                     Left            =   1320
                     TabIndex        =   39
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   2
                     Left            =   1680
                     TabIndex        =   38
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   3
                     Left            =   2010
                     TabIndex        =   37
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   4
                     Left            =   2340
                     TabIndex        =   36
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   5
                     Left            =   2700
                     TabIndex        =   35
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   6
                     Left            =   3030
                     TabIndex        =   34
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   7
                     Left            =   3360
                     TabIndex        =   33
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   8
                     Left            =   3720
                     TabIndex        =   32
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   9
                     Left            =   4170
                     TabIndex        =   31
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   10
                     Left            =   4650
                     TabIndex        =   30
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   11
                     Left            =   5280
                     TabIndex        =   29
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   12
                     Left            =   5610
                     TabIndex        =   28
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   13
                     Left            =   6000
                     TabIndex        =   27
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   14
                     Left            =   6360
                     TabIndex        =   26
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   15
                     Left            =   6810
                     TabIndex        =   25
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   16
                     Left            =   7140
                     TabIndex        =   24
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   17
                     Left            =   7590
                     TabIndex        =   23
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   18
                     Left            =   7950
                     TabIndex        =   22
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   19
                     Left            =   8280
                     TabIndex        =   21
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
                  Begin VB.CheckBox chkAnnoTool 
                     Height          =   285
                     Index           =   20
                     Left            =   8640
                     TabIndex        =   20
                     Top             =   660
                     Value           =   1  'Checked
                     Width           =   225
                  End
               End
               Begin VB.TextBox txtTitle 
                  Height          =   285
                  Left            =   1020
                  TabIndex        =   18
                  Top             =   240
                  Width           =   3915
               End
               Begin VB.TextBox txtNotes 
                  Height          =   285
                  Left            =   1020
                  TabIndex        =   17
                  Top             =   570
                  Width           =   3915
               End
               Begin VB.TextBox txtXAxis 
                  Height          =   285
                  Left            =   6240
                  TabIndex        =   16
                  Top             =   240
                  Width           =   3405
               End
               Begin VB.TextBox txtYAxis 
                  Height          =   285
                  Left            =   6240
                  TabIndex        =   15
                  Top             =   600
                  Width           =   3405
               End
               Begin VB.Frame FraMiscellenous 
                  Caption         =   "Miscellenous:"
                  Height          =   945
                  Left            =   90
                  TabIndex        =   9
                  Top             =   4680
                  Width           =   9915
                  Begin VB.CheckBox chkBorder 
                     Caption         =   "Border"
                     Height          =   285
                     Left            =   5490
                     TabIndex        =   113
                     Top             =   540
                     Width           =   1155
                  End
                  Begin VB.CheckBox chkZoom 
                     Caption         =   "Zoom"
                     Height          =   255
                     Left            =   4200
                     TabIndex        =   112
                     Top             =   570
                     Width           =   855
                  End
                  Begin VB.CheckBox chkCluster 
                     Caption         =   "Cluster"
                     Height          =   315
                     Left            =   2730
                     TabIndex        =   111
                     Top             =   510
                     Width           =   1065
                  End
                  Begin VB.CheckBox chkShowTips 
                     Caption         =   "Show tips"
                     Height          =   285
                     Left            =   1560
                     TabIndex        =   110
                     Top             =   540
                     Width           =   1155
                  End
                  Begin VB.CheckBox chkCrossHairs 
                     Caption         =   "Cross Hairs"
                     Height          =   285
                     Left            =   120
                     TabIndex        =   109
                     Top             =   540
                     Width           =   1155
                  End
                  Begin VB.ComboBox ComType 
                     Height          =   315
                     ItemData        =   "frmSpatialAnalysis.frx":1EF3E
                     Left            =   7140
                     List            =   "frmSpatialAnalysis.frx":1EF89
                     Style           =   2  'Dropdown List
                     TabIndex        =   100
                     Top             =   480
                     Width           =   2625
                  End
                  Begin VB.CheckBox chkContextMenus 
                     Caption         =   "Context menus"
                     Height          =   285
                     Left            =   120
                     TabIndex        =   14
                     Top             =   270
                     Width           =   1395
                  End
                  Begin VB.CheckBox chkMenuBar 
                     Caption         =   "Menu Bar"
                     Height          =   345
                     Left            =   1560
                     TabIndex        =   13
                     Top             =   240
                     Width           =   1065
                  End
                  Begin VB.CheckBox chkMultipleColors 
                     Caption         =   "Multiple Colors"
                     Height          =   255
                     Left            =   2730
                     TabIndex        =   12
                     Top             =   300
                     Width           =   1395
                  End
                  Begin VB.CheckBox chkPointLabelsGen 
                     Caption         =   "Point Labels"
                     Height          =   285
                     Left            =   4200
                     TabIndex        =   11
                     Top             =   270
                     Width           =   1245
                  End
                  Begin VB.CheckBox chkScrollable 
                     Caption         =   "Scrollable"
                     Height          =   255
                     Left            =   5490
                     TabIndex        =   10
                     Top             =   270
                     Width           =   1095
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Chart Type:"
                     Height          =   195
                     Index           =   9
                     Left            =   7140
                     TabIndex        =   101
                     Top             =   240
                     Width           =   825
                  End
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Title:"
                  Height          =   195
                  Index           =   5
                  Left            =   150
                  TabIndex        =   88
                  Top             =   300
                  Width           =   825
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Notes:"
                  Height          =   195
                  Index           =   6
                  Left            =   150
                  TabIndex        =   87
                  Top             =   630
                  Width           =   825
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "X Axis:"
                  Height          =   195
                  Index           =   7
                  Left            =   5490
                  TabIndex        =   86
                  Top             =   270
                  Width           =   690
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Y Axis:"
                  Height          =   195
                  Index           =   8
                  Left            =   5490
                  TabIndex        =   85
                  Top             =   630
                  Width           =   750
               End
            End
            Begin VB.Frame FraGeneralSettings 
               Caption         =   "General Settings:"
               Height          =   1215
               Left            =   90
               TabIndex        =   3
               Top             =   120
               Width           =   10095
               Begin VB.TextBox txtName 
                  Height          =   315
                  Left            =   1350
                  TabIndex        =   5
                  Top             =   270
                  Width           =   8565
               End
               Begin VB.TextBox txtDesc 
                  Height          =   465
                  Left            =   1350
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   4
                  Top             =   630
                  Width           =   8565
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Name:"
                  Height          =   195
                  Index           =   0
                  Left            =   150
                  TabIndex        =   7
                  Top             =   300
                  Width           =   1185
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Description:"
                  Height          =   345
                  Index           =   1
                  Left            =   120
                  TabIndex        =   6
                  Top             =   690
                  Width           =   1200
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   7620
            Left            =   -11220
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   330
            Width           =   10275
            _cx             =   18124
            _cy             =   13441
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
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Height          =   7440
               Left            =   90
               TabIndex        =   138
               Top             =   90
               Width           =   10095
               Begin VB.CommandButton cmdGenerateTemplate 
                  Caption         =   "Generate"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   8250
                  TabIndex        =   146
                  Top             =   2370
                  Width           =   1815
               End
               Begin VB.Frame frmTemplateDates 
                  Height          =   2115
                  Left            =   8250
                  TabIndex        =   141
                  Top             =   180
                  Width           =   1815
                  Begin XpressEditorsLibCtl.dxDateEdit dxDateFromTemplate 
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
                     Left            =   120
                     OleObjectBlob   =   "frmSpatialAnalysis.frx":1F037
                     TabIndex        =   142
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   1575
                  End
                  Begin XpressEditorsLibCtl.dxDateEdit dxDateTillTemplate 
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
                     Left            =   120
                     OleObjectBlob   =   "frmSpatialAnalysis.frx":1F0D7
                     TabIndex        =   143
                     Top             =   1440
                     Visible         =   0   'False
                     Width           =   1575
                  End
                  Begin VB.Label lblDateFrom 
                     Caption         =   "Date From:"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   145
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   855
                  End
                  Begin VB.Label lblDateTill 
                     Caption         =   "Date Until:"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   144
                     Top             =   1200
                     Visible         =   0   'False
                     Width           =   855
                  End
               End
               Begin VB.ListBox listTemplates 
                  Height          =   2790
                  Left            =   0
                  TabIndex        =   139
                  Top             =   240
                  Width           =   8175
               End
               Begin DXDBGRIDLibCtl.dxDBGrid ego 
                  Bindings        =   "frmSpatialAnalysis.frx":1F177
                  Height          =   4275
                  Left            =   0
                  OleObjectBlob   =   "frmSpatialAnalysis.frx":1F18D
                  TabIndex        =   147
                  Top             =   3120
                  Width           =   10095
               End
               Begin VB.Label lblAvailableTemplates 
                  Caption         =   "Available Templates:"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   140
                  Top             =   0
                  Width           =   2535
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   7620
            Left            =   -10920
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   330
            Width           =   10275
            _cx             =   18124
            _cy             =   13441
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
            Begin VB.CommandButton cmdResetRS 
               Caption         =   "Reset"
               Height          =   255
               Left            =   120
               TabIndex        =   128
               Top             =   3480
               Width           =   4935
            End
            Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
               Bindings        =   "frmSpatialAnalysis.frx":1FE35
               Height          =   3840
               Left            =   120
               OleObjectBlob   =   "frmSpatialAnalysis.frx":1FE4B
               TabIndex        =   127
               Top             =   3840
               Width           =   10095
            End
            Begin VB.Frame Frame4 
               Height          =   3375
               Left            =   5160
               TabIndex        =   122
               Top             =   0
               Width           =   4935
               Begin XpressEditorsLibCtl.dxMemoEdit txtOverlayFilter 
                  Height          =   495
                  Left            =   2640
                  OleObjectBlob   =   "frmSpatialAnalysis.frx":20AF3
                  TabIndex        =   126
                  Top             =   2760
                  Width           =   2175
               End
               Begin VB.ComboBox cmbDateParameter 
                  Height          =   315
                  ItemData        =   "frmSpatialAnalysis.frx":20BEF
                  Left            =   120
                  List            =   "frmSpatialAnalysis.frx":20BF6
                  Style           =   2  'Dropdown List
                  TabIndex        =   136
                  Top             =   2160
                  Width           =   2295
               End
               Begin VB.CheckBox chkUseDate 
                  Caption         =   "Use Date Parameter"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   137
                  Top             =   1920
                  Width           =   2295
               End
               Begin XpressEditorsLibCtl.dxDateEdit dxDateFrom 
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
                  Left            =   1200
                  OleObjectBlob   =   "frmSpatialAnalysis.frx":20C03
                  TabIndex        =   134
                  Top             =   2520
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.ComboBox txtSpatialOperation 
                  Height          =   315
                  ItemData        =   "frmSpatialAnalysis.frx":20CA3
                  Left            =   2640
                  List            =   "frmSpatialAnalysis.frx":20CAA
                  Style           =   2  'Dropdown List
                  TabIndex        =   130
                  Top             =   2160
                  Width           =   2175
               End
               Begin VB.CheckBox chkOverlayFilter 
                  Caption         =   "Filter"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   125
                  Top             =   2520
                  Width           =   2000
               End
               Begin VB.ListBox listOverlayLayer 
                  Height          =   1425
                  Left            =   120
                  Sorted          =   -1  'True
                  TabIndex        =   124
                  Top             =   360
                  Width           =   4695
               End
               Begin XpressEditorsLibCtl.dxDateEdit dxDateTill 
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
                  Left            =   1200
                  OleObjectBlob   =   "frmSpatialAnalysis.frx":20CB9
                  TabIndex        =   135
                  Top             =   2880
                  Visible         =   0   'False
                  Width           =   1215
               End
               Begin VB.Label lblDateTillSetup 
                  Caption         =   "Date Until:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   133
                  Top             =   3000
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.Label lblDateFromSetup 
                  Caption         =   "Date From:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   132
                  Top             =   2640
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.Label lblSpatialOp 
                  Caption         =   "Spatial Operation (DE-9IM)"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   129
                  Top             =   1920
                  Width           =   2175
               End
               Begin VB.Label lblOverlayLayer 
                  Caption         =   "Overlay Layer"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   123
                  Top             =   120
                  Width           =   2295
               End
            End
            Begin VB.Frame frameTarget 
               Height          =   3375
               Left            =   120
               TabIndex        =   115
               Top             =   0
               Width           =   4935
               Begin XpressEditorsLibCtl.dxMemoEdit txtTargetFilter 
                  Height          =   690
                  Left            =   120
                  OleObjectBlob   =   "frmSpatialAnalysis.frx":20D59
                  TabIndex        =   116
                  Top             =   2520
                  Width           =   2295
               End
               Begin VB.CheckBox chkTargetFilter 
                  Caption         =   "Filter"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   119
                  Top             =   2280
                  Width           =   2295
               End
               Begin VB.ComboBox cmbTargetID 
                  Height          =   315
                  Left            =   2520
                  Style           =   2  'Dropdown List
                  TabIndex        =   118
                  Top             =   2880
                  Width           =   2295
               End
               Begin VB.ListBox listTargetLayer 
                  Height          =   1815
                  Left            =   120
                  Sorted          =   -1  'True
                  TabIndex        =   117
                  Top             =   360
                  Width           =   4695
               End
               Begin VB.Label lblTargetLayerAtt 
                  Caption         =   "Target Layer Attribute"
                  Height          =   255
                  Left            =   2520
                  TabIndex        =   121
                  Top             =   2640
                  Width           =   2295
               End
               Begin VB.Label lblTargetLayer 
                  Caption         =   "Target Layer"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   120
                  Top             =   120
                  Width           =   2295
               End
            End
            Begin VB.CommandButton cmdAddSeries 
               Caption         =   "Add Series"
               Height          =   255
               Left            =   5160
               TabIndex        =   114
               Top             =   3480
               Width           =   4935
            End
         End
      End
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   3
      Left            =   0
      TabIndex        =   99
      Top             =   0
      Width           =   1305
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   2
      Left            =   0
      TabIndex        =   98
      Top             =   0
      Width           =   1305
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   1
      Left            =   0
      TabIndex        =   97
      Top             =   0
      Width           =   1305
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   0
      Left            =   4320
      TabIndex        =   92
      Top             =   4170
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmSpatialAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event GetSimpleSpatialAnalysis(sContraint As String, sOverLayLayer As String, sTargetLayer As String, sTargetID As String, sScope As String, sDE_9IM As String, ByRef sSpatialString As String)
Public Event GetAttributes(sLayerName As String, sAttribs() As String, bDateOnly As Boolean, bAnyAttribs As Boolean)
Private Declare Sub Sleep _
                Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim udtOASISChartO As OASISChartObj
Public sSpatialString As String
Private GISLocal As TatukGIS_XDK10.XGIS_Viewer

Dim RSChart As ADODB.Recordset
Dim RSSpatialAnalysisTemplate As ADODB.Recordset
Dim RSLoadedSpatialAnalysisTemplate As ADODB.Recordset
    
Dim doc As DOMDocument
Dim viewsNode As MSXML2.IXMLDOMElement
Dim relationsNode As MSXML2.IXMLDOMElement
Dim catalogsNode As IXMLDOMElement
Dim catalogNode As IXMLDOMElement
Dim m_oFRM As New frmOASISCharts

Public Sub SetGISViewer(GISPassed As TatukGIS_XDK10.XGIS_Viewer)
        '<EhHeader>
        On Error GoTo SetGISViewer_Err
        '</EhHeader>

100     Set GISLocal = GISPassed

        '<EhFooter>
        Exit Sub

SetGISViewer_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.SetGISViewer " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub InitForm()
        '<EhHeader>
        On Error GoTo InitForm_Err
        '</EhHeader>
        Dim i As Integer
        
100     i = GISLocal.Items.Count - 1
102     listTargetLayer.Clear
104     listOverlayLayer.Clear

106     Do Until i < 0
108         listOverlayLayer.AddItem GISLocal.get(GISLocal.Items.Item(i).Name).caption
110         listTargetLayer.AddItem GISLocal.get(GISLocal.Items.Item(i).Name).caption
112         i = i - 1
        Loop
    
114     listTargetLayer.ListIndex = 0
116     listOverlayLayer.ListIndex = 0
        '<EhFooter>
        Exit Sub

InitForm_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.InitForm " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetLayerName(sLayerCaption As String) As String
        '<EhHeader>
        On Error GoTo GetLayerName_Err
        '</EhHeader>

        Dim i As Integer
100     i = GISLocal.Items.Count - 1
    
102     GetLayerName = ""

104     Do Until i < 0

106         If GISLocal.get(GISLocal.Items.Item(i).Name).caption = sLayerCaption Then
108             GetLayerName = GISLocal.Items.Item(i).Name
            End If

110         i = i - 1
        Loop

        '<EhFooter>
        Exit Function

GetLayerName_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.GetLayerName " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function GetLayerCaption(sLayerName As String) As String
        '<EhHeader>
        On Error GoTo GetLayerCaption_Err
        '</EhHeader>

        Dim i As Integer
100     i = GISLocal.Items.Count - 1
    
102     GetLayerCaption = ""

104     Do Until i < 0

106         If GISLocal.Items.Item(i).Name = sLayerName Then
108             GetLayerCaption = GISLocal.get(GISLocal.Items.Item(i).Name).caption
            End If

110         i = i - 1
        Loop

        '<EhFooter>
        Exit Function

GetLayerCaption_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.GetLayerCaption " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub chkUseDate_Click()
        '<EhHeader>
        On Error GoTo chkUseDate_Click_Err
        '</EhHeader>
100     Call cmbDateParameter_Change
        '<EhFooter>
        Exit Sub

chkUseDate_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.chkUseDate_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmbDateParameter_Change()
        '<EhHeader>
        On Error GoTo cmbDateParameter_Change_Err
        '</EhHeader>

100     If chkUseDate = Unchecked Or Me.cmbDateParameter.ListCount = 0 Then
102         Me.dxDateFrom.Visible = False
104         Me.dxDateTill.Visible = False
106         Me.lblDateFromSetup.Visible = False
108         Me.lblDateTillSetup.Visible = False
110         chkUseDate = Unchecked
        Else
112         Me.dxDateFrom.Visible = True
114         Me.dxDateTill.Visible = True
116         Me.lblDateFromSetup.Visible = True
118         Me.lblDateTillSetup.Visible = True
        
        End If
    
        '<EhFooter>
        Exit Sub

cmbDateParameter_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.cmbDateParameter_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GenerateSeriesData(Optional sSeriesName As String)
        '<EhHeader>
        On Error GoTo GenerateSeriesData_Err
        '</EhHeader>

        Dim i As Integer
        Dim lEnd As Long
        Dim sData As String
        Dim sField1 As String
        Dim oRS As ADODB.Recordset
        Dim sAttribs() As String
        Dim iSeriesCount As Integer
        Dim iOldTab As Integer
        Dim sEnteredSeriesName As String
        
100     iOldTab = Me.C1TTab1Tab2.CurrTab
102     sSpatialString = ""
    
104     If Me.cmbTargetID.Text = "" Or Me.listTargetLayer = "" Or Me.listOverlayLayer = "" Or Me.txtSpatialOperation = "" Then
    
106         MsgBox "Please fill in the parameters!"
    
108     ElseIf Me.chkUseDate = vbChecked And (IsNull(dxDateFrom) Or IsNull(dxDateTill)) Then
    
110         MsgBox "Please specify the dates!"
    
        Else
        
112         Me.C1TTab1Tab2.TabVisible(3) = True
114         C1TTab1Tab2.CurrTab = 3
116         C1TTab1Tab2.Refresh
118         cmdAddSeries.Enabled = False
        
120         If Me.chkUseDate Then
122             If Me.chkOverlayFilter = vbChecked And Len(Me.txtOverlayFilter) > 0 Then
124                 RaiseEvent GetSimpleSpatialAnalysis(IIf(Me.chkTargetFilter = vbChecked, Me.txtTargetFilter, ""), GetLayerName(listOverlayLayer), GetLayerName(listTargetLayer), Me.cmbTargetID, Me.txtOverlayFilter & " AND " & Me.cmbDateParameter.Text & " BETWEEN #" & Me.dxDateFrom & "# AND #" & Me.dxDateTill & "#", DE9IM.Item(txtSpatialOperation.Text), sSpatialString)
                Else
126                 RaiseEvent GetSimpleSpatialAnalysis(IIf(Me.chkTargetFilter = vbChecked, Me.txtTargetFilter, ""), GetLayerName(listOverlayLayer), GetLayerName(listTargetLayer), Me.cmbTargetID, Me.cmbDateParameter.Text & " BETWEEN #" & dxDateFrom & "# AND #" & dxDateTill & "#", DE9IM.Item(txtSpatialOperation.Text), sSpatialString)
                End If

            Else
128             RaiseEvent GetSimpleSpatialAnalysis(IIf(Me.chkTargetFilter = vbChecked, Me.txtTargetFilter, ""), GetLayerName(listOverlayLayer), GetLayerName(listTargetLayer), Me.cmbTargetID, IIf(Me.chkOverlayFilter = vbChecked, Me.txtOverlayFilter, ""), DE9IM.Item(txtSpatialOperation.Text), sSpatialString)
            End If
        
130         If sSpatialString = "" Then
132             cmdAddSeries.Enabled = True
                Exit Sub
            End If
            
134         sSpatialString = Replace$(sSpatialString, "?", "`")
136         sSpatialString = Replace$(sSpatialString, "'", "`")
138         sSpatialString = Replace$(sSpatialString, """", "`")
140         sSpatialString = sSpatialString & ":::"

142         If RSChart.Fields.Count = 0 Then
    
144             RSChart.Fields.Append Me.cmbTargetID, adVarChar, 100

146             If sSeriesName = "" Then

148                 sEnteredSeriesName = InputBox("Please enter the name for this series", "Series Name", "Series1")

150                 If sEnteredSeriesName = "" Then
152                     RSChart.Fields.Append "Series1", adBigInt
                    Else
154                     RSChart.Fields.Append sEnteredSeriesName, adBigInt
                    End If

                Else
156                 RSChart.Fields.Append sSeriesName, adBigInt
                End If
            
158             RSChart.Open
    
160             lEnd = InStr(sSpatialString, ":::")

                Do

162                 If lEnd > 0 Then
164                     sData = Left$(sSpatialString, lEnd - 1)
166                     sSpatialString = Mid$(sSpatialString, lEnd + 3)
168                     sData = Replace$(sData, Chr(63), Chr(65))
170                     RSChart.AddNew
                        On Error Resume Next
172                     RSChart.Fields(0).Value = Left(sData, InStr(sData, ",,,") - 1)
174                     RSChart.Fields(1).Value = Right(sData, Len(sData) - Len(RSChart.Fields(0).Value) - 3)
                        On Error GoTo GenerateSeriesData_Err
176                     lEnd = InStr(sSpatialString, ":::")
                    End If

178             Loop While lEnd > 0

            Else
        
180             Set oRS = New ADODB.Recordset

182             i = 0
184             Do Until i = RSChart.Fields.Count

186                 If i = 0 Then
188                     oRS.Fields.Append RSChart.Fields(i).Name, adVarChar, 100
                    Else
190                     oRS.Fields.Append RSChart.Fields(i).Name, adBigInt
                    End If

192                 i = i + 1
                Loop
        
194             If sSeriesName = "" Then
                
196                 sEnteredSeriesName = InputBox("Please enter the name for this series", "Series Name", "Series" & RSChart.Fields.Count)

198                 If sEnteredSeriesName = "" Then
200                     oRS.Fields.Append "Series" & RSChart.Fields.Count, adBigInt
                    Else
202                     oRS.Fields.Append sEnteredSeriesName, adBigInt
                    End If
                
                Else
204                 oRS.Fields.Append sSeriesName, adBigInt
                End If
            
206             oRS.Open
        
208             SafeMoveFirst RSChart
        
210             Do Until RSChart.EOF
        
212                 oRS.AddNew
            
214                 i = 0
216                 Do Until i = RSChart.Fields.Count
218                     oRS.Fields(i).Value = RSChart.Fields(i).Value
220                     i = i + 1
                    Loop
            
222                 RSChart.MoveNext
                Loop
        
224             lEnd = InStr(sSpatialString, ":::")

                Do  ' Starts the loop

226                 If lEnd > 0 Then
228                     sData = Left$(sSpatialString, lEnd - 1)
230                     sSpatialString = Mid$(sSpatialString, lEnd + 3)
                        
                        'What if i cannot find field?
                        'RSChart.AddNew
232                     SafeMoveFirst oRS
234                     oRS.Find "[" & oRS.Fields(0).Name & "] = '" & Left(sData, InStr(sData, ",,,") - 1) & "'"
                
                        ' Fields(0).Value = Left(sData, InStr(sData, ",") - 1)
236                     If Not oRS.EOF Then oRS.Fields(oRS.Fields.Count - 1).Value = Right(sData, Len(sData) - Len(oRS.Fields(0).Value) - 3)
            
238                     lEnd = InStr(sSpatialString, ":::")
                        ' Gets Starting position of ASCII 13 in the new buffer
                    End If

240             Loop While lEnd > 0 ' Loop while ASCII 13 is still present in the buffer
       
242             Set RSChart = oRS.Clone
244             RSChart.Sort = RSChart.Fields(0).Name
246             oRS.Close
248             Set oRS = Nothing
            End If
    
250         RSChart.Sort = "[" & RSChart.Fields(0).Name & "]"
252         Me.dxDBGrid1.Columns.DestroyColumns
254         Set Me.dxDBGrid1.DataSource = RSChart
256         Me.dxDBGrid1.Columns.RetrieveFields
258         frameTarget.Enabled = False
260         Me.cmbTargetID.Enabled = False
262         Me.listTargetLayer.Enabled = False
264         Me.txtTargetFilter.Enabled = False
266         cmdAddSeries.Enabled = True
            'Me.C1TTab1Tab2.TabVisible(1) = False
            'Me.C1TTab1Tab2.TabVisible(2) = False
            
268         Me.cmdExportToHtml.Visible = True
270         Me.cmdExportToXLS.Visible = True
272         Me.cmdExportToXML.Visible = True
274         Me.cmdSaveTemplate.Visible = True
        
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
276         iSeriesCount = Me.dxDBGrid1.Columns.Count - 1
278         RSSpatialAnalysisTemplate.AddNew
        
280         RSSpatialAnalysisTemplate.Fields("SeriesName").Value = dxDBGrid1.Columns(iSeriesCount).caption
282         RSSpatialAnalysisTemplate.Fields("TargetLayer").Value = GetLayerName(listTargetLayer)
284         RSSpatialAnalysisTemplate.Fields("TargetLayerAttribute").Value = Me.cmbTargetID.Text
286         RSSpatialAnalysisTemplate.Fields("TargetLayerFilter").Value = Me.txtTargetFilter
288         RSSpatialAnalysisTemplate.Fields("OverlayLayer").Value = GetLayerName(listOverlayLayer)
290         RSSpatialAnalysisTemplate.Fields("OverlayLayerFilter").Value = Me.txtOverlayFilter
292         RSSpatialAnalysisTemplate.Fields("OverlayLayerSpatialOperation").Value = txtSpatialOperation.Text
294         RSSpatialAnalysisTemplate.Fields("OverlayLayerFilterUsed").Value = IIf(Me.chkOverlayFilter = vbChecked, True, False)
296         RSSpatialAnalysisTemplate.Fields("TargetLayerFilterUsed").Value = IIf(Me.chkTargetFilter = vbChecked, True, False)
298         RSSpatialAnalysisTemplate.Fields("OverlayLayerDateRange").Value = IIf(Me.chkUseDate = vbChecked, True, False)
       
300         If chkUseDate = vbChecked Then
        
302             RSSpatialAnalysisTemplate.Fields("OverlayLayerDateParameter").Value = Me.cmbDateParameter
304             RSSpatialAnalysisTemplate.Fields("OverlayLayerDateFrom").Value = Me.dxDateFrom
306             RSSpatialAnalysisTemplate.Fields("OverlayLayerDateTill").Value = Me.dxDateTill
    
            End If

308         C1TTab1Tab2.CurrTab = iOldTab
310         Me.C1TTab1Tab2.TabVisible(3) = False

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        End If
    
        '<EhFooter>
        Exit Sub

GenerateSeriesData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.GenerateSeriesData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAddSeries_Click()
        '<EhHeader>
        On Error GoTo cmdAddSeries_Click_Err
        '</EhHeader>

100     Call GenerateSeriesData

        '<EhFooter>
        Exit Sub

cmdAddSeries_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.cmdAddSeries_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
        
Private Sub AddChildTextNode(Parent As MSXML2.IXMLDOMNode, _
                             nodeName As String, _
                             Data As String)
        '<EhHeader>
        On Error GoTo AddChildTextNode_Err
        '</EhHeader>
        Dim node1 As IXMLDOMElement
        Dim node2 As IXMLDOMNode
        
100     Set node1 = doc.createElement(nodeName)
102     Parent.appendChild node1
104     Set node2 = doc.createTextNode(Data)
106     node1.appendChild node2

        '<EhFooter>
        Exit Sub

AddChildTextNode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.AddChildTextNode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadChartSettings()
        '<EhHeader>
        On Error GoTo LoadChartSettings_Err
        '</EhHeader>
        On Error Resume Next
        Dim sConnStr As String
        Dim i As Integer

100     With udtOASISChartO
        
            ' chkEnableChartTBR.Value = vbUnchecked
102         chkEnable.Value = vbUnchecked
104         chkAllowData.Value = vbUnchecked
106         chkDataEditor.Value = vbUnchecked
108         chkSeriesLegend.Value = vbUnchecked
110         chkPointLegend.Value = vbUnchecked
112         chkContextMenus.Value = vbUnchecked
114         chkMenuBar.Value = vbUnchecked
116         chkAllowEdit.Value = vbUnchecked
118         chkAllowDrag.Value = vbUnchecked
120         chkMultipleColors.Value = vbUnchecked
122         chkPointLabelsGen.Value = vbUnchecked
124         chkScrollable.Value = vbUnchecked
126         chkPointLabels.Value = vbUnchecked
128         chkDimmed.Value = vbUnchecked
130         chkShowTips.Value = vbUnchecked
132         chkCluster.Value = vbUnchecked
134         chkZoom.Value = vbUnchecked
136         chkBorder.Value = vbUnchecked
138         chkCrossHairs.Value = vbUnchecked

            '  If .bChartTBR Then chkEnableChartTBR.Value = vbChecked
140         If .bAnnoTBR Then chkEnable.Value = vbChecked
142         If .bDataEdtr Then chkDataEditor.Value = vbChecked
144         If .bSeriesLGD Then chkSeriesLegend.Value = vbChecked
146         If .bValueLGD Then chkPointLegend.Value = vbChecked
148         If .bDataHigls Then chkAllowData.Value = vbChecked
150         If .bContextMenu Then chkContextMenus.Value = vbChecked
152         If .bMenuBar Then chkMenuBar.Value = vbChecked
154         If .bDataEdtrAllowEdit Then chkAllowEdit.Value = vbChecked
156         If .bDataEdtrAllowDrag Then chkAllowDrag.Value = vbChecked
158         If .bMultipleColors Then chkMultipleColors.Value = vbChecked
160         If .bPointLabelsGen Then chkPointLabelsGen.Value = vbChecked
162         If .bScrollable Then chkScrollable.Value = vbChecked
164         If .bHGlsPointLabel Then chkPointLabels.Value = vbChecked
166         If .bHGlsDimmed Then chkDimmed.Value = vbChecked
168         If .bShowTips Then chkShowTips.Value = vbChecked
170         If .bCluster Then chkCluster.Value = vbChecked
172         If .bZoom Then chkZoom.Value = vbChecked
174         If .bBorder Then chkBorder.Value = vbChecked
176         If .bCrossHairs Then chkCrossHairs.Value = vbChecked

178         If .bChartTBR Then
180             chkEnableChartTBR.Value = vbChecked
182             ReDim .sChartTools(2)
184             .sChartTools(0) = 0
186             .sChartTools(1) = 20
            Else
188             chkEnableChartTBR.Value = vbUnchecked
            End If

190         If .iSerLgdAlign = oChartAlignment.CA_Docked_Left Then
192             OptLGDPlacement(0).Value = True
            Else
    
194             If .iSerLgdAlign = oChartAlignment.CA_Docked_Right Then
196                 OptLGDPlacement(1).Value = True
                Else

198                 If .iSerLgdAlign = oChartAlignment.CA_Docked_Top Then
200                     OptLGDPlacement(2).Value = True
                    Else
202                     OptLGDPlacement(3).Value = True
                    End If
                End If
    
            End If

204         If .iDataEdtrAlign = oChartAlignment.CA_Docked_Left Then
206             OptDEPlacement(0).Value = True
            Else
    
208             If .iDataEdtrAlign = oChartAlignment.CA_Docked_Right Then
210                 OptDEPlacement(1).Value = True
                Else

212                 If .iDataEdtrAlign = oChartAlignment.CA_Docked_Top Then
214                     OptDEPlacement(2).Value = True
                    Else
216                     OptDEPlacement(3).Value = True
                    End If
                End If
    
            End If

218         txtXAxis.Text = .sXAxis
220         txtYAxis.Text = .sYAxis
222         txtXAngle.Text = .XAxisAngle
        
224         .iHeight = 5000
226         .iWidth = 6000
228         .iParentHeight = 6500
230         .iParentWidth = 7060

232         With .udtTitle
        
234             txtTitle.Text = .CT_Text
236             .CT_BackColor = 2147483647
238             .CT_DrawingArea = True
240             .CT_Alignment = CA_StringAlignment_Center
242             .CT_DockArea = CDA_Area_Top
244             .CT_Flags = CF_TitleFlag_DrawingArea
246             .CT_LineAlignment = CA_StringAlignment_Far
248             .CT_LineGap = 2
            
250             With .CT_Font
252                 lblPrev(0).Font.Size = .CF_Size
254                 lblPrev(0).Font.Bold = .CF_Bold
256                 lblPrev(0).Font.Name = .CF_Name
258                 lblPrev(0).Font.Italic = .CF_Italic
260                 lblPrev(0).Font.Strikethrough = .CF_Strikethrough
262                 lblPrev(0).Font.Underline = .CF_Underline
264                 lblPrev(0).Font.Weight = .CF_Weight
                End With
            End With
   
266         With .udtNotes
        
268             txtNotes.Text = .CT_Text
270             .CT_BackColor = 2147483647
272             .CT_DrawingArea = True
274             .CT_Alignment = CA_StringAlignment_Far
276             .CT_DockArea = CDA_Area_Bottom
278             .CT_Flags = CF_TitleFlag_DrawingArea
280             .CT_LineAlignment = CA_StringAlignment_Far
282             .CT_LineGap = 2
            
284             With .CT_Font
286                 lblPrev(1).Font.Size = .CF_Size
288                 lblPrev(1).Font.Bold = .CF_Bold
290                 lblPrev(1).Font.Name = .CF_Name
292                 lblPrev(1).Font.Italic = .CF_Italic
294                 lblPrev(1).Font.Strikethrough = .CF_Strikethrough
296                 lblPrev(1).Font.Underline = .CF_Underline
298                 lblPrev(1).Font.Weight = .CF_Weight
                End With
            End With
   
300         ComType.ListIndex = .enmChartType - 1
    
302         .enmScheme = CS_Solid
304         .iAngleX = 30
306         .iAngleY = 30
308         .enmAxesStyle = CA_FlatFrame
310         .lngBackColor = 14935011
312         .bBorder = True
314         .lngBorderColor = 11053224
316         .enmBorderEffect = CBE_Dark

318         If .bCluster Then
320             chkCluster.Value = vbChecked
            Else
322             chkCluster.Value = vbUnchecked
            End If
            
324         If .bChart3D Then
326             chk3D.Value = vbChecked
            Else
328             chk3D.Value = vbUnchecked
            End If
            
330         If .bCrossHairs Then
332             chkCrossHairs = vbChecked
            Else
334             chkCrossHairs.Value = vbUnchecked
            End If
        
336         .sngCylSides = 0
338         .enmGrid = CG_None
340         .lngInsideColor = 16777215
342         .enmMarkerShape = CMS_Many
344         .iMarkerSize = 12
346         .sngMarkerStep = 3
348         .sngPerspective = 1
350         .bShowTips = True
352         .enmSmoothFlags = CSF_Fill
354         .enmStacked = CS_No
356         .iWallWidth = 4
358         .bZoom = False
    
        End With
    
        'm_oFRM.SetChart udtOASISChartO
        'm_oFRM.Show vbModeless, Me

        '<EhFooter>
        Exit Sub

LoadChartSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.LoadChartSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateChartSettings(Optional sPath As String, _
                                Optional sName As String)
        '<EhHeader>
        On Error GoTo CreateChartSettings_Err
        '</EhHeader>

        Dim i As Integer
        
100     With udtOASISChartO
        
102         If chkEnableChartTBR.Value = vbChecked Then
        
104             .bChartTBR = True
106             ReDim .sChartTools(0)
        
108             For i = chkMainTool.LBound To chkMainTool.UBound

110                 If chkMainTool(i).Value = vbUnchecked Then

112                     Select Case i
                
                            Case Is < 3
114                             .sChartTools(UBound(.sChartTools)) = i

116                         Case Is < 8
118                             .sChartTools(UBound(.sChartTools)) = i + 2

120                         Case Is < 13
122                             .sChartTools(UBound(.sChartTools)) = i + 3

124                         Case Else
126                             .sChartTools(UBound(.sChartTools)) = i + 4
                        End Select

128                     ReDim Preserve .sChartTools(UBound(.sChartTools) + 1)
                    End If
            
                Next
        
            Else
130             .bChartTBR = False
            End If

132         If chkEnable.Value = vbChecked Then

134             .bAnnoTBR = True
136             ReDim .sAnnoTools(0)

138             For i = chkAnnoTool.LBound To chkAnnoTool.UBound

140                 If chkAnnoTool(i).Value = vbUnchecked Then

142                     Select Case i
                
                            Case Is < 9
                    
144                             .sAnnoTools(UBound(.sAnnoTools)) = i

146                         Case Is < 12
148                             .sAnnoTools(UBound(.sAnnoTools)) = i + 2

150                         Case Is < 15
152                             .sAnnoTools(UBound(.sAnnoTools)) = i + 3

154                         Case Is < 18
156                             .sAnnoTools(UBound(.sAnnoTools)) = i + 4

158                         Case Is < 20
160                             .sAnnoTools(UBound(.sAnnoTools)) = i + 5

162                         Case Else
164                             .sAnnoTools(UBound(.sAnnoTools)) = i + 6
                        End Select

166                     ReDim Preserve .sAnnoTools(UBound(.sAnnoTools) + 1)
                    End If
            
                Next
    
            Else
168             .bAnnoTBR = False
            End If

170         If chkDataEditor.Value = vbChecked Then
172             .bDataEdtr = True
            Else
174             .bDataEdtr = False
            End If

176         If chkSeriesLegend.Value = vbChecked Then
178             .bSeriesLGD = True
            Else
180             .bSeriesLGD = False
            End If

182         If chkPointLegend.Value = vbChecked Then
184             .bValueLGD = True
            Else
186             .bValueLGD = False
            End If

188         If chkAllowData.Value = vbChecked Then
190             .bDataHigls = True
            Else
192             .bDataHigls = False
            End If
        
194         If chkContextMenus.Value = vbChecked Then
196             .bContextMenu = True
            Else
198             .bContextMenu = False
            End If
        
200         If chkMenuBar.Value = vbChecked Then
202             .bMenuBar = True
            Else
204             .bMenuBar = False
            End If

206         If chkAllowEdit.Value = vbChecked Then
208             .bDataEdtrAllowEdit = True
            Else
210             .bDataEdtrAllowEdit = False
            End If
    
212         If chkAllowDrag.Value = vbChecked Then
214             .bDataEdtrAllowDrag = True
            Else
216             .bDataEdtrAllowDrag = False
            End If
    
218         If chkMultipleColors.Value = vbChecked Then
220             .bMultipleColors = True
            Else
222             .bMultipleColors = False
            End If
    
224         If chkPointLabelsGen.Value = vbChecked Then
226             .bPointLabelsGen = True
            Else
228             .bPointLabelsGen = False
            End If
    
230         If chkScrollable.Value = vbChecked Then
232             .bScrollable = True
            Else
234             .bScrollable = False
            End If
    
236         If chkPointLabels.Value = vbChecked Then
238             .bHGlsPointLabel = True
            Else
240             .bHGlsPointLabel = False
            End If
    
242         If chkDimmed.Value = vbChecked Then
244             .bHGlsDimmed = True
            Else
246             .bHGlsDimmed = False
            End If
    
248         .sConnStr = "" 'Text1.Text
250         .sSQL = "" 'txtSQL.Text
    
252         If OptLGDPlacement(0).Value Then
254             .iSerLgdAlign = oChartAlignment.CA_Docked_Left '513
            Else
    
256             If OptLGDPlacement(1).Value Then
258                 .iSerLgdAlign = oChartAlignment.CA_Docked_Right ' 515
                Else

260                 If OptLGDPlacement(2).Value Then
262                     .iSerLgdAlign = oChartAlignment.CA_Docked_Top  ' 256
                    Else
264                     .iSerLgdAlign = oChartAlignment.CA_Docked_Bottom ' 258
                    End If
                End If
    
            End If
        
266         If OptValPlacement(0).Value Then
268             .iValLgdAlign = oChartAlignment.CA_Docked_Left '513
            Else
    
270             If OptValPlacement(1).Value Then
272                 .iValLgdAlign = oChartAlignment.CA_Docked_Right ' 515
                Else

274                 If OptValPlacement(2).Value Then
276                     .iValLgdAlign = oChartAlignment.CA_Docked_Top  ' 256
                    Else
278                     .iValLgdAlign = oChartAlignment.CA_Docked_Bottom ' 258
                    End If
                End If
    
            End If

280         If OptDEPlacement(0).Value Then
282             .iDataEdtrAlign = oChartAlignment.CA_Docked_Left '513
            Else
    
284             If OptDEPlacement(1).Value Then
286                 .iDataEdtrAlign = oChartAlignment.CA_Docked_Right ' 515
                Else

288                 If OptDEPlacement(2).Value Then
290                     .iDataEdtrAlign = oChartAlignment.CA_Docked_Top ' 256
                    Else
292                     .iDataEdtrAlign = oChartAlignment.CA_Docked_Bottom ' 258
                    End If
                End If
    
            End If

294         .sXAxis = txtXAxis.Text
296         .sYAxis = txtYAxis.Text
        
298         .iHeight = 5000
300         .iWidth = 6000
302         .iParentHeight = 6500
304         .iParentWidth = 7060

306         With .udtTitle
308             .CT_Text = txtTitle.Text
310             .CT_BackColor = 2147483647
312             .CT_DrawingArea = True
314             .CT_Alignment = CA_StringAlignment_Center
316             .CT_DockArea = CDA_Area_Top
318             .CT_Flags = CF_TitleFlag_DrawingArea
320             .CT_LineAlignment = CA_StringAlignment_Far
322             .CT_LineGap = 2
            
324             With .CT_Font
326                 .CF_Size = lblPrev(0).Font.Size
328                 .CF_Bold = lblPrev(0).Font.Bold
330                 .CF_Name = lblPrev(0).Font.Name
332                 .CF_Italic = lblPrev(0).Font.Italic
334                 .CF_Strikethrough = lblPrev(0).Font.Strikethrough
336                 .CF_Underline = lblPrev(0).Font.Underline
338                 .CF_Weight = lblPrev(0).Font.Weight ' 500
                End With
            End With
   
340         With .udtNotes
342             .CT_Text = txtNotes.Text
344             .CT_BackColor = 2147483647
346             .CT_DrawingArea = True
348             .CT_Alignment = CA_StringAlignment_Far
350             .CT_DockArea = CDA_Area_Bottom
352             .CT_Flags = CF_TitleFlag_DrawingArea
354             .CT_LineAlignment = CA_StringAlignment_Far
356             .CT_LineGap = 2
            
358             With .CT_Font
360                 .CF_Size = lblPrev(1).Font.Size
362                 .CF_Bold = lblPrev(1).Font.Bold
364                 .CF_Name = lblPrev(1).Font.Name
366                 .CF_Italic = lblPrev(1).Font.Italic
368                 .CF_Strikethrough = lblPrev(1).Font.Strikethrough
370                 .CF_Underline = lblPrev(1).Font.Underline
372                 .CF_Weight = lblPrev(1).Font.Weight ' 500
                End With
            End With
   
374         .enmChartType = ComType.ListIndex + 1 ' Chart_Bubble
    
376         .enmScheme = CS_Solid
378         .XAxisAngle = IIf(IsNumeric(txtXAngle.Text), txtXAngle.Text, 0)
380         .XAxisStaggered = IIf(chkStaggered.Value = vbChecked, True, False)
        
382         .iAngleX = .XAxisAngle
384         .iAngleY = 30
386         .enmAxesStyle = CA_FlatFrame
388         .lngBackColor = 14935011

390         If chkBorder.Value = vbChecked Then
392             .bBorder = True
            Else
394             .bBorder = False
            End If

396         .lngBorderColor = 11053224
398         .enmBorderEffect = CBE_Dark

400         If chkCluster.Value = vbChecked Then
402             .bCluster = True
            Else
404             .bCluster = False
            End If

406         If chk3D.Value = vbChecked Then
408             .bChart3D = True
            Else
410             .bChart3D = False
            End If

412         If chkCrossHairs.Value = vbChecked Then
414             .bCrossHairs = True
            Else
416             .bCrossHairs = False
            End If

418         .sngCylSides = 0
420         .enmGrid = CG_None
422         .lngInsideColor = 16777215
424         .enmMarkerShape = CMS_Many
426         .iMarkerSize = 12
428         .sngMarkerStep = 3
430         .sngPerspective = 1

432         If chkShowTips.Value = vbChecked Then
434             .bShowTips = True
            Else
436             .bShowTips = False
            End If
        
438         .enmSmoothFlags = CSF_Fill
440         .enmStacked = CS_No
442         .iWallWidth = 4

444         If chkZoom.Value = vbChecked Then
446             .bZoom = True
            Else
448             .bZoom = False
            End If
        
450         If Len(sPath) > 0 Then

452             With .udtExports(2)
454                 .bForceKill = True
456                 .enmExportFormat = tplBin
458                 .sFilename = sName
460                 .sPath = sPath
                End With
            
462             m_oFRM.SetChart udtOASISChartO
            
            End If

        End With

        '<EhFooter>
        Exit Sub

CreateChartSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.CreateChartSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdExportToHtml_Click()
        '<EhHeader>
        On Error GoTo cmdExportToHtml_Click_Err
        '</EhHeader>

        Dim c As New cCommonDialog

100     With c
102         .DialogTitle = "Export to location..."
104         .CancelError = False
106         .hWnd = Me.hWnd
            '.Flags = OFN_PATHMUSTEXIST
108         .InitDir = g_sAppPath
110         .Filter = "HTML File|*.html"
112         .FilterIndex = 1
114         .ShowSave
        
        End With
        
116     If Not c.Filename = "" And Not IsNull(c.Filename) Then Me.dxDBGrid1.M.ExportToHTML IIf(Not (Right(c.Filename, 5)) = ".html", c.Filename & ".html", c.Filename)

        '<EhFooter>
        Exit Sub

cmdExportToHtml_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.cmdExportToHtml_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdExportToXLS_Click()
        '<EhHeader>
        On Error GoTo cmdExportToXLS_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog

100     With c
102         .DialogTitle = "Export to location..."
104         .CancelError = False
106         .hWnd = Me.hWnd
            '.'Flags = OFN_PATHMUSTEXIST
108         .InitDir = g_sAppPath
110         .Filter = "Excel File|*.xls"
112         .FilterIndex = 1
114         .ShowSave
        
        End With
        
116     If Not c.Filename = "" And Not IsNull(c.Filename) Then Me.dxDBGrid1.M.ExportToXLS IIf(Not (Right(c.Filename, 4)) = ".xls", c.Filename & ".xls", c.Filename)

        '<EhFooter>
        Exit Sub

cmdExportToXLS_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.cmdExportToXLS_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdFontTitle_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo cmdFontTitle_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog

100     c.Font = lblPrev(Index).Font
102     c.ShowFont
    
104     With lblPrev(Index)
106         .FontBold = c.FontBold
108         .FontItalic = c.FontItalic
110         .FontName = c.FontName
112         .FontSize = c.FontSize
        End With
  
        '<EhFooter>
        Exit Sub

cmdFontTitle_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.cmdFontTitle_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdGenerateTemplate_Click()
        '<EhHeader>
        On Error GoTo cmdGenerateTemplate_Click_Err
        '</EhHeader>
    
100     If IsNull(dxDateFromTemplate) Or IsNull(dxDateTillTemplate) Then
            
102         MsgBox "You must specify dates!"

        Else
        
104         If Not IsNull(listTemplates.Text) And Not listTemplates.Text = "" Then
     
106             Call cmdResetRS_Click
108             RSLoadedSpatialAnalysisTemplate.Close
110             RSLoadedSpatialAnalysisTemplate.Open g_sAppPath & "\data\templates\spatialanalysis\" & Me.listTemplates & ".xml"
        
112             If RSLoadedSpatialAnalysisTemplate.State = adStateOpen Then

114                 Me.cmdGenerateTemplate.Enabled = False
116                 SafeMoveFirst RSLoadedSpatialAnalysisTemplate

118                 Do Until RSLoadedSpatialAnalysisTemplate.EOF
            
120                     listTargetLayer.Text = GetLayerCaption(RSLoadedSpatialAnalysisTemplate.Fields("TargetLayer").Value)
122                     cmbTargetID.Text = RSLoadedSpatialAnalysisTemplate.Fields("TargetLayerAttribute").Value
124                     txtTargetFilter = RSLoadedSpatialAnalysisTemplate.Fields("TargetLayerFilter").Value
126                     listOverlayLayer.Text = GetLayerCaption(RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayer").Value)
128                     txtOverlayFilter = RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerFilter").Value
130                     txtSpatialOperation.Text = RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerSpatialOperation").Value
                
132                     If RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerDateRange").Value = True Then
134                         cmbDateParameter = RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerDateParameter").Value
136                         dxDateFrom = dxDateFromTemplate 'RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerDateFrom").Value
138                         dxDateTill = dxDateTillTemplate 'RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerDateTill").Value
140                         chkUseDate = vbChecked
                        Else
142                         chkUseDate = vbUnchecked
                        End If
                
144                     If RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerFilterUsed").Value = True Then
146                         chkOverlayFilter = vbChecked
                        Else
148                         chkOverlayFilter = vbUnchecked
                        End If
                
150                     If RSLoadedSpatialAnalysisTemplate.Fields("TargetLayerFilterUsed").Value = True Then
152                         chkTargetFilter = vbChecked
                        Else
154                         chkTargetFilter = vbUnchecked
                        End If
                
156                     GenerateSeriesData RSLoadedSpatialAnalysisTemplate.Fields("SeriesName").Value
158                     RSLoadedSpatialAnalysisTemplate.MoveNext
                    Loop
            
160                 Me.ego.Columns.DestroyColumns
162                 Set Me.ego.DataSource = Me.dxDBGrid1.DataSource
164                 Me.ego.Columns.RetrieveFields
166                 Me.cmdGenerateTemplate.Enabled = True
168                 udtOASISChartO = m_oFRM.OpenChartTemplate(g_sAppPath & "\data\templates\spatialanalysis\" & Me.listTemplates & ".oct")
170                 LoadChartSettings
172                 m_oFRM.SetChartWithRS udtOASISChartO, RSChart
174                 m_oFRM.Show vbModeless, Me
                    'Call cmdSave_Click
            
                End If

            Else
        
176             MsgBox "Please select a template!"
            
            End If
    
        End If

        '<EhFooter>
        Exit Sub
cmdGenerateTemplate_Click_Err:
        MsgBox "There was an error generating this template.  Please contact your OASIS Administrator to get this template fixed", vbInformation, "Template in error"
        DebugPrint "There was an error generating this template.  Please contact your OASIS Administrator to get this template fixed"
        DebugPrint Err.Description & vbCrLf & "in OASISClient.frmSpatialAnalysis.cmdGenerateTemplate_Click " & "at line " & Erl
        Me.cmdGenerateTemplate.Enabled = True
        'Resume Next
        '</EhFooter>
End Sub

Private Sub InitialiseTemplateRS()
        '<EhHeader>
        On Error GoTo InitialiseTemplateRS_Err
        '</EhHeader>

100     If Not RSSpatialAnalysisTemplate Is Nothing Then
102         If RSSpatialAnalysisTemplate.State = adStateOpen Then RSSpatialAnalysisTemplate.Close
        End If

104     Set RSSpatialAnalysisTemplate = New ADODB.Recordset
106     RSSpatialAnalysisTemplate.Fields.Append "SeriesName", adVarChar, 255
108     RSSpatialAnalysisTemplate.Fields.Append "TargetLayer", adVarChar, 255
110     RSSpatialAnalysisTemplate.Fields.Append "TargetLayerAttribute", adVarChar, 255
112     RSSpatialAnalysisTemplate.Fields.Append "TargetLayerFilterUsed", adBoolean
114     RSSpatialAnalysisTemplate.Fields.Append "TargetLayerFilter", adVarChar, 255
116     RSSpatialAnalysisTemplate.Fields.Append "OverlayLayer", adVarChar, 255
118     RSSpatialAnalysisTemplate.Fields.Append "OverlayLayerFilterUsed", adBoolean
120     RSSpatialAnalysisTemplate.Fields.Append "OverlayLayerFilter", adVarChar, 255
122     RSSpatialAnalysisTemplate.Fields.Append "OverlayLayerSpatialOperation", adVarChar, 255
124     RSSpatialAnalysisTemplate.Fields.Append "OverlayLayerDateRange", adBoolean
126     RSSpatialAnalysisTemplate.Fields.Append "OverlayLayerDateParameter", adVarChar, 255
128     RSSpatialAnalysisTemplate.Fields.Append "OverlayLayerDateFrom", adDate
130     RSSpatialAnalysisTemplate.Fields.Append "OverlayLayerDateTill", adDate
132     RSSpatialAnalysisTemplate.Open
    
        '<EhFooter>
        Exit Sub

InitialiseTemplateRS_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmSpatialAnalysis.InitialiseTemplateRS " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdResetRS_Click()
        '<EhHeader>
        On Error GoTo cmdResetRS_Click_Err
        '</EhHeader>

100     If Not RSChart.State = adStateClosed Then RSChart.Close
102     Me.dxDBGrid1.Columns.DestroyColumns
104     Set RSChart = New ADODB.Recordset
106     frameTarget.Enabled = True
108     Me.cmbTargetID.Enabled = True
110     Me.listTargetLayer.Enabled = True
112     Me.txtTargetFilter.Enabled = True
114     Me.C1TTab1Tab2.TabVisible(2) = False
116     Me.cmdExportToHtml.Visible = False
118     Me.cmdExportToXLS.Visible = False
120     Me.cmdExportToXML.Visible = False
122     Me.cmdSaveTemplate.Visible = False
124     Call InitialiseTemplateRS
    
        '<EhFooter>
        Exit Sub

cmdResetRS_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.cmdResetRS_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>

100     CreateChartSettings
102     m_oFRM.SetChartWithRS udtOASISChartO, RSChart
104     m_oFRM.Show vbModeless, Me
        '<EhFooter>
        Exit Sub

cmdSave_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.cmdSave_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSaveTemplate_Click()
        '<EhHeader>
        On Error GoTo cmdSaveTemplate_Click_Err
        '</EhHeader>

        Dim sFilename As String
        Dim sFileNameXML As String
        Dim sFileNameOCT As String
        Dim sFileNamePATH As String
        Dim desc As String
        
        If RSChart.State = adStateClosed Then
            MsgBox "Please generate a spatial analysis query first!", vbInformation, "Spatial query not defined"
        ElseIf RSChart.RecordCount = 0 Then
            MsgBox "Please generate a spatial analysis query first!", vbInformation, "Spatial query not defined"
        Else
        
        desc = InputBox("Enter description for this template", "OASIS Spatial Analysis Templates")
        
118     If Not desc = "" And Not IsNull(desc) Then
        
120         If Not FileExists(g_sAppPath & "\data\templates\spatialanalysis\" & desc & ".xml") Then
        
122             If (Len(desc) < 4) Then
124                 sFilename = desc
                Else
126                 sFilename = IIf(Right(desc, 4) = ".xml", Left(desc, Len(desc) - 4), desc)
                End If
        
128             sFileNamePATH = g_sAppPath & "\data\templates\spatialanalysis\" 'Replace(c.Filename, c.FileTitle, "")"
130             sFileNameXML = sFilename & ".xml"
132             sFileNameOCT = sFilename & ".oct"
        
134             RSSpatialAnalysisTemplate.Save sFileNamePATH & sFileNameXML, adPersistXML
136             CreateChartSettings sFileNamePATH, sFilename
138             populateListView g_sAppPath & "\data\templates\spatialanalysis", "xml", Me.listTemplates
140             MsgBox "Two files (" & sFilename & ".xml & " & sFilename & ".oct) were created in directory:" & Chr(13) & Chr(13) & sFileNamePATH, vbInformation, "Template saved"
            Else
        
142             MsgBox "A template by this name already exists!", vbExclamation, "Template already exists"
        
            End If
        End If
        
        End If

        '<EhFooter>
        Exit Sub

cmdSaveTemplate_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.cmdSaveTemplate_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function FileExists(sFullPath As String) As Boolean
    Dim oFile As New Scripting.FileSystemObject
    FileExists = oFile.FileExists(sFullPath)
End Function

Private Sub cmdExportToXML_Click()
        '<EhHeader>
        On Error GoTo cmdExportToXML_Click_Err
        '</EhHeader>
        
        Dim c As New cCommonDialog

100     With c
102         .DialogTitle = "Export to location..."
104         .CancelError = False
106         .hWnd = Me.hWnd
108         .Flags = OFN_PATHMUSTEXIST
110         .InitDir = g_sAppPath
112         .Filter = "XML File|*.xml"
114         .FilterIndex = 1
116         .ShowSave
        
        End With
        
118     If Not c.Filename = "" And Not IsNull(c.Filename) Then Me.dxDBGrid1.M.ExportToXML IIf(Not (Right(c.Filename, 4)) = ".xml", c.Filename & ".xml", c.Filename)

        '<EhFooter>
        Exit Sub

cmdExportToXML_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.cmdExportToXML_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub populateListView(sPath As String, _
                             sFilter As String, _
                             lListBox As ListBox)
        '<EhHeader>
        On Error GoTo populateListView_Err
        '</EhHeader>

        Dim oItem As MSComctlLib.ListItem
        Dim sFile As String
   
100     lListBox.Clear
    
102     If Right(sPath, 1) <> "\" Then
104         sPath = sPath & "\"
        End If
   
106     sFile = Dir(sPath & "*." & sFilter)
   
108     While sFile <> Empty
110         lListBox.AddItem Left$(sFile, Len(sFile) - 4)
112         sFile = Dir
        Wend
   
        '<EhFooter>
        Exit Sub

populateListView_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.populateListView " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
        '<EhHeader>
        On Error GoTo Form_KeyDown_Err
        '</EhHeader>
    
100     If KeyCode = vbKeyF11 And Me.C1TTab1Tab2.TabVisible(1) Then
102         C1TTab1Tab2.CurrTab = 0
104         C1TTab1Tab2.TabVisible(1) = False
106         C1TTab1Tab2.TabVisible(2) = False
108     ElseIf KeyCode = vbKeyF11 Then
110         C1TTab1Tab2.TabVisible(1) = True
112         C1TTab1Tab2.TabVisible(2) = True
114         C1TTab1Tab2.CurrTab = 1
        End If
    
        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.Form_KeyDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
        Dim keyArray() As Variant
        Dim element As Variant
        
100     If Not g_sLanguage = "" Then
102         If Not m_Cnn.State = adStateClosed Then
104             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If
        
'106     Set Me.Picture3.Picture = g_PictureDialogLogo
    
108     ComType.ListIndex = 0

110     Set RSChart = New ADODB.Recordset
112     Me.C1TTab1Tab2.CurrTab = 0
114     Me.C1TTab1Tab2.TabVisible(1) = False
116     Me.C1TTab1Tab2.TabVisible(2) = False
118     Me.C1TTab1Tab2.TabVisible(3) = False

120     txtSpatialOperation.Clear
122     keyArray = DE9IM.Keys

124     For Each element In keyArray
126         Me.txtSpatialOperation.AddItem element
        Next

128     txtSpatialOperation.ListIndex = 0
130     Call InitialiseTemplateRS
132     populateListView g_sAppPath & "\data\templates\spatialanalysis", "xml", Me.listTemplates

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
        On Error Resume Next
        
100     Set RSChart = Nothing
102     Set RSSpatialAnalysisTemplate = Nothing
104     Set RSLoadedSpatialAnalysisTemplate = Nothing
106     Unload m_oFRM

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
 
Private Sub listOverlayLayer_Click()
        '<EhHeader>
        On Error GoTo listOverlayLayer_Click_Err
        '</EhHeader>

        Dim i As Integer
        Dim sAttribs() As String
        Dim bAnyAttribs As Boolean
    
100     If Not IsNull(listOverlayLayer) And Not listOverlayLayer = "" Then
    
102         RaiseEvent GetAttributes(GetLayerName(Me.listOverlayLayer), sAttribs, True, bAnyAttribs)
104         Me.cmbDateParameter.Clear
106         i = 0
        
108         If bAnyAttribs Then
            
110             Do Until i > UBound(sAttribs)
112                 Me.cmbDateParameter.AddItem sAttribs(i)
114                 i = i + 1
                Loop
         
116             If Not cmbDateParameter.ListCount = 0 Then cmbDateParameter.ListIndex = 0
            
            End If

118         Call cmbDateParameter_Change
        End If
    
        '<EhFooter>
        Exit Sub

listOverlayLayer_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.listOverlayLayer_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub listTargetLayer_Click()
        '<EhHeader>
        On Error GoTo listTargetLayer_Click_Err
        '</EhHeader>

        Dim i As Integer
        Dim sAttribs() As String
        Dim bAnyAttribs As Boolean
    
100     If Not IsNull(listTargetLayer) And Not listTargetLayer = "" Then
    
102         RaiseEvent GetAttributes(GetLayerName(listTargetLayer), sAttribs, False, bAnyAttribs)
         
104         Me.cmbTargetID.Clear
106         i = 0
        
108         If bAnyAttribs Then
            
110             Do Until i > UBound(sAttribs)
             
112                 Me.cmbTargetID.AddItem sAttribs(i)
114                 i = i + 1
                Loop
         
116             If Not cmbTargetID.ListCount = 0 Then cmbTargetID.ListIndex = 0
            
            End If
        End If

        '<EhFooter>
        Exit Sub

listTargetLayer_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.listTargetLayer_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub listTemplates_Click()
        '<EhHeader>
        On Error GoTo listTemplates_Click_Err
        '</EhHeader>

        Dim bDatesUsed As Boolean
        Dim dDateFrom As Date
        Dim dDateTill As Date
    
100     bDatesUsed = False
    
102     If Not IsNull(listTemplates.Text) Then

104         If Not RSLoadedSpatialAnalysisTemplate Is Nothing Then RSLoadedSpatialAnalysisTemplate.Close
106         Set RSLoadedSpatialAnalysisTemplate = New ADODB.Recordset
108         RSLoadedSpatialAnalysisTemplate.Open g_sAppPath & "\data\templates\spatialanalysis\" & Me.listTemplates & ".xml"
        
110         If RSLoadedSpatialAnalysisTemplate.State = adStateOpen Then
        
112             Do Until RSLoadedSpatialAnalysisTemplate.EOF
            
114                 If RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerDateRange").Value = True Then
116                     bDatesUsed = True
118                     dDateFrom = RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerDateFrom").Value
120                     dDateTill = RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerDateTill").Value
                    End If
            
122                 RSLoadedSpatialAnalysisTemplate.MoveNext
                Loop

124             SafeMoveFirst RSLoadedSpatialAnalysisTemplate
        
126             If Not bDatesUsed Then
128                 Me.dxDateFromTemplate.Visible = False
130                 Me.dxDateTillTemplate.Visible = False
132                 Me.lblDateFrom.Visible = False
134                 Me.lblDateTill.Visible = False
                Else
136                 Me.dxDateFromTemplate.Visible = True
138                 Me.dxDateTillTemplate.Visible = True
140                 Me.lblDateFrom.Visible = True
142                 Me.lblDateTill.Visible = True
144                 Me.dxDateFromTemplate = dDateFrom
146                 Me.dxDateTillTemplate = dDateTill
                End If
        
            End If

        End If

        '<EhFooter>
        Exit Sub

listTemplates_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSpatialAnalysis.listTemplates_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
