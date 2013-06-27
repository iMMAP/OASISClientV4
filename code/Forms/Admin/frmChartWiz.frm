VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{605925BE-4799-4093-A2E7-39323147E70E}#1.0#0"; "C1Query8.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChartWiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Chart Template Creator"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9840
   Icon            =   "frmChartWiz.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Template"
      Height          =   345
      Left            =   5910
      TabIndex        =   110
      Top             =   8460
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Template"
      Height          =   345
      Left            =   7230
      TabIndex        =   109
      Top             =   8460
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   8550
      TabIndex        =   108
      Top             =   8460
      Width           =   1245
   End
   Begin C1SizerLibCtl.C1Elastic C1EMain 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9825
      _cx             =   17330
      _cy             =   14843
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
         Height          =   8235
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   9645
         _cx             =   17013
         _cy             =   14526
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   7860
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   9555
            _cx             =   16854
            _cy             =   13864
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
            Begin VB.Frame FraChartSettings 
               Caption         =   "Chart Settings:"
               Height          =   6150
               Left            =   0
               TabIndex        =   8
               Top             =   1560
               Width           =   9555
               Begin VB.Frame FraYAxislblFormat 
                  Caption         =   "Y Axis Label Format:"
                  Height          =   705
                  Left            =   5010
                  TabIndex        =   138
                  Top             =   2940
                  Width           =   4425
                  Begin VB.TextBox txtYDecPlace 
                     Height          =   315
                     Left            =   2700
                     TabIndex        =   141
                     Text            =   "0"
                     Top             =   240
                     Width           =   465
                  End
                  Begin VB.ComboBox ComYLBLFormat 
                     Height          =   315
                     ItemData        =   "frmChartWiz.frx":6852
                     Left            =   150
                     List            =   "frmChartWiz.frx":6871
                     Style           =   2  'Dropdown List
                     TabIndex        =   139
                     Top             =   240
                     Width           =   1755
                  End
                  Begin VB.Label lblDecimal 
                     AutoSize        =   -1  'True
                     Caption         =   "Decimal:"
                     Height          =   195
                     Left            =   2040
                     TabIndex        =   140
                     Top             =   270
                     Width           =   615
                  End
               End
               Begin VB.Frame FraXAxis 
                  Caption         =   "X Axis:"
                  Height          =   1035
                  Left            =   7950
                  TabIndex        =   133
                  Top             =   930
                  Width           =   1485
                  Begin VB.TextBox txtXAngle 
                     Height          =   285
                     Left            =   630
                     TabIndex        =   135
                     Text            =   "45"
                     Top             =   210
                     Width           =   795
                  End
                  Begin VB.CheckBox chkStaggered 
                     Caption         =   "Staggered"
                     Height          =   285
                     Left            =   60
                     TabIndex        =   134
                     Top             =   570
                     Width           =   1395
                  End
                  Begin VB.Label lblAngle 
                     AutoSize        =   -1  'True
                     Caption         =   "Angle:"
                     Height          =   195
                     Left            =   60
                     TabIndex        =   136
                     Top             =   210
                     Width           =   450
                  End
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   3
                  Left            =   9180
                  TabIndex        =   115
                  Top             =   570
                  Width           =   255
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   2
                  Left            =   9180
                  TabIndex        =   114
                  Top             =   240
                  Width           =   255
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   1
                  Left            =   4980
                  TabIndex        =   113
                  Top             =   540
                  Width           =   255
               End
               Begin VB.CommandButton cmdFontTitle 
                  Caption         =   "..."
                  Height          =   315
                  Index           =   0
                  Left            =   4980
                  TabIndex        =   112
                  Top             =   210
                  Width           =   255
               End
               Begin VB.Frame FraChartTools 
                  Caption         =   "Chart Tools:"
                  Height          =   1035
                  Left            =   90
                  TabIndex        =   65
                  Top             =   930
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
                        Picture         =   "frmChartWiz.frx":68C5
                        ScaleHeight     =   465
                        ScaleWidth      =   6705
                        TabIndex        =   107
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
                  Width           =   4725
                  Begin VB.Frame Frame2 
                     BorderStyle     =   0  'None
                     Caption         =   "Frame2"
                     Height          =   285
                     Left            =   1650
                     TabIndex        =   121
                     Top             =   540
                     Width           =   2955
                     Begin VB.OptionButton OptValPlacement 
                        Caption         =   "Bottom"
                        Height          =   255
                        Index           =   3
                        Left            =   2130
                        TabIndex        =   125
                        Top             =   30
                        Width           =   795
                     End
                     Begin VB.OptionButton OptValPlacement 
                        Caption         =   "Top"
                        Height          =   255
                        Index           =   2
                        Left            =   1440
                        TabIndex        =   124
                        Top             =   30
                        Width           =   615
                     End
                     Begin VB.OptionButton OptValPlacement 
                        Caption         =   "Right"
                        Height          =   255
                        Index           =   1
                        Left            =   690
                        TabIndex        =   123
                        Top             =   30
                        Width           =   705
                     End
                     Begin VB.OptionButton OptValPlacement 
                        Caption         =   "Left"
                        Height          =   255
                        Index           =   0
                        Left            =   30
                        TabIndex        =   122
                        Top             =   30
                        Value           =   -1  'True
                        Width           =   615
                     End
                  End
                  Begin VB.CheckBox chkSeriesLegend 
                     Caption         =   "Series"
                     Height          =   285
                     Left            =   30
                     TabIndex        =   62
                     Top             =   240
                     Width           =   765
                  End
                  Begin VB.CheckBox chkPointLegend 
                     Caption         =   "Value"
                     Height          =   285
                     Left            =   30
                     TabIndex        =   61
                     Top             =   570
                     Width           =   735
                  End
                  Begin VB.OptionButton OptLGDPlacement 
                     Caption         =   "Left"
                     Height          =   255
                     Index           =   0
                     Left            =   1680
                     TabIndex        =   60
                     Top             =   270
                     Value           =   -1  'True
                     Width           =   615
                  End
                  Begin VB.OptionButton OptLGDPlacement 
                     Caption         =   "Right"
                     Height          =   255
                     Index           =   1
                     Left            =   2340
                     TabIndex        =   59
                     Top             =   270
                     Width           =   705
                  End
                  Begin VB.OptionButton OptLGDPlacement 
                     Caption         =   "Top"
                     Height          =   255
                     Index           =   2
                     Left            =   3090
                     TabIndex        =   58
                     Top             =   270
                     Width           =   615
                  End
                  Begin VB.OptionButton OptLGDPlacement 
                     Caption         =   "Bottom"
                     Height          =   255
                     Index           =   3
                     Left            =   3780
                     TabIndex        =   57
                     Top             =   270
                     Width           =   795
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Alignement:"
                     Height          =   195
                     Index           =   2
                     Left            =   810
                     TabIndex        =   64
                     Top             =   270
                     Width           =   825
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Alignement:"
                     Height          =   195
                     Index           =   3
                     Left            =   810
                     TabIndex        =   63
                     Top             =   600
                     Width           =   825
                  End
               End
               Begin VB.Frame FraDataEditor 
                  Caption         =   "Data Editor:"
                  Height          =   945
                  Left            =   5010
                  TabIndex        =   47
                  Top             =   1980
                  Width           =   4425
                  Begin VB.CheckBox chkDataEditor 
                     Caption         =   "Show"
                     Height          =   255
                     Left            =   90
                     TabIndex        =   54
                     Top             =   240
                     Width           =   765
                  End
                  Begin VB.OptionButton OptDEPlacement 
                     Caption         =   "Left"
                     Height          =   255
                     Index           =   0
                     Left            =   990
                     TabIndex        =   53
                     Top             =   570
                     Width           =   615
                  End
                  Begin VB.OptionButton OptDEPlacement 
                     Caption         =   "Right"
                     Height          =   255
                     Index           =   1
                     Left            =   1650
                     TabIndex        =   52
                     Top             =   570
                     Width           =   705
                  End
                  Begin VB.OptionButton OptDEPlacement 
                     Caption         =   "Top"
                     Height          =   255
                     Index           =   2
                     Left            =   2400
                     TabIndex        =   51
                     Top             =   570
                     Width           =   615
                  End
                  Begin VB.OptionButton OptDEPlacement 
                     Caption         =   "Bottom"
                     Height          =   255
                     Index           =   3
                     Left            =   3090
                     TabIndex        =   50
                     Top             =   570
                     Value           =   -1  'True
                     Width           =   795
                  End
                  Begin VB.CheckBox chkAllowEdit 
                     Caption         =   "Allow Edit"
                     Height          =   255
                     Left            =   960
                     TabIndex        =   49
                     Top             =   240
                     Width           =   1065
                  End
                  Begin VB.CheckBox chkAllowDrag 
                     Caption         =   "Allow Drag"
                     Height          =   285
                     Left            =   2040
                     TabIndex        =   48
                     Top             =   210
                     Width           =   1095
                  End
                  Begin VB.Label lblNoys 
                     AutoSize        =   -1  'True
                     Caption         =   "Alignement:"
                     Height          =   195
                     Index           =   4
                     Left            =   60
                     TabIndex        =   55
                     Top             =   570
                     Width           =   825
                  End
               End
               Begin VB.Frame FraDataHighlighting 
                  Caption         =   "Data Highlighting:"
                  Height          =   705
                  Left            =   90
                  TabIndex        =   43
                  Top             =   2940
                  Width           =   4725
                  Begin VB.CheckBox chkAllowData 
                     Caption         =   "Allow"
                     Height          =   255
                     Left            =   90
                     TabIndex        =   46
                     Top             =   300
                     Width           =   765
                  End
                  Begin VB.CheckBox chkDimmed 
                     Caption         =   "Dimmed"
                     Height          =   255
                     Left            =   1050
                     TabIndex        =   45
                     Top             =   300
                     Width           =   945
                  End
                  Begin VB.CheckBox chkPointLabels 
                     Caption         =   "Point Labels"
                     Height          =   255
                     Left            =   2190
                     TabIndex        =   44
                     Top             =   300
                     Width           =   1215
                  End
               End
               Begin VB.Frame FraAnnotations 
                  Caption         =   "Annotations:"
                  Height          =   1035
                  Left            =   90
                  TabIndex        =   19
                  Top             =   3690
                  Width           =   9375
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
                     Picture         =   "frmChartWiz.frx":10BC7
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
                  Left            =   570
                  TabIndex        =   18
                  Top             =   240
                  Width           =   4365
               End
               Begin VB.TextBox txtNotes 
                  Height          =   285
                  Left            =   900
                  TabIndex        =   17
                  Top             =   570
                  Width           =   4035
               End
               Begin VB.TextBox txtXAxis 
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   16
                  Top             =   240
                  Width           =   3255
               End
               Begin VB.TextBox txtYAxis 
                  Height          =   285
                  Left            =   5880
                  TabIndex        =   15
                  Top             =   600
                  Width           =   3255
               End
               Begin VB.Frame FraMiscellenous 
                  Caption         =   "Miscellenous:"
                  Height          =   1305
                  Left            =   90
                  TabIndex        =   9
                  Top             =   4740
                  Width           =   9375
                  Begin VB.CheckBox chk3D 
                     Caption         =   "3D"
                     Height          =   285
                     Left            =   6630
                     TabIndex        =   137
                     Top             =   540
                     Width           =   555
                  End
                  Begin VB.CheckBox chkBorder 
                     Caption         =   "Border"
                     Height          =   285
                     Left            =   5490
                     TabIndex        =   130
                     Top             =   540
                     Width           =   1155
                  End
                  Begin VB.CheckBox chkZoom 
                     Caption         =   "Zoom"
                     Height          =   255
                     Left            =   4200
                     TabIndex        =   129
                     Top             =   570
                     Width           =   855
                  End
                  Begin VB.CheckBox chkCluster 
                     Caption         =   "Cluster"
                     Height          =   315
                     Left            =   2730
                     TabIndex        =   128
                     Top             =   510
                     Width           =   1065
                  End
                  Begin VB.CheckBox chkShowTips 
                     Caption         =   "Show tips"
                     Height          =   285
                     Left            =   1560
                     TabIndex        =   127
                     Top             =   540
                     Width           =   1155
                  End
                  Begin VB.CheckBox chkCrossHairs 
                     Caption         =   "Cross Hairs"
                     Height          =   285
                     Left            =   120
                     TabIndex        =   126
                     Top             =   540
                     Width           =   1155
                  End
                  Begin VB.ComboBox ComType 
                     Height          =   315
                     ItemData        =   "frmChartWiz.frx":1EFB1
                     Left            =   7440
                     List            =   "frmChartWiz.frx":1EFFC
                     Style           =   2  'Dropdown List
                     TabIndex        =   119
                     Top             =   240
                     Width           =   1815
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
                     Left            =   6600
                     TabIndex        =   120
                     Top             =   270
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
                  Width           =   345
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Notes:"
                  Height          =   195
                  Index           =   6
                  Left            =   180
                  TabIndex        =   87
                  Top             =   660
                  Width           =   465
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "X Axis:"
                  Height          =   195
                  Index           =   7
                  Left            =   5340
                  TabIndex        =   86
                  Top             =   270
                  Width           =   480
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Y Axis:"
                  Height          =   195
                  Index           =   8
                  Left            =   5370
                  TabIndex        =   85
                  Top             =   600
                  Width           =   480
               End
            End
            Begin VB.Frame FraGeneralSettings 
               Caption         =   "General Settings:"
               Height          =   1395
               Left            =   30
               TabIndex        =   3
               Top             =   120
               Width           =   9495
               Begin VB.CheckBox chkShowPreview 
                  Caption         =   "Show Preview"
                  Height          =   285
                  Left            =   6480
                  TabIndex        =   132
                  Top             =   240
                  Width           =   1515
               End
               Begin VB.CommandButton cmdAssign 
                  Caption         =   "Update"
                  Height          =   345
                  Left            =   8220
                  TabIndex        =   131
                  Top             =   210
                  Width           =   1065
               End
               Begin VB.TextBox txtName 
                  Height          =   315
                  Left            =   1110
                  TabIndex        =   5
                  Top             =   270
                  Width           =   2895
               End
               Begin VB.TextBox txtDesc 
                  Height          =   645
                  Left            =   1110
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   4
                  Top             =   630
                  Width           =   8205
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Name:"
                  Height          =   195
                  Index           =   0
                  Left            =   150
                  TabIndex        =   7
                  Top             =   300
                  Width           =   465
               End
               Begin VB.Label lblNoys 
                  AutoSize        =   -1  'True
                  Caption         =   "Description:"
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   6
                  Top             =   690
                  Width           =   840
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   7860
            Left            =   10590
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   330
            Width           =   9555
            _cx             =   16854
            _cy             =   13864
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
            Height          =   7860
            Left            =   10290
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   330
            Width           =   9555
            _cx             =   16854
            _cy             =   13864
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
            Begin VB.TextBox Text1 
               Height          =   375
               Left            =   4470
               TabIndex        =   98
               Top             =   990
               Width           =   4335
            End
            Begin VB.ListBox List2 
               Height          =   1230
               Left            =   4470
               TabIndex        =   97
               Top             =   1650
               Width           =   4335
            End
            Begin VB.ListBox List1 
               Height          =   1230
               Left            =   30
               TabIndex        =   96
               Top             =   1650
               Width           =   4335
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Load database schema from this database:"
               Height          =   405
               Left            =   30
               TabIndex        =   95
               Top             =   960
               Width           =   4335
            End
            Begin VB.CommandButton cmd 
               Caption         =   "..."
               Height          =   375
               Left            =   8850
               TabIndex        =   94
               Top             =   990
               Width           =   405
            End
            Begin VB.TextBox txtSQL 
               Height          =   1185
               Left            =   30
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   93
               Top             =   5280
               Width           =   8895
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Build SQL"
               Height          =   525
               Left            =   8970
               TabIndex        =   91
               Top             =   3120
               Width           =   495
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   1305
               Left            =   60
               TabIndex        =   92
               Top             =   6510
               Width           =   8925
               _ExtentX        =   15743
               _ExtentY        =   2302
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
            Begin C1Query80Ctl.C1QueryFrame C1QueryFrame2 
               Height          =   2115
               Left            =   4500
               TabIndex        =   99
               Top             =   3120
               Width           =   4395
               _cx             =   7752
               _cy             =   3731
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
            Begin C1Query80Ctl.C1QueryFrame C1QueryFrame1 
               Height          =   2115
               Left            =   30
               TabIndex        =   100
               Top             =   3090
               Width           =   4335
               _cx             =   7646
               _cy             =   3731
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
            Begin C1Query80Ctl.C1Query C1Query1 
               Height          =   540
               Left            =   5100
               TabIndex        =   101
               Top             =   3870
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
               SchemaData      =   "frmChartWiz.frx":1F0AA
            End
            Begin VB.Label Label5 
               Caption         =   "Query fields and conditions (fill after selecting main table):"
               Height          =   255
               Left            =   30
               TabIndex        =   106
               Top             =   2880
               Width           =   4815
            End
            Begin VB.Label Label4 
               Caption         =   "has natural joins with:"
               Height          =   375
               Left            =   4470
               TabIndex        =   105
               Top             =   1410
               Width           =   4575
            End
            Begin VB.Label Label3 
               Caption         =   "Main table:"
               Height          =   255
               Left            =   30
               TabIndex        =   104
               Top             =   1410
               Width           =   1575
            End
            Begin VB.Label Label2 
               Caption         =   $"frmChartWiz.frx":1F0CA
               Height          =   735
               Left            =   270
               TabIndex        =   103
               Top             =   330
               Width           =   8055
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
               TabIndex        =   102
               Top             =   90
               Width           =   3135
            End
         End
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -30
      Top             =   8640
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   3
      Left            =   0
      TabIndex        =   118
      Top             =   0
      Width           =   1305
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   2
      Left            =   0
      TabIndex        =   117
      Top             =   0
      Width           =   1305
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   1
      Left            =   0
      TabIndex        =   116
      Top             =   0
      Width           =   1305
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   0
      Left            =   4320
      TabIndex        =   111
      Top             =   4170
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmChartWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim doc As DOMDocument
Dim viewsNode As MSXML2.IXMLDOMElement
Dim relationsNode As MSXML2.IXMLDOMElement
Dim catalogsNode As IXMLDOMElement
Dim catalogNode As IXMLDOMElement
Dim m_oFRM As New frmChartTool
'Dim RSLocalUserGroups As New ADODB.Recordset
'Dim m_frmSelectUserGroup As frmSelectUserGroup

        
Private Sub AddChildTextNode(parent As MSXML2.IXMLDOMNode, _
                             nodeName As String, _
                             Data As String)
    Dim node1 As IXMLDOMElement
    Set node1 = doc.createElement(nodeName)
    parent.appendChild node1
    Dim node2 As IXMLDOMNode
    Set node2 = doc.createTextNode(Data)
    node1.appendChild node2
End Sub

Public Sub setUserGroupsRS(ByRef PassedRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo setUserGroupsRS_Err
        '</EhHeader>
        
100     'Set RSLocalUserGroups = PassedRS

        '<EhFooter>
        Exit Sub

setUserGroupsRS_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmWizards.setUserGroupsRS " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


'Private Sub LoadUGs()
'
'    '</EhHeader>
'    If Not m_frmSelectUserGroup Is Nothing Then Exit Sub
'    Set m_frmSelectUserGroup = New frmSelectUserGroup
'    Set m_frmSelectUserGroup.dxUG.DataSource = RSLocalUserGroups
'    m_frmSelectUserGroup.dxUG.Columns.RetrieveFields
'    m_frmSelectUserGroup.dxUG.Dataset.Refresh
'    m_frmSelectUserGroup.Show vbModal
'
'    'm_frmFeedsWizard.setUserGroupsRS RSLocalUserGroups
'    'm_frmFeedsWizard.Show vbModal, Me
'    'Unload m_frmFeedsWizard
'    Unload m_frmSelectUserGroup
'    Set m_frmSelectUserGroup = Nothing
'
'End Sub

Private Sub chkShowPreview_Click()
    If chkShowPreview.Value = vbChecked Then
        m_oFRM.Show vbModeless, Me
    Else
        m_oFRM.Hide
    End If
End Sub

Private Sub cmd_Click()
        '<EhHeader>
        On Error GoTo cmd_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog

100     c.Filter = "*.mdb"
102     c.DefaultExt = "*.mdb"
104     c.ShowOpen
    
106     If Not c.FileTitle = "" Then
108         Text1.Text = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & c.fileName & ";Persist Security Info=False"
110         Command1_Click
        End If
    
        '<EhFooter>
        Exit Sub

cmd_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISChartCfreator.frmMain.cmd_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateChartSettings(Optional sPath As String, _
                                Optional sName As String)
        '<EhHeader>
        On Error GoTo CreateChartSettings_Err
        '</EhHeader>
        Dim sConnStr As String
        Dim i As Integer
        Dim udtOASISChartO As OASISChartObj

100     sConnStr = Adodc1.ConnectionString

102     With udtOASISChartO
        
104         If chkEnableChartTBR.Value = vbChecked Then
        
106             .bChartTBR = True
108             ReDim .sChartTools(0)
        
110             For i = chkMainTool.LBound To chkMainTool.UBound

112                 If chkMainTool(i).Value = vbUnchecked Then

114                     Select Case i
                
                            Case Is < 3
                    
116                             .sChartTools(UBound(.sChartTools)) = i

118                         Case Is < 8
120                             .sChartTools(UBound(.sChartTools)) = i + 2

122                         Case Is < 13
124                             .sChartTools(UBound(.sChartTools)) = i + 3

126                         Case Else
128                             .sChartTools(UBound(.sChartTools)) = i + 4
                        End Select

130                     ReDim Preserve .sChartTools(UBound(.sChartTools) + 1)
                    End If
            
                Next
        
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
    
            End If

168         If chkDataEditor.Value = vbChecked Then .bDataEdtr = True

170         If chkSeriesLegend.Value = vbChecked Then .bSeriesLGD = True

172         If chkPointLegend.Value = vbChecked Then .bValueLGD = True

174         If chkAllowData.Value = vbChecked Then .bDataHigls = True
        
176         If chkContextMenus.Value = vbChecked Then .bContextMenu = True
        
178         If chkMenuBar.Value = vbChecked Then .bMenuBar = True

180         If chkAllowEdit.Value = vbChecked Then .bDataEdtrAllowEdit = True
    
182         If chkAllowDrag.Value = vbChecked Then .bDataEdtrAllowDrag = True
    
184         If chkMultipleColors.Value = vbChecked Then .bMultipleColors = True
    
186         If chkPointLabelsGen.Value = vbChecked Then .bPointLabelsGen = True
    
188         If chkScrollable.Value = vbChecked Then .bScrollable = True
    
190         If chkPointLabels.Value = vbChecked Then .bHGlsPointLabel = True
    
192         If chkDimmed.Value = vbChecked Then .bHGlsDimmed = True
    
194         .sConnStr = Text1.Text
196         .sSQL = txtSQL.Text
    
198         If OptLGDPlacement(0).Value Then
200             .iSerLgdAlign = oChartAlignment.CA_Docked_Left '513
            Else
    
202             If OptLGDPlacement(1).Value Then
204                 .iSerLgdAlign = oChartAlignment.CA_Docked_Right ' 515
                Else

206                 If OptLGDPlacement(2).Value Then
208                     .iSerLgdAlign = oChartAlignment.CA_Docked_Top  ' 256
                    Else
210                     .iSerLgdAlign = oChartAlignment.CA_Docked_Bottom ' 258
                    End If
                End If
    
            End If
        
212         If OptValPlacement(0).Value Then
214             .iValLgdAlign = oChartAlignment.CA_Docked_Left '513
            Else
    
216             If OptValPlacement(1).Value Then
218                 .iValLgdAlign = oChartAlignment.CA_Docked_Right ' 515
                Else

220                 If OptValPlacement(2).Value Then
222                     .iValLgdAlign = oChartAlignment.CA_Docked_Top  ' 256
                    Else
224                     .iValLgdAlign = oChartAlignment.CA_Docked_Bottom ' 258
                    End If
                End If
    
            End If

226         If OptDEPlacement(0).Value Then
228             .iDataEdtrAlign = oChartAlignment.CA_Docked_Left '513
            Else
    
230             If OptDEPlacement(1).Value Then
232                 .iDataEdtrAlign = oChartAlignment.CA_Docked_Right ' 515
                Else

234                 If OptDEPlacement(2).Value Then
236                     .iDataEdtrAlign = oChartAlignment.CA_Docked_Top ' 256
                    Else
238                     .iDataEdtrAlign = oChartAlignment.CA_Docked_Bottom ' 258
                    End If
                End If
    
            End If

240         .sXAxis = txtXAxis.Text
242         .sYAxis = txtYAxis.Text
        
244         If IsNumeric(txtYDecPlace.Text) Then
'TODO This Is fucked up
246           '  .YAxisLabelDecimals = txtYDecPlace.Text
            End If
        'TODO This Is Also Fucked up
248       '  .YAxisFormat = ComYLBLFormat.ListIndex
        
250         .iHeight = 5000
252         .iWidth = 6000
254         .iParentHeight = 6500
256         .iParentWidth = 7060

258         With .udtTitle
260             .CT_Text = txtTitle.Text
262             .CT_BackColor = 2147483647
264             .CT_DrawingArea = True
266             .CT_Alignment = CA_StringAlignment_Center
268             .CT_DockArea = CDA_Area_Top
270             .CT_Flags = CF_TitleFlag_DrawingArea
272             .CT_LineAlignment = CA_StringAlignment_Far
274             .CT_LineGap = 2
            
276             With .CT_Font
278                 .CF_Size = lblPrev(0).Font.Size
280                 .CF_Bold = lblPrev(0).Font.Bold
282                 .CF_Name = lblPrev(0).Font.Name
284                 .CF_Italic = lblPrev(0).Font.Italic
286                 .CF_Strikethrough = lblPrev(0).Font.Strikethrough
288                 .CF_Underline = lblPrev(0).Font.Underline
290                 .CF_Weight = lblPrev(0).Font.Weight ' 500
                End With
            End With
   
292         With .udtNotes
294             .CT_Text = txtNotes.Text
296             .CT_BackColor = 2147483647
298             .CT_DrawingArea = True
300             .CT_Alignment = CA_StringAlignment_Far
302             .CT_DockArea = CDA_Area_Bottom
304             .CT_Flags = CF_TitleFlag_DrawingArea
306             .CT_LineAlignment = CA_StringAlignment_Far
308             .CT_LineGap = 2
            
310             With .CT_Font
312                 .CF_Size = lblPrev(1).Font.Size
314                 .CF_Bold = lblPrev(1).Font.Bold
316                 .CF_Name = lblPrev(1).Font.Name
318                 .CF_Italic = lblPrev(1).Font.Italic
320                 .CF_Strikethrough = lblPrev(1).Font.Strikethrough
322                 .CF_Underline = lblPrev(1).Font.Underline
324                 .CF_Weight = lblPrev(1).Font.Weight ' 500
                End With
            End With
   
326         .enmChartType = ComType.ListIndex + 1 ' Chart_Bubble
    
328         .enmScheme = CS_Solid
330         .XAxisAngle = IIf(IsNumeric(txtXAngle.Text), txtXAngle.Text, 0)
332         .XAxisStaggered = IIf(chkStaggered.Value = vbChecked, True, False)
        
334         .iAngleX = .XAxisAngle
336         .iAngleY = 30
338         .enmAxesStyle = CA_FlatFrame
340         .lngBackColor = 14935011

342         If chkBorder.Value = vbChecked Then .bBorder = True
344         .lngBorderColor = 11053224
346         .enmBorderEffect = CBE_Dark

348         If chkCluster.Value = vbChecked Then .bCluster = False
350         If chk3D.Value = vbChecked Then .bChart3D = True
352         If chkCrossHairs.Value = vbChecked Then .bCrossHairs = False
354         .sngCylSides = 0
356         .enmGrid = CG_None
358         .lngInsideColor = 16777215
360         .enmMarkerShape = CMS_Many
362         .iMarkerSize = 12
364         .sngMarkerStep = 3
366         .sngPerspective = 1

368         If chkShowTips.Value = vbChecked Then .bShowTips = True
370         .enmSmoothFlags = CSF_Fill
372         .enmStacked = CS_No
374         .iWallWidth = 4

376         If chkZoom.Value = vbChecked Then .bZoom = False
                
378         If Len(sPath) > 0 Then

                'Binary Template
380             With .udtExports(2)
382                 .bForceKill = True
384                 .enmExportFormat = tplBin
386                 .sFileName = sName
388                 .sPath = sPath
                End With

            End If
    
        End With
    
390     m_oFRM.SetChart udtOASISChartO
    
        '    If Len(txtSQL.Text) > 0 Then
        '        If Len(sPath) > 0 Then
        '            If FileExists(sPath & "\SQL_" & Replace$(sName, ".oct", ".xml")) Then
        '                C1QueryFrame1.SaveToXMLFile sPath & "\SQL_" & Replace$(sName, ".oct", ".xml")
        '
        '            End If
        '        End If
        '    End If
    
392     If chkShowPreview.Value = vbChecked Then
394         m_oFRM.Show vbModeless, Me
        End If

        'C1QueryFrame1.SaveToXML
        '<EhFooter>
        Exit Sub

CreateChartSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmChartWiz.CreateChartSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadChartSettings(udtOASISChartO As OASISChartObj)
        '<EhHeader>
        On Error GoTo LoadChartSettings_Err
        '</EhHeader>
        Dim sConnStr As String
        Dim i As Integer

100     With udtOASISChartO
        
102         If .bChartTBR Then
            
104             chkEnableChartTBR.Value = vbChecked
        
    '            ReDim .sChartTools(0)
    '
    '            For i = chkMainTool.LBound To chkMainTool.UBound
    '
    '                If chkMainTool(i).Value = vbUnchecked Then
    '
    '                    Select Case i
    '
    '                        Case Is < 3
    '
    '                            .sChartTools(UBound(.sChartTools)) = i
    '
    '                        Case Is < 8
    '                            .sChartTools(UBound(.sChartTools)) = i + 2
    '
    '                        Case Is < 13
    '                            .sChartTools(UBound(.sChartTools)) = i + 3
    '
    '                        Case Else
    '                            .sChartTools(UBound(.sChartTools)) = i + 4
    '                    End Select
    '
    '                    ReDim Preserve .sChartTools(UBound(.sChartTools) + 1)
    '                End If
    '
    '            Next
            Else
106             chkEnableChartTBR.Value = vbUnchecked
            End If

108         If .bAnnoTBR Then
            
110             chkEnable.Value = vbChecked
    
    '            ReDim .sAnnoTools(0)
    '
    '            For i = chkAnnoTool.LBound To chkAnnoTool.UBound
    '
    '                If chkAnnoTool(i).Value = vbUnchecked Then
    '
    '                    Select Case i
    '
    '                        Case Is < 9
    '
    '                            .sAnnoTools(UBound(.sAnnoTools)) = i
    '
    '                        Case Is < 12
    '                            .sAnnoTools(UBound(.sAnnoTools)) = i + 2
    '
    '                        Case Is < 15
    '                            .sAnnoTools(UBound(.sAnnoTools)) = i + 3
    '
    '                        Case Is < 18
    '                            .sAnnoTools(UBound(.sAnnoTools)) = i + 4
    '
    '                        Case Is < 20
    '                            .sAnnoTools(UBound(.sAnnoTools)) = i + 5
    '
    '                        Case Else
    '                            .sAnnoTools(UBound(.sAnnoTools)) = i + 6
    '                    End Select
    '
    '                    ReDim Preserve .sAnnoTools(UBound(.sAnnoTools) + 1)
    '                End If
    '
    '            Next
            Else
112             chkEnable.Value = vbUnchecked
            End If

114         chkAllowData.Value = vbUnchecked
116         chkDataEditor.Value = vbUnchecked
118         chkSeriesLegend.Value = vbUnchecked
120         chkPointLegend.Value = vbUnchecked
122         chkContextMenus.Value = vbUnchecked
124         chkMenuBar.Value = vbUnchecked
126         chkAllowEdit.Value = vbUnchecked
128         chkAllowDrag.Value = vbUnchecked
130         chkMultipleColors.Value = vbUnchecked
132         chkPointLabelsGen.Value = vbUnchecked
134         chkScrollable.Value = vbUnchecked
136         chkPointLabels.Value = vbUnchecked
138         chkDimmed.Value = vbUnchecked
        
140         If .bDataEdtr Then chkDataEditor.Value = vbChecked
142         If .bSeriesLGD Then chkSeriesLegend.Value = vbChecked
144         If .bValueLGD Then chkPointLegend.Value = vbChecked
146         If .bDataHigls Then chkAllowData.Value = vbChecked
148         If .bContextMenu Then chkContextMenus.Value = vbChecked
150         If .bMenuBar Then chkMenuBar.Value = vbChecked
152         If .bDataEdtrAllowEdit Then chkAllowEdit.Value = vbChecked
154         If .bDataEdtrAllowDrag Then chkAllowDrag.Value = vbChecked
156         If .bMultipleColors Then chkMultipleColors.Value = vbChecked
158         If .bPointLabelsGen Then chkPointLabelsGen.Value = vbChecked
160         If .bScrollable Then chkScrollable.Value = vbChecked
162         If .bHGlsPointLabel Then chkPointLabels.Value = vbChecked
164         If .bHGlsDimmed Then chkDimmed.Value = vbChecked

    
166         If .iSerLgdAlign = oChartAlignment.CA_Docked_Left Then
168              OptLGDPlacement(0).Value = True
            Else
    
170             If .iSerLgdAlign = oChartAlignment.CA_Docked_Right Then
172                 OptLGDPlacement(1).Value = True
                Else

174                 If .iSerLgdAlign = oChartAlignment.CA_Docked_Top Then
176                     OptLGDPlacement(2).Value = True
                    Else
178                     OptLGDPlacement(3).Value = True
                    End If
                End If
    
            End If

180         If .iDataEdtrAlign = oChartAlignment.CA_Docked_Left Then
182             OptDEPlacement(0).Value = True
            Else
    
184             If .iDataEdtrAlign = oChartAlignment.CA_Docked_Right Then
186                 OptDEPlacement(1).Value = True
                Else

188                 If .iDataEdtrAlign = oChartAlignment.CA_Docked_Top Then
190                     OptDEPlacement(2).Value = True
                    Else
192                     OptDEPlacement(3).Value = True
                    End If
                End If
    
            End If

194         txtXAxis.Text = .sXAxis
196         txtYAxis.Text = .sYAxis
            'TODO This Is Fucked up again!
            'txtYDecPlace.Text = .YAxisLabelDecimals
            'ComYLBLFormat.ListIndex = .YAxisFormat
        
198         .iHeight = 5000
200         .iWidth = 6000
202         .iParentHeight = 6500
204         .iParentWidth = 7060

206         With .udtTitle
208             .CT_Text = txtTitle.Text
210             .CT_BackColor = 2147483647
212             .CT_DrawingArea = True
214             .CT_Alignment = CA_StringAlignment_Center
216             .CT_DockArea = CDA_Area_Top
218             .CT_Flags = CF_TitleFlag_DrawingArea
220             .CT_LineAlignment = CA_StringAlignment_Far
222             .CT_LineGap = 2
            
224             With .CT_Font
226                 .CF_Size = lblPrev(0).Font.Size
228                 .CF_Bold = lblPrev(0).Font.Bold
230                 .CF_Name = lblPrev(0).Font.Name
232                 .CF_Italic = lblPrev(0).Font.Italic
234                 .CF_Strikethrough = lblPrev(0).Font.Strikethrough
236                 .CF_Underline = lblPrev(0).Font.Underline
238                 .CF_Weight = lblPrev(0).Font.Weight ' 500
                End With
            End With
   
240         With .udtNotes
242             .CT_Text = txtNotes.Text
244             .CT_BackColor = 2147483647
246             .CT_DrawingArea = True
248             .CT_Alignment = CA_StringAlignment_Far
250             .CT_DockArea = CDA_Area_Bottom
252             .CT_Flags = CF_TitleFlag_DrawingArea
254             .CT_LineAlignment = CA_StringAlignment_Far
256             .CT_LineGap = 2
            
258             With .CT_Font
260                 .CF_Size = lblPrev(1).Font.Size
262                 .CF_Bold = lblPrev(1).Font.Bold
264                 .CF_Name = lblPrev(1).Font.Name
266                 .CF_Italic = lblPrev(1).Font.Italic
268                 .CF_Strikethrough = lblPrev(1).Font.Strikethrough
270                 .CF_Underline = lblPrev(1).Font.Underline
272                 .CF_Weight = lblPrev(1).Font.Weight ' 500
                End With
            End With
   
274         ComType.ListIndex = .enmChartType - 1 ' Chart_Bubble
    
276         .enmScheme = CS_Solid
278         .iAngleX = 30
280         .iAngleY = 30
282         .enmAxesStyle = CA_FlatFrame
284         .lngBackColor = 14935011
286         .bBorder = True
288         .lngBorderColor = 11053224
290         .enmBorderEffect = CBE_Dark
292         .bCluster = False
294         .bChart3D = False
296         .bCrossHairs = False
298         .sngCylSides = 0
300         .enmGrid = CG_None
302         .lngInsideColor = 16777215
304         .enmMarkerShape = CMS_Many
306         .iMarkerSize = 12
308         .sngMarkerStep = 3
310         .sngPerspective = 1
312         .bShowTips = True
314         .enmSmoothFlags = CSF_Fill
316         .enmStacked = CS_No
318         .iWallWidth = 4
320         .bZoom = False
    
        End With
    
322     m_oFRM.SetChart udtOASISChartO
324     m_oFRM.Show vbModeless, Me

        '<EhFooter>
        Exit Sub

LoadChartSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmChartWiz.LoadChartSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAssign_Click()
    '<EhHeader>
    On Error GoTo cmdAssign_Click_Err
    '</EhHeader>
    CreateChartSettings
    '<EhFooter>
    Exit Sub

cmdAssign_Click_Err:
    MsgBox Err.Description
        
    '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload m_oFRM
    Unload Me
End Sub

Private Sub cmdFontTitle_Click(Index As Integer)
    Dim c As New cCommonDialog

    c.Font = lblPrev(Index).Font
    c.ShowFont
    
     With lblPrev(Index)
        .FontBold = c.FontBold
        .FontItalic = c.FontItalic
        .FontName = c.FontName
        .FontSize = c.FontSize
    End With
  
End Sub

Private Sub cmdOpen_Click()
        Dim c As New cCommonDialog
        
        On Error Resume Next
100     c.DefaultExt = "*.oct"
102     c.DialogTitle = "Open OASIS Chart Template File"
104     c.Filter = "OASIS Chart Templates Files (*.oct;)|*.oct;"
106     c.ShowOpen
    
108     If Not c.fileName = "" Then
110         LoadChartSettings m_oFRM.OpenChartTemplate(c.FileTitle)
        End If

End Sub

Private Sub cmdSave_Click()
        Dim c As New cCommonDialog
        
        On Error Resume Next
100     c.DefaultExt = "*.oct"
102     c.DialogTitle = "Save OASIS Chart Template File"
104     c.Filter = "OASIS Chart Templates Files (*.oct;*.xml)|*.oct;*.xml"
106     c.ShowSave
    
108     If Not c.fileName = "" Then
110         CreateChartSettings Replace(c.fileName, c.FileTitle, ""), Left(c.FileTitle, Len(c.FileTitle) - 4)
        End If

        If Len(txtSQL.Text) > 0 Then
            If Len(c.fileName) > 0 Then
               ' If FileExists(c.File & "\SQL_" & Replace$(c.FileTitle, ".oct", ".xml")) Then
                    C1QueryFrame1.SaveToXMLFile Replace$(c.fileName, c.FileTitle, "") & "SQL_" & Replace$(c.FileTitle, ".oct", ".xml")

                'End If
            End If
        End If

End Sub

Private Function FileExists(sFullPath As String) As Boolean
        '<EhHeader>
        On Error GoTo FileExists_Err
        '</EhHeader>

        Dim oFile As New Scripting.FileSystemObject
100     FileExists = oFile.FileExists(sFullPath)

        '<EhFooter>
        Exit Function

FileExists_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISChartOCTfiles.FileExists " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub Command1_Click()
        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>
   
        Dim Cnxn As ADODB.Connection
        Dim rstSchema As ADODB.Recordset
        Dim strCnxn As String
        Dim schemaNode As IXMLDOMElement
        Dim rootFolderNode As IXMLDOMElement
        Dim foldersNode As IXMLDOMElement
        Dim viewNode As IXMLDOMElement
        Dim folderNode As IXMLDOMElement
        Dim TableName As String
        Dim viewFieldNode As IXMLDOMElement
        Dim tableInfosNode As IXMLDOMElement
        Dim tableInfoNode As IXMLDOMElement
        Dim viewFieldsNode As IXMLDOMElement
        Dim folderFieldsNode As IXMLDOMElement
        Dim rstColumns As ADODB.Recordset
        Dim folderFieldNode As IXMLDOMElement
        Dim leng As String

        If Len(Text1.Text) < 10 Then
            MsgBox "Could not find a valid connection string... Open a correct Access mdb and try again."
            Exit Sub
        End If
   
100     Adodc1.ConnectionString = Text1.Text
      
102     Set Cnxn = New ADODB.Connection
104     strCnxn = Adodc1.ConnectionString
106     Cnxn.open strCnxn
    
        ' Open database schema
108     Set rstSchema = Cnxn.OpenSchema(adSchemaTables)
   
110     Set doc = New DOMDocument
   
112     Set schemaNode = doc.createElement("SCHEMA")

114     doc.appendChild schemaNode
   
116     Set rootFolderNode = doc.createElement("FOLDER")
118     schemaNode.appendChild rootFolderNode
        
120     Set foldersNode = doc.createElement("SUBFOLDERS")
122     rootFolderNode.appendChild foldersNode
  
124     Set viewsNode = doc.createElement("VIEWS")
126     schemaNode.appendChild viewsNode
   
128     Set relationsNode = doc.createElement("JOINRELATIONS")
130     schemaNode.appendChild relationsNode
   
132     Set catalogsNode = doc.createElement("CATALOGS")
134     schemaNode.appendChild catalogsNode
   
136     AddChildTextNode schemaNode, "IMPORTEDFROM", Adodc1.ConnectionString
138     AddChildTextNode schemaNode, "CP", "0"
   
        ' creating views and folders
140     Do Until rstSchema.EOF

142         If (rstSchema!TABLE_TYPE = "TABLE") Then
144             Set folderNode = doc.createElement("FOLDER")
146             foldersNode.appendChild folderNode
148             Set viewNode = doc.createElement("VIEW")
150             TableName = rstSchema!TABLE_NAME
152             AddChildTextNode viewNode, "VIEWNAME", TableName
154             AddChildTextNode folderNode, "FOLDERNAME", TableName
156             viewsNode.appendChild viewNode
158             List1.AddItem TableName
                ' table info in a view
160             Set tableInfosNode = doc.createElement("VIEWTABLES")
162             viewNode.appendChild tableInfosNode
164             Set tableInfoNode = doc.createElement("VIEWTABLE")
166             tableInfosNode.appendChild tableInfoNode
168             AddChildTextNode tableInfoNode, "TABLENAME", TableName
170             AddChildTextNode tableInfoNode, "REQUIRED", "1"
                ' for each table column create a view field and a folder field
172             Set viewFieldsNode = doc.createElement("VIEWFIELDS")
174             viewNode.appendChild viewFieldsNode
176             Set folderFieldsNode = doc.createElement("FOLDERFIELDS")
178             folderNode.appendChild folderFieldsNode
                
180             Set rstColumns = Cnxn.OpenSchema(adSchemaColumns, Array(Empty, Empty, TableName))

182             Do Until rstColumns.EOF
184                 Set viewFieldNode = doc.createElement("VIEWFIELD")
186                 viewFieldsNode.appendChild viewFieldNode
188                 Set folderFieldNode = doc.createElement("FOLDERFIELD")
190                 folderFieldsNode.appendChild folderFieldNode
192                 AddChildTextNode viewFieldNode, "FIELDNAME", rstColumns!COLUMN_NAME
194                 AddChildTextNode viewFieldNode, "FIELDTABLENAME", TableName
196                 AddChildTextNode viewFieldNode, "FIELDTYPE", rstColumns!DATA_TYPE
198                 AddChildTextNode folderFieldNode, "NAME", rstColumns!COLUMN_NAME
200                 AddChildTextNode folderFieldNode, "VIEW", TableName
202                 AddChildTextNode folderFieldNode, "TABLENAME", TableName
204                 leng = "0"

206                 If rstColumns!CHARACTER_MAXIMUM_LENGTH <> Empty Then
208                     leng = rstColumns!CHARACTER_MAXIMUM_LENGTH
                    End If

210                 AddChildTextNode viewFieldNode, "FIELDSIZE", leng
212                 leng = "0"

214                 If rstColumns!NUMERIC_PRECISION <> Empty Then
216                     leng = rstColumns!NUMERIC_PRECISION
                    End If

218                 AddChildTextNode viewFieldNode, "FIELDPREC", leng
220                 leng = "0"

222                 If rstColumns!NUMERIC_SCALE <> Empty Then
224                     leng = rstColumns!NUMERIC_SCALE
                    End If

226                 AddChildTextNode viewFieldNode, "FIELDSCALE", leng
228                 rstColumns.MoveNext
                Loop

            End If

230         rstSchema.MoveNext
        Loop

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISChartCfreator.frmMain.Command1_Click " & "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub Command2_Click()
        '<EhHeader>
        On Error GoTo Command2_Click_Err
        '</EhHeader>
        Dim RS As New ADODB.Recordset
        Dim cn As New ADODB.Connection

100     If C1Query1.BuildSQL = True Then
102         txtSQL.Text = C1Query1.SQL
104         cn.open Adodc1.ConnectionString
106         RS.CursorLocation = adUseClient
108         RS.open txtSQL.Text, cn, adOpenDynamic, adLockBatchOptimistic
110         Set DataGrid1.DataSource = RS
        
        End If

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        MsgBox Err.Description
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     Text1.Text = Adodc1.ConnectionString '
        ComType.ListIndex = 0
        ComYLBLFormat.ListIndex = 0

        
        'LoadUGs
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISChartCfreator.frmMain.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub List1_Click()
        Dim view1 As String
        Dim viewFieldsNode As IXMLDOMElement
        Dim tableCandidate As String
        Dim fieldsNode As IXMLDOMNode
        Dim added As Boolean
        Dim fieldCandidateName As String
        Dim relationNode As IXMLDOMElement
        Dim viewLinksNode As IXMLDOMElement
        Dim relateViewsNode As IXMLDOMElement
        Dim viewLinkNode As IXMLDOMNode
        Dim joinsNode As IXMLDOMElement
        Dim joinNode As IXMLDOMElement
        Dim candidate As String
        Dim found As Boolean
        Dim newCatNode As IXMLDOMNode
        Dim i As Integer
        Dim j As Integer
        Dim K As Integer
        
        '<EhHeader>
        On Error GoTo List1_Click_Err
        '</EhHeader>

100     If List1.Text = Empty Or List1.Text = "" Then
            Exit Sub
        End If

102     List2.Clear
104     While relationsNode.childNodes.Length > 0
106         relationsNode.removeChild relationsNode.childNodes(0)
        Wend
108     While catalogsNode.childNodes.Length > 0
110         catalogsNode.removeChild catalogsNode.childNodes(0)
        Wend
        ' Add 1st view group and main view name (first item in the view group)
112     Set catalogNode = doc.createElement("CATALOG")
114     catalogsNode.appendChild catalogNode
116     AddChildTextNode catalogNode, "VIEW", List1.Text

        ' Add join relationships and catalogs
118     For i = 0 To viewsNode.childNodes.Length - 1

120         If viewsNode.childNodes(i).childNodes(0).nodeTypedValue = List1.Text Then
122             view1 = List1.Text
124             Set viewFieldsNode = viewsNode.childNodes(i).childNodes(2)
                Exit For
            End If

        Next

126     For i = 0 To viewsNode.childNodes.Length - 1

128         tableCandidate = viewsNode.childNodes(i).childNodes(0).nodeTypedValue

130         If tableCandidate <> List1.Text Then
132             Set fieldsNode = viewsNode.childNodes(i).childNodes(2) ' 2nd node is VIEWFIELDS

134             For j = 0 To fieldsNode.childNodes.Length - 1
136                 added = False
138                 fieldCandidateName = fieldsNode.childNodes(j).childNodes(0).nodeTypedValue

140                 For K = 0 To viewFieldsNode.childNodes.Length - 1

142                     If viewFieldsNode.childNodes(K).childNodes(0).nodeTypedValue = fieldCandidateName Then
144                         List2.AddItem tableCandidate & " (using " & fieldCandidateName & " field)"
                            ' Add relationship
146                         Set relationNode = doc.createElement("JOINRELATION")
148                         relationsNode.appendChild relationNode
150                         Set relateViewsNode = doc.createElement("RELATEVIEWS")
152                         AddChildTextNode relateViewsNode, "VIEWNAME", List1.Text
154                         AddChildTextNode relateViewsNode, "VIEWNAME", tableCandidate
156                         relationNode.appendChild relateViewsNode
158                         Set viewLinksNode = doc.createElement("VIEWLINKS")
160                         relationNode.appendChild viewLinksNode
162                         Set viewLinkNode = doc.createElement("VIEWLINK")
164                         viewLinksNode.appendChild viewLinkNode
166                         AddChildTextNode viewLinkNode, "VIEW1", List1.Text
168                         AddChildTextNode viewLinkNode, "JOINTYPE", "0"
170                         AddChildTextNode viewLinkNode, "VIEW2", tableCandidate
                            ' Joins
172                         Set joinsNode = doc.createElement("JOINS")
174                         relationNode.appendChild joinsNode
176                         Set joinNode = doc.createElement("JOIN")
178                         joinsNode.appendChild joinNode
180                         AddChildTextNode joinNode, "FIELD1", fieldCandidateName
182                         AddChildTextNode joinNode, "TABLE1", List1.Text
184                         AddChildTextNode joinNode, "VIEW1", List1.Text
186                         AddChildTextNode joinNode, "OPERATOR", "0"  '' is equal to
188                         AddChildTextNode joinNode, "FIELD2", fieldCandidateName
190                         AddChildTextNode joinNode, "TABLE2", tableCandidate
192                         AddChildTextNode joinNode, "VIEW2", tableCandidate
                            ' Add this table to the view group
194                         AddChildTextNode catalogNode, "VIEW", tableCandidate
196                         added = True
                            Exit For
                        End If

                    Next

198                 If added = True Then
                        Exit For
                    End If

                Next

            End If

        Next
  
        ' Create the remaining (2nd, etc) view groups (catalogs).
200     For i = 0 To viewsNode.childNodes.Length - 1
202         candidate = viewsNode.childNodes(i).childNodes(0).nodeTypedValue
204         found = False

206         For j = 0 To catalogNode.childNodes.Length - 1

208             If catalogNode.childNodes(j).nodeTypedValue = candidate Then
210                 found = True
                    Exit For
                End If

            Next

212         If Not found Then
214             Set newCatNode = doc.createElement("CATALOG")
216             catalogsNode.appendChild newCatNode
218             AddChildTextNode newCatNode, "VIEW", candidate
            End If

        Next
 
        ' Assign the schema to the control
220     C1Query1.schema = doc.xml

        ' Initialize the C1QueryFrame controls
222     C1QueryFrame1.Clear
224     C1QueryFrame1.CurrentItemID = 0
226     C1QueryFrame1.Render
228     C1QueryFrame2.Clear
230     C1QueryFrame2.CurrentItemID = 0
232     C1QueryFrame2.Render
        '<EhFooter>
        Exit Sub

List1_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISChartCfreator.frmMain.List1_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

