VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{605925BE-4799-4093-A2E7-39323147E70E}#1.0#0"; "C1Query8.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmChartProviderSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Security Charts"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9555
   Icon            =   "frmChartProviderSetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   7755
      Left            =   -30
      TabIndex        =   5
      Top             =   0
      Width           =   9615
      _cx             =   16960
      _cy             =   13679
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
      Caption         =   "Templates|Template Designer"
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7380
         Left            =   -10170
         TabIndex        =   79
         Top             =   330
         Width           =   9525
         Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
            Height          =   3945
            Left            =   90
            OleObjectBlob   =   "frmChartProviderSetting.frx":6852
            TabIndex        =   88
            Top             =   3180
            Width           =   9315
         End
         Begin VB.ListBox listTemplates 
            Height          =   2790
            Left            =   90
            TabIndex        =   86
            Top             =   300
            Width           =   7395
         End
         Begin VB.CommandButton cmdGenerateTemplate 
            Caption         =   "Generate Chart"
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
            Left            =   7590
            TabIndex        =   85
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Frame frmTemplateDates 
            Height          =   2115
            Left            =   7590
            TabIndex        =   80
            Top             =   240
            Width           =   1815
            Begin VB.CheckBox chkDateFilter 
               Caption         =   "Use date filter"
               Height          =   315
               Left            =   180
               TabIndex        =   89
               Top             =   180
               Width           =   1485
            End
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
               OleObjectBlob   =   "frmChartProviderSetting.frx":74FA
               TabIndex        =   81
               Top             =   960
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
               OleObjectBlob   =   "frmChartProviderSetting.frx":759A
               TabIndex        =   82
               Top             =   1620
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.Label lblDateTill 
               Caption         =   "Date Until:"
               Height          =   255
               Left            =   120
               TabIndex        =   84
               Top             =   1380
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblDateFrom 
               Caption         =   "Date From:"
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   720
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin VB.Label lblAvailableTemplates 
            Caption         =   "Available Templates:"
            Height          =   255
            Left            =   90
            TabIndex        =   87
            Top             =   60
            Width           =   2535
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7380
         Left            =   45
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   330
         Width           =   9525
         _cx             =   16801
         _cy             =   13018
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
         Begin C1SizerLibCtl.C1Elastic C1EMain 
            Height          =   7575
            Left            =   -510
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   -270
            Width           =   10125
            _cx             =   17859
            _cy             =   13361
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
               Height          =   7395
               Left            =   90
               TabIndex        =   8
               Top             =   90
               Width           =   9945
               _cx             =   17542
               _cy             =   13044
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
                  Height          =   7365
                  Left            =   15
                  TabIndex        =   9
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   9915
                  _cx             =   17489
                  _cy             =   12991
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
                  Begin VB.CheckBox chkUseCustom 
                     Caption         =   "Store custom SQL changes when saving"
                     Height          =   375
                     Left            =   600
                     TabIndex        =   90
                     Top             =   7020
                     Width           =   3315
                  End
                  Begin VB.CommandButton cmdAssign 
                     Caption         =   "View"
                     Height          =   315
                     Left            =   8400
                     TabIndex        =   10
                     Top             =   7050
                     Width           =   1425
                  End
                  Begin VB.CommandButton cmdSave 
                     Caption         =   "Save Template"
                     Height          =   315
                     Left            =   6900
                     TabIndex        =   66
                     Top             =   7050
                     Width           =   1425
                  End
                  Begin VB.Frame FraChartSettings 
                     Caption         =   "Chart Settings:"
                     Height          =   3300
                     Left            =   450
                     TabIndex        =   21
                     Top             =   150
                     Width           =   9495
                     Begin VB.Frame FraXAxis 
                        Caption         =   "X Axis:"
                        Height          =   945
                        Left            =   7920
                        TabIndex        =   58
                        Top             =   1590
                        Width           =   1485
                        Begin VB.CheckBox chkStaggered 
                           Caption         =   "Staggered"
                           Height          =   285
                           Left            =   60
                           TabIndex        =   60
                           Top             =   570
                           Width           =   1395
                        End
                        Begin VB.TextBox txtXAngle 
                           Height          =   285
                           Left            =   630
                           TabIndex        =   59
                           Text            =   "45"
                           Top             =   210
                           Width           =   795
                        End
                        Begin VB.Label lblAngle 
                           AutoSize        =   -1  'True
                           Caption         =   "Angle:"
                           Height          =   195
                           Left            =   60
                           TabIndex        =   61
                           Top             =   210
                           Width           =   450
                        End
                     End
                     Begin VB.CommandButton cmdFontTitle 
                        Caption         =   "..."
                        Height          =   315
                        Index           =   3
                        Left            =   9150
                        TabIndex        =   57
                        Top             =   570
                        Width           =   255
                     End
                     Begin VB.CommandButton cmdFontTitle 
                        Caption         =   "..."
                        Height          =   315
                        Index           =   2
                        Left            =   9150
                        TabIndex        =   56
                        Top             =   240
                        Width           =   255
                     End
                     Begin VB.CommandButton cmdFontTitle 
                        Caption         =   "..."
                        Height          =   315
                        Index           =   1
                        Left            =   4980
                        TabIndex        =   55
                        Top             =   540
                        Width           =   255
                     End
                     Begin VB.CommandButton cmdFontTitle 
                        Caption         =   "..."
                        Height          =   315
                        Index           =   0
                        Left            =   4980
                        TabIndex        =   54
                        Top             =   210
                        Width           =   255
                     End
                     Begin VB.Frame FraChartTools 
                        Caption         =   "Chart:"
                        Height          =   645
                        Left            =   90
                        TabIndex        =   48
                        Top             =   930
                        Width           =   9315
                        Begin VB.ComboBox ComType 
                           Height          =   315
                           ItemData        =   "frmChartProviderSetting.frx":763A
                           Left            =   480
                           List            =   "frmChartProviderSetting.frx":7685
                           Style           =   2  'Dropdown List
                           TabIndex        =   52
                           Top             =   180
                           Width           =   1815
                        End
                        Begin VB.CheckBox chkEnable 
                           Caption         =   "Enable Annotations"
                           Height          =   285
                           Left            =   3720
                           TabIndex        =   51
                           Top             =   240
                           Width           =   1695
                        End
                        Begin VB.CheckBox chk3D 
                           Caption         =   "Use 3D"
                           Height          =   255
                           Left            =   5490
                           TabIndex        =   50
                           Top             =   240
                           Width           =   945
                        End
                        Begin VB.CheckBox chkEnableChartTBR 
                           Caption         =   "Enable Tools"
                           Height          =   285
                           Left            =   2400
                           TabIndex        =   49
                           Top             =   240
                           Width           =   1275
                        End
                        Begin VB.Label lblNoys 
                           AutoSize        =   -1  'True
                           Caption         =   "Type:"
                           Height          =   195
                           Index           =   9
                           Left            =   30
                           TabIndex        =   53
                           Top             =   270
                           Width           =   405
                        End
                     End
                     Begin VB.Frame FraLegend 
                        Caption         =   "Legend:"
                        Height          =   945
                        Left            =   90
                        TabIndex        =   41
                        Top             =   1590
                        Width           =   3075
                        Begin VB.CheckBox chkSeriesLegend 
                           Caption         =   "Series"
                           Height          =   285
                           Left            =   30
                           TabIndex        =   45
                           Top             =   240
                           Width           =   765
                        End
                        Begin VB.CheckBox chkPointLegend 
                           Caption         =   "Value"
                           Height          =   285
                           Left            =   30
                           TabIndex        =   44
                           Top             =   570
                           Width           =   735
                        End
                        Begin VB.ComboBox ComAlignment 
                           Height          =   315
                           Index           =   0
                           ItemData        =   "frmChartProviderSetting.frx":7733
                           Left            =   1710
                           List            =   "frmChartProviderSetting.frx":7743
                           Style           =   2  'Dropdown List
                           TabIndex        =   43
                           Top             =   180
                           Width           =   1305
                        End
                        Begin VB.ComboBox ComAlignment 
                           Height          =   315
                           Index           =   1
                           ItemData        =   "frmChartProviderSetting.frx":7761
                           Left            =   1710
                           List            =   "frmChartProviderSetting.frx":7771
                           Style           =   2  'Dropdown List
                           TabIndex        =   42
                           Top             =   510
                           Width           =   1305
                        End
                        Begin VB.Label lblNoys 
                           AutoSize        =   -1  'True
                           Caption         =   "Alignment:"
                           Height          =   195
                           Index           =   2
                           Left            =   810
                           TabIndex        =   47
                           Top             =   270
                           Width           =   735
                        End
                        Begin VB.Label lblNoys 
                           AutoSize        =   -1  'True
                           Caption         =   "Alignment:"
                           Height          =   195
                           Index           =   3
                           Left            =   810
                           TabIndex        =   46
                           Top             =   600
                           Width           =   735
                        End
                     End
                     Begin VB.Frame FraDataEditor 
                        Caption         =   "Data Editor:"
                        Height          =   945
                        Left            =   3180
                        TabIndex        =   37
                        Top             =   1590
                        Width           =   2325
                        Begin VB.CheckBox chkDataEditor 
                           Caption         =   "Show"
                           Height          =   255
                           Left            =   90
                           TabIndex        =   39
                           Top             =   240
                           Width           =   765
                        End
                        Begin VB.ComboBox ComAlignment 
                           Height          =   315
                           Index           =   2
                           ItemData        =   "frmChartProviderSetting.frx":778F
                           Left            =   930
                           List            =   "frmChartProviderSetting.frx":779F
                           Style           =   2  'Dropdown List
                           TabIndex        =   38
                           Top             =   510
                           Width           =   1305
                        End
                        Begin VB.Label lblNoys 
                           AutoSize        =   -1  'True
                           Caption         =   "Alignment:"
                           Height          =   195
                           Index           =   4
                           Left            =   60
                           TabIndex        =   40
                           Top             =   570
                           Width           =   735
                        End
                     End
                     Begin VB.Frame FraDataHighlighting 
                        Caption         =   "Data Highlighting:"
                        Height          =   945
                        Left            =   5520
                        TabIndex        =   33
                        Top             =   1590
                        Width           =   2385
                        Begin VB.CheckBox chkAllowData 
                           Caption         =   "Allow"
                           Height          =   255
                           Left            =   90
                           TabIndex        =   36
                           Top             =   210
                           Width           =   765
                        End
                        Begin VB.CheckBox chkDimmed 
                           Caption         =   "Dimmed"
                           Height          =   255
                           Left            =   90
                           TabIndex        =   35
                           Top             =   510
                           Width           =   945
                        End
                        Begin VB.CheckBox chkPointLabels 
                           Caption         =   "Point Labels"
                           Height          =   255
                           Left            =   1110
                           TabIndex        =   34
                           Top             =   210
                           Width           =   1215
                        End
                     End
                     Begin VB.TextBox txtTitle 
                        Height          =   285
                        Left            =   720
                        TabIndex        =   32
                        Top             =   240
                        Width           =   4215
                     End
                     Begin VB.TextBox txtNotes 
                        Height          =   285
                        Left            =   720
                        TabIndex        =   31
                        Top             =   570
                        Width           =   4215
                     End
                     Begin VB.TextBox txtXAxis 
                        Height          =   285
                        Left            =   5880
                        TabIndex        =   30
                        Top             =   240
                        Width           =   3255
                     End
                     Begin VB.TextBox txtYAxis 
                        Height          =   285
                        Left            =   5880
                        TabIndex        =   29
                        Top             =   600
                        Width           =   3255
                     End
                     Begin VB.Frame FraMiscellenous 
                        Caption         =   "Miscellenous:"
                        Height          =   675
                        Left            =   90
                        TabIndex        =   22
                        Top             =   2520
                        Width           =   9315
                        Begin VB.CheckBox chkBorder 
                           Caption         =   "Border"
                           Height          =   255
                           Left            =   3720
                           TabIndex        =   28
                           Top             =   300
                           Width           =   825
                        End
                        Begin VB.CheckBox chkCluster 
                           Caption         =   "Cluster"
                           Height          =   255
                           Left            =   4590
                           TabIndex        =   27
                           Top             =   300
                           Width           =   825
                        End
                        Begin VB.CheckBox chkShowTips 
                           Caption         =   "Show tips"
                           Height          =   255
                           Left            =   1260
                           TabIndex        =   26
                           Top             =   300
                           Width           =   1155
                        End
                        Begin VB.CheckBox chkCrossHairs 
                           Caption         =   "Cross Hairs"
                           Height          =   255
                           Left            =   120
                           TabIndex        =   25
                           Top             =   300
                           Width           =   1155
                        End
                        Begin VB.CheckBox chkMultipleColors 
                           Caption         =   "Multiple Colors"
                           Height          =   255
                           Left            =   5460
                           TabIndex        =   24
                           Top             =   300
                           Width           =   1335
                        End
                        Begin VB.CheckBox chkPointLabelsGen 
                           Caption         =   "Point Labels"
                           Height          =   255
                           Left            =   2430
                           TabIndex        =   23
                           Top             =   300
                           Width           =   1245
                        End
                     End
                     Begin VB.Label lblNoys 
                        AutoSize        =   -1  'True
                        Caption         =   "Title:"
                        Height          =   195
                        Index           =   5
                        Left            =   150
                        TabIndex        =   65
                        Top             =   300
                        Width           =   345
                     End
                     Begin VB.Label lblNoys 
                        AutoSize        =   -1  'True
                        Caption         =   "Notes:"
                        Height          =   195
                        Index           =   6
                        Left            =   180
                        TabIndex        =   64
                        Top             =   600
                        Width           =   465
                     End
                     Begin VB.Label lblNoys 
                        AutoSize        =   -1  'True
                        Caption         =   "X Axis:"
                        Height          =   195
                        Index           =   7
                        Left            =   5340
                        TabIndex        =   63
                        Top             =   270
                        Width           =   480
                     End
                     Begin VB.Label lblNoys 
                        AutoSize        =   -1  'True
                        Caption         =   "Y Axis:"
                        Height          =   195
                        Index           =   8
                        Left            =   5370
                        TabIndex        =   62
                        Top             =   600
                        Width           =   480
                     End
                  End
                  Begin VB.Frame Frame3 
                     Caption         =   "Data Source"
                     Height          =   3585
                     Left            =   420
                     TabIndex        =   11
                     Top             =   3450
                     Width           =   9495
                     Begin VB.CommandButton cmdSwapTable 
                        Caption         =   "..."
                        Height          =   225
                        Left            =   8940
                        TabIndex        =   12
                        Top             =   120
                        Visible         =   0   'False
                        Width           =   465
                     End
                     Begin MSDataGridLib.DataGrid DataGrid1 
                        Height          =   2985
                        Left            =   4830
                        TabIndex        =   13
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
                        TabIndex        =   14
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
                           TabIndex        =   15
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
                              TabIndex        =   16
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
                           TabIndex        =   17
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
                              TabIndex        =   18
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
                           TabIndex        =   19
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
                              TabIndex        =   20
                              Top             =   30
                              Width           =   4365
                           End
                        End
                     End
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   7365
                  Left            =   10560
                  TabIndex        =   67
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   9915
                  _cx             =   17489
                  _cy             =   12991
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
               Begin C1SizerLibCtl.C1Elastic C1ElasticDude 
                  Height          =   7365
                  Left            =   10860
                  TabIndex        =   68
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   9915
                  _cx             =   17489
                  _cy             =   12991
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
                     TabIndex        =   73
                     Top             =   990
                     Width           =   4335
                  End
                  Begin VB.ListBox List2 
                     Height          =   840
                     Left            =   4470
                     TabIndex        =   72
                     Top             =   1650
                     Width           =   4335
                  End
                  Begin VB.ListBox List1 
                     Height          =   840
                     Left            =   30
                     TabIndex        =   71
                     Top             =   1650
                     Width           =   4335
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "Load database schema from this database:"
                     Height          =   405
                     Left            =   30
                     TabIndex        =   70
                     Top             =   960
                     Width           =   4335
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "..."
                     Height          =   375
                     Left            =   8850
                     TabIndex        =   69
                     Top             =   990
                     Width           =   405
                  End
                  Begin C1Query80Ctl.C1Query C1Query1 
                     Height          =   540
                     Left            =   10
                     TabIndex        =   74
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
                     SchemaData      =   "frmChartProviderSetting.frx":77BD
                  End
                  Begin VB.Label Label4 
                     Caption         =   "has natural joins with:"
                     Height          =   375
                     Left            =   4470
                     TabIndex        =   78
                     Top             =   1410
                     Width           =   4575
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Main table:"
                     Height          =   255
                     Left            =   30
                     TabIndex        =   77
                     Top             =   1410
                     Width           =   1575
                  End
                  Begin VB.Label Label2 
                     Caption         =   $"frmChartProviderSetting.frx":77DD
                     Height          =   735
                     Left            =   270
                     TabIndex        =   76
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
                     TabIndex        =   75
                     Top             =   90
                     Width           =   3135
                  End
               End
            End
         End
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   0
      Top             =   7470
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   714
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin C1Query80Ctl.C1Query C1Query2 
      Height          =   540
      Left            =   30
      TabIndex        =   4
      Top             =   7470
      Width           =   540
      _cx             =   952
      _cy             =   952
      DesignTemplates =   ""
      MainViewName    =   ""
      DataMember      =   ""
      FilterMode      =   0   'False
      ApplyExtensions =   2
      NameSubstitute  =   ""
      SaveSchemaAsString=   0   'False
      PathSeparator   =   "."
      SchemaData      =   "frmChartProviderSetting.frx":78CC
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblPrev 
      Caption         =   "Prev"
      Height          =   525
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "frmChartProviderSettings"
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
Dim RSLocalUserGroups As New ADODB.Recordset

Private XMLValues() As String
Private XMLValuesQRY() As String
Private XMLValuesC As Integer
Private XMLValuesQC As Integer
Dim udtOASISChartO As OASISChartObj
Dim m_frmOASISCharts As frmOASISCharts

Private Sub XMLValuesQRYClear()
        '<EhHeader>
        On Error GoTo XMLValuesQRYClear_Err
        '</EhHeader>
100     XMLValuesQC = 0
102     ReDim XMLValuesQRY(0) As String
        '<EhFooter>
        Exit Sub

XMLValuesQRYClear_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.XMLValuesQRYClear " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub XMLValuesAddQRY(XMLValue As String)
        '<EhHeader>
        On Error GoTo XMLValuesAddQRY_Err
        '</EhHeader>
100     XMLValuesQRY(XMLValuesQC) = XMLValue
102     XMLValuesQC = XMLValuesQC + 1
104     ReDim Preserve XMLValuesQRY(0 To XMLValuesQC) As String
        '<EhFooter>
        Exit Sub

XMLValuesAddQRY_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.XMLValuesAddQRY " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub XMLValuesClear()
        '<EhHeader>
        On Error GoTo XMLValuesClear_Err
        '</EhHeader>
100     XMLValuesC = 0
102     ReDim XMLValues(0) As String
        '<EhFooter>
        Exit Sub

XMLValuesClear_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.XMLValuesClear " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub XMLValuesAdd(XMLValue As String)
        '<EhHeader>
        On Error GoTo XMLValuesAdd_Err
        '</EhHeader>
100     XMLValues(XMLValuesC) = XMLValue
102     XMLValuesC = XMLValuesC + 1
104     ReDim Preserve XMLValues(0 To XMLValuesC) As String
        '<EhFooter>
        Exit Sub

XMLValuesAdd_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.XMLValuesAdd " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function LoadQuery(sFile As String) As Boolean
        '<EhHeader>
        On Error GoTo LoadQuery_Err
        '</EhHeader>
    
        Dim xmlDoc, Node
        Dim i As Integer
        Dim fs As New FileSystemObject

100     If fs.FileExists(g_sAppPath & "\data\templates\SecurityChartTemplates\" & sFile & "_query.xml") Then

102         XMLValuesClear

104         Set xmlDoc = CreateObject("Msxml.DOMDocument")
106         xmlDoc.async = False
108         xmlDoc.Load (g_sAppPath & "\data\templates\SecurityChartTemplates\" & sFile & "_query.xml")

110         If Not xmlDoc.documentElement Is Nothing Then

112             For i = 1 To xmlDoc.documentElement.childNodes.Length - 1
114                 Set Node = xmlDoc.documentElement.childNodes.Item(i)
116                 XMLValuesAdd (Node.childNodes.Item(1).Text)
                Next

118             XMLValuesQRYClear

120             Set xmlDoc = CreateObject("Msxml.DOMDocument")
122             xmlDoc.async = False
124             xmlDoc.Load (g_sAppPath & "\data\templates\SecurityChartTemplates\" & sFile & "_constraints.xml")

126             If Not xmlDoc.documentElement Is Nothing Then

128                 For i = 1 To xmlDoc.documentElement.childNodes.Length - 1
130                     Set Node = xmlDoc.documentElement.childNodes.Item(i)
132                     XMLValuesAddQRY (Node.childNodes.Item(1).Text)
                    Next
            
                End If
            End If
        ElseIf fs.FileExists(g_sAppPath & "\data\templates\SecurityChartTemplates\" & sFile & ".ini") Then
            
            Dim oIni As New clIniReader
            oIni.Path = g_sAppPath & "\data\templates\SecurityChartTemplates\" & sFile & ".ini"
            oIni.Section = "default"
            oIni.Key = "SQL"
            txtSQL.Text = oIni.Value
            LoadQuery = True
        End If

        '<EhFooter>
        Exit Function

LoadQuery_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.LoadQuery " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
        
Private Sub AddChildTextNode(Parent As MSXML2.IXMLDOMNode, _
                             nodeName As String, _
                             Data As String)
        '<EhHeader>
        On Error GoTo AddChildTextNode_Err
        '</EhHeader>
        Dim node1 As IXMLDOMElement
100     Set node1 = doc.createElement(nodeName)
102     Parent.appendChild node1
        Dim node2 As IXMLDOMNode
104     Set node2 = doc.createTextNode(Data)
106     node1.appendChild node2
        '<EhFooter>
        Exit Sub

AddChildTextNode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.AddChildTextNode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub C1Query1_Error(ByVal ErrorNumber As Long, _
                           Description As String, _
                           CancelDisplay As Boolean)
        '<EhHeader>
        On Error GoTo C1Query1_Error_Err
        '</EhHeader>
100     CancelDisplay = True
        '<EhFooter>
        Exit Sub

C1Query1_Error_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.C1Query1_Error " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub C1QueryFrame1_Change(ByVal ItemID As Long, _
                                 ByVal ItemElement As C1Query80Ctl.ElementTypeEnum)
        '<EhHeader>
        On Error GoTo C1QueryFrame1_Change_Err
        '</EhHeader>
    
        On Error Resume Next

100     C1Query1.BuildSQL

102     If Len(C1Query1.SQL) > 0 Then
104         Adodc1.ConnectionString = m_Cnn.ConnectionString
106         Adodc1.CommandType = adCmdText
108         Adodc1.RecordSource = C1Query1.SQL
110         txtSQL.Text = C1Query1.SQL
112         Adodc1.Refresh
114         Set DataGrid1.DataSource = Adodc1
        End If

        '<EhFooter>
        Exit Sub

C1QueryFrame1_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.C1QueryFrame1_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub C1QueryFrame1_Error(ByVal ErrorNumber As Long, _
                                Description As String, _
                                CancelDisplay As Boolean)
        '<EhHeader>
        On Error GoTo C1QueryFrame1_Error_Err
        '</EhHeader>
100     CancelDisplay = True
        '<EhFooter>
        Exit Sub

C1QueryFrame1_Error_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.C1QueryFrame1_Error " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub C1QueryFrame2_Change(ByVal ItemID As Long, _
                                 ByVal ItemElement As C1Query80Ctl.ElementTypeEnum)
        '<EhHeader>
        On Error GoTo C1QueryFrame2_Change_Err
        '</EhHeader>
    
        On Error Resume Next
    
100     C1Query1.BuildSQL

102     If Len(C1Query1.SQL) > 0 Then
104         Adodc1.ConnectionString = m_Cnn.ConnectionString
106         Adodc1.CommandType = adCmdText
108         Adodc1.RecordSource = C1Query1.SQL
110         txtSQL.Text = C1Query1.SQL
112         Adodc1.Refresh
114         Set DataGrid1.DataSource = Adodc1
        End If

        '<EhFooter>
        Exit Sub

C1QueryFrame2_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.C1QueryFrame2_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub C1QueryFrame2_Error(ByVal ErrorNumber As Long, _
                                Description As String, _
                                CancelDisplay As Boolean)
        '<EhHeader>
        On Error GoTo C1QueryFrame2_Error_Err
        '</EhHeader>
100     CancelDisplay = True
        '<EhFooter>
        Exit Sub

C1QueryFrame2_Error_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.C1QueryFrame2_Error " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkDateFilter_Click()
        '<EhHeader>
        On Error GoTo chkDateFilter_Click_Err
        '</EhHeader>

100     If chkDateFilter.Value = vbChecked Then
102         Me.dxDateFromTemplate.Visible = True
104         Me.dxDateTillTemplate.Visible = True
106         Me.lblDateFrom.Visible = True
108         Me.lblDateTill.Visible = True
110         dxDateFromTemplate = Format(Now() - 365, "Medium Date")
112         dxDateTillTemplate = Format(Now(), "Medium Date")
        Else
114         Me.dxDateFromTemplate.Visible = False
116         Me.dxDateTillTemplate.Visible = False
118         Me.lblDateFrom.Visible = False
120         Me.lblDateTill.Visible = False
        End If
        
        '<EhFooter>
        Exit Sub

chkDateFilter_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.chkDateFilter_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkUseCustom_Click()
    DebugPrint ""
    'C1Query1.SQL = txtSQL.Text
End Sub

Private Sub cmdAssign_Click()
        '<EhHeader>
        On Error GoTo cmdAssign_Click_Err
        '</EhHeader>
    
        'On Error Resume Next
    
100     Adodc1.ConnectionString = m_Cnn.ConnectionString
102     Adodc1.CommandType = adCmdText
    
104     If Len(txtSQL.Text) > 0 Then
106         Adodc1.RecordSource = txtSQL.Text
108         Adodc1.Refresh
110         Set DataGrid1.DataSource = Adodc1
        End If
    
112     CreateChartSettings , , Adodc1.Recordset
114     m_frmOASISCharts.Show vbModeless, Me
        '<EhFooter>
        Exit Sub

cmdAssign_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.cmdAssign_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateChartSettings(Optional sPath As String, _
                                Optional sName As String, _
                                Optional oRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo CreateChartSettings_Err
        '</EhHeader>
        
        Dim sConnStr As String
        Dim i As Integer

100     With udtOASISChartO
        
102         If chkEnableChartTBR.Value = vbChecked Then
104             .bChartTBR = True
106             ReDim .sChartTools(2)
108             .sChartTools(0) = 0
110             .sChartTools(1) = 20
            End If
        
112         If chkEnable.Value = vbChecked Then
114             .bAnnoTBR = True 'This does not work in the dll
116             ReDim .sAnnoTools(2)
            End If

118         .bDataEdtr = IIf(chkDataEditor.Value = vbChecked, True, False)
120         .bSeriesLGD = IIf(chkSeriesLegend.Value = vbChecked, True, False)
122         .bValueLGD = IIf(chkPointLegend.Value = vbChecked, True, False)
124         .bDataHigls = IIf(chkAllowData.Value = vbChecked, True, False)
126         .bMultipleColors = IIf(chkMultipleColors.Value = vbChecked, True, False)
128         .bPointLabelsGen = IIf(chkPointLabelsGen.Value = vbChecked, True, False)
130         .bHGlsPointLabel = IIf(chkPointLabels.Value = vbChecked, True, False)
132         .bHGlsDimmed = IIf(chkDimmed.Value = vbChecked, True, False)
134         .bCluster = IIf(chkCluster.Value = vbChecked, True, False)
136         .bChart3D = IIf(chk3D.Value = vbChecked, True, False)
138         .bCrossHairs = IIf(chkCrossHairs.Value = vbChecked, True, False)
140         .bBorder = IIf(chkBorder.Value = vbChecked, True, False)
142         .bShowTips = IIf(chkShowTips.Value = vbChecked, True, False)
144         .XAxisStaggered = IIf(chkStaggered.Value = vbChecked, True, False)
            
146         .bContextMenu = False
148         .bMenuBar = False
150         .bDataEdtrAllowEdit = False
152         .bDataEdtrAllowDrag = False
154         .bScrollable = False
156         .iSerLgdAlign = GetAlignment(ComAlignment(0).ListIndex)
158         .iValLgdAlign = GetAlignment(ComAlignment(1).ListIndex)
160         .iDataEdtrAlign = GetAlignment(ComAlignment(2).ListIndex)
162         .sXAxis = txtXAxis.Text
164         .XAxisAngle = IIf(IsNumeric(txtXAngle.Text), txtXAngle.Text, 0)
166         .sYAxis = txtYAxis.Text
168         .iHeight = 5000
170         .iWidth = 6000
172         .iParentHeight = 6500
174         .iParentWidth = 7060
176         .enmChartType = ComType.ListIndex + 1 ' Chart_Bubble
178         .enmScheme = CS_Solid
180         .iAngleX = 30
182         .iAngleY = 30
184         .enmAxesStyle = CA_FlatFrame
186         .lngBackColor = 14935011
188         .lngBorderColor = 11053224
190         .enmBorderEffect = CBE_Dark
192         .sngCylSides = 0
194         .enmGrid = CG_None
196         .lngInsideColor = 16777215
198         .enmMarkerShape = CMS_Many
200         .iMarkerSize = 12
202         .sngMarkerStep = 3
204         .sngPerspective = 1
206         .enmSmoothFlags = CSF_Fill
208         .enmStacked = CS_No
210         .iWallWidth = 4
212         .bZoom = False
214         .sConnStr = m_Cnn.ConnectionString
216         .sSQL = txtSQL.Text

218         With .udtTitle
220             .CT_Text = txtTitle.Text
222             .CT_BackColor = 2147483647
224             .CT_DrawingArea = True
226             .CT_Alignment = CA_StringAlignment_Center
228             .CT_DockArea = CDA_Area_Top
230             .CT_Flags = CF_TitleFlag_DrawingArea
232             .CT_LineAlignment = CA_StringAlignment_Far
234             .CT_LineGap = 2
            
236             With .CT_Font
238                 .CF_Size = lblPrev(0).Font.Size
240                 .CF_Bold = lblPrev(0).Font.Bold
242                 .CF_Name = lblPrev(0).Font.Name
244                 .CF_Italic = lblPrev(0).Font.Italic
246                 .CF_Strikethrough = lblPrev(0).Font.Strikethrough
248                 .CF_Underline = lblPrev(0).Font.Underline
250                 .CF_Weight = lblPrev(0).Font.Weight ' 500
                End With
            End With
   
252         With .udtNotes
254             .CT_Text = txtNotes.Text
256             .CT_BackColor = 2147483647
258             .CT_DrawingArea = True
260             .CT_Alignment = CA_StringAlignment_Far
262             .CT_DockArea = CDA_Area_Bottom
264             .CT_Flags = CF_TitleFlag_DrawingArea
266             .CT_LineAlignment = CA_StringAlignment_Far
268             .CT_LineGap = 2
            
270             With .CT_Font
272                 .CF_Size = lblPrev(1).Font.Size
274                 .CF_Bold = lblPrev(1).Font.Bold
276                 .CF_Name = lblPrev(1).Font.Name
278                 .CF_Italic = lblPrev(1).Font.Italic
280                 .CF_Strikethrough = lblPrev(1).Font.Strikethrough
282                 .CF_Underline = lblPrev(1).Font.Underline
284                 .CF_Weight = lblPrev(1).Font.Weight ' 500
                End With
            End With
   
286         If Len(sPath) > 0 Then

288             With .udtExports(0)
290                 .bForceKill = True
292                 .enmExportFormat = tplBin
294                 .sFilename = sName
296                 .sPath = sPath
                End With

            Else
            
298             With .udtExports(0)
300                 .bForceKill = False
302                 .enmExportFormat = tplBin
304                 .sFilename = ""
306                 .sPath = ""
                End With

            End If
            
        End With
    
308     If oRS Is Nothing Then
310         m_frmOASISCharts.SetChart udtOASISChartO
        Else
312         m_frmOASISCharts.SetChartWithRS udtOASISChartO, oRS
        End If
    
        '<EhFooter>
        Exit Sub

CreateChartSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.CreateChartSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadChartSettings()
        '<EhHeader>
        On Error GoTo LoadChartSettings_Err
        '</EhHeader>

        Dim sConnStr As String
        Dim i As Integer
    
100     If Not IsNull(listTemplates.Text) Then

102         udtOASISChartO = m_frmOASISCharts.OpenChartTemplate(g_sAppPath & "\data\templates\SecurityChartTemplates\" & Me.listTemplates.Text & "_chart.oct")

            On Error Resume Next

104         With udtOASISChartO
        
106             If .bChartTBR Then
108                 chkEnableChartTBR.Value = vbChecked
110                 ReDim .sChartTools(2)
112                 .sChartTools(0) = 0
114                 .sChartTools(1) = 20
                Else
116                 chkEnableChartTBR.Value = vbUnchecked
                End If
        
118             If .bAnnoTBR Then
120                 chkEnable.Value = vbChecked
122                 ReDim .sAnnoTools(2)
                Else
124                 chkEnable.Value = vbUnchecked
                End If
        
126             If .bDataEdtr Then
128                 chkDataEditor.Value = vbChecked
                Else
130                 chkDataEditor.Value = vbUnchecked
                End If
        
132             If .bSeriesLGD Then
134                 chkSeriesLegend.Value = vbChecked
                Else
136                 chkSeriesLegend.Value = vbUnchecked
                End If
        
138             If .bValueLGD Then
140                 chkPointLegend.Value = vbChecked
                Else
142                 chkPointLegend.Value = vbUnchecked
                End If
        
144             If .bDataHigls Then
146                 chkAllowData.Value = vbChecked
                Else
148                 chkAllowData.Value = vbUnchecked
                End If
        
150             .bContextMenu = False
152             .bMenuBar = False
154             .bDataEdtrAllowEdit = False
156             .bDataEdtrAllowDrag = False
        
158             If .bMultipleColors Then
160                 chkMultipleColors.Value = vbChecked
                Else
162                 chkMultipleColors.Value = vbUnchecked
                End If
        
164             If .bPointLabelsGen Then
166                 chkPointLabelsGen.Value = vbChecked
                Else
168                 chkPointLabelsGen.Value = vbUnchecked
                End If

170             .bScrollable = False
        
172             If .bHGlsPointLabel Then
174                 chkPointLabels.Value = vbChecked
                Else
176                 chkPointLabels.Value = vbUnchecked
                End If
        
178             If .bHGlsDimmed Then
180                 chkDimmed.Value = vbChecked
                Else
182                 chkDimmed.Value = vbUnchecked
                End If
        
184             SetAlignment 0, .iSerLgdAlign
186             SetAlignment 1, .iValLgdAlign
188             SetAlignment 2, .iDataEdtrAlign
        
190             txtXAxis.Text = .sXAxis
192             txtXAngle.Text = .XAxisAngle
194             chkStaggered.Value = IIf(.XAxisStaggered = True, vbChecked, vbUnchecked)
196             txtYAxis.Text = .sYAxis
        
198             txtTitle.Text = .udtTitle.CT_Text
        
200             lblPrev(0).Font.Size = .udtTitle.CT_Font.CF_Size
202             lblPrev(0).Font.Bold = .udtTitle.CT_Font.CF_Bold
204             lblPrev(0).Font.Name = .udtTitle.CT_Font.CF_Name
206             lblPrev(0).Font.Italic = .udtTitle.CT_Font.CF_Italic
208             lblPrev(0).Font.Strikethrough = .udtTitle.CT_Font.CF_Strikethrough
210             lblPrev(0).Font.Underline = .udtTitle.CT_Font.CF_Underline
212             lblPrev(0).Font.Weight = .udtTitle.CT_Font.CF_Weight

214             txtNotes.Text = .udtNotes.CT_Text
        
216             lblPrev(1).Font.Size = .udtNotes.CT_Font.CF_Size
218             lblPrev(1).Font.Bold = .udtNotes.CT_Font.CF_Bold
220             lblPrev(1).Font.Name = .udtNotes.CT_Font.CF_Name
222             lblPrev(1).Font.Italic = .udtNotes.CT_Font.CF_Italic
224             lblPrev(1).Font.Strikethrough = .udtNotes.CT_Font.CF_Strikethrough
226             lblPrev(1).Font.Underline = .udtNotes.CT_Font.CF_Underline
228             lblPrev(1).Font.Weight = .udtNotes.CT_Font.CF_Weight

230             ComType.ListIndex = .enmChartType - 1
        
232             If .bBorder Then
234                 chkBorder.Value = vbChecked
                Else
236                 chkBorder.Value = vbUnchecked
                End If

238             If .bCluster Then
240                 chkCluster.Value = vbChecked
                Else
242                 chkCluster.Value = vbUnchecked
                End If
        
244             If .bChart3D Then
246                 chk3D.Value = vbChecked
                Else
248                 chk3D.Value = vbUnchecked
                End If
        
250             If .bCrossHairs Then
252                 chkCrossHairs.Value = vbChecked
                Else
254                 chkCrossHairs.Value = vbUnchecked
                End If
        
256             If .bShowTips Then
258                 chkShowTips.Value = vbChecked
                Else
260                 chkShowTips.Value = vbUnchecked
                End If
                
264             txtSQL.Text = .sSQL
    
            End With
    
        End If
    
        '<EhFooter>
        Exit Sub

LoadChartSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.LoadChartSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetAlignment(iIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo GetAlignment_Err
        '</EhHeader>
        
100     GetAlignment = 513
        
102     Select Case iIndex

            Case 0
104             GetAlignment = 513

106         Case 1
108             GetAlignment = 515

110         Case 2
112             GetAlignment = 256

114         Case 3
116             GetAlignment = 258
        End Select
        
        '<EhFooter>
        Exit Function

GetAlignment_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.GetAlignment " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub SetAlignment(iIndex As Integer, _
                         iValue As Integer)
        '<EhHeader>
        On Error GoTo SetAlignment_Err
        '</EhHeader>
        
        'GetAlignment = 513
        'GetAlignment (ComAlignment(0).ListIndex)
        
100     ComAlignment(iIndex).ListIndex = 0

102     Select Case iValue

            Case 513
                'GetAlignment = 513
104             ComAlignment(iIndex).ListIndex = 0

106         Case 515
                'GetAlignment = 515
108             ComAlignment(iIndex).ListIndex = 1

110         Case 256
                'GetAlignment = 256
112             ComAlignment(iIndex).ListIndex = 2

114         Case 258
                'GetAlignment = 258
116             ComAlignment(iIndex).ListIndex = 3
        End Select
        
        '<EhFooter>
        Exit Sub

SetAlignment_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.SetAlignment " & _
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
               "in OASISClient.frmChartProviderSettings.cmdFontTitle_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SaveSQLSettings(sFile As String, _
                            desc As String, _
                            xml As String)
        '<EhHeader>
        On Error GoTo SaveSQLSettings_Err
        '</EhHeader>
        Dim i As Integer
        Dim found As Boolean
        Dim xmlDoc, Node
        Dim Res As VbMsgBoxResult

100     If Not desc = "" Then

102         Set xmlDoc = CreateObject("Msxml.DOMDocument")
104         xmlDoc.async = False
106         xmlDoc.documentElement = xmlDoc.createElement("storage")
        
108         Set Node = xmlDoc.documentElement.appendChild(xmlDoc.createElement("query"))
110         Call Node.setAttribute("id", xmlDoc.documentElement.childNodes.Length)
    
112         Set Node = xmlDoc.documentElement.appendChild(xmlDoc.createElement("query"))
114         Call Node.setAttribute("id", xmlDoc.documentElement.childNodes.Length)

116         While (Node.childNodes.Length > 0)
118             Node.removeChild (Node.childNodes.Item(0))
            Wend

120         Node.appendChild xmlDoc.createElement("description")
122         Node.lastChild.Text = desc
124         Node.appendChild xmlDoc.createElement("value")
126         Node.lastChild.Text = xml
128         xmlDoc.Save (g_sAppPath & "\data\templates\SecurityChartTemplates\" & sFile)
        
        End If

        '<EhFooter>
        Exit Sub

SaveSQLSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.SaveSQLSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GenerateTemplate()
        '<EhHeader>
        On Error GoTo GenerateTemplate_Err
        '</EhHeader>
        Dim i As Integer
        Dim l As Long

100     If Not IsNull(listTemplates.Text) Then
            
            On Error Resume Next
        
102         If Not LoadQuery(listTemplates.Text) Then
104             C1QueryFrame1.LoadFromXML XMLValues(0) ' (XMLValues(ComChartSQL.ListIndex - 1))
106             C1QueryFrame1.Render
108             C1QueryFrame2.LoadFromXML XMLValuesQRY(0) '(XMLValuesQRY(ComChartSQL.ListIndex - 1))
110             C1QueryFrame2.Render
112             C1Query1.BuildSQL

114             If Len(C1Query1.SQL) > 0 Then
116                 Adodc1.ConnectionString = m_Cnn.ConnectionString
118                 Adodc1.CommandType = adCmdText
120                 Adodc1.RecordSource = C1Query1.SQL
122                 txtSQL.Text = C1Query1.SQL
124                 Adodc1.Refresh

126                 Set DataGrid1.DataSource = Adodc1
128                 dxDBGrid1.Columns.DestroyColumns
130                 dxDBGrid1.KeyField = Adodc1.Recordset.Fields(0).Name
132                 Set dxDBGrid1.DataSource = Adodc1
134                 dxDBGrid1.Columns.RetrieveFields
                
136                 l = (dxDBGrid1.Width / Adodc1.Recordset.Fields.Count)
138                 i = 0

140                 Do Until i = dxDBGrid1.Columns.Count
142                     dxDBGrid1.Columns(i).Width = l
144                     i = i + 1
                    Loop
                
                End If
            
146             LoadChartSettings

            Else
        
148             Adodc1.ConnectionString = m_Cnn.ConnectionString
150             Adodc1.CommandType = adCmdText
152             Adodc1.RecordSource = txtSQL.Text
154             Adodc1.Refresh

156             Set DataGrid1.DataSource = Adodc1
158             dxDBGrid1.Columns.DestroyColumns
160             dxDBGrid1.KeyField = Adodc1.Recordset.Fields(0).Name
162             Set dxDBGrid1.DataSource = Adodc1
164             dxDBGrid1.Columns.RetrieveFields
                
166             l = (dxDBGrid1.Width / Adodc1.Recordset.Fields.Count)
168             i = 0

170             Do Until i = dxDBGrid1.Columns.Count
172                 dxDBGrid1.Columns(i).Width = l
174                 i = i + 1
                Loop

176             C1QueryFrame1.Clear
178             C1QueryFrame2.Clear
180             C1QueryFrame1.Render
182             C1QueryFrame2.Render

184             LoadChartSettings
            End If

        Else
186         C1QueryFrame1.Clear
188         C1QueryFrame2.Clear
190         C1QueryFrame1.Render
192         C1QueryFrame2.Render
194         txtSQL.Text = ""
196         Set DataGrid1.DataSource = Nothing
        End If

        '<EhFooter>
        Exit Sub

GenerateTemplate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.GenerateTemplate " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdGenerateTemplate_Click()
        '<EhHeader>
        On Error GoTo cmdGenerateTemplate_Click_Err
        '</EhHeader>
    
        Dim sSQL As String
        Dim sSQL1 As String
        Dim sSQL2 As String
    
100     Call GenerateTemplate
    
102     If chkDateFilter.Value = vbChecked And Not IsNull(dxDateFromTemplate) And Not IsNull(dxDateTillTemplate) Then
    
104         sSQL = C1Query1.SQL 'txtSQL.Text

106         If InStr(sSQL, "WHERE") <> 0 Then

108             sSQL1 = Left$(sSQL, InStr(sSQL, "WHERE") + 5)
110             sSQL2 = Right$(sSQL, Len(sSQL) - InStr(sSQL, "WHERE") - 5)
112             sSQL = sSQL1 & " (Incident_DATE BETWEEN #" & dxDateFromTemplate & "# AND #" & dxDateTillTemplate & "#) AND " & sSQL2
        
114         ElseIf InStr(sSQL, "GROUP BY") <> 0 Then

116             sSQL1 = Left$(sSQL, InStr(sSQL, "GROUP BY") - 1)
118             sSQL2 = Right$(sSQL, Len(sSQL) - InStr(sSQL, "GROUP BY") + 1)
120             sSQL = sSQL1 & " WHERE (Incident_DATE BETWEEN #" & dxDateFromTemplate & "# AND #" & dxDateTillTemplate & "#) " & sSQL2
             
            Else
    
122             sSQL = sSQL & " WHERE Incident_DATE BETWEEN #" & dxDateFromTemplate & "# AND #" & dxDateTillTemplate & "#" '& sSQL2
        
            End If

            On Error GoTo theend
            'MsgBox sSQL
124         Adodc1.RecordSource = sSQL
126         Adodc1.Refresh
128         txtSQL.Text = sSQL
        
        End If

130     Call cmdAssign_Click
theend:
    
        '<EhFooter>
        Exit Sub

cmdGenerateTemplate_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.cmdGenerateTemplate_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>
        Dim xml As String
        Dim desc As String

100     desc = InputBox("Enter a name for this template", "OASIS Charting Templates")
    
102     If Not desc = "" Then

104         If Not FileExists(g_sAppPath & "\data\templates\SecurityChartTemplates\" & desc & "_query.xml") And Not FileExists(g_sAppPath & "\data\templates\SecurityChartTemplates\" & desc & ".ini") Then

                If Not chkUseCustom.Value = vbChecked Then
106                 xml = C1QueryFrame1.SaveToXML
108                 SaveSQLSettings desc & "_query.xml", desc, xml
110                 xml = C1QueryFrame2.SaveToXML
112                 SaveSQLSettings desc & "_constraints.xml", desc, xml
                Else
                
                    Dim oIni As New clIniReader
    
                     CreateNewINI g_sAppPath & "\data\templates\SecurityChartTemplates\" & desc & ".ini", oIni, txtSQL.Text
                
                End If

114             CreateChartSettings g_sAppPath & "\data\templates\SecurityChartTemplates\", desc & "_chart", Adodc1.Recordset
116             populateListView g_sAppPath & "\data\templates\SecurityChartTemplates", "xml", "ini", Me.listTemplates
                '118             MsgBox "Three files (" & desc & "_query.xml & " & desc & "_constraints.xml & " & desc & "_chart.oct) were created in directory:" & Chr(13) & Chr(13) & g_sAppPath & "\data\templates\SecurityChartTemplates", vbInformation, "Template saved"
    
            Else
        
120             MsgBox "A template by this name already exists!", vbExclamation, "Template already exists"
        
            End If

        End If

        '<EhFooter>
        Exit Sub

cmdSave_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmChartProviderSettings.cmdSave_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CreateNewINI(StrIniFile As String, _
                              oINIReader As clIniReader, _
                              sVal As String) As Boolean
        '<EhHeader>
        On Error GoTo CreateNewINI_Err
        '</EhHeader>
        Dim fs As New FileSystemObject
    
100     If Not fs.FileExists(StrIniFile) Then

102         fs.CreateTextFile StrIniFile
104         Set fs = Nothing
        End If
        
        sVal = Replace(sVal, vbCrLf, " ")
        sVal = Replace(sVal, vbNewLine, " ")
        sVal = Replace(sVal, vbLf, " ")
        
        With oINIReader
            .Path = StrIniFile
            .Section = "default"
            
            .Key = "SQL"
            .Value = sVal
            .AddKeyWithValue

        End With
        
        '106     With oINIReader
        '            .Path = StrIniFile
        '108         .Section = "default"
        '110         .Key = "starts"
        '112         .Value = sVal
        '114         .AddNewSection
        '        End With
    
116     CreateNewINI = True
        '<EhFooter>
        Exit Function

CreateNewINI_Err:
         
        '</EhFooter>
End Function


Private Function FileExists(sFullPath As String) As Boolean
    Dim oFile As New Scripting.FileSystemObject
    FileExists = oFile.FileExists(sFullPath)
End Function

Private Sub cmdSwapTable_Click()
        '<EhHeader>
        On Error GoTo cmdSwapTable_Click_Err
        '</EhHeader>
100     frmTableChooser.Show vbModal, Me
        '<EhFooter>
        Exit Sub

cmdSwapTable_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.cmdSwapTable_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = vbKeyF11 Then
    
        If C1Tab1.TabVisible(1) Then
            C1Tab1.TabVisible(1) = False
            C1Tab1.CurrTab = 0
        Else
        
            C1Tab1.TabVisible(1) = True
            C1Tab1.CurrTab = 1
        End If
    
    End If

End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        
100     Set m_frmOASISCharts = New frmOASISCharts
        
102     ComAlignment(0).ListIndex = 0
104     ComAlignment(1).ListIndex = 0
106     ComAlignment(2).ListIndex = 0
108     ComType.ListIndex = 1
110     populateListView g_sAppPath & "\data\templates\SecurityChartTemplates", "xml", "ini", Me.listTemplates
112     LoadSchema
114     LoadTable "oincidents_FEA"
116     XMLValuesClear
        C1Tab1.CurrTab = 0

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub populateListView(sPath As String, _
                             sFilter As String, sFilter1 As String, _
                             lListBox As ListBox)
        '<EhHeader>
        On Error GoTo populateListView_Err
        '</EhHeader>

        Dim oItem As MSComctlLib.ListItem
        Dim sFile As String
        Dim sString As String
   
100     lListBox.Clear
    
102     If Right(sPath, 1) <> "\" Then
104         sPath = sPath & "\"
        End If
   
106     sFile = Dir(sPath & "*." & sFilter)
   
108     While sFile <> Empty
110         sString = Left$(sFile, Len(sFile) - 4)

112         If Right$(sString, 6) = "_query" Then
114             sString = Left$(sString, Len(sString) - 6)
116             lListBox.AddItem sString
            End If

118         sFile = Dir
        Wend
        
        
306     sFile = Dir(sPath & "*." & sFilter1)
   
308     While sFile <> Empty
310         sString = Left$(sFile, Len(sFile) - 4)

312        ' If Right$(sString, 6) = "_query" Then
314        '     sString = Left$(sString, Len(sString) - 6)
316             lListBox.AddItem sString
            ' End If

318         sFile = Dir
        Wend

        
        
'        Dim oIni As New clIniReader
'
'202     With oIni
'
'204         .Path = g_sAppPath & "\data\templates\SecurityChartTemplates\sup.ini"
'206         .Section = "default"
'208         .Key = "SQL"
'210         DebugPrint .Value
'
'        End With

        '<EhFooter>
        Exit Sub

populateListView_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmChartProviderSettings.populateListView " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadSchema()
        '<EhHeader>
        On Error GoTo LoadSchema_Err
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
      
100     Set Cnxn = New ADODB.Connection
102     strCnxn = m_Cnn.ConnectionString
104     Cnxn.Open strCnxn
    
        ' Open database schema
106     Set rstSchema = Cnxn.OpenSchema(adSchemaTables)
   
108     Set doc = New DOMDocument
   
110     Set schemaNode = doc.createElement("SCHEMA")

112     doc.appendChild schemaNode
   
114     Set rootFolderNode = doc.createElement("FOLDER")
116     schemaNode.appendChild rootFolderNode
        
118     Set foldersNode = doc.createElement("SUBFOLDERS")
120     rootFolderNode.appendChild foldersNode
  
122     Set viewsNode = doc.createElement("VIEWS")
124     schemaNode.appendChild viewsNode
   
126     Set relationsNode = doc.createElement("JOINRELATIONS")
128     schemaNode.appendChild relationsNode
   
130     Set catalogsNode = doc.createElement("CATALOGS")
132     schemaNode.appendChild catalogsNode
   
134     AddChildTextNode schemaNode, "IMPORTEDFROM", m_Cnn.ConnectionString
136     AddChildTextNode schemaNode, "CP", "0"
   
        ' creating views and folders
        
138     frmTableChooser.List1.Clear
        
140     rstSchema.Filter = "TABLE_NAME = 'oincidents_FEA'"
        
142     Do Until rstSchema.EOF

144         If (rstSchema!TABLE_TYPE = "TABLE") Then
146             Set folderNode = doc.createElement("FOLDER")
148             foldersNode.appendChild folderNode
150             Set viewNode = doc.createElement("VIEW")
152             TableName = rstSchema!TABLE_NAME
154             AddChildTextNode viewNode, "VIEWNAME", TableName
156             AddChildTextNode folderNode, "FOLDERNAME", TableName
158             viewsNode.appendChild viewNode
160             List1.AddItem TableName
162             frmTableChooser.List1.AddItem TableName
                ' table info in a view
164             Set tableInfosNode = doc.createElement("VIEWTABLES")
166             viewNode.appendChild tableInfosNode
168             Set tableInfoNode = doc.createElement("VIEWTABLE")
170             tableInfosNode.appendChild tableInfoNode
172             AddChildTextNode tableInfoNode, "TABLENAME", TableName
174             AddChildTextNode tableInfoNode, "REQUIRED", "1"
                ' for each table column create a view field and a folder field
176             Set viewFieldsNode = doc.createElement("VIEWFIELDS")
178             viewNode.appendChild viewFieldsNode
180             Set folderFieldsNode = doc.createElement("FOLDERFIELDS")
182             folderNode.appendChild folderFieldsNode
                
184             Set rstColumns = Cnxn.OpenSchema(adSchemaColumns, Array(Empty, Empty, TableName))

186             Do Until rstColumns.EOF
188                 Set viewFieldNode = doc.createElement("VIEWFIELD")
190                 viewFieldsNode.appendChild viewFieldNode
192                 Set folderFieldNode = doc.createElement("FOLDERFIELD")
194                 folderFieldsNode.appendChild folderFieldNode
196                 AddChildTextNode viewFieldNode, "FIELDNAME", rstColumns!COLUMN_NAME
198                 AddChildTextNode viewFieldNode, "FIELDTABLENAME", TableName
200                 AddChildTextNode viewFieldNode, "FIELDTYPE", rstColumns!DATA_TYPE
202                 AddChildTextNode folderFieldNode, "NAME", rstColumns!COLUMN_NAME
204                 AddChildTextNode folderFieldNode, "VIEW", TableName
206                 AddChildTextNode folderFieldNode, "TABLENAME", TableName
208                 leng = "0"

210                 If rstColumns!CHARACTER_MAXIMUM_LENGTH <> Empty Then
212                     leng = rstColumns!CHARACTER_MAXIMUM_LENGTH
                    End If

214                 AddChildTextNode viewFieldNode, "FIELDSIZE", leng
216                 leng = "0"

218                 If rstColumns!NUMERIC_PRECISION <> Empty Then
220                     leng = rstColumns!NUMERIC_PRECISION
                    End If

222                 AddChildTextNode viewFieldNode, "FIELDPREC", leng
224                 leng = "0"

226                 If rstColumns!NUMERIC_SCALE <> Empty Then
228                     leng = rstColumns!NUMERIC_SCALE
                    End If

230                 AddChildTextNode viewFieldNode, "FIELDSCALE", leng
232                 rstColumns.MoveNext
                Loop

            End If

234         rstSchema.MoveNext
        Loop

        '<EhFooter>
        Exit Sub

LoadSchema_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.LoadSchema " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadSchemaEX()
        '<EhHeader>
        On Error GoTo LoadSchemaEX_Err
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
      
100     Set Cnxn = New ADODB.Connection
102     strCnxn = m_Cnn.ConnectionString
104     Cnxn.Open strCnxn
    
        ' Open database schema
106     Set rstSchema = Cnxn.OpenSchema(adSchemaTables)
   
108     Set doc = New DOMDocument
   
110     Set schemaNode = doc.createElement("SCHEMA")

112     doc.appendChild schemaNode
   
114     Set rootFolderNode = doc.createElement("FOLDER")
116     schemaNode.appendChild rootFolderNode
        
118     Set foldersNode = doc.createElement("SUBFOLDERS")
120     rootFolderNode.appendChild foldersNode
  
122     Set viewsNode = doc.createElement("VIEWS")
124     schemaNode.appendChild viewsNode
   
126     Set relationsNode = doc.createElement("JOINRELATIONS")
128     schemaNode.appendChild relationsNode
   
130     Set catalogsNode = doc.createElement("CATALOGS")
132     schemaNode.appendChild catalogsNode
   
134     AddChildTextNode schemaNode, "IMPORTEDFROM", m_Cnn.ConnectionString
136     AddChildTextNode schemaNode, "CP", "0"
   
        ' creating views and folders
        
138     frmTableChooser.List1.Clear
        
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
160             frmTableChooser.List1.AddItem TableName
                ' table info in a view
162             Set tableInfosNode = doc.createElement("VIEWTABLES")
164             viewNode.appendChild tableInfosNode
166             Set tableInfoNode = doc.createElement("VIEWTABLE")
168             tableInfosNode.appendChild tableInfoNode
170             AddChildTextNode tableInfoNode, "TABLENAME", TableName
172             AddChildTextNode tableInfoNode, "REQUIRED", "1"
                ' for each table column create a view field and a folder field
174             Set viewFieldsNode = doc.createElement("VIEWFIELDS")
176             viewNode.appendChild viewFieldsNode
178             Set folderFieldsNode = doc.createElement("FOLDERFIELDS")
180             folderNode.appendChild folderFieldsNode
                
182             Set rstColumns = Cnxn.OpenSchema(adSchemaColumns, Array(Empty, Empty, TableName))

184             Do Until rstColumns.EOF
186                 Set viewFieldNode = doc.createElement("VIEWFIELD")
188                 viewFieldsNode.appendChild viewFieldNode
190                 Set folderFieldNode = doc.createElement("FOLDERFIELD")
192                 folderFieldsNode.appendChild folderFieldNode
194                 AddChildTextNode viewFieldNode, "FIELDNAME", rstColumns!COLUMN_NAME
196                 AddChildTextNode viewFieldNode, "FIELDTABLENAME", TableName
198                 AddChildTextNode viewFieldNode, "FIELDTYPE", rstColumns!DATA_TYPE
200                 AddChildTextNode folderFieldNode, "NAME", rstColumns!COLUMN_NAME
202                 AddChildTextNode folderFieldNode, "VIEW", TableName
204                 AddChildTextNode folderFieldNode, "TABLENAME", TableName
206                 leng = "0"

208                 If rstColumns!CHARACTER_MAXIMUM_LENGTH <> Empty Then
210                     leng = rstColumns!CHARACTER_MAXIMUM_LENGTH
                    End If

212                 AddChildTextNode viewFieldNode, "FIELDSIZE", leng
214                 leng = "0"

216                 If rstColumns!NUMERIC_PRECISION <> Empty Then
218                     leng = rstColumns!NUMERIC_PRECISION
                    End If

220                 AddChildTextNode viewFieldNode, "FIELDPREC", leng
222                 leng = "0"

224                 If rstColumns!NUMERIC_SCALE <> Empty Then
226                     leng = rstColumns!NUMERIC_SCALE
                    End If

228                 AddChildTextNode viewFieldNode, "FIELDSCALE", leng
230                 rstColumns.MoveNext
                Loop

            End If

232         rstSchema.MoveNext
        Loop

        '<EhFooter>
        Exit Sub

LoadSchemaEX_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.LoadSchemaEX " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadTable(sTable As String)
        '<EhHeader>
        On Error GoTo LoadTable_Err
        '</EhHeader>
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
        Dim k As Integer

100     List2.Clear
102     While relationsNode.childNodes.Length > 0
104         relationsNode.removeChild relationsNode.childNodes(0)
        Wend
106     While catalogsNode.childNodes.Length > 0
108         catalogsNode.removeChild catalogsNode.childNodes(0)
        Wend
        ' Add 1st view group and main view name (first item in the view group)
110     Set catalogNode = doc.createElement("CATALOG")
112     catalogsNode.appendChild catalogNode
114     AddChildTextNode catalogNode, "VIEW", sTable

        ' Add join relationships and catalogs
116     For i = 0 To viewsNode.childNodes.Length - 1

118         If viewsNode.childNodes(i).childNodes(0).nodeTypedValue = sTable Then
120             view1 = sTable
122             Set viewFieldsNode = viewsNode.childNodes(i).childNodes(2)
                Exit For
            End If

        Next

124     For i = 0 To viewsNode.childNodes.Length - 1

126         tableCandidate = viewsNode.childNodes(i).childNodes(0).nodeTypedValue

128         If tableCandidate <> sTable Then
130             Set fieldsNode = viewsNode.childNodes(i).childNodes(2) ' 2nd node is VIEWFIELDS

132             For j = 0 To fieldsNode.childNodes.Length - 1
134                 added = False
136                 fieldCandidateName = fieldsNode.childNodes(j).childNodes(0).nodeTypedValue

138                 For k = 0 To viewFieldsNode.childNodes.Length - 1

140                     If viewFieldsNode.childNodes(k).childNodes(0).nodeTypedValue = fieldCandidateName Then
142                         List2.AddItem tableCandidate & " (using " & fieldCandidateName & " field)"
                            ' Add relationship
144                         Set relationNode = doc.createElement("JOINRELATION")
146                         relationsNode.appendChild relationNode
148                         Set relateViewsNode = doc.createElement("RELATEVIEWS")
150                         AddChildTextNode relateViewsNode, "VIEWNAME", sTable
152                         AddChildTextNode relateViewsNode, "VIEWNAME", tableCandidate
154                         relationNode.appendChild relateViewsNode
156                         Set viewLinksNode = doc.createElement("VIEWLINKS")
158                         relationNode.appendChild viewLinksNode
160                         Set viewLinkNode = doc.createElement("VIEWLINK")
162                         viewLinksNode.appendChild viewLinkNode
164                         AddChildTextNode viewLinkNode, "VIEW1", sTable
166                         AddChildTextNode viewLinkNode, "JOINTYPE", "0"
168                         AddChildTextNode viewLinkNode, "VIEW2", tableCandidate
                            ' Joins
170                         Set joinsNode = doc.createElement("JOINS")
172                         relationNode.appendChild joinsNode
174                         Set joinNode = doc.createElement("JOIN")
176                         joinsNode.appendChild joinNode
178                         AddChildTextNode joinNode, "FIELD1", fieldCandidateName
180                         AddChildTextNode joinNode, "TABLE1", sTable
182                         AddChildTextNode joinNode, "VIEW1", sTable
184                         AddChildTextNode joinNode, "OPERATOR", "0"  '' is equal to
186                         AddChildTextNode joinNode, "FIELD2", fieldCandidateName
188                         AddChildTextNode joinNode, "TABLE2", tableCandidate
190                         AddChildTextNode joinNode, "VIEW2", tableCandidate
                            ' Add this table to the view group
192                         AddChildTextNode catalogNode, "VIEW", tableCandidate
194                         added = True
                            Exit For
                        End If

                    Next

196                 If added = True Then
                        Exit For
                    End If

                Next

            End If

        Next
  
        ' Create the remaining (2nd, etc) view groups (catalogs).
198     For i = 0 To viewsNode.childNodes.Length - 1
200         candidate = viewsNode.childNodes(i).childNodes(0).nodeTypedValue
202         found = False

204         For j = 0 To catalogNode.childNodes.Length - 1

206             If catalogNode.childNodes(j).nodeTypedValue = candidate Then
208                 found = True
                    Exit For
                End If

            Next

210         If Not found Then
212             Set newCatNode = doc.createElement("CATALOG")
214             catalogsNode.appendChild newCatNode
216             AddChildTextNode newCatNode, "VIEW", candidate
            End If

        Next
 
        ' Assign the schema to the control
218     C1Query1.Schema = doc.xml

        ' Initialize the C1QueryFrame controls
220     C1QueryFrame1.Clear
222     C1QueryFrame1.CurrentItemID = 0
224     C1QueryFrame1.Render
226     C1QueryFrame2.Clear
228     C1QueryFrame2.CurrentItemID = 0
230     C1QueryFrame2.Render

        '<EhFooter>
        Exit Sub

LoadTable_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.LoadTable " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

        On Error Resume Next
100     Unload m_frmOASISCharts
102     RSLocalUserGroups.Close
104     Set RSLocalUserGroups = Nothing
106     Set m_frmOASISCharts = Nothing

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.Form_Unload " & _
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
102     Call GenerateTemplate
        
        'MsgBox Adodc1.Recordset.Source

        'Stop
        'Adodc1.Recordset
    
        '    If Not IsNull(listTemplates.Text) Then
        '
        '        If Not RSLoadedSpatialAnalysisTemplate Is Nothing Then RSLoadedSpatialAnalysisTemplate.Close
        '        Set RSLoadedSpatialAnalysisTemplate = New ADODB.Recordset
        '        RSLoadedSpatialAnalysisTemplate.Open g_sAppPath & "\data\templates\SecurityChartTemplates\" & Me.listTemplates & ".xml"
        '
        '        If RSLoadedSpatialAnalysisTemplate.State = adStateOpen Then
        '
        '            Do Until RSLoadedSpatialAnalysisTemplate.EOF
        '
        '                If RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerDateRange").Value = True Then
        '                    bDatesUsed = True
        '                    dDateFrom = RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerDateFrom").Value
        '                    dDateTill = RSLoadedSpatialAnalysisTemplate.Fields("OverlayLayerDateTill").Value
        '                End If
        '
        '                RSLoadedSpatialAnalysisTemplate.MoveNext
        '            Loop
        '
        '            SafeMoveFirst RSLoadedSpatialAnalysisTemplate
        '
        '            If Not bDatesUsed Then
        '                Me.dxDateFromTemplate.Visible = False
        '                Me.dxDateTillTemplate.Visible = False
        '                Me.lblDateFrom.Visible = False
        '                Me.lblDateFrom.Visible = False
        '            Else
        '                Me.dxDateFromTemplate.Visible = True
        '                Me.dxDateTillTemplate.Visible = True
        '                Me.lblDateFrom.Visible = True
        '                Me.lblDateFrom.Visible = True
        '                Me.dxDateFromTemplate = dDateFrom
        '                Me.dxDateTillTemplate = dDateTill
        '            End If
        '
        '        End If
        '
        '    End If

        '<EhFooter>
        Exit Sub

listTemplates_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChartProviderSettings.listTemplates_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
