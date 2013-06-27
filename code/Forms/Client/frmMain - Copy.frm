VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{13C02136-A0FB-4F12-9894-5D9DC001170C}#6.0#0"; "ctTree.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{E760686B-BC9E-4802-9ECF-175FDF4062CE}#5.0#0"; "MAPX50.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{98C28BA7-605C-46CF-927B-9AA96B3230E0}#1.1#0"; "Scoring_Module.ocx"
Object = "{978D0654-1960-4687-8496-61E47D034439}#1.0#0"; "OASISWebBrowser.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TatukGIS_XDK10_WIN32.ocx"
Begin VB.Form frmMain 
   Caption         =   "s"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12840
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8595
   ScaleWidth      =   12840
   WindowState     =   2  'Maximized
   Begin MapXLib.Map Map1 
      Height          =   255
      Left            =   10140
      TabIndex        =   147
      Top             =   6300
      Visible         =   0   'False
      Width           =   555
      _Version        =   500012
      _ExtentX        =   979
      _ExtentY        =   450
      _StockProps     =   1
      MapCatalog.GeoDictionary=   "GeoDictionary"
      GeoSet          =   "Empty GeosetName {9A9AC2F4-8375-44d1-BCEB-476AE986F190}"
      DefaultStyle.TextFontBackColor=   16777215
      DefaultStyle.SupportsBitmapSymbols=   -1  'True
      DefaultStyle.SymbolChar=   55
      DefaultStyle.SymbolFontBackColor=   16777215
      BeginProperty DefaultStyle.TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DefaultStyle.SymbolFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Map Symbols"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DefaultStyle.LineStyle=   1
      DefaultStyle.LineWidth=   1
      DefaultStyle.RegionColor=   16777215
      DefaultStyle.LinePattern=   2
      DefaultStyle.RegionBackColor=   16777215
      DefaultStyle.RegionBorderStyle=   1
      DefaultStyle.RegionBorderWidth=   1
      Title.Visible   =   0   'False
      Title.Text      =   "Empty Title {01A9504B-CE13-4415-A5A0-51D8C2F15204}"
      Title.Style.TextFontBackColor=   16777215
      Title.Style.TextFontOpaque=   -1  'True
      Title.Style.SymbolChar=   0
      BeginProperty Title.Style.TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   23.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Title.Style.SymbolFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   23.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Title.X         =   184
      Title.Y         =   17
      Map.NumericCoordSys.ProjectionInfo=   "frmMain.frx":6852
      Map.DisplayCoordSys.ProjectionInfo=   "frmMain.frx":6982
   End
   Begin VB.CommandButton cmdCommand6 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5160
      TabIndex        =   136
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAttachments 
      Caption         =   "Attachments"
      Height          =   495
      Left            =   4080
      TabIndex        =   135
      Top             =   4980
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSendMess 
      Caption         =   "SendMess"
      Height          =   255
      Left            =   4080
      TabIndex        =   134
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAttmtLister 
      Caption         =   "AttmtLister"
      Height          =   255
      Left            =   4020
      TabIndex        =   133
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   195
      Left            =   3960
      TabIndex        =   132
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrToolTip 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4620
      Top             =   7740
   End
   Begin VB.PictureBox picToolTip 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   1185
      TabIndex        =   130
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label lblToolTip 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   131
         Top             =   30
         Width           =   45
      End
   End
   Begin VB.Frame FraFrmConvert 
      Caption         =   "frmConvert"
      Height          =   765
      Left            =   13260
      TabIndex        =   64
      Top             =   960
      Visible         =   0   'False
      Width           =   3105
      Begin VB.TextBox txtUTMY 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   2460
         Width           =   2235
      End
      Begin VB.TextBox txtUTMX 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   129
         Top             =   2040
         Width           =   2235
      End
      Begin VB.TextBox txtDDY 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   128
         Top             =   1680
         Width           =   2235
      End
      Begin VB.TextBox txtDDX 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   1320
         Width           =   2235
      End
      Begin VB.Frame FraCoords 
         Caption         =   "Coords"
         Height          =   945
         Left            =   300
         TabIndex        =   66
         Top             =   300
         Width           =   2355
         Begin VB.TextBox txtYOrg 
            Height          =   315
            Left            =   390
            TabIndex        =   68
            Top             =   510
            Width           =   1815
         End
         Begin VB.TextBox txtXOrg 
            Height          =   285
            Left            =   390
            TabIndex        =   67
            Top             =   180
            Width           =   1785
         End
         Begin VB.Label lblY 
            AutoSize        =   -1  'True
            Caption         =   "Y:"
            Height          =   195
            Left            =   180
            TabIndex        =   70
            Top             =   510
            Width           =   150
         End
         Begin VB.Label lblX 
            AutoSize        =   -1  'True
            Caption         =   "X:"
            Height          =   195
            Left            =   210
            TabIndex        =   69
            Top             =   180
            Width           =   150
         End
      End
      Begin VB.CommandButton cmdDoConversion 
         Caption         =   "DoConversion"
         Height          =   285
         Left            =   1410
         TabIndex        =   65
         Top             =   150
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdShowSelectedInfo 
      Caption         =   "ShowSelectedInfo"
      Height          =   285
      Left            =   3960
      TabIndex        =   61
      Top             =   3900
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdTracker 
      Caption         =   "Tracker"
      Height          =   285
      Left            =   3960
      TabIndex        =   59
      Top             =   3600
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdSingleSelect 
      Caption         =   "Single Select"
      Height          =   285
      Left            =   3960
      TabIndex        =   57
      Top             =   3180
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame FraRotation 
      Caption         =   "Rotation"
      Height          =   555
      Left            =   12240
      TabIndex        =   49
      Top             =   2400
      Visible         =   0   'False
      Width           =   915
      Begin VB.TextBox edRotationAngle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   60
         TabIndex        =   50
         Top             =   210
         Width           =   240
      End
      Begin MSComCtl2.UpDown udRotationAngle 
         Height          =   285
         Left            =   301
         TabIndex        =   51
         Top             =   210
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "edRotationAngle"
         BuddyDispid     =   196632
         OrigLeft        =   1680
         OrigRight       =   1920
         OrigBottom      =   255
         Increment       =   5
         Max             =   180
         Min             =   -180
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
   End
   Begin VB.CommandButton cmdCommand5 
      Caption         =   "tool Test"
      Height          =   285
      Left            =   3960
      TabIndex        =   47
      Top             =   810
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdCommand4 
      Caption         =   "Gen Edit lyr"
      Height          =   285
      Left            =   3960
      TabIndex        =   46
      Top             =   2400
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSScriptControlCtl.ScriptControl SSC 
      Left            =   15160
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "MY OPS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   2
      Top             =   2940
      Visible         =   0   'False
      Width           =   1170
   End
   Begin OASISMAScoring.MineActionScoring MineActionScoring1 
      Left            =   3870
      Top             =   6570
      _ExtentX        =   1693
      _ExtentY        =   661
   End
   Begin VB.Timer tmrInternetCheck 
      Enabled         =   0   'False
      Left            =   4095
      Top             =   7740
   End
   Begin VB.CommandButton cmdSetScope 
      Caption         =   "Scope"
      Height          =   285
      Left            =   4020
      TabIndex        =   41
      Top             =   1020
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdCreateThematics 
      Caption         =   "Themat"
      Height          =   285
      Left            =   4020
      TabIndex        =   40
      Top             =   1260
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdLoadSQLLyr 
      Caption         =   "LoadSQLLyr"
      Height          =   285
      Left            =   10890
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdSelectCircle 
      Caption         =   "Circle Select"
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Top             =   2100
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.ComboBox comActiveLyr 
      Height          =   315
      Left            =   6030
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cmbSnap 
      Height          =   315
      Left            =   7830
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   8460
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   82182146
      CurrentDate     =   39232
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   285
      Left            =   3990
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComctlLib.ImageList MapTools 
      Left            =   5160
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8421376
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6BC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":700C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":711E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7230
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7342
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7454
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7566
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7678
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":778A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":789C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ActiveBar3LibraryCtl.ActiveBar3 AB 
      Align           =   1  'Align Top
      Height          =   8595
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12840
      _LayoutVersion  =   2
      _ExtentX        =   22648
      _ExtentY        =   15161
      _DataPath       =   ""
      Bands           =   "frmMain.frx":801A
      Begin MSComctlLib.ImageList imlHeader 
         Left            =   4320
         Top             =   5580
         _ExtentX        =   794
         _ExtentY        =   794
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":22435
               Key             =   ""
               Object.Tag             =   "Adress"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":22547
               Key             =   ""
               Object.Tag             =   "Check"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":22659
               Key             =   ""
               Object.Tag             =   "Company"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2276B
               Key             =   ""
               Object.Tag             =   "Count"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2287D
               Key             =   ""
               Object.Tag             =   "Time"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2298F
               Key             =   ""
               Object.Tag             =   "Man"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":22AA1
               Key             =   ""
               Object.Tag             =   "Phone"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":22BB3
               Key             =   ""
               Object.Tag             =   "Apple"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":22CC5
               Key             =   ""
               Object.Tag             =   "Salary"
            EndProperty
         EndProperty
      End
      Begin C1SizerLibCtl.C1Elastic elDynamicReports 
         Height          =   855
         Left            =   4680
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   6240
         Visible         =   0   'False
         Width           =   1335
         _cx             =   2355
         _cy             =   1508
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
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
         AutoSizeChildren=   8
         BorderWidth     =   0
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
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmMain.frx":22DD7
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin OASISClient.DynamicDataReports DynamicDataReports1 
            Height          =   855
            Left            =   0
            TabIndex        =   139
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1508
         End
      End
      Begin C1SizerLibCtl.C1Elastic elDynamicData 
         Height          =   705
         Left            =   4200
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
         _cx             =   1931
         _cy             =   1244
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
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
         AutoSizeChildren=   8
         BorderWidth     =   0
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
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmMain.frx":22E0B
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin OASISClient.DynamicDataModule DynamicDataModule1 
            Height          =   705
            Left            =   0
            TabIndex        =   140
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1244
         End
      End
      Begin C1SizerLibCtl.C1Elastic elAddons 
         Height          =   1725
         Left            =   2295
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1935
         Visible         =   0   'False
         Width           =   2220
         _cx             =   3916
         _cy             =   3043
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
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
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmMain.frx":22E3F
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin OASISWebBrowser.Browser WebBrowser2 
            Height          =   1725
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   3043
         End
      End
      Begin C1SizerLibCtl.C1Elastic elOasisProfile 
         Height          =   915
         Left            =   13920
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   6660
         _cx             =   11748
         _cy             =   1614
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
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
         BorderWidth     =   10
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
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmMain.frx":22E71
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin OASISWebBrowser.Browser WebBrowser1 
            Height          =   615
            Left            =   150
            TabIndex        =   45
            Top             =   150
            Width           =   6360
            _ExtentX        =   11218
            _ExtentY        =   1085
         End
      End
      Begin C1SizerLibCtl.C1Elastic elRSSTool 
         Height          =   7350
         Left            =   14880
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1740
         Width           =   5445
         _cx             =   9604
         _cy             =   12965
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
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
         BorderWidth     =   10
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
         GridRows        =   1
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmMain.frx":22EA6
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin OASISClient.OASISRSSBrowser RSSBrowser1 
            Height          =   7050
            Left            =   150
            TabIndex        =   141
            Top             =   150
            Width           =   5145
            _ExtentX        =   9075
            _ExtentY        =   12435
         End
      End
      Begin C1SizerLibCtl.C1Elastic elMap 
         Height          =   7260
         Left            =   4320
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   10935
         _cx             =   19288
         _cy             =   12806
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
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
         AutoSizeChildren=   8
         BorderWidth     =   10
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
         _GridInfo       =   $"frmMain.frx":22EDD
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab C1TTab1Tab2 
            Height          =   2325
            Left            =   150
            TabIndex        =   88
            Top             =   4785
            Width           =   10635
            _cx             =   18759
            _cy             =   4101
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
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
            Caption         =   "Data|Selector|Select"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   3
            Position        =   2
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
            TabHeight       =   250
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Flags(1)        =   2
            Begin OASISClient.ctlSelector ctlSelector1 
               Height          =   2235
               Left            =   11835
               TabIndex        =   146
               Top             =   45
               Width           =   10290
               _ExtentX        =   18150
               _ExtentY        =   3942
            End
            Begin C1SizerLibCtl.C1Elastic elSelector 
               Height          =   2235
               Left            =   11535
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   45
               Width           =   10290
               _cx             =   18150
               _cy             =   3942
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
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
               _GridInfo       =   $"frmMain.frx":22F22
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic elSelTools 
                  Height          =   375
                  Left            =   15
                  TabIndex        =   109
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   10260
                  _cx             =   18098
                  _cy             =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
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
                  Begin OASISClient.OASISButton OASSelect 
                     Height          =   315
                     Left            =   3780
                     TabIndex        =   142
                     Top             =   0
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   556
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
                  Begin VB.ComboBox ComFeatureLayer 
                     Height          =   315
                     Left            =   6420
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   112
                     Top             =   0
                     Width           =   2535
                  End
                  Begin VB.ComboBox ComSelLayer 
                     Height          =   315
                     Left            =   1620
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   110
                     Top             =   0
                     Width           =   2115
                  End
                  Begin VB.Label lblFeatureLayer 
                     AutoSize        =   -1  'True
                     Caption         =   "Feature Layer:"
                     Height          =   195
                     Left            =   5340
                     TabIndex        =   113
                     Top             =   60
                     Width           =   1020
                  End
                  Begin VB.Label lblActiveInfo 
                     AutoSize        =   -1  'True
                     Caption         =   "Selection/Reporting:"
                     Height          =   195
                     Left            =   60
                     TabIndex        =   111
                     Top             =   60
                     Width           =   1470
                  End
               End
               Begin C1SizerLibCtl.C1Tab C1TTabSelector 
                  Height          =   1830
                  Left            =   15
                  TabIndex        =   96
                  Top             =   390
                  Width           =   10260
                  _cx             =   18098
                  _cy             =   3228
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
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
                  Caption         =   "Attributes|Geo Sum|Reports"
                  Align           =   0
                  CurrTab         =   0
                  FirstTab        =   0
                  Style           =   2
                  Position        =   3
                  AutoSwitch      =   -1  'True
                  AutoScroll      =   -1  'True
                  TabPreview      =   0   'False
                  ShowFocusRect   =   -1  'True
                  TabsPerPage     =   3
                  BorderWidth     =   0
                  BoldCurrent     =   0   'False
                  DogEars         =   0   'False
                  MultiRow        =   -1  'True
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
                  Begin C1SizerLibCtl.C1Elastic elAttribCloseHolder 
                     Height          =   1800
                     Left            =   15
                     TabIndex        =   117
                     TabStop         =   0   'False
                     Top             =   15
                     Width           =   10230
                     _cx             =   18045
                     _cy             =   3175
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   204
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
                     GridRows        =   1
                     GridCols        =   2
                     Frame           =   3
                     FrameStyle      =   0
                     FrameWidth      =   1
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   $"frmMain.frx":22F64
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin C1SizerLibCtl.C1Elastic elSelCloseTools 
                        Height          =   1800
                        Left            =   9810
                        TabIndex        =   120
                        TabStop         =   0   'False
                        Top             =   0
                        Width           =   420
                        _cx             =   741
                        _cy             =   3175
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   204
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
                        Begin VB.CommandButton cmdSelTool 
                           Caption         =   "CL"
                           BeginProperty Font 
                              Name            =   "Terminal"
                              Size            =   9
                              Charset         =   255
                              Weight          =   400
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   375
                           Index           =   4
                           Left            =   0
                           Picture         =   "frmMain.frx":22FA3
                           TabIndex        =   126
                           ToolTipText     =   "Clear"
                           Top             =   1440
                           Width           =   420
                        End
                        Begin VB.CommandButton cmdSelTool 
                           BeginProperty Font 
                              Name            =   "Times New Roman"
                              Size            =   14.25
                              Charset         =   0
                              Weight          =   700
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   375
                           Index           =   3
                           Left            =   0
                           Picture         =   "frmMain.frx":297F5
                           Style           =   1  'Graphical
                           TabIndex        =   43
                           ToolTipText     =   "Selector Settings"
                           Top             =   1080
                           Width           =   420
                        End
                        Begin VB.CommandButton cmdSelTool 
                           BeginProperty Font 
                              Name            =   "Times New Roman"
                              Size            =   14.25
                              Charset         =   0
                              Weight          =   700
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   375
                           Index           =   2
                           Left            =   0
                           Picture         =   "frmMain.frx":30047
                           Style           =   1  'Graphical
                           TabIndex        =   123
                           ToolTipText     =   "Do reports"
                           Top             =   720
                           Width           =   420
                        End
                        Begin VB.CommandButton cmdSelTool 
                           BeginProperty Font 
                              Name            =   "Times New Roman"
                              Size            =   14.25
                              Charset         =   0
                              Weight          =   700
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   375
                           Index           =   1
                           Left            =   0
                           Picture         =   "frmMain.frx":36899
                           Style           =   1  'Graphical
                           TabIndex        =   122
                           ToolTipText     =   "Show Geo Summary"
                           Top             =   360
                           Width           =   420
                        End
                        Begin VB.CommandButton cmdSelTool 
                           BeginProperty Font 
                              Name            =   "Times New Roman"
                              Size            =   14.25
                              Charset         =   0
                              Weight          =   700
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   375
                           Index           =   0
                           Left            =   0
                           Picture         =   "frmMain.frx":3D0EB
                           Style           =   1  'Graphical
                           TabIndex        =   121
                           ToolTipText     =   "Close Current Info"
                           Top             =   0
                           Width           =   420
                        End
                     End
                     Begin C1SizerLibCtl.C1Tab AttribTabs 
                        Height          =   1800
                        Left            =   0
                        TabIndex        =   118
                        Top             =   0
                        Width           =   9810
                        _cx             =   17304
                        _cy             =   3175
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   204
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
                        Caption         =   ""
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
                        TabHeight       =   240
                        TabCaptionPos   =   4
                        TabPicturePos   =   0
                        CaptionEmpty    =   ""
                        Separators      =   0   'False
                        AccessibleName  =   ""
                        AccessibleDescription=   ""
                        AccessibleValue =   ""
                        AccessibleRole  =   37
                        Begin C1SizerLibCtl.C1Elastic elDynHolder 
                           Height          =   1470
                           Index           =   0
                           Left            =   45
                           TabIndex        =   119
                           TabStop         =   0   'False
                           Top             =   45
                           Width           =   9720
                           _cx             =   17145
                           _cy             =   2593
                           BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                              Name            =   "Arial"
                              Size            =   8.25
                              Charset         =   204
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
                           _GridInfo       =   $"frmMain.frx":4393D
                           AccessibleName  =   ""
                           AccessibleDescription=   ""
                           AccessibleValue =   ""
                           AccessibleRole  =   9
                           Begin TatukGIS_XDK10.XGIS_ControlAttributes SelAttributes1 
                              Height          =   1290
                              Index           =   0
                              Left            =   90
                              TabIndex        =   138
                              Top             =   90
                              Width           =   9540
                              ReadOnly        =   -1  'True
                              AllowRestructure=   -1  'True
                              ColorHeader     =   -16777201
                              ColorGrid       =   -16777211
                              Align           =   0
                              BevelInner      =   0
                              BevelOuter      =   0
                              Ctl3D           =   0   'False
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
                              Object.Visible         =   -1  'True
                              DoubleBuffered  =   0   'False
                              AllowNull       =   0   'False
                              BevelWidth      =   1
                              BorderWidth     =   0
                              HelpContextId   =   0
                              TabOrder        =   -1
                              TabStop         =   0   'False
                              UnitsEPSG       =   904202
                           End
                        End
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic elSelReports 
                     Height          =   1800
                     Left            =   10875
                     TabIndex        =   97
                     TabStop         =   0   'False
                     Top             =   15
                     Width           =   10230
                     _cx             =   18045
                     _cy             =   3175
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   204
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
                     Begin C1SizerLibCtl.C1Elastic C1Elastic22 
                        Height          =   1620
                        Left            =   420
                        TabIndex        =   98
                        TabStop         =   0   'False
                        Top             =   -60
                        Width           =   6255
                        _cx             =   11033
                        _cy             =   2858
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   204
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
                        Begin VB.Frame FraSelectionSettings 
                           Caption         =   "Selection Settings:"
                           Height          =   1035
                           Left            =   3780
                           TabIndex        =   114
                           Top             =   480
                           Width           =   2445
                           Begin VB.ComboBox txtSpatialOperation 
                              Height          =   315
                              ItemData        =   "frmMain.frx":4397D
                              Left            =   150
                              List            =   "frmMain.frx":43984
                              Style           =   2  'Dropdown List
                              TabIndex        =   116
                              Top             =   540
                              Width           =   2175
                           End
                           Begin VB.ComboBox ComBuffLevel 
                              Height          =   315
                              ItemData        =   "frmMain.frx":43993
                              Left            =   150
                              List            =   "frmMain.frx":439AF
                              Style           =   2  'Dropdown List
                              TabIndex        =   115
                              Top             =   210
                              Width           =   1335
                           End
                        End
                        Begin VB.CheckBox chkIncludeGeo 
                           Caption         =   "Include Geo ID"
                           Height          =   285
                           Left            =   120
                           TabIndex        =   106
                           Top             =   990
                           Width           =   1575
                        End
                        Begin VB.CommandButton cmdPrint 
                           Caption         =   "Print"
                           Height          =   285
                           Left            =   3780
                           TabIndex        =   105
                           Top             =   120
                           Width           =   1005
                        End
                        Begin VB.TextBox txtTitle 
                           Height          =   255
                           Left            =   1080
                           TabIndex        =   104
                           Top             =   120
                           Width           =   2655
                        End
                        Begin VB.CheckBox chkIncludeArea 
                           Caption         =   "Include Area"
                           Height          =   255
                           Left            =   120
                           TabIndex        =   103
                           Top             =   1260
                           Width           =   1305
                        End
                        Begin VB.CheckBox chkIncludeLength 
                           Caption         =   "Include Length"
                           Height          =   225
                           Left            =   1740
                           TabIndex        =   102
                           Top             =   1020
                           Width           =   1605
                        End
                        Begin VB.CheckBox chkIncludeCentroid 
                           Caption         =   "Include Centroid"
                           Height          =   285
                           Left            =   1740
                           TabIndex        =   101
                           Top             =   1260
                           Width           =   1485
                        End
                        Begin VB.TextBox txtMapTitle 
                           Height          =   315
                           Left            =   1020
                           TabIndex        =   100
                           Top             =   720
                           Width           =   2625
                        End
                        Begin VB.CheckBox chkIncludeMap 
                           Caption         =   "Include Map"
                           Height          =   285
                           Left            =   120
                           TabIndex        =   99
                           Top             =   420
                           Width           =   1245
                        End
                        Begin VB.Label lblTitle 
                           AutoSize        =   -1  'True
                           Caption         =   "Report Title:"
                           Height          =   195
                           Left            =   60
                           TabIndex        =   108
                           Top             =   180
                           Width           =   870
                        End
                        Begin VB.Label lblMapTitle 
                           AutoSize        =   -1  'True
                           Caption         =   "Map Title:"
                           Height          =   195
                           Left            =   90
                           TabIndex        =   107
                           Top             =   750
                           Width           =   705
                        End
                     End
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic elGisAttr 
               Height          =   2235
               Left            =   300
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   45
               Width           =   10290
               _cx             =   18150
               _cy             =   3942
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
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
               _GridInfo       =   $"frmMain.frx":439DB
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin ActiveBar3LibraryCtl.ActiveBar3 abGridTools 
                  Height          =   390
                  Left            =   0
                  TabIndex        =   90
                  Top             =   0
                  Width           =   10290
                  _LayoutVersion  =   2
                  _ExtentX        =   18150
                  _ExtentY        =   688
                  _DataPath       =   ""
                  Bands           =   "frmMain.frx":43A19
                  Begin TatukGIS_XDK10.XGIS_ControlScale Scale1 
                     Height          =   375
                     Left            =   8640
                     TabIndex        =   145
                     ToolTipText     =   "Scalebar"
                     Top             =   0
                     Width           =   1875
                     Dividers        =   5
                     Align           =   0
                     BevelInner      =   0
                     BevelOuter      =   0
                     BorderStyle     =   0
                     Color           =   -16777201
                     Ctl3D           =   -1  'True
                     Enabled         =   -1  'True
                     FullRepaint     =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     FontColor       =   -2147483630
                     Object.Visible         =   -1  'True
                     DoubleBuffered  =   0   'False
                     UnitsEPSG       =   904201
                  End
                  Begin VB.CheckBox chkFilterIn 
                     Caption         =   "Filter In Map"
                     Height          =   285
                     Left            =   6015
                     TabIndex        =   93
                     Top             =   75
                     Width           =   1185
                  End
                  Begin VB.CheckBox chkSelectIn 
                     Caption         =   "Select In Map"
                     Height          =   315
                     Left            =   7275
                     TabIndex        =   92
                     Top             =   60
                     Width           =   1305
                  End
                  Begin VB.CheckBox chkOnlyVisible 
                     Caption         =   "Load Visible"
                     Height          =   255
                     Left            =   4815
                     TabIndex        =   91
                     Top             =   90
                     Value           =   1  'Checked
                     Width           =   1215
                  End
               End
               Begin DXDBGRIDLibCtl.dxDBGrid dxGISDataGrid 
                  Height          =   1845
                  Left            =   0
                  OleObjectBlob   =   "frmMain.frx":45765
                  TabIndex        =   94
                  Top             =   390
                  Width           =   10290
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elTatukGIS 
            Height          =   4605
            Left            =   150
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   150
            Width           =   10635
            _cx             =   18759
            _cy             =   8123
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
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
            BorderWidth     =   0
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
            Begin TatukGIS_XDK10.XGIS_ControlNorthArrow watermark 
               Height          =   1035
               Left            =   6660
               TabIndex        =   144
               Top             =   2520
               Visible         =   0   'False
               Width           =   1275
               Symbol          =   0
               Transparent     =   -1  'True
               Path            =   ""
               Align           =   0
               BevelInner      =   0
               BevelOuter      =   0
               BorderStyle     =   0
               Color           =   -16777201
               Ctl3D           =   -1  'True
               Enabled         =   -1  'True
               FullRepaint     =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontColor       =   -2147483630
               Object.Visible         =   -1  'True
               DoubleBuffered  =   -1  'True
               Color2          =   0
               Color1          =   0
            End
            Begin TatukGIS_XDK10.XGIS_ViewerWnd GIS10 
               Height          =   3195
               Left            =   60
               TabIndex        =   137
               Top             =   0
               Width           =   3255
               BigExtentMargin =   -10
               RestrictedDrag  =   -1  'True
               CachedPaint     =   -1  'True
               IncrementalPaint=   -1  'True
               FullPaint       =   -1  'True
               CodePage        =   0
               OutCodePage     =   0
               CharSet         =   0
               UseRTree        =   0   'False
               PrinterTileSize =   2700
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
               SelectionPattern=   "frmMain.frx":463ED
               SelectionTransparency=   100
               SelectionWidth  =   100
               SelectionOutlineOnly=   0   'False
               OldCachedPaint  =   0   'False
               PrinterModeDraft=   0   'False
               PrinterModeForceBitmap=   0   'False
               GDIType         =   1
               ScaleAsFloat    =   1
               Mode            =   2
               BorderStyle     =   0
               CursorForDrag   =   0
               CursorForSelect =   0
               CursorForZoom   =   0
               CursorForEdit   =   0
               MinZoomSize     =   -5
               ScrollBars      =   0
               AutoCenter      =   0   'False
               Align           =   0
               Ctl3D           =   -1  'True
               Object.Visible         =   -1  'True
               Cursor          =   18
               DoubleBuffered  =   0   'False
               ModeMouseButton =   0
               CursorForUserDefined=   0
               View3D          =   0   'False
            End
            Begin TatukGIS_XDK10.XGIS_ControlNorthArrow NArrow 
               Height          =   1095
               Left            =   6960
               TabIndex        =   143
               Top             =   0
               Width           =   1575
               Symbol          =   7
               Transparent     =   -1  'True
               Path            =   ""
               Align           =   0
               BevelInner      =   0
               BevelOuter      =   2
               BorderStyle     =   0
               Color           =   -16777201
               Ctl3D           =   -1  'True
               Enabled         =   -1  'True
               FullRepaint     =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontColor       =   -2147483630
               Object.Visible         =   -1  'True
               DoubleBuffered  =   -1  'True
               Color2          =   0
               Color1          =   0
            End
            Begin OASISClient.ctlZoomSlider ctlZoomSlider1 
               Height          =   2235
               Left            =   3720
               TabIndex        =   148
               Top             =   0
               Width           =   250
               _ExtentX        =   873
               _ExtentY        =   3942
            End
            Begin C1SizerLibCtl.C1Tab C1TabFastFunction 
               Height          =   3915
               Left            =   0
               TabIndex        =   12
               Top             =   0
               Width           =   2835
               _cx             =   5001
               _cy             =   6906
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   204
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
               Caption         =   "Themes|Geomarks|GeoConvert|Go To|Settings|Magnify|<<"
               Align           =   0
               CurrTab         =   2
               FirstTab        =   0
               Style           =   3
               Position        =   3
               AutoSwitch      =   -1  'True
               AutoScroll      =   -1  'True
               TabPreview      =   -1  'True
               ShowFocusRect   =   -1  'True
               TabsPerPage     =   7
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
               Flags(3)        =   2
               Begin C1SizerLibCtl.C1Elastic elGeoCalc 
                  Height          =   3825
                  Left            =   45
                  TabIndex        =   34
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   2475
                  _cx             =   4366
                  _cy             =   6747
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
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
                  Begin VB.Frame frmdddddd 
                     Caption         =   "DDD.DDDDD"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   6.75
                        Charset         =   204
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   960
                     Left            =   30
                     TabIndex        =   83
                     Top             =   690
                     Width           =   2310
                     Begin MSMask.MaskEdBox me2Long 
                        Height          =   300
                        Left            =   990
                        TabIndex        =   84
                        Top             =   510
                        Width           =   1215
                        _ExtentX        =   2143
                        _ExtentY        =   529
                        _Version        =   393216
                        MaxLength       =   9
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   204
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Mask            =   "###.#####"
                        PromptChar      =   "_"
                     End
                     Begin MSMask.MaskEdBox me2Lat 
                        Height          =   300
                        Left            =   990
                        TabIndex        =   85
                        Top             =   225
                        Width           =   1215
                        _ExtentX        =   2143
                        _ExtentY        =   529
                        _Version        =   393216
                        MaxLength       =   8
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   204
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Mask            =   "##.#####"
                        PromptChar      =   "_"
                     End
                     Begin VB.Label Label2 
                        AutoSize        =   -1  'True
                        Caption         =   "Longitude:"
                        Height          =   195
                        Index           =   1
                        Left            =   135
                        TabIndex        =   87
                        Top             =   630
                        Width           =   750
                     End
                     Begin VB.Label Label1 
                        AutoSize        =   -1  'True
                        Caption         =   "Latitude:"
                        Height          =   195
                        Index           =   1
                        Left            =   135
                        TabIndex        =   86
                        Top             =   315
                        Width           =   615
                     End
                  End
                  Begin VB.Frame frMGRS 
                     Caption         =   "MGRS"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   6.75
                        Charset         =   204
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   555
                     Left            =   0
                     TabIndex        =   81
                     Top             =   3030
                     Visible         =   0   'False
                     Width           =   2265
                     Begin VB.TextBox txtMGRS 
                        Height          =   285
                        Left            =   90
                        TabIndex        =   82
                        Text            =   "MGRS"
                        Top             =   225
                        Width           =   2130
                     End
                  End
                  Begin VB.Frame frddmmss 
                     Caption         =   "DDD:MM:SS.SS"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   6.75
                        Charset         =   204
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   960
                     Left            =   30
                     TabIndex        =   76
                     Top             =   1770
                     Visible         =   0   'False
                     Width           =   2265
                     Begin MSMask.MaskEdBox me1Long 
                        Height          =   330
                        Left            =   945
                        TabIndex        =   77
                        Top             =   540
                        Width           =   1260
                        _ExtentX        =   2223
                        _ExtentY        =   582
                        _Version        =   393216
                        MaxLength       =   12
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   204
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Mask            =   "###:##:##.##"
                        PromptChar      =   "_"
                     End
                     Begin MSMask.MaskEdBox me1Lat 
                        Height          =   330
                        Left            =   945
                        TabIndex        =   78
                        Top             =   225
                        Width           =   1260
                        _ExtentX        =   2223
                        _ExtentY        =   582
                        _Version        =   393216
                        MaxLength       =   11
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   204
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Mask            =   "##:##:##.##"
                        PromptChar      =   "_"
                     End
                     Begin VB.Label Label1 
                        AutoSize        =   -1  'True
                        Caption         =   "Latitude:"
                        Height          =   195
                        Index           =   0
                        Left            =   135
                        TabIndex        =   80
                        Top             =   315
                        Width           =   615
                     End
                     Begin VB.Label Label2 
                        AutoSize        =   -1  'True
                        Caption         =   "Longitude:"
                        Height          =   195
                        Index           =   0
                        Left            =   135
                        TabIndex        =   79
                        Top             =   630
                        Width           =   750
                     End
                  End
                  Begin VB.Frame frddmmmm 
                     Caption         =   "DDD:MM.MMMM"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   6.75
                        Charset         =   204
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   960
                     Left            =   0
                     TabIndex        =   71
                     Top             =   2070
                     Visible         =   0   'False
                     Width           =   2265
                     Begin MSMask.MaskEdBox me3Long 
                        Height          =   330
                        Left            =   945
                        TabIndex        =   72
                        Top             =   540
                        Width           =   1275
                        _ExtentX        =   2249
                        _ExtentY        =   582
                        _Version        =   393216
                        MaxLength       =   11
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   204
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Mask            =   "###:##.####"
                        PromptChar      =   "_"
                     End
                     Begin MSMask.MaskEdBox me3Lat 
                        Height          =   330
                        Left            =   945
                        TabIndex        =   73
                        Top             =   225
                        Width           =   1275
                        _ExtentX        =   2249
                        _ExtentY        =   582
                        _Version        =   393216
                        MaxLength       =   10
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "Arial"
                           Size            =   8.25
                           Charset         =   204
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Mask            =   "##:##.####"
                        PromptChar      =   "_"
                     End
                     Begin VB.Label Label1 
                        AutoSize        =   -1  'True
                        Caption         =   "Latitude:"
                        Height          =   195
                        Index           =   2
                        Left            =   90
                        TabIndex        =   75
                        Top             =   315
                        Width           =   615
                     End
                     Begin VB.Label Label2 
                        AutoSize        =   -1  'True
                        Caption         =   "Longitude:"
                        Height          =   195
                        Index           =   2
                        Left            =   90
                        TabIndex        =   74
                        Top             =   630
                        Width           =   750
                     End
                  End
                  Begin VB.CommandButton cmdPolyEdit 
                     Height          =   330
                     Left            =   930
                     MaskColor       =   &H0000C000&
                     Picture         =   "frmMain.frx":4657F
                     Style           =   1  'Graphical
                     TabIndex        =   63
                     ToolTipText     =   "Create scribble Area"
                     Top             =   3480
                     UseMaskColor    =   -1  'True
                     Width           =   450
                  End
                  Begin VB.CommandButton cmdZoomToSettings 
                     Caption         =   "..."
                     Height          =   255
                     Left            =   1800
                     TabIndex        =   62
                     Top             =   450
                     Width           =   525
                  End
                  Begin VB.TextBox txtFont 
                     Appearance      =   0  'Flat
                     BackColor       =   &H8000000F&
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "Symbol"
                        Size            =   15.75
                        Charset         =   2
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   345
                     Index           =   0
                     Left            =   1230
                     Locked          =   -1  'True
                     TabIndex        =   54
                     Text            =   "5"
                     Top             =   360
                     Width           =   405
                  End
                  Begin VB.Frame FraPinColor 
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   0  'None
                     Height          =   225
                     Index           =   0
                     Left            =   930
                     TabIndex        =   53
                     Top             =   420
                     Width           =   255
                  End
                  Begin VB.CommandButton cmdRemove 
                     Height          =   330
                     Left            =   1410
                     MaskColor       =   &H000000FF&
                     Picture         =   "frmMain.frx":468C1
                     Style           =   1  'Graphical
                     TabIndex        =   52
                     ToolTipText     =   "Remove Markers"
                     Top             =   3480
                     Width           =   450
                  End
                  Begin VB.CheckBox chkUseMarker 
                     Caption         =   "Marker"
                     Height          =   195
                     Left            =   90
                     TabIndex        =   48
                     Top             =   450
                     Width           =   825
                  End
                  Begin VB.TextBox txtConversionResutl 
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   6
                        Charset         =   204
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1770
                     Left            =   60
                     Locked          =   -1  'True
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   38
                     Top             =   1650
                     Width           =   2265
                  End
                  Begin VB.ComboBox comConversionType 
                     Height          =   315
                     ItemData        =   "frmMain.frx":4D113
                     Left            =   660
                     List            =   "frmMain.frx":4D123
                     Style           =   2  'Dropdown List
                     TabIndex        =   37
                     Top             =   60
                     Width           =   1710
                  End
                  Begin VB.CommandButton cmdZoomTo 
                     Height          =   330
                     Left            =   1890
                     Picture         =   "frmMain.frx":4D14F
                     Style           =   1  'Graphical
                     TabIndex        =   36
                     ToolTipText     =   "Zoom to/Create Marker"
                     Top             =   3480
                     Width           =   450
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Format:"
                     Height          =   195
                     Left            =   90
                     TabIndex        =   39
                     Top             =   150
                     Width           =   585
                  End
               End
               Begin C1SizerLibCtl.C1Elastic elMagnifier 
                  Height          =   3825
                  Left            =   4080
                  TabIndex        =   35
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   2475
                  _cx             =   4366
                  _cy             =   6747
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
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
                  Begin VB.Timer tmrUpdate 
                     Enabled         =   0   'False
                     Interval        =   50
                     Left            =   540
                     Top             =   315
                  End
               End
               Begin VB.Frame FraFmThemeSettings 
                  Caption         =   "Analysis"
                  Height          =   3825
                  Left            =   3780
                  TabIndex        =   22
                  Top             =   45
                  Width           =   2475
                  Begin VB.CommandButton cmdLocationAnalysis 
                     Caption         =   "Location Analysis"
                     Height          =   240
                     Left            =   90
                     TabIndex        =   30
                     Top             =   3510
                     Visible         =   0   'False
                     Width           =   2235
                  End
                  Begin VB.OptionButton OptThemeBoundary 
                     Caption         =   "OASIS Security Cells"
                     Height          =   375
                     Index           =   2
                     Left            =   90
                     TabIndex        =   29
                     Top             =   1350
                     Visible         =   0   'False
                     Width           =   2220
                  End
                  Begin VB.CheckBox chkDynamicUpdate 
                     Caption         =   "Dynamic Update"
                     Height          =   420
                     Left            =   90
                     TabIndex        =   28
                     Top             =   1755
                     Visible         =   0   'False
                     Width           =   1605
                  End
                  Begin VB.OptionButton OptThemeBoundary 
                     Caption         =   "Urban Administrative Areas"
                     Height          =   555
                     Index           =   1
                     Left            =   90
                     TabIndex        =   27
                     Top             =   675
                     Visible         =   0   'False
                     Width           =   2190
                  End
                  Begin VB.OptionButton OptThemeBoundary 
                     Caption         =   "Administrave borders"
                     Height          =   375
                     Index           =   0
                     Left            =   90
                     TabIndex        =   26
                     Top             =   225
                     Value           =   -1  'True
                     Visible         =   0   'False
                     Width           =   2190
                  End
                  Begin VB.CommandButton cmdCommand1 
                     Caption         =   "Activate"
                     Height          =   465
                     Left            =   90
                     TabIndex        =   25
                     Top             =   2655
                     Visible         =   0   'False
                     Width           =   2235
                  End
                  Begin VB.CommandButton cmdSetScoring 
                     Caption         =   "Set Scoring"
                     Height          =   465
                     Left            =   90
                     TabIndex        =   24
                     Top             =   3240
                     Width           =   2235
                  End
                  Begin VB.CommandButton cmdTimeAnalysis 
                     Caption         =   "Time Analysis"
                     Height          =   240
                     Left            =   90
                     TabIndex        =   23
                     Top             =   3285
                     Visible         =   0   'False
                     Width           =   2235
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic1 
                     Height          =   3810
                     Left            =   0
                     TabIndex        =   42
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   2475
                     _cx             =   4366
                     _cy             =   6720
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   204
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
               End
               Begin C1SizerLibCtl.C1Elastic Themes 
                  Height          =   3825
                  Left            =   -3390
                  TabIndex        =   13
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   2475
                  _cx             =   4366
                  _cy             =   6747
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
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
                  Begin VB.CommandButton cmdAddBookmarks 
                     Caption         =   "*"
                     Height          =   210
                     Left            =   120
                     TabIndex        =   14
                     Top             =   60
                     Width           =   480
                  End
                  Begin TreeViewLibCtl.ctTree ctTreeBookmrks 
                     Height          =   3225
                     Left            =   90
                     TabIndex        =   31
                     Top             =   315
                     Width           =   2325
                     _Version        =   393216
                     _ExtentX        =   4101
                     _ExtentY        =   5689
                     _StockProps     =   68
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   204
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PictureClose    =   "frmMain.frx":539A1
                     PictureMinus    =   "frmMain.frx":539BD
                     PictureOpen     =   "frmMain.frx":539D9
                     PicturePlus     =   "frmMain.frx":539F5
                     PictureLeaf     =   "frmMain.frx":53A11
                     TitleBackImage  =   "frmMain.frx":53A2D
                     CheckPicDown    =   "frmMain.frx":53A49
                     CheckPicUp      =   "frmMain.frx":53A65
                     CheckPicDisabled=   "frmMain.frx":53A81
                     RadioPicDown    =   "frmMain.frx":53A9D
                     RadioPicUp      =   "frmMain.frx":53AB9
                     BackImage       =   "frmMain.frx":53AD5
                     MouseIcon       =   "frmMain.frx":53AF1
                     NoFocusBackColor=   -2147483633
                     HeaderData      =   "frmMain.frx":53B0D
                     PicArray0       =   "frmMain.frx":53B35
                     PicArray1       =   "frmMain.frx":53B51
                     PicArray2       =   "frmMain.frx":53B6D
                     PicArray3       =   "frmMain.frx":53B89
                     PicArray4       =   "frmMain.frx":53BA5
                     PicArray5       =   "frmMain.frx":53BC1
                     PicArray6       =   "frmMain.frx":53BDD
                     PicArray7       =   "frmMain.frx":53BF9
                     PicArray8       =   "frmMain.frx":53C15
                     PicArray9       =   "frmMain.frx":53C31
                     PicArray10      =   "frmMain.frx":53C4D
                     PicArray11      =   "frmMain.frx":53C69
                     PicArray12      =   "frmMain.frx":53C85
                     PicArray13      =   "frmMain.frx":53CA1
                     PicArray14      =   "frmMain.frx":53CBD
                     PicArray15      =   "frmMain.frx":53CD9
                     PicArray16      =   "frmMain.frx":53CF5
                     PicArray17      =   "frmMain.frx":53D11
                     PicArray18      =   "frmMain.frx":53D2D
                     PicArray19      =   "frmMain.frx":53D49
                     PicArray20      =   "frmMain.frx":53D65
                     PicArray21      =   "frmMain.frx":53D81
                     PicArray22      =   "frmMain.frx":53D9D
                     PicArray23      =   "frmMain.frx":53DB9
                     PicArray24      =   "frmMain.frx":53DD5
                     PicArray25      =   "frmMain.frx":53DF1
                     PicArray26      =   "frmMain.frx":53E0D
                     PicArray27      =   "frmMain.frx":53E29
                     PicArray28      =   "frmMain.frx":53E45
                     PicArray29      =   "frmMain.frx":53E61
                     PicArray30      =   "frmMain.frx":53E7D
                     PicArray31      =   "frmMain.frx":53E99
                     PicArray32      =   "frmMain.frx":53EB5
                     PicArray33      =   "frmMain.frx":53ED1
                     PicArray34      =   "frmMain.frx":53EED
                     PicArray35      =   "frmMain.frx":53F09
                     PicArray36      =   "frmMain.frx":53F25
                     PicArray37      =   "frmMain.frx":53F41
                     PicArray38      =   "frmMain.frx":53F5D
                     PicArray39      =   "frmMain.frx":53F79
                     PicArray40      =   "frmMain.frx":53F95
                     PicArray41      =   "frmMain.frx":53FB1
                     PicArray42      =   "frmMain.frx":53FCD
                     PicArray43      =   "frmMain.frx":53FE9
                     PicArray44      =   "frmMain.frx":54005
                     PicArray45      =   "frmMain.frx":54021
                     PicArray46      =   "frmMain.frx":5403D
                     PicArray47      =   "frmMain.frx":54059
                     PicArray48      =   "frmMain.frx":54075
                     PicArray49      =   "frmMain.frx":54091
                     PicArray50      =   "frmMain.frx":540AD
                     PicArray51      =   "frmMain.frx":540C9
                     PicArray52      =   "frmMain.frx":540E5
                     PicArray53      =   "frmMain.frx":54101
                     PicArray54      =   "frmMain.frx":5411D
                     PicArray55      =   "frmMain.frx":54139
                     PicArray56      =   "frmMain.frx":54155
                     PicArray57      =   "frmMain.frx":54171
                     PicArray58      =   "frmMain.frx":5418D
                     PicArray59      =   "frmMain.frx":541A9
                     PicArray60      =   "frmMain.frx":541C5
                     PicArray61      =   "frmMain.frx":541E1
                     PicArray62      =   "frmMain.frx":541FD
                     PicArray63      =   "frmMain.frx":54219
                     PicArray64      =   "frmMain.frx":54235
                     PicArray65      =   "frmMain.frx":54251
                     PicArray66      =   "frmMain.frx":5426D
                     PicArray67      =   "frmMain.frx":54289
                     PicArray68      =   "frmMain.frx":542A5
                     PicArray69      =   "frmMain.frx":542C1
                     PicArray70      =   "frmMain.frx":542DD
                     PicArray71      =   "frmMain.frx":542F9
                     PicArray72      =   "frmMain.frx":54315
                     PicArray73      =   "frmMain.frx":54331
                     PicArray74      =   "frmMain.frx":5434D
                     PicArray75      =   "frmMain.frx":54369
                     PicArray76      =   "frmMain.frx":54385
                     PicArray77      =   "frmMain.frx":543A1
                     PicArray78      =   "frmMain.frx":543BD
                     PicArray79      =   "frmMain.frx":543D9
                     PicArray80      =   "frmMain.frx":543F5
                     PicArray81      =   "frmMain.frx":54411
                     PicArray82      =   "frmMain.frx":5442D
                     PicArray83      =   "frmMain.frx":54449
                     PicArray84      =   "frmMain.frx":54465
                     PicArray85      =   "frmMain.frx":54481
                     PicArray86      =   "frmMain.frx":5449D
                     PicArray87      =   "frmMain.frx":544B9
                     PicArray88      =   "frmMain.frx":544D5
                     PicArray89      =   "frmMain.frx":544F1
                     PicArray90      =   "frmMain.frx":5450D
                     PicArray91      =   "frmMain.frx":54529
                     PicArray92      =   "frmMain.frx":54545
                     PicArray93      =   "frmMain.frx":54561
                     PicArray94      =   "frmMain.frx":5457D
                     PicArray95      =   "frmMain.frx":54599
                     PicArray96      =   "frmMain.frx":545B5
                     PicArray97      =   "frmMain.frx":545D1
                     PicArray98      =   "frmMain.frx":545ED
                     PicArray99      =   "frmMain.frx":54609
                  End
                  Begin VB.Label lblGeoMarks 
                     Caption         =   "Geo Marks"
                     Height          =   180
                     Left            =   780
                     TabIndex        =   15
                     Top             =   60
                     Width           =   990
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Bookmarks 
                  Height          =   3825
                  Left            =   -3690
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   2475
                  _cx             =   4366
                  _cy             =   6747
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
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
                  Begin VB.ComboBox ComThemes 
                     Height          =   315
                     ItemData        =   "frmMain.frx":54625
                     Left            =   45
                     List            =   "frmMain.frx":54656
                     Style           =   2  'Dropdown List
                     TabIndex        =   20
                     Top             =   345
                     Width           =   2325
                  End
                  Begin VB.Frame FraDetails 
                     Caption         =   "Details"
                     Height          =   2220
                     Left            =   75
                     TabIndex        =   18
                     Top             =   720
                     Width           =   2295
                     Begin VB.TextBox txtThemeDescription 
                        Height          =   1875
                        Left            =   105
                        MultiLine       =   -1  'True
                        TabIndex        =   19
                        Text            =   "frmMain.frx":547BC
                        Top             =   270
                        Width           =   2100
                     End
                  End
                  Begin VB.CommandButton cmdActivateTheme 
                     Caption         =   "Activate Theme"
                     Height          =   360
                     Left            =   90
                     TabIndex        =   17
                     Top             =   3150
                     Width           =   2205
                  End
                  Begin VB.Label lblAvailableThemes 
                     Caption         =   "Available Themes:"
                     Height          =   270
                     Left            =   120
                     TabIndex        =   21
                     Top             =   90
                     Width           =   1500
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1EGOTO 
                  Height          =   3825
                  Left            =   3480
                  TabIndex        =   32
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   2475
                  _cx             =   4366
                  _cy             =   6747
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
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
                  Caption         =   "GOTO"
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
            End
            Begin C1SizerLibCtl.C1Elastic elScroller 
               Height          =   345
               Left            =   30
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   3870
               Width           =   10185
               _cx             =   17965
               _cy             =   609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
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
               GridRows        =   1
               GridCols        =   2
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmMain.frx":547CE
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin OASISClient.MsgScroller MsgScroll 
                  Height          =   345
                  Left            =   270
                  Top             =   0
                  Width           =   9915
                  _ExtentX        =   17489
                  _ExtentY        =   609
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
               Begin VB.CommandButton cmdMSGScroller 
                  Appearance      =   0  'Flat
                  Caption         =   "5"
                  BeginProperty Font 
                     Name            =   "Marlett"
                     Size            =   15.75
                     Charset         =   2
                     Weight          =   500
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   0
                  TabIndex        =   56
                  Top             =   0
                  Width           =   270
               End
            End
         End
      End
      Begin CONTROLSLibCtl.dxProgressBar dxProgressBar1 
         Height          =   495
         Left            =   5580
         TabIndex        =   60
         Top             =   6480
         Visible         =   0   'False
         Width           =   975
         _Version        =   65536
         _cx             =   1720
         _cy             =   873
         ForeColor       =   0
         BackColor       =   15790320
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MinPos          =   0
         MaxPos          =   100
         Pos             =   50
         Step            =   10
         ShowText        =   -1  'True
         Orientation     =   0
         StartColor      =   16711680
         EndColor        =   16777215
         DrawBorderStyle =   1
         ShowTextStyle   =   0
         DrawBarStyle    =   2
         DrawBarBorderStyle=   2
      End
   End
   Begin VB.Menu mnuMAPPopUp 
      Caption         =   "mapPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuAddToClipboard 
         Caption         =   "Add To Clipboard"
      End
      Begin VB.Menu mnuZoomOUT 
         Caption         =   "Zoom Out"
      End
      Begin VB.Menu mnuZoomIN 
         Caption         =   "Zoom In"
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "Previous View"
      End
      Begin VB.Menu mnuGetCoords 
         Caption         =   "Get Coordinates"
      End
      Begin VB.Menu mnuSendToSMS 
         Caption         =   "Send To SMS"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOPSSetting 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuMapTip 
         Caption         =   "Map Tip"
      End
      Begin VB.Menu mnumapToolSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoomRectanletool 
         Caption         =   "Zoom Rectangle Tool"
      End
      Begin VB.Menu mnuPanTool 
         Caption         =   "Pan Tool"
      End
      Begin VB.Menu mnuOPSView 
         Caption         =   "Full Screen"
      End
      Begin VB.Menu mnuMapSpes 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLayer 
         Caption         =   "View Layer in Grid"
         Begin VB.Menu mnuLayer 
            Caption         =   "--Nothing--"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuGridAction 
      Caption         =   "dxGridAction"
      Visible         =   0   'False
      Begin VB.Menu mnuMapAction 
         Caption         =   "Map Actions"
         Begin VB.Menu mnuZoomTo 
            Caption         =   "ZoomTo"
         End
         Begin VB.Menu mnuSelectInMap 
            Caption         =   "Select In Map"
         End
         Begin VB.Menu mnuClearSelections 
            Caption         =   "Clear Selections"
         End
         Begin VB.Menu mnuFlash 
            Caption         =   "Flash"
         End
         Begin VB.Menu mnuAutoZoom 
            Caption         =   "Auto Zoom"
         End
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShowSelected 
         Caption         =   "Show Selected in Map View"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHideSelected 
         Caption         =   "Hide Selected from Map View"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGridSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDataViewActions 
         Caption         =   "Data View Actions"
         Begin VB.Menu mnuRec2Clip 
            Caption         =   "Add record to clipboard"
         End
         Begin VB.Menu mnuLoadAll 
            Caption         =   "Load All"
         End
         Begin VB.Menu mnuVisibleExtent 
            Caption         =   "Load Only Visible"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuShowRecordCount 
            Caption         =   "Show Record Counts"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuAutoHeight 
            Caption         =   "Auto Size Memo Field Height"
         End
         Begin VB.Menu mnuAutoSize 
            Caption         =   "Auto Size Field Width"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuadvGisGrid 
            Caption         =   "Advanced settings"
            Begin VB.Menu mnu4Rows 
               Caption         =   "4 Rows"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnu8rows 
               Caption         =   "8 Rows"
            End
            Begin VB.Menu mnu16rows 
               Caption         =   "16 Rows"
            End
            Begin VB.Menu mnuUnlimitedRows 
               Caption         =   "Unlimited"
            End
            Begin VB.Menu mnuShowCentroid 
               Caption         =   "Show centroid in data view"
            End
         End
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGridSettings 
         Caption         =   "Settings"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuSelector 
      Caption         =   "selector menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSelFlash 
         Caption         =   "Flash"
      End
      Begin VB.Menu mnuSelZoomTo 
         Caption         =   "Zoom To"
      End
      Begin VB.Menu mnuSelSelect 
         Caption         =   "Select/Deselect in Map"
      End
      Begin VB.Menu mnuSelEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuSelSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAutoClear 
         Caption         =   "Automatic Clear"
      End
      Begin VB.Menu mnuSelAutoFlash 
         Caption         =   "Auto Flash"
      End
      Begin VB.Menu mnuSelAutoSelect 
         Caption         =   "Auto Select"
      End
      Begin VB.Menu mnuSelAutoZoom 
         Caption         =   "Auto Zoom"
      End
      Begin VB.Menu mnuSelSel1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelSettings 
         Caption         =   "Selection Settings..."
      End
   End
   Begin VB.Menu mnuSelToolPopUp 
      Caption         =   "SelToolPopUp"
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'TODO Remove Below''
Dim lLOG As Long

Private Declare Function GetWindowRect _
                Lib "user32" (ByVal hWnd As Long, _
                              lpRect As RECT) As Long
Private Declare Function GetTickCount _
                Lib "kernel32" () As Long
                                
Private m_SubmittedPending As Boolean
Private WithEvents m_frmSpatialAnalysis As frmSpatialAnalysis
Attribute m_frmSpatialAnalysis.VB_VarHelpID = -1
'Attribute m_frmSpatialAnalysis.VB_VarHelpID = -1
Private WithEvents EventLayer As TatukGIS_XDK10.XGIS_LayerVector
Attribute EventLayer.VB_VarHelpID = -1
'Attribute EventLayer.VB_VarHelpID = -1
Private WithEvents m_frmOPSView As frmOPSView
Attribute m_frmOPSView.VB_VarHelpID = -1
'Attribute m_frmOPSView.VB_VarHelpID = -1
Private WithEvents m_frmOvMap As frmOVMap
Attribute m_frmOvMap.VB_VarHelpID = -1
'Attribute m_frmOvMap.VB_VarHelpID = -1
Private WithEvents m_fmrAddIncident As frmAddIncident
Attribute m_fmrAddIncident.VB_VarHelpID = -1
'Attribute m_fmrAddIncident.VB_VarHelpID = -1
Private m_frmCannedReports As frmCannedReports
'Attribute m_frmCannedReports.VB_VarHelpID = -1
'Private WithEvents m_frmW3Wizard As frmW3Wizard
'Attribute m_frmW3Wizard.VB_VarHelpID = -1
Private WithEvents m_frmMnuOASISProfile As frmMnuOASISProfile
Attribute m_frmMnuOASISProfile.VB_VarHelpID = -1
'Attribute m_frmMnuOASISProfile.VB_VarHelpID = -1
Private WithEvents m_frmAddOns As frmAddons
Attribute m_frmAddOns.VB_VarHelpID = -1
'Attribute m_frmAddOns.VB_VarHelpID = -1
Private WithEvents m_frmChangeTracer As frmChangeTracer
Attribute m_frmChangeTracer.VB_VarHelpID = -1
'Attribute m_frmChangeTracer.VB_VarHelpID = -1
Private WithEvents m_frmFreqSettings As frmFreqSettings
Attribute m_frmFreqSettings.VB_VarHelpID = -1
'Attribute m_frmFreqSettings.VB_VarHelpID = -1
Private WithEvents m_frmMnuOperations As frmMnuOperations
Attribute m_frmMnuOperations.VB_VarHelpID = -1
'Attribute m_frmMnuOperations.VB_VarHelpID = -1
Private WithEvents m_frmMnuDynamicDataModule As frmMnuDynamicDataModule
Attribute m_frmMnuDynamicDataModule.VB_VarHelpID = -1
'Attribute m_frmMnuDynamicDataModule.VB_VarHelpID = -1
Private WithEvents m_frmMnuDynamicReportsModule As frmMnuDynamicReportsModule
Attribute m_frmMnuDynamicReportsModule.VB_VarHelpID = -1
'Attribute m_frmMnuDynamicReportsModule.VB_VarHelpID = -1
Private m_frmLocator As frmLocator
'Attribute m_frmLocator.VB_VarHelpID = -1
Private WithEvents m_frmOASISCharts As frmOASISCharts
Attribute m_frmOASISCharts.VB_VarHelpID = -1
'Attribute m_frmOASISCharts.VB_VarHelpID = -1
Private WithEvents m_frmUpdateSettings As frmUpdateSettings
Attribute m_frmUpdateSettings.VB_VarHelpID = -1
'Attribute m_frmUpdateSettings.VB_VarHelpID = -1
Private WithEvents g_clsHotKey As cRegHotKey
Attribute g_clsHotKey.VB_VarHelpID = -1
'Attribute g_clsHotKey.VB_VarHelpID = -1
Private WithEvents m_frmTextAnnoSettings As frmTextAnnoSettings
Attribute m_frmTextAnnoSettings.VB_VarHelpID = -1
'Attribute m_frmTextAnnoSettings.VB_VarHelpID = -1
Private WithEvents m_oClipboardViewer As cClipboardViewer
Attribute m_oClipboardViewer.VB_VarHelpID = -1
'Attribute m_oClipboardViewer.VB_VarHelpID = -1
Private WithEvents m_frmMainSettings As frmMainSettings
Attribute m_frmMainSettings.VB_VarHelpID = -1
'Attribute m_frmMainSettings.VB_VarHelpID = -1
Private WithEvents m_frmSelectorReports As frmSelectorReports
Attribute m_frmSelectorReports.VB_VarHelpID = -1
'Attribute m_frmSelectorReports.VB_VarHelpID = -1
Private WithEvents m_frmSelectorSettings As frmSelectorSettings
Attribute m_frmSelectorSettings.VB_VarHelpID = -1
'Attribute m_frmSelectorSettings.VB_VarHelpID = -1
Private WithEvents m_frmDynamicContent As frmDynamicContent
Attribute m_frmDynamicContent.VB_VarHelpID = -1
'Attribute m_frmDynamicContent.VB_VarHelpID = -1
Private WithEvents m_frmSpatialiseDD As frmSpatialiseDD
Attribute m_frmSpatialiseDD.VB_VarHelpID = -1
Private m_frmIncidentsV2SitrepGenerator As frmIncidentsV2SitrepGenerator
Private WithEvents m_frmIncidentsV2DataEntry As frmIncidentsV2DataEntry
Attribute m_frmIncidentsV2DataEntry.VB_VarHelpID = -1
Private WithEvents m_frmIncidentsV2ControlPanel As frmIncidentsV2ControlPanel
Attribute m_frmIncidentsV2ControlPanel.VB_VarHelpID = -1
'Attribute m_frmSpatialiseDD.VB_VarHelpID = -1
Private ClipBoard_xt As cCustomClipboard

Private WithEvents m_frmMapLibraryDLG As frmMapLibraryDLG
Attribute m_frmMapLibraryDLG.VB_VarHelpID = -1

Private m_frmAttributes As frmAttributes
Private m_frmSearch As frmSearch
Private m_bSelectRadiusTool As Boolean
Private m_oCosmeticLayer As TatukGIS_XDK10.XGIS_LayerVector
Private m_oDrawLyr As TatukGIS_XDK10.XGIS_LayerSqlAdo
Private m_oBufferLyr As TatukGIS_XDK10.XGIS_LayerVector

'Both for spatialise DD tool
Private m_oSpatialiseLayer As TatukGIS_XDK10.XGIS_LayerVector
Private mDDLayer As TatukGIS_XDK10.XGIS_LayerVector


Private m_oIncidentLyr As TatukGIS_XDK10.XGIS_LayerVector
Private m_oW3Lyr As TatukGIS_XDK10.XGIS_LayerVector
Private m_oW3WHOLyr As TatukGIS_XDK10.XGIS_LayerVector
Private m_oSQLIncLyr As TatukGIS_XDK10.XGIS_LayerSqlAdo
Private m_oSQLOpsLyr As TatukGIS_XDK10.XGIS_LayerSqlAdo
Private m_oIncRS As ADODB.Recordset
Private FirstRun As Boolean
Private g_oEditLayer As TatukGIS_XDK10.XGIS_LayerAbstract
Private m_oSQLGenericLyrs() As TatukGIS_XDK10.XGIS_LayerSqlAdo

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private m_prevTool As OASIS_TOOLS
Private m_bUseDistrictOnly As Boolean
Private m_oPtRubberStart As New TatukGIS_XDK10.XGIS_Point
Private m_oLineRubber As New TatukGIS_XDK10.XGIS_ShapeArc
Private m_bShowRubberBand As Boolean
Private m_ShapeTemp As XGIS_Shape
Private m_bLOADING As Boolean

'OASIS FEATURES
Private m_oShpArc As TatukGIS_XDK10.XGIS_ShapeArc
Private m_oShpPoint As TatukGIS_XDK10.XGIS_ShapePoint
Private m_oShpPolygon As TatukGIS_XDK10.XGIS_ShapePolygon
Private m_oShpUnknown As TatukGIS_XDK10.XGIS_Shape
Private m_oIncShpPt As TatukGIS_XDK10.XGIS_ShapePoint

Private m_sAdmVal1 As String
Private m_sAdmVal2 As String
Private m_sAdmLoc As String

Private m_lZoomSelectColour As Long
        Private m_lZoomBackColour As Long
        Private m_lZoomCrosshairColour As Long

Private m_navBar As NavigationBarExtension
Private m_styleCombo As ToolbarStyleABCombo

Private m_DefaultViewz As Double
Private m_DefaultViewx As Double
Private m_DefaultViewy As Double
Private m_DefaultViewName As String
Private m_sInitialMapName As String
Private m_bMapInitialized As Boolean
Private m_sSecAnalysisLyrName As String
Private m_sSecAnalysisFieldName As String

Private m_udtCoord As udtCoord
Private m_bOnTheMap As Boolean
Private m_bGeneralSpatialAnalysis As Boolean
Private m_oXgisThemeWiz As New TatukGIS_XDK10.XGIS_ControlLegendVectorWiz
Private m_bThematicsDone As Boolean
Private m_sInternetCon As String
Private m_bInternetCon As Boolean
Private m_cUploadThreads() As clsThreads

Private m_dPTG As TatukGIS_XDK10.XGIS_Point
'Private m_bZoomControlChangingZoom As Boolean

Dim WithEvents FormThread As MThreadVB.Thread
Attribute FormThread.VB_VarHelpID = -1
'Attribute FormThread.VB_VarHelpID = -1
Dim WithEvents pThread As MThreadVB.Thread
Attribute pThread.VB_VarHelpID = -1
'Attribute pThread.VB_VarHelpID = -1
Dim WithEvents pInitThread As MThreadVB.Thread
Attribute pInitThread.VB_VarHelpID = -1
'Attribute pInitThread.VB_VarHelpID = -1
Dim WithEvents pSubmitGeoMarksThread As MThreadVB.Thread
Attribute pSubmitGeoMarksThread.VB_VarHelpID = -1
'Attribute pSubmitGeoMarksThread.VB_VarHelpID = -1
Dim WithEvents pLoadLyrAttrToGrdThread As MThreadVB.Thread
Attribute pLoadLyrAttrToGrdThread.VB_VarHelpID = -1
'Attribute pLoadLyrAttrToGrdThread.VB_VarHelpID = -1
Dim WithEvents pUpdateCheckerThread As MThreadVB.Thread
Attribute pUpdateCheckerThread.VB_VarHelpID = -1
'Attribute pUpdateCheckerThread.VB_VarHelpID = -1
Dim WithEvents pSynchThread As MThreadVB.Thread
Attribute pSynchThread.VB_VarHelpID = -1
'Attribute pSynchThread.VB_VarHelpID = -1
Dim WithEvents pCheckSynchThread As MThreadVB.Thread
Attribute pCheckSynchThread.VB_VarHelpID = -1
'Attribute pCheckSynchThread.VB_VarHelpID = -1
Dim WithEvents incThread As MThreadVB.Thread
Attribute incThread.VB_VarHelpID = -1
'Attribute incThread.VB_VarHelpID = -1
Dim WithEvents pGetIncThread As MThreadVB.Thread
Attribute pGetIncThread.VB_VarHelpID = -1
'Attribute pGetIncThread.VB_VarHelpID = -1
Dim WithEvents pLoadMap As MThreadVB.Thread
Attribute pLoadMap.VB_VarHelpID = -1
'Attribute pLoadMap.VB_VarHelpID = -1
Dim WithEvents pCheckInet As MThreadVB.Thread
Attribute pCheckInet.VB_VarHelpID = -1
'Attribute pCheckInet.VB_VarHelpID = -1

'For the Geometry Edit
Dim vkControl As Boolean
Private menuPos As TatukGIS_XDK10.XGIS_Point

'For the GIS Grid
Dim GISGridRS As ADODB.Recordset
Dim GISGridExtent As XGIS_Extent
Dim GISGridLayerName As String
Dim GlobalGISGridLayer As TatukGIS_XDK10.XGIS_LayerVector

Const VK_CONTROL = 17
Const VK_DELETE = 46

Private m_bPoint As Boolean
Private m_bLine As Boolean
Private m_bPoly As Boolean
Private m_bMPoint As Boolean
Private bDebugMode As Boolean
Private btns(5) As ActiveBar3LibraryCtl.Tool

Private oVB As Object

Private WithEvents oSync As SynchWorker
Attribute oSync.VB_VarHelpID = -1
'Attribute oSync.VB_VarHelpID = -1
'Private oSync As SynchWorker
Private WithEvents m_InternetCheck As SynchWorker
Attribute m_InternetCheck.VB_VarHelpID = -1
'Attribute m_InternetCheck.VB_VarHelpID = -1
Private WithEvents m_oSQLLyrSynch As SynchWorker
Attribute m_oSQLLyrSynch.VB_VarHelpID = -1
'Attribute m_oSQLLyrSynch.VB_VarHelpID = -1

Private WithEvents oClientInterCom As OASISInterComm.IClient
Attribute oClientInterCom.VB_VarHelpID = -1
'Attribute oClientInterCom.VB_VarHelpID = -1

Private m_PrevExt() As TatukGIS_XDK10.XGIS_Extent
Private m_bPrevActionUsed As Boolean

Private m_bScrib As Boolean
Private m_bPolyL As Boolean
Private m_bPolyG As Boolean
Private m_bRECT As Boolean

Private m_lyrFtrSelector As TatukGIS_XDK10.XGIS_LayerVector
Private m_lyrFtrTarget As TatukGIS_XDK10.XGIS_LayerVector

Private gdpts() As POINTAPI
Private ptsSel() As TatukGIS_XDK10.XGIS_Point
Private g_lHw As Long
Private m_SelUID As Long
Private m_lBufUID As Long
Private m_bDrawFinished As Boolean
Public m_SelLyrCol As Collection
Private m_shpProps() As ShpProps
Private m_udtSelectorSettings As SelectorSettings
Private Const Pi As Double = 3.14159265359 'Pi constant, used in radians/degrees conversion
Private objProj As GeoMercator
Private ptTip As POINTAPI
Private m_oToolTipSHP As TatukGIS_XDK10.XGIS_Shape
Private m_LyrCol As Collection
Private m_sIncidentIni As String
Private m_bAlreadyinitedProfile As Boolean

Private SecurityLayerDateFrom As Date
Private SecurityLayerDateTill As Date
Private m_eZoomberMaxExtent As XGIS_Extent
Private m_ColZoomScales As New Collection
Private m_sCurrentMapPath As String

Private Sub AB_BandOpen(ByVal Band As ActiveBar3LibraryCtl.Band, ByVal Cancel As ActiveBar3LibraryCtl.ReturnBool)
    If Band.Name = "SysCustomize" Then
        Cancel = True
    End If
End Sub

Private Sub AB_Resize(ByVal left As Long, ByVal top As Long, ByVal Width As Long, ByVal Height As Long)
    Scale1.left = abGridTools.Width - (Scale1.Width + 10)
    RealaignTB
End Sub

Private Sub abGridTools_Resize(ByVal left As Long, ByVal top As Long, ByVal Width As Long, ByVal Height As Long)
    Scale1.left = abGridTools.Width - (Scale1.Width + 10)
End Sub

Private Sub AttribTabs_Switch(OldTab As Integer, _
                              NewTab As Integer, _
                              Cancel As Integer)
        '<EhHeader>
        On Error GoTo AttribTabs_Switch_Err
        '</EhHeader>
100     'm_bProgress = True
            
102     With m_udtSelectorSettings
    
104         If .bAutoZoom Then
106             SelAutoZoom m_shpProps(NewTab).uID, m_shpProps(NewTab).sLayerName
            End If

108         If .bAutoSelect Then
110             SelAutoSelect m_shpProps(NewTab).uID, m_shpProps(NewTab).sLayerName
            End If

112         If .bAutoFlash Then
114             SelAutoFlash m_shpProps(NewTab).uID, m_shpProps(NewTab).sLayerName
            End If
    
        End With
    
116     'm_bProgress = False

        '<EhFooter>
        Exit Sub

AttribTabs_Switch_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.AttribTabs_Switch " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SelAutoZoom(uID As Long, sLayerName As String)
        'GIS109 Done
        '<EhHeader>
        On Error GoTo SelAutoZoom_Err
        '</EhHeader>
        Dim oLyr9 As TatukGIS_XDK10.XGIS_LayerVector
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
    
100     Set oLyr9 = GIS10.get(sLayerName)

102     For Each oShp9 In oLyr9.Loop(oLyr9.Extent, "GIS_UID = " & uID, Nothing, "", True)
104         If Not oShp9 Is Nothing Then GIS10.VisibleExtent = oShp9.Extent: Exit For
        Next
        '<EhFooter>
        Exit Sub

SelAutoZoom_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SelAutoZoom " & _
               "at line " & Erl
       
        '</EhFooter>
End Sub

Private Sub SelAutoSelect(uID As Long, sLayerName As String)
        'GIS109 Done
        '<EhHeader>
        On Error GoTo SelAutoSelect_Err
        '</EhHeader>
        Dim oLyr9 As TatukGIS_XDK10.XGIS_LayerVector
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
    
100     Set oLyr9 = GIS10.get(sLayerName)

102     For Each oShp9 In oLyr9.Loop(oLyr9.Extent, "GIS_UID = " & uID, Nothing, "", True)
104         If Not oShp9 Is Nothing Then
            
106             With oShp9
108                 .Lock TatukGIS_XDK10.XgisLockExtent
110                 Set oShp9 = .MakeEditable
112                 .IsSelected = True
114                 .Unlock
116                 .draw
                End With
            
118             GIS10.UpDate
            
                Exit For
            
            End If
        Next
        '<EhFooter>
        Exit Sub

SelAutoSelect_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SelAutoSelect " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub chkFilterIn_Click()
        '<EhHeader>
        On Error GoTo chkFilterIn_Click_Err
        '</EhHeader>

        Dim FilterText As String
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector

        Dim lStart As Long
        Dim lEnd As Long
        Dim lEndOld As Long
        Dim sField As String
        Dim sDate As String
        
100     FilterText = dxGISDataGrid.Filter.FilterText
102     lEndOld = 1

104     If Not FilterText = "" Then

106         If InStr(FilterText, "Date]") > 0 Then
                
108             lStart = InStr(FilterText, "Date]") + 5
110             lStart = InStr(lStart, FilterText, "'", vbTextCompare) + 1
112             lEnd = InStr(lStart + 1, FilterText, "'", vbTextCompare)
114             sDate = Mid(FilterText, lStart, lEnd - lStart)
116             FilterText = Replace(FilterText, sDate, Format(sDate, "yyyymmdd"), , 1, vbTextCompare)
                
118             lStart = InStr(FilterText, "Date]") + 5
120             lStart = InStrRev(FilterText, "[", lStart)
122             lEndOld = lEnd
124             lEnd = InStr(lStart + 1, FilterText, "]", vbTextCompare)
126             sField = Mid(FilterText, lStart, lEnd - lStart + 1)
128             FilterText = Replace(FilterText, sField, LCase("Format(" & sField & ", 'yyyymmdd')"), , 1, vbTextCompare)
            
            End If
            
130         If InStr(lEndOld, FilterText, "Date]", vbTextCompare) > 0 Then
            
132             lStart = InStr(lEnd, FilterText, "Date]") + 5
134             lStart = InStr(lStart, FilterText, "'", vbTextCompare) + 1
136             lEndOld = lEnd
138             lEnd = InStr(lStart + 1, FilterText, "'", vbTextCompare)
140             sDate = Mid(FilterText, lStart, lEnd - lStart)
142             FilterText = Replace(FilterText, sDate, Format(sDate, "yyyymmdd"), , 1, vbTextCompare)
                
144             lStart = InStr(lEndOld, FilterText, "Date]") + 5
146             lStart = InStrRev(FilterText, "[", lStart)
148             lEnd = InStr(lStart + 1, FilterText, "]", vbTextCompare)
150             sField = Mid(FilterText, lStart, lEnd - lStart + 1)
152             FilterText = Replace(FilterText, sField & " ", "Format(" & sField & ", 'yyyymmdd') ")
            
            End If
        
        End If

154     FilterText = Replace(FilterText, "[", "")
156     FilterText = Replace(FilterText, "]", "")

158     If Not chkFilterIn.value = vbChecked Then FilterText = ""

160     If Not abGridTools.Tools.Item("comLyr").Text = "---Nothing---" Then

162         If Not m_oSQLIncLyr Is Nothing Then

164             If abGridTools.Tools.Item("comLyr").Text = m_oSQLIncLyr.Name Then
            
166                 LoadOASISIncidentsEX False, False
                
168                 If Not Len(FilterText) = 0 Then

170                     If Not Len(m_oSQLIncLyr.scope) = 0 Then
172                         FilterText = "(" & FilterText & ") AND "
                        End If
                    
                    End If
                
174                 If Not Len(m_oSQLIncLyr.scope) = 0 Then
                
176                     FilterText = FilterText & "(" & m_oSQLIncLyr.scope & ")"
                
                    End If
                
                End If
            
            End If
    
178         GIS10.Lock
                      
180         Set lL = GIS10.get(GISGridLayerName)
182         DoEvents
184         lL.scope = FilterText
186         lL.Params.Visible = True
188         GIS10.UpDate
190         GIS10.Unlock
    
        End If

        '<EhFooter>
        Exit Sub

chkFilterIn_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.chkFilterIn_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkOnlyVisible_Click()
    mnuLoadAll.Checked = Not IIf(chkOnlyVisible.value = vbChecked, True, False)
    mnuVisibleExtent.Checked = IIf(chkOnlyVisible.value = vbChecked, True, False)
    
End Sub

Private Sub chkSelectIn_Click()
        '<EhHeader>
        On Error GoTo chkSelectIn_Click_Err
        '</EhHeader>
        Dim FilterText As String
        Dim oLyr9 As TatukGIS_XDK10.XGIS_LayerVector
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
            
100     If Not abGridTools.Tools.Item("comLyr").Text = "---Nothing---" Then

102         FilterText = dxGISDataGrid.Filter.FilterText
104         FilterText = Replace(FilterText, "[", "")
106         FilterText = Replace(FilterText, "]", "")

108         If Not chkSelectIn.value = vbChecked Then
110             mnusep.Visible = False
112             mnuHideSelected.Visible = False
114             mnuShowSelected.Visible = False
116             FilterText = ""
            Else
118             mnusep.Visible = True
120             mnuHideSelected.Visible = True
122             mnuShowSelected.Visible = False
            End If
            
124         GIS10.Lock

126         Call mnuClearSelections_Click
128         Set oLyr9 = GIS10.get(GISGridLayerName)
130         oLyr9.DeselectAll
            
132         For Each oShp9 In oLyr9.Loop(oLyr9.Extent, FilterText, Nothing, "", True)

134             If Not oShp9 Is Nothing Then
136                 Set oShp9 = oShp9.MakeEditable

138                 If Not chkSelectIn.value = vbChecked Then
140                     oShp9.IsSelected = False
                    Else
142                     oShp9.IsSelected = True
                    End If
                End If

            Next

144         GIS10.Unlock
            
        End If

        '<EhFooter>
        Exit Sub

chkSelectIn_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.chkSelectIn_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAttachments_Click()
    AddAttachmentRecord "incident", "attachment", "desc", "sguiD", "sFileType", "FileSize", "subdomain"
End Sub

Private Sub AddAttachmentRecord(incidentID As String, _
                                AttachmentTable As String, _
                                Description As String, _
                                sGUID As String, _
                                FileType As String, _
                                FileSize As String, _
                                subdomain As String)


    If DoTableExists("Attachments", m_Cnn) Then
        
        DebugPrint "INSERT INTO Attachments (incidentID,AttachmentTable,Description,sGUID,FileType,FileSize,other) VALUES ('" & incidentID & "','" & AttachmentTable & "','" & Description & "','" & sGUID & "','" & FileType & "','" & FileSize & "')"
        m_Cnn.Execute "INSERT INTO Attachments (incidentID,AttachmentTable,Description,sGUID,FileType,FileSize,other) VALUES ('" & incidentID & "','" & AttachmentTable & "','" & Description & "','" & sGUID & "','" & FileType & "','" & FileSize & "','" & subdomain & "')"

    End If

End Sub

Private Sub cmdAttmtLister_Click()
    frmAttachmentUploader.Show vbModal, Me
    frmAttachmentViewer.Show
End Sub

Private Sub cmdCommand6_Click()
        frmMapPrint.Init GIS10.viewer, m_frmMnuOperations.Legend1.Legend, g_bMapTplUpdate, g_sRemoteTablePrefix
        frmMapPrint.Show vbModeless, Me
End Sub

Private Sub cmdSearch_Click()

    If m_frmSearch Is Nothing Then
        Set m_frmSearch = New frmSearch
    End If
                
    m_frmSearch.Init GIS10.viewer
                
    m_frmSearch.Show vbModeless, Me
End Sub

Private Sub cmdSendMess_Click()
    frmSendMess.Show
End Sub

Private Sub ctlSelector1_ChangeSelector(sTool As String)
    'g_CurrentTool = oTool
    'g_CurrentFeatureType =
    'GIS10.Mode = TatukGIS_XDK10.XgisSelect
    g_CurrentFeatureType = Location
    
    ReDim ptsSel(0)
    ReDim gdpts(0)
    
    Select Case sTool

        Case "Circle (specify km)"
        
            g_CurrentTool = oPointBuffer ' oCreateLocationPoint 'oCreateLocationPoint
            g_CurrentFeatureType = Custom
            m_bPolyL = False
            m_bPolyG = False
            SetOASISTool

        Case "Polyline"
        
            GIS10.Mode = TatukGIS_XDK10.XgisEdit
            g_CurrentTool = oLineSelect
            m_bPolyL = True
            m_bPolyG = False
            SetOASISTool
        
        Case "Circle (draw)"

            GIS10.Mode = TatukGIS_XDK10.XgisEdit ' TatukGIS_XDK10.XGISUserDefined
            g_CurrentTool = oCircleSelect
            m_bPolyL = False
            m_bPolyG = False
            SetOASISTool

        Case "Polygon"

            GIS10.Mode = TatukGIS_XDK10.XgisEdit ' TatukGIS_XDK10.XGISUserDefined
            g_CurrentTool = oAreaSelect
            m_bPolyL = False
            m_bPolyG = True
            SetOASISTool

        Case "Pan"
            
            GIS10.Mode = TatukGIS_XDK10.XgisDrag
            'endEdit
            g_CurrentTool = oPan
            m_bPolyL = False
            m_bPolyG = False
            SetOASISTool
    End Select
    
    With dxGISDataGrid

        If mnuAutoSize.Checked Then
            .Options.Set (egoAutoWidth)
        Else
            .Options.Unset (egoAutoWidth)
        End If
            
        If mnuAutoHeight.Checked Then
            .Options.Set (egoRowAutoHeight)
        Else
            .Options.Unset (egoRowAutoHeight)
        End If

        .Options.Set (egoShowGroupPanel)
        .Options.Set (egoBandMoving)
        .Options.Set (egoColumnMoving)
        .Options.Set (egoMultiSort)
        .Options.Set (egoShowFooter)
        .Options.Set (egoAutoSort)
        .Options.Set (egoShowButtons)
        .Options.Set (egoShowRowFooter)
        .Options.Set (egoAutoSearch)
        .Options.Set (egoAutoExpandOnSearch)
        .Options.Set (egoAnsiSort)
        .Options.Set (egoLoadAllRecords)
        .Options.Set (egoAutoSearch)
        .Options.Unset (egoCanNavigation)
        .Options.Unset (egoDynamicLoad)
        .DatasetType = dtADODataset
        .Filter.FilterActive = True
        .Filter.FilterStatus = fsAlways
    
    End With
    
End Sub

Private Sub ctlSelector1_DisableSelections(sLayerName As String)
    
    Dim lL As TatukGIS_XDK10.XGIS_LayerVector
    Set lL = GIS10.get(sLayerName)
    
    If Not lL Is Nothing Then
        lL.DeselectAll
    End If
    
End Sub

Private Sub ctlSelector1_ExportData(oRS As ADODB.Recordset)
    
    Dim m_frmReportsFromRS As frmReportsFromRS
    Set m_frmReportsFromRS = New frmReportsFromRS

    frmGridRptProprties.Init oRS
    frmGridRptProprties.txtTitle.Text = "OASIS Resources Finder"
    frmGridRptProprties.Show vbModal, Me
                        
    If frmGridRptProprties.chkIncludeMap = vbChecked Then
        Clipboard.Clear
        GIS10.viewer.PrintClipboard

        m_frmReportsFromRS.SetReportRS frmGridRptProprties.txtTitle.Text, oRS, "", Clipboard.GetData(vbCFEMetafile), frmGridRptProprties.txtMapTitle.Text
    Else
        m_frmReportsFromRS.SetReportRS frmGridRptProprties.txtTitle.Text, oRS, ""
    End If

    Unload frmGridRptProprties
    m_frmReportsFromRS.ShowReport
    m_frmReportsFromRS.Show vbModal, Me
    'sPDFPath = m_frmReportsFromRS.PDFPath
                
    Unload m_frmReportsFromRS
    Set m_frmReportsFromRS = Nothing
    'Set RSIncidentForOASISReports = Nothing

End Sub

Private Sub ctlSelector1_FlashShape(sUID As String)
    Dim oLayer As XGIS_LayerVector
    Set oLayer = GIS10.get(ctlSelector1.GetActiveLayer)
    oLayer.GetShape(CInt(sUID)).Flash 4
End Sub

Private Sub ctlSelector1_GetLayers(sLayers As String)

    Dim i As Long

    With m_frmMnuOperations.Legend1
    
        For i = 0 To GIS10.items.Count - 1

            If InStr(g_sPermanentLyrs, GIS10.items.Item(i).Name & ",") < 1 Then
                If GisUtils.IsInherited(GIS10.items.Item(i), "XGIS_LayerVector") Then
                    sLayers = sLayers & ";" & GIS10.items.Item(i).Name
                End If
            End If

        Next

    End With

End Sub

Private Sub ctlSelector1_MergeOtherSelections(sOtherLayer As String)

    Dim bErrorPoint As Boolean
    Dim oLayerOther As XGIS_LayerVector
    Dim oLayerAnalyzed As XGIS_LayerVector
    Dim oShape As XGIS_Shape
    Dim oShapeMerged As XGIS_Shape
    
    Set oLayerOther = GIS10.get(sOtherLayer)
    m_oBufferLyr.RevertAll
    Set oLayerAnalyzed = m_oBufferLyr
    Set oShapeMerged = oLayerAnalyzed.CreateShape(XgisShapeTypeUnknown)
    
    Dim i As Long
    i = 0
    
    Do Until i = oLayerOther.items.Count
        
        Set oShape = oLayerOther.items.Item(i)
        
        If oShape.IsSelected Then
        
            If oShape.ShapeType = XgisShapeTypePoint Then
                If Not bErrorPoint Then MsgBox "NOTE: Point features will be ignored"
                bErrorPoint = True
            Else
                Set oShapeMerged = oShapeMerged.Join(oShape)
            End If
    
        End If
        
        i = i + 1
    Loop
       
    m_oBufferLyr.AddShape oShapeMerged
    oLayerOther.DeselectAll
    If Not bErrorPoint Then MsgBox "New selection layer created." & vbCrLf & vbCrLf & "Please select an analysis layer now from the legend and click 'GO'"
   
End Sub

Private Sub ctlSelector1_UseOtherLayer(sOtherLayer As String)

    AddShapeWithBufferAndHighlightIntersect m_oBufferLyr.GetShape(m_oBufferLyr.GetLastUid), 0

End Sub

Private Sub ctlZoomSlider1_ZoomChanged(oExtent As TatukGIS_XDK10.XGIS_Extent)
'm_bZoomControlChangingZoom = True
    Dim oPt As XGIS_Point
    GIS10.Lock
    Set oPt = GIS10.CenterPtg
    GIS10.VisibleExtent = oExtent
    Call GIS10.CenterViewport(oPt)
    GIS10.Unlock
    Set oPt = Nothing
    DoEvents
 '   m_bZoomControlChangingZoom = False

    
End Sub

Private Sub dxGISDataGrid_OnShowCellTip(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, _
                                        ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, _
                                        TipText As String, _
                                        l As Single, _
                                        T As Single, _
                                        R As Single, _
                                        b As Single, _
                                        NeedShowTip As Boolean)
    
    Dim i As Integer
    
    On Error Resume Next

    DebugPrint "BUGFIX: NeedShowTip:" & NeedShowTip
    'BugFix Load Visible, no bug works as designed

    If mnuGridAction.Visible Then
        NeedShowTip = False
        Exit Sub
    Else
        NeedShowTip = True
    End If
    
    If Node.HasChildren Then
        TipText = "Grouping:" & vbNewLine & Node.Strings(0)
    Else
        TipText = ""

        For i = 1 To Node.ValuesCount - 1
            TipText = TipText & Node.values(i) & vbNewLine
        Next

    End If

    DebugPrint TipText & " ERR.Code" & Err.Description

End Sub

Private Sub DynamicDataModule1_GetSpatialLoc(oLayer As TatukGIS_XDK10.XGIS_LayerVector)

    m_prevTool = oPan
    elMap.Visible = True
    AB.ClientAreaControl = elMap
    elDynamicData.Visible = False
    
    AB.Enabled = False

    If Not m_frmSpatialiseDD Is Nothing Then
        If Not m_frmSpatialiseDD.Visible Then
            Set m_frmSpatialiseDD = New frmSpatialiseDD
        End If
    Else
        Set m_frmSpatialiseDD = New frmSpatialiseDD
    End If
    
    Set mDDLayer = oLayer
    m_frmSpatialiseDD.Tag = "SPATIAL TABLE"
    m_frmSpatialiseDD.Show vbModeless, Me
    m_frmSpatialiseDD.Init False
    m_frmSpatialiseDD.SetLayer mDDLayer
    
End Sub

Private Sub DynamicDataModule1_GetSpatialLocEx(oPoint As TatukGIS_XDK10.XGIS_Point)

    Dim oShape As TatukGIS_XDK10.XGIS_Shape
    
    m_prevTool = oPan
    elMap.Visible = True
    AB.ClientAreaControl = elMap
    elDynamicData.Visible = False
    
    AB.Enabled = False

    If Not m_frmSpatialiseDD Is Nothing Then
        If Not m_frmSpatialiseDD.Visible Then
            Set m_frmSpatialiseDD = New frmSpatialiseDD
        End If
    Else
        Set m_frmSpatialiseDD = New frmSpatialiseDD
    End If
    
    Set mDDLayer = GIS10.get("Draw_Layer")
    Set oShape = mDDLayer.CreateShape(XgisShapeTypePoint)
    oShape.AddPart
    oShape.AddPoint oPoint
    
    m_frmSpatialiseDD.Tag = "NON SPATIAL TABLE"
    m_frmSpatialiseDD.Show vbModeless, Me
    m_frmSpatialiseDD.Init True
    m_frmSpatialiseDD.SetLayer mDDLayer
    
End Sub

Private Sub elGisAttr_ResizeChildren()
    Scale1.left = abGridTools.Width - (Scale1.Width + 10)
End Sub

Private Sub EventLayer_OnPaintLayer(translated As Boolean, ByVal layer As Object)
 '   translated = True
'    GIS10.get(layer.Name).draw
End Sub

Private Sub Form_Resize()
    Scale1.left = abGridTools.Width - (Scale1.Width + 10)
    RealaignTB
    ctlZoomSlider1.ZOrder 0
    'AB.Bands.Item("tbrHelp").DockingOffset = Me.Width - 600
End Sub

Private Sub GIS10_OnClick(translated As Boolean)
    GIS10.SetFocus
End Sub

Private Sub GIS10_OnExtentChange(translated As Boolean)
'Stop
End Sub

Private Sub GIS10_OnMouseWheel(translated As Boolean, _
                               ByVal Shift As TatukGIS_XDK10.XShiftState, _
                               ByVal wheelDelta As Long, _
                               ByVal mousePos As TatukGIS_XDK10.IXPoint, _
                               handled As Boolean)
        '<EhHeader>
        On Error GoTo GIS10_OnMouseWheel_Err
        '</EhHeader>
        
        If g_CurrentTool = oInfo Or g_CurrentTool = oPan Or g_CurrentTool = oZoom Or g_CurrentTool = oZoom2FullExtent Or g_CurrentTool = oZoomEx Or g_CurrentTool = oZoomIn Or g_CurrentTool = oZoomOut Then
        
            Dim Pt     As New xPoint
            Dim ptg1   As New XGIS_Point
   
            'Set pt = GIS10.ScreenToClient(mousePos)
100         translated = True

102         Set ptg1 = GIS10.ScreenToMap(GIS10.ScreenToClient(mousePos))
        
            GIS10.Lock
        
104         If wheelDelta > 0 Then
106             ' GIS10.zoom = GIS10.zoom * 2
                ctlZoomSlider1.ZoomIn
            Else
108             'GIS10.zoom = GIS10.zoom / 2
                ctlZoomSlider1.ZoomOut
            End If
        
            GIS10.CenterViewport ptg1
            GIS10.Unlock
        End If
        
        '<EhFooter>
        Exit Sub

GIS10_OnMouseWheel_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GIS10_OnMouseWheel " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS10_OnZoomChange(translated As Boolean)

'   ctlZoomSlider1.SetZoomLevelWithoutUpdate (ctlZoomSlider1.GetZoomBarPercentageFromCurrentExtentOnly(GIS10.VisibleExtent))
    
    '

'        '<EhHeader>
'        On Error GoTo GIS10_OnZoomChange_Err
'        '</EhHeader>
'        Dim i As Integer
'
'100     For i = 0 To 99
'
'102         If GisUtils.GisIsSameExtent(GIS10.VisibleExtent, m_oExtents100(i)) Then
'104             ctlZoomSlider1.SetZoomLevelWithoutUpdate i
'                Exit Sub
'            End If
'
'        Next
'
'        '<EhFooter>
'        Exit Sub
'
'GIS10_OnZoomChange_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.GIS10_OnZoomChange " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub m_frmDynamicContent_CategoryClick()
        '<EhHeader>
        On Error GoTo m_frmDynamicContent_CategoryClick_Err
        '</EhHeader>
100     RSSBrowser1.UpdateFeedList
        '<EhFooter>
        Exit Sub

m_frmDynamicContent_CategoryClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmDynamicContent_CategoryClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmDynamicContent_FeedsClick(sURL As String, bUseGeo As Boolean, oItems As MSXML2.IXMLDOMNodeList, sCountryISO As String)
        '<EhHeader>
        On Error GoTo m_frmDynamicContent_FeedsClick_Err
        '</EhHeader>
        If Not bUseGeo Then Exit Sub
100     RSSBrowser1.UpdateHeaderList sURL, bUseGeo, oItems, sCountryISO
        '<EhFooter>
        Exit Sub

m_frmDynamicContent_FeedsClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmDynamicContent_FeedsClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmDynamicContent_HeadlinesClick(oNode As MSXML2.IXMLDOMNode)
        '<EhHeader>
        On Error GoTo m_frmDynamicContent_HeadlinesClick_Err
        '</EhHeader>
        RSSBrowser1.ImageURL = m_frmDynamicContent.ImageURL
100     RSSBrowser1.LoadHeader oNode
        '<EhFooter>
        Exit Sub

m_frmDynamicContent_HeadlinesClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmDynamicContent_HeadlinesClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmDynamicContent_StatusMessage(sMess As String)
        '<EhHeader>
        On Error GoTo m_frmDynamicContent_StatusMessage_Err
        '</EhHeader>
100     RSSBrowser1.UpdateStatusBar sMess
        '<EhFooter>
        Exit Sub

m_frmDynamicContent_StatusMessage_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmDynamicContent_StatusMessage " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmIncidentsV2DataEntry_RefreshMap()

    'tatuk workaround (again) - otherwise the new incidentv2 will not display after save
    m_frmMnuOperations_ActivateSecurityIncidents (False)
    m_frmMnuOperations_ActivateSecurityIncidents (True)

End Sub

Private Sub ApplyMainSettings()
        '<EhHeader>
        On Error GoTo ApplyMainSettings_Err
        '</EhHeader>

100     With m_frmMainSettings
102         NArrow.Visible = IIf(.chkUseNorth = vbChecked, True, False)
            'If Not oMapObjects.UseNorthArrow Then
            '    NArrow.TRANSPARENT = True ' Tatuk bug,
            'Else
104             NArrow.TRANSPARENT = IIf(.chkArrowTransparent = vbChecked, True, False)
            'End If
            
            
            If .chkUseNorth = vbChecked Then
                NArrow.GIS_Viewer = GIS10.viewer
106         If .chkUseCustom = vbChecked Then
108             NArrow.Path = .chkUseCustom.Tag
            ElseIf .cboArrow.Text = "iMMAP" Then
                NArrow.Path = .Arrow.Path
            Else
110             NArrow.Path = ""
112             NArrow.Symbol = .cboArrow.List(.cboArrow.ListIndex)
114             NArrow.Color1 = .Arrow.Color1
116             NArrow.Color2 = .Arrow.Color2
            End If

118         NArrow.UpDate
            Else
            NArrow.GIS_Viewer = Nothing
            NArrow.Visible = False
            End If
            
120         GIS10.ScrollBars = .ComScroll.ListIndex
            
122         If .chkShowWater = vbChecked Then
                watermark.GIS_Viewer = GIS10.viewer
124             watermark.Visible = True
                watermark.Path = .chkShowWater.Tag
                watermark.TRANSPARENT = True
                watermark.ZOrder 0
            Else
126             watermark.Visible = False
                watermark.GIS_Viewer = Nothing
            End If
        
            watermark.UpDate
        
128
130          .edRotationAngle.Text = .udRotationAngle.value
132             GIS10.RotationPoint.x = GIS10.viewer.CenterPtg.x
134             GIS10.RotationPoint.y = GIS10.viewer.CenterPtg.y
136             GIS10.RotationAngle = DegToRad(.udRotationAngle.value)
138             GIS10.UpDate
  
        End With
        
        tmrToolTip.Enabled = False
        tmrToolTip.Interval = 1000 * oMapTipSetting.TipDelay
        tmrToolTip.Enabled = oMapTipSetting.Enabled
        
        '<EhFooter>
        Exit Sub

ApplyMainSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ApplyMainSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ApplyLocalSettings()
        '<EhHeader>
        On Error GoTo ApplyLocalSettings_Err
        '</EhHeader>

100     With oMapObjects
102         NArrow.Visible = IIf(.UseNorthArrow, True, False)

            If Not .UseNorthArrow Then
                NArrow.TRANSPARENT = True 'Tatuk bug
            Else
104             NArrow.TRANSPARENT = IIf(.NorthArrowTransparency, True, False)
            End If
            
108         If .UseNorthArrow Then
110             NArrow.GIS_Viewer = GIS10.viewer

112             If .NorthArrowPicture <> "" Then
114                 NArrow.Path = .NorthArrowPicture
                End If

122             NArrow.Symbol = .NorthArrowType
124             NArrow.Color1 = .NorthArrowColor
126             NArrow.Color2 = .NorthArrowColor
128             NArrow.UpDate
            Else
130             NArrow.GIS_Viewer = Nothing
132             NArrow.Visible = False
            End If
                        
136         If .UseWaterMark Then
138             watermark.GIS_Viewer = GIS10.viewer
140             watermark.Visible = True
142             watermark.Path = .WaterMarkPath
144             watermark.TRANSPARENT = True
146             watermark.ZOrder 0
            Else
148             watermark.Visible = False
150             watermark.GIS_Viewer = Nothing
            End If
        
152         watermark.UpDate
        End With
        
        With oMapSettings
            GIS10.ScrollBars = .ScrollBars

            If .MapRotation <> 0 Then
156             GIS10.RotationPoint.x = GIS10.viewer.CenterPtg.x
158             GIS10.RotationPoint.y = GIS10.viewer.CenterPtg.y
160             GIS10.RotationAngle = DegToRad(CDbl(.MapRotation))
162             GIS10.UpDate
            End If

        End With

        '<EhFooter>
        Exit Sub

ApplyLocalSettings_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.ApplyLocalSettings " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMainSettings_doApply()
    ApplyMainSettings
End Sub

Private Sub m_frmMainSettings_doOK()
    ApplyMainSettings
End Sub

Private Sub m_frmMainSettings_GetFields(sName As String)
    Dim i As Integer
    Dim oVec As TatukGIS_XDK10.XGIS_LayerVector

    With m_frmMainSettings
        .ComTipField.Clear
        
        Set oVec = GIS10.get(m_LyrCol.Item(sName))
        
        If Not oVec Is Nothing Then
                .ComTipField.AddItem "UID"
            For i = 0 To oVec.Fields.Count - 1
                .ComTipField.AddItem oVec.Fields.Item(i).Name
            Next

        End If
        
        FindIndexStrEx .ComTipField, oMapTipSetting.MapTipField

        If .ComTipField.ListIndex = -1 Then FindIndexStrEx .ComTipField, "UID"

    End With
    
End Sub



Private Sub m_frmMnuDynamicReportsModule_DRGroupClicked()
    DynamicDataReports1.GroupClicked
End Sub

Private Sub m_frmMnuOperations_ChangeActiveLayer(sName As String)
    
    Dim sFields As String
    Dim oLayer As XGIS_LayerVector
    Dim oLayerAbstract As XGIS_LayerAbstract
    
    Dim i As Long
    i = 0
    
    Set oLayerAbstract = GIS10.get(sName)
    
    If Not oLayerAbstract Is Nothing Then
    
        If GisUtils.IsInherited(oLayerAbstract, "XGIS_LayerVector") And oLayerAbstract.Active Then
        
            Set oLayer = GIS10.get(sName)
        
            Do Until i = (oLayer.Fields.Count)
                sFields = sFields & "," & oLayer.Fields.Item(i).Name
                i = i + 1
            Loop
        
            ctlSelector1.SetLayer sName, GIS10.get(sName).caption, sFields
        Else
            ctlSelector1.SetLayer "", GIS10.get(sName).caption, sFields
        End If
      
        Set oLayer = Nothing
   
    End If
    
    Set oLayerAbstract = Nothing

End Sub

Private Sub m_frmMnuOperations_ExportLayer(oLyrName As String, sFilter As String, sDialogTitle As String, sFileExtention As String, fType As OASIS_GIS_DATA_TYPE)

    Dim c As New cCommonDialog
    Dim xportLyr As TatukGIS_XDK10.XGIS_LayerVector
    Dim eFile As OASIS_GIS_DATA_TYPE
    Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
    
    Set oLyr = GIS10.get(oLyrName)
    
    With c

        Select Case fType
    
            Case OASIS_GIS_DATA_TYPE.vDXF
                .Filter = "Autocad DXF files (*.dxf)|*.dxf"
                .CancelError = False
                .DialogTitle = "Export to Autocad DXF file"
                .InitDir = g_sAppPath & "\data\user\Maps\"
                .ShowSave
                    
                If Len(.Filename) < 1 Then
                    Exit Sub
                End If
                    
                Set xportLyr = New TatukGIS_XDK10.XGIS_LayerDXF
                    
                xportLyr.Path = .Filename & ".dxf"
                oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
                xportLyr.SaveAll

            Case OASIS_GIS_DATA_TYPE.vGML
                .Filter = "OGIS GML files (*.GML)|*.GML"
                .CancelError = False
                .DialogTitle = "Export to OGIS GML  file"
                .InitDir = g_sAppPath & "\data\user\Maps\"
                .ShowSave
                                    
                If Len(.Filename) < 1 Then
                    Exit Sub
                End If
                    
                Set xportLyr = New TatukGIS_XDK10.XGIS_LayerGML
                    
                xportLyr.Path = .Filename & ".gml"
                            
                oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
                xportLyr.SaveAll

            Case OASIS_GIS_DATA_TYPE.vGPX
                .Filter = "GPS Export files (*.gpx)|*.gpx"
                .CancelError = False
                .DialogTitle = "Export to GPS Export file"
                .InitDir = g_sAppPath & "\data\user\Maps\"
                .ShowSave
                    
                If Len(.Filename) < 1 Then
                    frmLayerSelection.ClearItem
                    Exit Sub
                End If
                                        
                Set xportLyr = New TatukGIS_XDK10.XGIS_LayerGPX
                    
                xportLyr.Path = .Filename & ".gpx"
                oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
                xportLyr.SaveAll

            Case OASIS_GIS_DATA_TYPE.vKML
                .Filter = "GOOGLE KML files (*.kml)|*.kml"
                .CancelError = False
                .DialogTitle = "Export to Google KML file"
                .InitDir = g_sAppPath & "\data\user\Maps\"
                .ShowSave
                    
                If Len(.Filename) < 1 Then
                    frmLayerSelection.ClearItem
                    Exit Sub
                End If
                                        
                Set xportLyr = New TatukGIS_XDK10.XGIS_LayerKML
                    
                xportLyr.Path = .Filename & ".kml"
                oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
                xportLyr.SaveAll

            Case OASIS_GIS_DATA_TYPE.vNSQL

            Case OASIS_GIS_DATA_TYPE.vOSQL

            Case OASIS_GIS_DATA_TYPE.vSHP
                .Filter = "ESRI Shape files (*.shp)|*.shp"
                .CancelError = False
                .DialogTitle = "Export to ESRI shape file"
                .InitDir = g_sAppPath & "\data\user\Maps\"
                .ShowSave
                    
                Set xportLyr = New TatukGIS_XDK10.XGIS_LayerSHP
                                    
                If Len(.Filename) < 1 Then
                    frmLayerSelection.ClearItem
                    Exit Sub
                End If
                                    
                xportLyr.Path = .Filename & ".shp"
                oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
                xportLyr.SaveAll

            Case OASIS_GIS_DATA_TYPE.vTAB
                .Filter = "GPS Export files (*.tab)|*.tab"
                .CancelError = False
                .DialogTitle = "Export to Mapinfo TAB file"
                .InitDir = g_sAppPath & "\data\user\Maps\"
                .ShowSave
                    
                If Len(.Filename) < 1 Then
                    frmLayerSelection.ClearItem
                    Exit Sub
                End If
                                        
                Set xportLyr = New TatukGIS_XDK10.XGIS_LayerTAB
                    
                xportLyr.Path = .Filename & ".tab"
                oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
                xportLyr.SaveAll
        End Select

    End With

End Sub

Private Sub m_frmMnuOperations_LoadLayerToGrid(Index As Integer)
    abGridTools.Tools.Item("comLyr").CBListIndex = Index
End Sub

Private Sub m_frmMnuOperations_LoadMap(sMapName As String, oExtent As TatukGIS_XDK10.XGIS_Extent)
m_frmMnuOperations.Visible = False
DoEvents
  InitMap sMapName, oExtent
       m_frmMnuOperations.Show
End Sub

Private Sub m_frmMnuOperations_LoadSetting(sLayer As String)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_LoadSetting_Err
        '</EhHeader>
        Dim c As New cCommonDialog
                
100     With c
102         .DialogTitle = "Load layer settings file"
104         .CancelError = False
106         .hWnd = Me.hWnd
108         .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
110         .InitDir = g_sAppPath & "\data\user\"
112         .Filter = "OASIS layer settings File (*.ini)"
114         .FilterIndex = 1
116         .ShowOpen
            '.Filename
        End With
    
        Dim oStream As New ADODB.Stream
        Dim oSList As New TatukGIS_XDK10.XStringList
        
196     oStream.Open
198     oStream.Type = 2
200     oStream.Charset = "ascii"
202     oStream.LoadFromFile c.Filename
204     oSList.Text = oStream.ReadText
206     oStream.Close
    
118     With GIS10.get(sLayer)
120         .ConfigName = c.Filename
            .ParamsList.LoadFromStrings oSList
122         .UseConfig = True
124         .ReadConfig
        End With
    
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_LoadSetting_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.m_frmMnuOperations_LoadSetting " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_MapPreview(oPic As stdole.StdPicture)
    DebugPrint "TODO"
End Sub

Private Sub m_frmMnuOperations_OpenNewMap(sMAP As String)
    DebugPrint "TODO, Check this"
    
    InitMap sMAP
    LoadLayerAttrDataToGridInit
    
End Sub

Private Sub m_frmMnuOperations_SaveAllSettings()
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_SaveAllSettings_Err
        '</EhHeader>
        Dim oLyrAbs As TatukGIS_XDK10.XGIS_LayerAbstract
        Dim i As Integer
        Dim sLayers As String
        Dim bDefVal As Boolean
        Dim oStream As New ADODB.Stream
        Dim oSList As New TatukGIS_XDK10.XStringList
        
100     sLayers = "The Setting files were written in: " & g_sAppPath & "\Data\gis\temp\" & vbCrLf

102     For i = 0 To GIS10.items.Count - 1
        
104         Set oStream = New ADODB.Stream
106         oStream.Open
108         oStream.Type = 2
110         oStream.Charset = "ascii"
        
112         Set oLyrAbs = GIS10.get(GIS10.items.Item(i).Name)
        
            On Error Resume Next
114         Kill g_sAppPath & "\Data\gis\temp\" & oLyrAbs.Name & ".ini"
            On Error GoTo m_frmMnuOperations_SaveAllSettings_Err
        
116         oLyrAbs.ConfigName = g_sAppPath & "\Data\gis\temp\" & oLyrAbs.Name & ".ini"
118         oLyrAbs.ParamsList.SaveToStrings oSList
120         oStream.WriteText oSList.Text
122         oStream.SaveToFile (g_sAppPath & "\Data\user\Exports\" & oLyrAbs.Name & ".ini")
124         oStream.Close
126         Set oStream = Nothing
        
128         sLayers = sLayers & "Setting file written: " & oLyrAbs.Name & ".ini" & vbCrLf
        Next
    
130     MsgBox sLayers
  
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_SaveAllSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_SaveAllSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_SaveMyMap()
    SavePrivateAppSettings True
End Sub

Private Sub m_frmMnuOperations_SaveSetting(sLayer As String)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_SaveSetting_Err
        '</EhHeader>
        Dim c As New cCommonDialog
                                
100     With c
102         .DialogTitle = "Save layer settings file"
104         .CancelError = False
106         .hWnd = Me.hWnd
            '.Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
108         .InitDir = g_sAppPath & "\data\user\"
110         .Filter = "OASIS layer settings File (*.ini)"
112         .FilterIndex = 1
114         .ShowSave
        End With
    
        Dim oStream As New ADODB.Stream
        Dim oSList As New TatukGIS_XDK10.XStringList
        On Error Resume Next
116     Kill c.Filename
        On Error GoTo m_frmMnuOperations_SaveSetting_Err
118     oStream.Open
120     oStream.Type = 2
122     oStream.Charset = "ascii"
    
124     With GIS10.get(sLayer)
126         .ParamsList.SaveToStrings oSList
128         oStream.WriteText oSList.Text
130         oStream.SaveToFile (c.Filename) & ".ini"
132         oStream.Close
134         Set oStream = Nothing
136         .ConfigName = c.Filename
138         .WriteConfig
        End With
    
140     MsgBox "Your file was saved to: " & c.Filename & ".ini"

        '<EhFooter>
        Exit Sub

m_frmMnuOperations_SaveSetting_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_SaveSetting " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_ShowMapLibraryDLG(sGUID As String, bInfoOnly As Boolean)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_ShowMapLibraryDLG_Err
        '</EhHeader>
        Dim RS As New ADODB.Recordset

100     If Not m_frmMapLibraryDLG.Visible Then

102         With m_frmMapLibraryDLG
            
                .bPreview = False
                
                .cmdCancel.Visible = Not bInfoOnly
            
                .FraFrmMainDetails.Visible = True
                .FraMapInformation.Visible = False
            
104             If sGUID = "" Then
        
106                 .txtMapInfo(0).Text = "My Map: " & Format(Now(), "dd-MMM-yy")
108                 .txtMapInfo(1).Text = "Created By..."
110                 .txtMapInfo(2).Text = "Copyright..."
112                 .txtMapInfo(3).Text = "Url..."
114                 .txtMapInfo(4).Text = "Source..."
116                 .txtMapInfo(5).Text = "Contact..."
118                 .txtMapInfo(6).Text = "Description..."
120                 .lblGISDetails(0).caption = "X Min: " & GIS10.VisibleExtent.xmin
122                 .lblGISDetails(1).caption = "X Max: " & GIS10.VisibleExtent.xmax
124                 .lblGISDetails(2).caption = "Y Min: " & GIS10.VisibleExtent.ymin
126                 .lblGISDetails(3).caption = "Y Max: " & GIS10.VisibleExtent.ymax
128                 .lblGISDetails(4).caption = "Center X: " & GIS10.CenterPtg.x
130                 .lblGISDetails(5).caption = "Center Y: " & GIS10.CenterPtg.y
132                 .lblGISDetails(6).caption = "Scale: " & GIS10.ScaleAsText
134                 .lblGISDetails(7).caption = "EPSG: " & GIS10.CS.EPSG
136                 .DTPicker1.value = Now()
            
                Else
                    'Fill Default Values
138                 RS.Open "SELECT * FROM [ttkGISProjectDef] WHERE sGUID = '" & sGUID & "'", m_Cnn, adOpenDynamic, adLockOptimistic
140                 .txtMapInfo(0).Text = RS.Fields.Item("sName").value
                
142                 If Not IsNull(RS.Fields.Item("CreatedBy").value) Then
144                     .txtMapInfo(1).Text = RS.Fields.Item("CreatedBy").value
                    Else
146                     .txtMapInfo(1).Text = ""
                    End If
                
148                 If Not IsNull(RS.Fields.Item("Copyright").value) Then
150                     .txtMapInfo(2).Text = RS.Fields.Item("Copyright").value
                    Else
152                     .txtMapInfo(2).Text = ""
                    End If

154                 If Not IsNull(RS.Fields.Item("url").value) Then
156                     .txtMapInfo(3).Text = RS.Fields.Item("url").value
                    Else
158                     .txtMapInfo(3).Text = ""
                    End If

160                 If Not IsNull(RS.Fields.Item("Source").value) Then
162                     .txtMapInfo(4).Text = RS.Fields.Item("Source").value
                    Else
164                     .txtMapInfo(4).Text = ""
                    End If

166                 If Not IsNull(RS.Fields.Item("sInfo").value) Then
168                     .txtMapInfo(6).Text = RS.Fields.Item("sInfo").value
                    Else
170                     .txtMapInfo(6).Text = ""
                    End If

172                 If Not IsNull(RS.Fields.Item("XMIN").value) Then
174                     .lblGISDetails(0).caption = "X Min: " & RS.Fields.Item("XMIN").value
                    Else
176                     .lblGISDetails(0).caption = ""
                    End If

178                 If Not IsNull(RS.Fields.Item("XMAX").value) Then
180                     .lblGISDetails(1).caption = "X Max: " & RS.Fields.Item("XMAX").value
                    Else
182                     .lblGISDetails(1).caption = ""
                    End If

184                 If Not IsNull(RS.Fields.Item("YMIN").value) Then
186                     .lblGISDetails(2).caption = "Y Min: " & RS.Fields.Item("YMIN").value
                    Else
188                     .lblGISDetails(2).caption = ""
                    End If

190                 If Not IsNull(RS.Fields.Item("YMAX").value) Then
192                     .lblGISDetails(3).caption = "Y Max: " & RS.Fields.Item("YMAX").value
                    Else
194                     .lblGISDetails(3).caption = ""
                    End If

196                 If Not IsNull(RS.Fields.Item("centerX").value) Then
198                     .lblGISDetails(4).caption = "Center X: " & RS.Fields.Item("centerX").value
                    Else
200                     .lblGISDetails(4).caption = ""
                    End If

202                 If Not IsNull(RS.Fields.Item("centerY").value) Then
204                     .lblGISDetails(5).caption = "Center Y: " & RS.Fields.Item("centerY").value
                    Else
206                     .lblGISDetails(5).caption = ""
                    End If

208                 If Not IsNull(RS.Fields.Item("scale").value) Then
210                     .lblGISDetails(6).caption = "Scale: " & RS.Fields.Item("scale").value
                    Else
212                     .lblGISDetails(6).caption = ""
                    End If
                
214                 If Not IsNull(RS.Fields.Item("EPSG").value) Then
216                     .lblGISDetails(7).caption = "EPSG: " & RS.Fields.Item("EPSG").value
                    Else
218                     .lblGISDetails(7).caption = ""
                    End If
                
220                 If Not IsNull(RS.Fields.Item("CreatedDate").value) Then
222                     .DTPicker1.value = RS.Fields.Item("CreatedDate").value
                    Else
                        ' .DTPicker1.Value = ""
                    End If

                    If bInfoOnly Then
                        .FraMapInformation.Move .FraFrmMainDetails.left, .FraFrmMainDetails.top, .FraFrmMainDetails.Width, .FraFrmMainDetails.Height
                        .FraMapInformation.Visible = True
                        .FraFrmMainDetails.Visible = False
                        
                        .bPreview = True
                        .populatePreview
                        
                        Dim pb As PropertyBag
                        Set pb = New PropertyBag
                        pb.Contents = RS.Fields("oImagePreview").GetChunk(RS.Fields("oImagePreview").ActualSize)
                        .picMapPreview.Picture = pb.ReadProperty("MyImage")
                        Set pb = Nothing
                    End If

                End If

224             .Show vbModal, Me

                If Not bInfoOnly Then
226                 If .bOK Then
228                     If sGUID = "" Then
230                         m_frmMnuOperations.NewMapLibraryItem .txtMapInfo(0).Text, .txtMapInfo(6).Text, .txtMapInfo(1).Text, .txtMapInfo(2).Text, .txtMapInfo(3).Text, .txtMapInfo(4).Text, GIS10.VisibleExtent.xmin, GIS10.VisibleExtent.xmax, GIS10.VisibleExtent.ymin, GIS10.VisibleExtent.ymax, GIS10.CenterPtg.x, GIS10.CenterPtg.y, GIS10.ScaleAsText, GIS10.CS.EPSG, .DTPicker1.value
                        Else
232                         m_frmMnuOperations.MapLibSaveMap Replace(g_sAppPath & "\data\user\maps\" & sGUID & ".ttkgp", "Database Driven:", ""), sGUID, .txtMapInfo(0).Text, .txtMapInfo(6).Text, .txtMapInfo(1).Text, .txtMapInfo(2).Text, .txtMapInfo(3).Text, .txtMapInfo(4).Text, .DTPicker1.value
                        End If

234                     .bOK = False
                    End If
                End If

                .bPreview = False
        
            End With

        End If

        '<EhFooter>
        Exit Sub

m_frmMnuOperations_ShowMapLibraryDLG_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.m_frmMnuOperations_ShowMapLibraryDLG " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_ZoomLyrExtent(oLyrName As String)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_ZoomLyrExtent_Err
        '</EhHeader>

100     GIS10.VisibleExtent = GIS10.get(oLyrName).Extent
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_ZoomLyrExtent_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_ZoomLyrExtent " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmSpatialiseDD_FlashIt()

    If m_oSpatialiseLayer.items.Count > 0 Then
        
        If Not m_oSpatialiseLayer.GetShape(m_oSpatialiseLayer.GetLastUid).IsInsideExtent(GIS10.VisibleExtent, TatukGIS_XDK10.XgisInsideTypeFull) Then
            GIS10.VisibleExtent = m_oSpatialiseLayer.GetShape(m_oSpatialiseLayer.GetLastUid).Extent
        End If

        m_oSpatialiseLayer.GetShape(m_oSpatialiseLayer.GetLastUid).Flash
    
    Else
        MsgBox "there is no shape to flash!"
    End If

End Sub

Private Sub m_frmSpatialiseDD_PanLeftRight(bLeft As Boolean)
    Dim dDiffer As Double
    Dim eExtent As New TatukGIS_XDK10.XGIS_Extent
    
    With GIS10.VisibleExtent
    
        dDiffer = GIS10.VisibleExtent.ymax - GIS10.VisibleExtent.ymin
        
        If bLeft Then
        
            eExtent.Prepare .xmin - dDiffer / 4, .ymin, .xmax - dDiffer / 4, .ymax
        
        Else
        
            eExtent.Prepare .xmin + dDiffer / 4, .ymin, .xmax + dDiffer / 4, .ymax
        
        End If
    
    End With
    
    GIS10.VisibleExtent = eExtent
End Sub

Private Sub m_frmSpatialiseDD_PanUpDown(bUp As Boolean)

    Dim dDiffer As Double
    Dim eExtent As New TatukGIS_XDK10.XGIS_Extent
    
    With GIS10.VisibleExtent
    
        dDiffer = GIS10.VisibleExtent.ymax - GIS10.VisibleExtent.ymin
        
        If Not bUp Then
        
            eExtent.Prepare .xmin, .ymin - dDiffer / 4, .xmax, .ymax - dDiffer / 4
        
        Else
        
            eExtent.Prepare .xmin, .ymin + dDiffer / 4, .xmax, .ymax + dDiffer / 4
        
        End If
    
    End With
    
    GIS10.VisibleExtent = eExtent

End Sub

Private Sub m_frmSpatialiseDD_UpdateShape(oShape As TatukGIS_XDK10.XGIS_Shape)
    
    If Not oShape.IsInsideExtent(GIS10.VisibleExtent, TatukGIS_XDK10.XgisInsideTypeFull) Then
        
        MsgBox "Click flash to zoom to the shape", vbInformation, "Spatial point off the screen"
       
    End If
    
    If m_oSpatialiseLayer Is Nothing Then
        Set m_oSpatialiseLayer = New TatukGIS_XDK10.XGIS_LayerVector
        m_oSpatialiseLayer.Name = GUIDGen
    End If
    
    If GIS10.get(m_oSpatialiseLayer.Name) Is Nothing Then
        GIS10.Add m_oSpatialiseLayer
    End If
    
    Do Until Not m_oSpatialiseLayer.items.Count > 0
        m_oSpatialiseLayer.Delete m_oSpatialiseLayer.GetLastUid
        m_oSpatialiseLayer.SaveData
    Loop
    
    If Not mDDLayer.FileInfo = "Generic Vector Layer" Then mDDLayer.SaveAll
    m_oSpatialiseLayer.SaveAll
    Set oShape = mDDLayer.AddShape(oShape, True)
    If Not mDDLayer.FileInfo = "Generic Vector Layer" Then mDDLayer.SaveAll
    UpdateSpatialiseDDList oShape

End Sub

Private Sub m_frmSpatialiseDD_ZoomIn(bZoomIn As Boolean)

    If bZoomIn Then
        GIS10.zoom = GIS10.zoom * 2
    Else
        GIS10.zoom = GIS10.zoom / 2
    End If

End Sub

Private Sub m_frmSpatialiseDD_CreateObject(sObjectType As String)

    If m_oSpatialiseLayer Is Nothing Then
        Set m_oSpatialiseLayer = New TatukGIS_XDK10.XGIS_LayerVector
        m_oSpatialiseLayer.Name = GUIDGen
    Else
        m_oSpatialiseLayer.SaveAll
    End If
    
    If GIS10.get(m_oSpatialiseLayer.Name) Is Nothing Then
        GIS10.Add m_oSpatialiseLayer
    End If
    
    Do Until Not m_oSpatialiseLayer.items.Count > 0
        m_oSpatialiseLayer.Delete m_oSpatialiseLayer.GetLastUid
        m_oSpatialiseLayer.SaveData
    Loop

    ReDim ptsSel(0)
    ReDim gdpts(0)
    
    Select Case sObjectType

        Case "Point"
        
            g_CurrentTool = oCreateLocationPoint 'oCreateLocationPoint
            g_CurrentFeatureType = Custom
            m_bPolyL = False
            m_bPolyG = False
            SetOASISTool

        Case "Polyline"
        
            GIS10.Mode = TatukGIS_XDK10.XgisEdit
            g_CurrentTool = oLineSelect
            m_bPolyL = True
            m_bPolyG = False
            SetOASISTool
        
        Case "Polygon"

            GIS10.Mode = TatukGIS_XDK10.XgisEdit ' TatukGIS_XDK10.XGISUserDefined
            g_CurrentTool = oAreaSelect
            m_bPolyL = False
            m_bPolyG = True
            SetOASISTool

        Case Else
            
            GIS10.Mode = TatukGIS_XDK10.XgisDrag
            EndEdit
            
    End Select
    
End Sub

Public Sub DynamicDataModule1_ConvertMGRStoPT(sMGRS As String, _
                                              x As Double, _
                                              y As Double)

    ConvMGRStoPT sMGRS, x, y

End Sub

Public Sub m_frmSpatialiseDD_ConvertMGRS(sMGRS As String, x As Double, y As Double)
ConvMGRStoPT sMGRS, x, y

End Sub

Public Sub m_frmSpatialiseDD_RestoreDDWindow(bCommitShape As Boolean)
    
    Dim x As Double
    Dim y As Double
    'Dim sMGRS As String
    
    If m_frmSpatialiseDD.Tag = "NON SPATIAL TABLE" Then
        If bCommitShape Then
            m_frmSpatialiseDD.List1.ListIndex = 0
            x = m_frmSpatialiseDD.List1.Text
            m_frmSpatialiseDD.List2.ListIndex = 0
            y = m_frmSpatialiseDD.List2.Text
            DynamicDataModule1.SaveNonSpatialXY x, y, ConvPTtoMGRS(x, y)
        End If
        
        mDDLayer.GetShape(mDDLayer.GetLastUid).Delete
    End If
    
    If Not m_oSpatialiseLayer Is Nothing Then
        GIS10.Delete m_oSpatialiseLayer.Name
        Set m_oSpatialiseLayer = Nothing
    End If
    
    EndEdit
    m_bPolyG = False
    m_bPolyL = False
    g_CurrentTool = m_prevTool
    
    elMap.Visible = False
    elDynamicData.Visible = True
    AB.ClientAreaControl = elDynamicData
    AB.Enabled = True

    If Not m_frmSpatialiseDD.Tag = "NON SPATIAL TABLE" Then

        If Not bCommitShape Then
            Set mDDLayer = Nothing
            DynamicDataModule1.SaveChanges False, Nothing, m_frmSpatialiseDD.GetKMLForShape
        Else
            DynamicDataModule1.SaveChanges True, mDDLayer, m_frmSpatialiseDD.GetKMLForShape
        
        End If
    End If
    
End Sub

Private Sub m_frmMnuDynamicDataModule_DDTableClicked()
    DynamicDataModule1.ListDataElements_Click
End Sub

Private Sub m_frmMnuDynamicReportsModule_DatabaseClicked()
    DynamicDataReports1.cmbDatabases_Click
End Sub

Private Sub m_frmMnuDynamicReportsModule_DRListClicked()
    DynamicDataReports1.listQueries_Click
End Sub

Private Sub m_frmMnuDynamicReportsModule_DRfilterClicked()
    DynamicDataReports1.DisplayChartSelected False
End Sub


Public Sub m_frmMnuDynamicDataModule_DDDatabaseClicked()
    DynamicDataModule1.ListDatabases_Click ' ComboDatabases_Click
End Sub

Private Sub cmdDoConversion_Click()
    Dim oPt As TatukGIS_XDK10.XGIS_Point
    Dim oLyr As TatukGIS_XDK10.XGIS_LayerAbstract
    

    
    DebugPrint ""
End Sub

Private Sub DoLineSelect()
    ReDim ptsSel(0)
    ReDim gdpts(0)
    g_CurrentTool = oLineSelect
    m_bPolyL = True
End Sub

Private Sub DoPolyLineSelect()
    ReDim ptsSel(0)
    ReDim gdpts(0)
    GIS10.Mode = TatukGIS_XDK10.XgisUserDefined
    g_CurrentTool = oAreaSelect
    m_bPolyG = True
End Sub

Private Sub cmdPolyEdit_Click()
' enmCurTool = GeoFencing

    DoPolyLineSelect

End Sub

Private Sub cmdSelectCircle_Click()
    g_CurrentTool = oRadiusSelect
    GIS10.Mode = TatukGIS_XDK10.XgisSelect
End Sub

Private Sub LoadWMS_WFS(Optional sWMSURL As String)
        '<EhHeader>
        On Error GoTo LoadWMS_WFS_Err
        '</EhHeader>
    Dim oLyr As New TatukGIS_XDK10.XGIS_LayerWMS
    'TODO 1) Add a check for internet connections 2) A nicer way to choose what layers are opened... Now all are opened
    '3) Add to check Clipboard...
        
        If sWMSURL = "" Then
            sWMSURL = Clipboard.GetText
            
            If InStr(sWMSURL, "http") = 0 Then sWMSURL = ""
            
            If Not sWMSURL = "" Then
                If MsgBox("You have some text in the cliboard, do you want use this as the url..." & vbCrLf & sWMSURL, vbYesNo, "Loading WMS URL from clipboard") = vbYes Then
                    
                Else
                    sWMSURL = InputBox("Enter URL to the WMS", "Add WMS Layers to OASIS")
                End If
            Else
                sWMSURL = InputBox("Enter URL to the WMS", "Add WMS Layers to OASIS", "http://demo.mapserver.org/cgi-bin/wms?SERVICE=WMS&VERSION=1.1.1&REQUEST=GetCapabilities")
            End If
        End If
        
100     If sWMSURL = "" Then sWMSURL = "http://demo.mapserver.org/cgi-bin/wms?SERVICE=WMS&VERSION=1.1.1&REQUEST=GetCapabilities"

102     With oLyr
104         .Path = sWMSURL
106         .Name = .FoundLayers.Text
108         .Open
        End With
    
110     GIS10.Add oLyr

        '<EhFooter>
        Exit Sub

LoadWMS_WFS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.LoadWMS_WFS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdMSGScroller_Click()

    If cmdMSGScroller.caption = "6" Then 'Expand
        cmdMSGScroller.caption = "5"
        elScroller.Height = 255
        elScroller.Width = elTatukGIS.Width
        If Not mnuOPSView.Checked Then GIS10.Height = elTatukGIS.Height + 150
        MsgScroll.AutoScroll = True
    Else
        cmdMSGScroller.caption = "6"
        elScroller.Height = 255
        elScroller.Width = cmdMSGScroller.Width
        MsgScroll.AutoScroll = False
        If Not mnuOPSView.Checked Then GIS10.Height = elTatukGIS.Height + 400
    End If

End Sub

Private Sub cmdRemove_Click()
    If g_lPinUID > 0 Then
        m_oDrawLyr.Delete g_lPinUID
        GIS10.UpDate
        g_lPinUID = 0
    End If
End Sub

Private Sub popInArray(ByRef aArray As Variant, xElement As Variant)
    Dim nLen As Long
    Dim aTemp() As Variant
    Dim i As Long
    
    nLen = aLen(aArray)
    
    For i = 0 To nLen - 1
        If i <> xElement Then
            aAdd aTemp, aArray(i)
        End If
    Next i
    
    aArray = aTemp
End Sub

Private Sub cmdSelTool_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo cmdSelTool_Click_Err
        '</EhHeader>
        Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
        Dim shp As TatukGIS_XDK10.XGIS_Shape

        Dim aTemp() As ShpProps
        Dim i As Long
    
100     Select Case Index
    
            Case 0
                'Close
102             If AttribTabs.NumTabs = 0 Then
104                  elDynHolder(0).Visible = False
                    Exit Sub
                End If
            
106             For i = 0 To UBound(m_shpProps)

108                 If AttribTabs.CurrTab = 0 And i = 0 Then ReDim aTemp(0)

110                 If i <> AttribTabs.CurrTab Then
112                     If i = 0 Then
114                         ReDim aTemp(0)
                        Else
116                         ReDim Preserve aTemp(UBound(aTemp) + 1)
                        End If
                    
118                     aTemp(UBound(aTemp)) = m_shpProps(i)
                    End If

120             Next i
            
122             m_shpProps = aTemp
            
124             AttribTabs.RemoveTab AttribTabs.CurrTab
            
126             If AttribTabs.NumTabs = 0 Then
128                 elDynHolder(0).Visible = False
                Else
130                 AttribTabs.CurrTab = AttribTabs.NumTabs - 1
                End If

132         Case 1

134             If Not frmGeoSummary.Visible Then
136                 frmGeoSummary.Show vbModeless, Me
                End If
 
138             Set oLyr = GIS10.get(m_shpProps(AttribTabs.CurrTab).sLayerName)
            
140             If Not oLyr Is Nothing Then
142                 frmGeoSummary.SetGeoSum oLyr
144                 frmGeoSummary.SetFocus
                End If
 
146         Case 2
148             m_frmSelectorReports.Show vbModal, Me

150         Case 3

152             With m_udtSelectorSettings
154                 m_frmSelectorSettings.Init .dBuffeLevel, .sSpatialOperator, .bAutoZoom, .bAutoSelect, .bAutoFlash, .bAutoClear, .bEdit
156                 m_frmSelectorSettings.Move cmdSelTool(0).Container.left$, cmdSelTool(0).Container.top
158                 m_frmSelectorSettings.Show vbModal, Me
                End With

160         Case 4
        
         '   Dim i As Integer
        
162             For i = 0 To GIS10.items.Count - 1
164                 If GisUtils.IsInherited(GIS10.items.Item(i), "XGIS_LayerVector") Then
166                     GIS10.items.Item(i).Deselect.All
                    End If
                Next

        End Select

        '<EhFooter>
        Exit Sub

cmdSelTool_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdSelTool_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdShowSelectedInfo_Click()
    
    m_frmAttributes.GEO1.ShowSelected GIS10.get(GISGridLayerName)
    
End Sub

Private Sub cmdSingleSelect_Click()
    GIS10.Mode = TatukGIS_XDK10.XgisSelect
    g_CurrentTool = oSingleSelect
End Sub



Private Sub GetNewIncidents(DummyArgument As Variant)
'        'Dim MsXmlHttp As New MSXML2.XMLHTTP
'        'Dim MsXmlDoc As New MSXML2.DOMDocument
'        '<EhHeader>
'        On Error GoTo GetNewIncidents_Err
'        '</EhHeader>
'        Dim sSQL As String
'        Dim oRSSynchHistory As ADODB.Recordset
'        Dim oRSServer As ADODB.Recordset
'        Dim oRSLocalInc As ADODB.Recordset
'        Dim lMAX As Long
'        Dim oRSMax As New ADODB.Recordset
'        Dim sWebsite As String
'        Dim sString As String
'        Dim RSUpdater As ADODB.Recordset
'
'100     Set oRSSynchHistory = New ADODB.Recordset
'
'102     oRSMax.Open "SELECT MAX(UID) FROM oincidents_GEO", m_Cnn, adOpenDynamic, adLockReadOnly
'
'104     If Not oRSMax.EOF And Not oRSMax.BOF Then
'106         If Not IsNull(oRSMax.Fields(0).Value) Then
'108             lMAX = oRSMax.Fields(0).Value
'            End If
'        End If
'
'110     With oRSSynchHistory
'112         .Open "SELECT * FROM SynchHistory WHERE [updates] = 'true'", m_Cnn, adOpenDynamic, adLockReadOnly
'
'114         If Not .EOF And Not .BOF Then
'116             SafeMoveFirst oRSSynchHistory
'
'118             Do While Not .EOF
'120                 sSQL = sSQL & IIf(Len(sSQL) > 0, " AND ID <> '" & .Fields.Item("sID").Value & "'", "ID <> '" & .Fields.Item("sID").Value & "'")
'122                 .MoveNext
'                Loop
'
'            End If
'
'124         Set oRSServer = New ADODB.Recordset
'
'126         sWebsite = g_sAppServerPath
'
'128         If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
'130             sWebsite = sWebsite & "/"
'            End If
'
'132         If Len(sSQL) > 1 Then
'134             sString = sWebsite & "oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM oincidents WHERE " & sSQL)
'                'Set oRSServer = OpenSilentHttpCommsRS(sString, True)
'136             Set oRSServer = New ADODB.Recordset
'138             oRSServer.Open sString
'
'            Else
'140             sString = sWebsite & "oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM oincidents")
'                'Set oRSServer = OpenSilentHttpCommsRS(sString, True)
'142             Set oRSServer = New ADODB.Recordset
'144             oRSServer.Open sString
'            End If
'
'146         If oRSServer.State = 0 Then
'                Exit Sub
'            End If
'
'        End With
'
'148     With oRSServer
'
'150         If SafeMoveFirst(oRSServer) Then
'
'152             Do While Not .EOF
'
'154                 lMAX = lMAX + 1
'
'156                 Set RSUpdater = New ADODB.Recordset
'158                 With RSUpdater
'
'160                     .Open "SELECT * FROM oincidents_FEA", m_Cnn, adOpenDynamic, adLockBatchOptimistic
'162                     .AddNew
'
'164                     .Fields("UID").Value = lMAX
'166                     .Fields("ID").Value = oRSServer.Fields("ID").Value
'168                     .Fields("NAME").Value = oRSServer.Fields("NAME").Value
'170                     .Fields("TYPE").Value = oRSServer.Fields("TYPE").Value
'172                     .Fields("TARGET").Value = oRSServer.Fields("TARGET").Value
'
'174                     If Not IsNull(oRSServer.Fields("Dead").Value) Then
'176                         .Fields("Dead").Value = oRSServer.Fields("Dead").Value
'                        Else
'178                         .Fields("Dead").Value = 0
'                        End If
'
'180                     If Not IsNull(oRSServer.Fields("Affected").Value) Then
'182                         .Fields("Affected").Value = oRSServer.Fields("Affected").Value
'                        Else
'184                         .Fields("Affected").Value = 0
'                        End If
'
'186                     If Not IsNull(oRSServer.Fields("Violent").Value) Then
'188                         .Fields("Violent").Value = oRSServer.Fields("Violent").Value
'                        Else
'190                         .Fields("Violent").Value = 0
'                        End If
'
'192                     If Not IsNull(oRSServer.Fields("Injured").Value) Then
'194                         .Fields("Injured").Value = oRSServer.Fields("Injured").Value
'                        Else
'196                         .Fields("Injured").Value = 0
'                        End If
'
'198                     .Fields("Incident_DATE").Value = oRSServer.Fields("Incident_DATE").Value
'200                     .Fields("TIME00").Value = oRSServer.Fields("TIME00").Value
'202                     .Fields("Town").Value = oRSServer.Fields("Town").Value
'204                     .Fields("District").Value = oRSServer.Fields("District").Value
'206                     .Fields("Province").Value = oRSServer.Fields("Province").Value
'208                     .Fields("Description").Value = oRSServer.Fields("Description").Value
'
'210                     If Not IsNull(oRSServer.Fields("Scoring").Value) Then
'212                         .Fields("Scoring").Value = oRSServer.Fields("Scoring").Value
'                        Else
'214                         .Fields("Scoring").Value = 0
'                        End If
'
'216                     .Fields("Incident_DATESERIAL").Value = oRSServer.Fields("Incident_DATESERIAL").Value
'
'218                     .UpdateBatch adAffectCurrent
'220                     .Close
'
'                    End With
'222                 Set RSUpdater = Nothing
'
'224                 Set RSUpdater = New ADODB.Recordset
'226                 With RSUpdater
'
'228                     .Open "SELECT * FROM oincidents_GEO", m_Cnn, adOpenDynamic, adLockBatchOptimistic
'
'230                     .AddNew
'232                     .Fields("UID").Value = lMAX
'234                     .Fields("XMIN").Value = oRSServer.Fields("XMIN").Value
'236                     .Fields("XMAX").Value = oRSServer.Fields("XMAX").Value
'238                     .Fields("YMIN").Value = oRSServer.Fields("YMIN").Value
'240                     .Fields("YMAX").Value = oRSServer.Fields("YMAX").Value
'242                     .Fields("SHAPETYPE").Value = 2
'
'244                     .UpdateBatch adAffectCurrent
'246                     .Close
'
'                    End With
'248                 Set RSUpdater = Nothing
'
'250                 SetNewSynchDBElement GetGuid, oRSServer!ID, "Synched Incident", "", g_sRemoteTablePrefix & " UG", RFC3339DateTime, "oincidents", True, "'true'"
'252                 .MoveNext
'
'                Loop
'
'            End If
'
'        End With
'
'        '<EhFooter>
'        Exit Sub
'
'GetNewIncidents_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.GetNewIncidents " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub cmdTracker_Click()
'            ShellExecute g_lHw, vbNullString, g_sAppPath & "\bin\OASISCommsMon.exe", "1^" & m_Cnn.ConnectionString & "^" & g_sRemoteTablePrefix & "^" & g_sAppServerPath & "^" & g_bHasEncrypt & "^" & g_sKey & "^" & Me.Hwnd & "^" & tmrInternetCheck.Interval, "C:\", 1
 
    frmOASISTracker.Show
    
Dim lHandle As Long


    'lHandle = ShellExecute(Me.Hwnd, vbNullString, "C:\Program Files\Globalsat\TR Management Center\OASIS_Tracker.exe", "", "c:\", 1)
    
   lHandle = FindWindow(vbNullString, "TR Management Center")
    
   SetParent lHandle, frmOASISTracker.hWnd
   
   MoveWindow lHandle, 0, -60, 1100, 750, 1
   
End Sub

Private Sub cmdZoomToSettings_Click()
    frmZoomToSettings.Show vbModal, Me
End Sub

Private Sub dxGISDataGrid_OnDragEndHeader(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal x As Single, ByVal y As Single, accept As Boolean)
        '<EhHeader>
        On Error GoTo dxGISDataGrid_OnDragEndHeader_Err
        '</EhHeader>
100     If dxGISDataGrid.Filter.FilterActive Then
102         If InStr(dxGISDataGrid.Filter.FilterCaption, Column.caption) = 0 Then
104             dxGISDataGrid.M.AddGroupColumn Column
106             dxGISDataGrid.M.LoadColumnValues Column
108             dxGISDataGrid.M.RefreshGroupColumns
110             dxGISDataGrid.M.FreeColumnValues Column
112             dxGISDataGrid.M.FullRefresh
114             dxGISDataGrid.M.Invalidate
116             'dxGISDataGrid.m.DeleteGroupColumn Column.GroupIndex
            End If
        End If
        '<EhFooter>
        Exit Sub

dxGISDataGrid_OnDragEndHeader_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.dxGISDataGrid_OnDragEndHeader " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub FraPinColor_Click(Index As Integer)
Dim c As cCommonDialog
    Set c = New cCommonDialog
    c.ShowColor

    FraPinColor(Index).BackColor = c.color

End Sub

Private Sub GIS10_OnDblClick(translated As Boolean)
        '<EhHeader>
        On Error GoTo GIS10_OnDblClick_Err
        '</EhHeader>
        
        Dim tmp As TatukGIS_XDK10.IXGIS_Shape
        Dim lL As TatukGIS_XDK10.IXGIS_LayerVector
        Dim oshp As TatukGIS_XDK10.XGIS_Shape
        Dim bRunUrl As Boolean
        
100     DebugPrint "----- **DBL Click Start** ----"
102     DebugPrint "gdpts:" & UBound(gdpts)
104     DebugPrint "ptsSel:" & UBound(ptsSel)
        
106     m_bScrib = False
108     m_bRECT = False
    
110     If m_bPolyG Or m_bPolyL Then
    
112         RemoveAlltabs
        
114         If UBound(gdpts) = 0 Then Exit Sub
116         ReDim Preserve gdpts(UBound(gdpts) - 1)
118         ReDim Preserve ptsSel(UBound(ptsSel) - 1)
120         GIS10.PaintMinimum
        
122         If m_bPolyG Then
124             m_SelUID = CreatePolygonEX
            Else
126             m_SelUID = CreatePolyLine
            End If
        
128         m_bPolyG = False
130         m_bPolyL = False
132         EndEdit
        
134         ReDim gdpts(0)
136         ReDim ptsSel(0)
        
138         m_bPolyL = False
140         m_bPolyG = False
        
142         m_bDrawFinished = True
144         translated = True

146         AddShapeWithBufferAndHighlightIntersect m_oBufferLyr.GetShape(m_SelUID), ctlSelector1.GetBuffer
            
            Exit Sub
        End If
        
160     Set oshp = GIS10.Locate(m_dPTG, 5 / GIS10.zoom)

162     If Not oshp Is Nothing Then

164         If Not g_RSGISGridTableSettings.EOF Or Not g_RSGISGridTableSettings.Bof Then
                        
166             SafeMoveFirst g_RSGISGridTableSettings
168             g_RSGISGridTableSettings.Find "name = '" & oshp.layer.Name & "'"
                    
170             If Not g_RSGISGridTableSettings.EOF Then
172                 If g_RSGISGridTableSettings.Fields.Item("isURLLayer").value Then
174                     If g_RSGISGridTableSettings.Fields.Item("autoRunUrls").value Then
176                         bRunUrl = True
                        End If
                    End If
                End If
    
178             If bRunUrl Then
180                 If oshp.GetField(g_RSGISGridTableSettings.Fields.Item("URLLayerField").value) <> "" Then
182                     ShellExecute Me.hWnd, vbNullString, oshp.GetField(g_RSGISGridTableSettings.Fields.Item("URLLayerField").value), vbNullString, "C:\", 1
                        Exit Sub
                    End If
                End If
                
            End If
                        
        End If
    
        If g_CurrentTool = oInfo Or g_CurrentTool = oPan Or g_CurrentTool = oZoom Or g_CurrentTool = oZoom2FullExtent Or g_CurrentTool = oZoomEx Or g_CurrentTool = oZoomIn Or g_CurrentTool = oZoomOut Then
            
            GIS10.Lock
            GIS10.viewer.CenterViewport m_dPTG
            If ctlZoomSlider1.GetZoom < 2 Then
                GIS10.zoom = GIS10.zoom * 2
            Else
                ctlZoomSlider1.ZoomIn
            End If
            
            GIS10.Unlock
    
        End If

184     DebugPrint "----- **DblClick End** ----"
186     DebugPrint "gdpts:" & UBound(gdpts)
188     DebugPrint "ptsSel:" & UBound(ptsSel)
        
        '<EhFooter>
        Exit Sub

GIS10_OnDblClick_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GIS10_OnDblClick " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS_OnLayerAdd(translated As Boolean, ByVal layer As Object)
'    On Error GoTo Hell
'
'    Dim i As Integer
'    Dim j As Long
'
'    For i = 0 To 10000
'        j = j + 1
'    Next
'
'
'    Exit Sub
'Hell:
'
'    MsgBox Err.Description
End Sub

Private Sub m_frmMainSettings_ScoringSettings()
    m_frmMnuOperations_DoSecurityAnalysis
End Sub

Private Sub m_frmMnuOperations_ResetSecurityAnalysis()
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_ResetSecurityAnalysis_Err
        '</EhHeader>
    
100     If Not EventLayer Is Nothing Then
102         ResetScoring GIS10.get(EventLayer.Name)
114         GIS10.UpDate
            DoEvents
        End If
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_ResetSecurityAnalysis_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_ResetSecurityAnalysis " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_ScoringSettings()
    cmdSetScoring_Click
End Sub

Private Sub m_frmOPSView_Closing()
    
    SetParent GIS10.hWnd, elTatukGIS.hWnd
    elTatukGIS.Refresh
    C1TabFastFunction.Width = 340
    GIS10.Width = elTatukGIS.Width
    GIS10.Height = elTatukGIS.Height
    GIS10.ZOrder 0
    C1TabFastFunction.Width = 340
    C1TabFastFunction.Height = 3900
    C1TabFastFunction.ZOrder 0
    mnuOPSView.Checked = False
    
    On Error Resume Next
    Unload m_frmOPSView
    Set m_frmOPSView = Nothing
    Call elTatukGIS_RealignFrame
    
End Sub


Private Sub GetSHP(uID As Long, sLayerName As String, oshp As TatukGIS_XDK10.XGIS_Shape)
    Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
    'GIS109 Done
    Set oLyr = GIS10.get(sLayerName)
    
    For Each oshp In oLyr.Loop(oLyr.Extent, "GIS_UID = " & uID, Nothing, "", True)
        If oshp Is Nothing Then Exit For
    Next

End Sub

Private Function GetNearestShape(PassedPoint As TatukGIS_XDK10.XGIS_Point, _
                                 sLayerName As String, Optional bFlash As Boolean = False) As Double
    
    Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
    Dim oPoint As TatukGIS_XDK10.XGIS_Point
    Dim oshp As TatukGIS_XDK10.XGIS_Shape
    Dim oDistance As Double
    Dim oPart As Long
    Dim oUtil As New TatukGIS_XDK10.XGIS_Utils
    Dim oShapeForPoint As New TatukGIS_XDK10.XGIS_Shape

    Set oLyr = GIS10.get(sLayerName)
    GetNearestShape = 0.00006666666
    If Not oLyr Is Nothing Then
    
        Set oshp = oLyr.Locate(PassedPoint, 1000000000, False)
        
        If Not oshp Is Nothing Then
        
            'oDistance = HaversineDistance(PassedPoint.x, oshp.Centroid.x, PassedPoint.y, oshp.Centroid.y)
            
            Dim oShapeFromPoint As XGIS_Shape
           Set oShapeFromPoint = GisUtils.GisCreateShapeFromWKT("POINT(" & PassedPoint.x & " " & PassedPoint.y & ")")
            
            
            oDistance = oshp.Distance2Shape(oShapeFromPoint)
            
            Set oShapeFromPoint = Nothing
            
            'oshp.distance(PassedPoint, 1000000000)
            'oDistance = oDistance * 111.3199
        
            'If Not oLyr.Locate(PassedPoint, 5 / GIS10.zoom, False) Is Nothing Then
            '    GetNearestShape = 0
            'Else
                'oShp.Flash
                If bFlash Then oshp.Flash 6
                GetNearestShape = oDistance
            'End If
            
        End If
    End If
    
End Function

Private Function DistanceBetween2Points(oCentroidPoint As TatukGIS_XDK10.XGIS_Point, _
                                        oshp As TatukGIS_XDK10.XGIS_Shape) As Double
'Exit Function

    Dim oCentriodShape As TatukGIS_XDK10.XGIS_Shape
    DistanceBetween2Points = 0.00006666666
    
    Set oCentriodShape = GIS10.get("Buffers").CreateShape(XgisShapeTypeUnknown)
    oCentriodShape.Lock TatukGIS_XDK10.XgisLockExtent
    oCentriodShape.AddPart
    oCentriodShape.AddPoint oCentroidPoint
    oCentriodShape.Unlock
    oCentriodShape.Delete
    
    If Not oshp Is Nothing Then
        If oshp.Contains(oCentriodShape) Then
            DistanceBetween2Points = 0
        Else
            'DistanceBetween2Points = oshp.distance(oCentroidPoint, 1000000000)
            'DistanceBetween2Points = DistanceBetween2Points * 111.3199
            
            DistanceBetween2Points = HaversineDistance(oCentroidPoint.x, oshp.Centroid.x, oCentroidPoint.y, oshp.Centroid.y)
            
        End If
    End If
    
End Function

Private Function GetDistanceToAllShapes(PassedPoint As TatukGIS_XDK10.XGIS_Point, _
                                        sLayerName As String, sDescField As String, lIndexOfFirstItem As Long, dMultiplyFactor As Double) As String
    
    Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
    Dim oPoint As TatukGIS_XDK10.XGIS_Point
    Dim oshp As TatukGIS_XDK10.XGIS_Shape
    Dim oDistance As Double
    Dim oPart As Long
    Dim oUtil As New TatukGIS_XDK10.XGIS_Utils
    Dim oShapeForPoint As New TatukGIS_XDK10.XGIS_Shape
    Dim sDistance As String
    Dim i As Long

    Dim oShp9 As TatukGIS_XDK10.XGIS_Shape

    Set oLyr = GIS10.get(sLayerName)
    
    For Each oShp9 In oLyr.Loop(oLyr.Extent, "", Nothing, "", True)
        If Not GetDistanceToAllShapes = "" Then GetDistanceToAllShapes = GetDistanceToAllShapes & "|||"
        oDistance = Round(oShp9.distance(PassedPoint, 1000000000) * dMultiplyFactor, 0)
        sDistance = CStr(oDistance)
        If Len(sDistance) < 4 Then sDistance = "0" & sDistance
        If Len(sDistance) < 4 Then sDistance = "0" & sDistance
        If Len(sDistance) < 4 Then sDistance = "0" & sDistance
        If Len(sDistance) < 4 Then sDistance = "0" & sDistance
        If Len(sDistance) < 4 Then sDistance = "0" & sDistance
        GetDistanceToAllShapes = GetDistanceToAllShapes & sDistance & "km to " & oShp9.GetField(sDescField) & ".   "
        
        i = lIndexOfFirstItem
        Do Until i >= oLyr.Fields.Count
        
            GetDistanceToAllShapes = GetDistanceToAllShapes & "[" & Replace(oLyr.FieldInfo(i).Name, "_", " ") & ": "
            GetDistanceToAllShapes = GetDistanceToAllShapes & oShp9.GetField(oLyr.FieldInfo(i).Name) & "]    "
            i = i + 1
            
        Loop
    Next
        
    GetDistanceToAllShapes = GetDistanceToAllShapes & "|||"
    
End Function


Private Sub ChangeSelectionType(enmtool As OASIS_TOOLS)
        '<EhHeader>
        On Error GoTo ChangeSelectionType_Err
        '</EhHeader>
        
100     If m_SelUID > 0 Then
102         m_oDrawLyr.Delete m_SelUID
104         m_SelUID = 0
        End If
        
106     If m_lBufUID > 0 Then
108         GIS10.get("Buffers").Delete m_lBufUID
        End If
        
110     Select Case enmtool
        
            Case OASIS_TOOLS.oAreaSelect

112             DoPolyLineSelect

114         Case OASIS_TOOLS.oCircleSelect

116             doCircleSelect

118         Case OASIS_TOOLS.oFeatureSelect

120             doFeatureSelect2

122         Case OASIS_TOOLS.oLineSelect

124             DoLineSelect
            
        End Select

        ChangeToolCursor

        '<EhFooter>
        Exit Sub

ChangeSelectionType_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ChangeSelectionType " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub doFeatureSelect2()
    g_CurrentTool = oFeatureSelect
    GIS10.Mode = TatukGIS_XDK10.XgisUserDefined
End Sub

Private Sub doCircleSelect()
    g_CurrentTool = oCircleSelect
    GIS10.Mode = TatukGIS_XDK10.XgisSelect
End Sub

Public Sub m_frmSpatialAnalysis_GetAttributes(sLayerName As String, _
                                                  sAttribs() As String, _
                                                  bDateOnly As Boolean, _
                                                  bAnyAttribs As Boolean)
        '<EhHeader>
        On Error GoTo m_frmSpatialAnalysis_GetAttributes_Err
        '</EhHeader>

        Dim i As Integer
        Dim j As Integer
        Dim jj As Integer
        Dim lL As TatukGIS_XDK10.XGIS_LayerAbstract
        Dim sLocalAttribs() As String
        Dim iNumberOfAttribs As Integer
    
100     iNumberOfAttribs = 0
         
102     Set lL = GIS10.get(sLayerName)
    
        'rDim sLocalAttribs(0)

104     If Not lL.SubType = 1 Then
106         j = 0
108         jj = lL.Fields.Count

110         Do Until j = jj
                'MsgBox ll.Fields.Item(j).FieldType & " for " & ll.Fields.Item(j).Name

112             If Not bDateOnly Or (lL.Fields.Item(j).FieldType = 4 And bDateOnly) Then
            
114                 iNumberOfAttribs = iNumberOfAttribs + 1
116                 ReDim Preserve sLocalAttribs(iNumberOfAttribs - 1)
118                 sLocalAttribs(iNumberOfAttribs - 1) = lL.Fields.Item(j).Name
                End If

                'If j = 0 Then MsgBox ll.Fields.Item(j).Name
120             j = j + 1
            Loop

            'If Not UBound(sLocalAttribs) = 0 Then ReDim Preserve sLocalAttribs(UBound(sLocalAttribs) - 1)
122         If iNumberOfAttribs = 0 Then
124             bAnyAttribs = False
126             sAttribs = sLocalAttribs
            Else
128             bAnyAttribs = True
130             sAttribs = sLocalAttribs
            End If

        End If
        
        '<EhFooter>
        Exit Sub

m_frmSpatialAnalysis_GetAttributes_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmSpatialAnalysis_GetAttributes " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ZoomToAdminLyr(sLayer As String, sField As String, sValue As String)
        '<EhHeader>
        On Error GoTo ZoomToAdminLyr_Err
        '</EhHeader>

    'GIS109 Done
    Dim oLyr9 As TatukGIS_XDK10.XGIS_LayerVector
    Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
    
100 Set oLyr9 = GIS10.get(sLayer)

102 For Each oShp9 In oLyr9.Loop(oLyr9.Extent, sField & " = '" & sValue & "'", Nothing, "", True)
104     If Not oShp9 Is Nothing Then
106         GIS10.VisibleExtent = oShp9.Extent
108         Set oShp9 = oShp9.MakeEditable
110         oShp9.Flash
            Exit For
        End If
    Next
    
        '<EhFooter>
        Exit Sub

ZoomToAdminLyr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ZoomToAdminLyr " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub AddBtn(btnIndex As Integer, _
                   dStyle As ActiveBar3LibraryCtl.ToolStyles, _
                   sCaption As String, _
                   sName As String, _
                   lngToolID As Long, _
                   sBandName As String)
        '<EhHeader>
        On Error GoTo AddBtn_Err
        '</EhHeader>

100     Set btns(btnIndex) = New ActiveBar3LibraryCtl.Tool
    
102     Set btns(btnIndex) = AB.Bands(sBandName).Tools.Add(lngToolID, sName)

104     With btns(btnIndex)
106         .caption = sCaption
            '.Category = "File"
108         .ControlType = ddTTButton
110         .Style = dStyle
            '      .ControlType = ddTTButton
            '   .Style = ddSIcon
            '   .caption = " File Save"
            '    .Category = "File"
            '   .SetPicture  ddITNormal, LoadPicture(g_sAppPath & "\FSave.bmp")
        End With

112     AB.RecalcLayout
    
        ' AB.Bands.Item("").Tools.Add
        '<EhFooter>
        Exit Sub

AddBtn_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.AddBtn " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub AB_ComboSelChange(ByVal Tool As ActiveBar3LibraryCtl.Tool)
        '<EhHeader>
        On Error GoTo AB_ComboSelChange_Err
        '</EhHeader>
        
        On Error Resume Next
        Dim sProc As String
        
100     Select Case Tool.Name
    
            Case "comGenLyr"
    
        End Select
        
        '<EhFooter>
        Exit Sub

AB_ComboSelChange_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.AB_ComboSelChange " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CreateGenericEditLyr(oGenericLyr As TatukGIS_XDK10.XGIS_LayerSqlAdo, _
                                      sLyrName As String, _
                                      sDisplayName As String, _
                                      Optional bShowPointsNumber As Boolean, _
                                      Optional bShowTracking As Boolean, _
                                      Optional bShowFast As Boolean, _
                                      Optional bHideFromLgd As Boolean) As Boolean
        '<EhHeader>
        On Error GoTo CreateGenericEditLyr_Err
        '</EhHeader>

100     With oGenericLyr
102         .Name = sDisplayName
104         .caption = sDisplayName
106         .HideFromLegend = bHideFromLgd
        End With
      
108     Set g_oEditLayer = oGenericLyr
110     GIS10.Mode = TatukGIS_XDK10.XgisEdit
    
112     With GIS10.Editor
114         .ShowTracking = bShowTracking
116         .ShowFast = bShowFast
118         .ShowPointsNumber = bShowPointsNumber
        End With
    
120     g_CurrentFeatureType = Location
122     m_prevTool = g_CurrentTool
        'g_CurrentTool = oCreateLocationPolyline

        '<EhFooter>
        Exit Function

CreateGenericEditLyr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CreateGenericEditLyr " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdCommand4_Click()
        '<EhHeader>
        On Error GoTo cmdCommand4_Click_Err
        
        '</EhHeader>
100     CreateGenericEditLyr m_oSQLGenericLyrs(0), m_oSQLGenericLyrs(0).Name, m_oSQLGenericLyrs(0).Name, True, True, True, False
        '<EhFooter>
        Exit Sub

cmdCommand4_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.cmdCommand4_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCommand5_Click()
        '<EhHeader>
        On Error GoTo cmdCommand5_Click_Err
        '</EhHeader>

        If m_frmTextAnnoSettings Is Nothing Then Set m_frmTextAnnoSettings = New frmTextAnnoSettings
                
        ListAllAnnotationShps
        m_frmTextAnnoSettings.Show vbModeless, Me
        g_CurrentTool = oCreateLocationText
        GIS10.Mode = TatukGIS_XDK10.XgisSelect
        '<EhFooter>
        Exit Sub

cmdCommand5_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdCommand5_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdExportData_Click()
        '<EhHeader>
        On Error GoTo cmdExportData_Click_Err
        '</EhHeader>
100             frmExportFormats.FraExportFormats.Visible = False
102             frmExportFormats.Show vbModal, Me

104             If frmExportFormats.bExport Then
106                 If frmExportFormats.chkFormats(0).value = vbChecked Then
108                     dxGISDataGrid.M.ExportToXLS g_sAppPath & "\data\user\Exports\W3_OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xls"
                    End If
110                 If frmExportFormats.chkFormats(1).value = vbChecked Then
112                     dxGISDataGrid.M.ExportToXML g_sAppPath & "\data\user\Exports\W3_OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xml"
                    End If
114                 If frmExportFormats.chkFormats(2).value = vbChecked Then
116                     dxGISDataGrid.M.ExportToHTML g_sAppPath & "\data\user\Exports\W3_OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".htm"
                    End If
118                 If frmExportFormats.chkFormats(3).value = vbChecked Then
120                     dxGISDataGrid.M.SaveAllToTextFile g_sAppPath & "\data\user\Exports\W3_OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".txt"
                    End If
                End If

 

        '<EhFooter>
        Exit Sub

cmdExportData_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdExportData_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub cmdSetScope_Click()
    
    'm_frmChangeTracer.Init
    'm_frmChangeTracer.Show vbModeless, Me

End Sub

Private Sub cmdZoomTo_Click()
        '<EhHeader>
        On Error GoTo cmdZoomTo_Click_Err
        '</EhHeader>
        Dim ptg As New TatukGIS_XDK10.XGIS_Point
        Dim oshp As TatukGIS_XDK10.XGIS_Shape
        Dim sFont As String
        Dim SymbolList As TatukGIS_XDK10.XGIS_SymbolList
        Dim x As Double
        Dim y As Double
        
        x = CDbl(IIf(IsNumeric(Replace(me2Long.Text, "_", "")), Replace(me2Long.Text, "_", ""), 0))
        y = CDbl(IIf(IsNumeric(Replace(me2Lat.Text, "_", "")), Replace(me2Lat.Text, "_", ""), 0))

        If x + y = 0 Then
            MsgBox "Coordinate values do not seem to be correct! Please Check and try again", vbInformation, "OASIS Coordinate Utils"
            Exit Sub
        End If
    
100     ptg.Prepare x, y

       ' me2Lat.Text = Y
       ' me2Long.Text = X
        
102     If chkUseMarker Then
            
            
104         If Not m_oDrawLyr Is Nothing Then

                If Not g_ZoomToSettings.UseMultiple Then
                If g_lPinUID > 0 Then
                    m_oDrawLyr.Delete g_lPinUID
                    g_lPinUID = 0
                End If
                Else
                    g_lPinUID = 0
                End If
                
106             If g_lPinUID > 0 Then
                    
108                 Set oshp = m_oDrawLyr.GetShape(g_lPinUID)
                 
110                 If oshp Is Nothing Then Exit Sub
                
112                 With oshp
114                     .Lock TatukGIS_XDK10.XgisLockExtent
116                     .SetPosition ptg, m_oDrawLyr, 0
118                     .Unlock
                    End With
                
                Else
120                 Set oshp = m_oDrawLyr.CreateShape(XgisShapeTypePoint)
                    
122                 If oshp Is Nothing Then Exit Sub
                        
124                 With oshp
126                     .Lock TatukGIS_XDK10.XgisLockExtent
128                     .AddPart
130                     .AddPoint ptg

132                     With .Params.labels
134                         .Alignment = TatukGIS_XDK10.XgisLabelAlignmentCenter
136                         .color = vbRed
138                         .FontColor = vbGreen
140                         .Allocator = False
142                         .Duplicates = True
144                         .Font.Size = 45
146                         .OutlineWidth = 0
148                         .Pattern = XbsClear
150                         .Position = TatukGIS_XDK10.XgisLabelPositionMiddleCenter
152                         .value = ""
                        End With
                                        
154                     sFont = txtFont(0).FontName
156                     sFont = sFont & ":" & Asc(txtFont(0).Text) & ":NORMAL"

158                     Set SymbolList = New TatukGIS_XDK10.XGIS_SymbolList

160                     .Params.Marker.color = FraPinColor(0).BackColor
162                     .Params.Marker.OutlineColor = 16711680
164                     .Params.Marker.Symbol = SymbolList.Prepare(sFont)
166                     .Params.Marker.Size = 500
'168                     .Params.Marker.ShowLegend = 1
'170                     .Params.Legend = "Uncategorized"
                
'172                     .Params.Marker.OutlineStyle = XpsSolid
'174                     .Params.Marker.OutlineWidth = 20
'176                     .Params.Marker.OutlineColor = vbRed
'178                     .Params.Marker.Style = TatukGIS_XDK10.XGISMarkerStyleTriangleDown
'180                     .Params.Marker.Color = vbBlue
'182                     .Params.Marker.Size = 200
                        Dim ostr As TatukGIS_XDK10.XStringList
                        
                        Set ostr = New TatukGIS_XDK10.XStringList
                       'TATUK BUG in 9.x .Params.Marker.SaveToStrings ostr
                       
                        On Error Resume Next
                        
                        .SetField "StringStyle", ostr.Text
                        DebugPrint ostr.Text
184                     .Unlock
186                     g_lPinUID = .uID
                    End With

                End If

            End If
        End If

188     GIS10.CenterViewport ptg
        If Not oshp Is Nothing Then oshp.Flash 12
        
       ' m_oDrawLyr.SaveAll
        
        
190     GIS10.UpDate
    
    
        '<EhFooter>
        Exit Sub

cmdZoomTo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdZoomTo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub comConversionType_Click()
    '
    'dd.dddd
    'dd: mm: ss.ss
    'dd: mm.mmmm
    'MGRS
        '<EhHeader>
        On Error GoTo comConversionType_Click_Err
        '</EhHeader>

100     frmdddddd.Visible = False
102     frddmmmm.Visible = False
104     frddmmss.Visible = False
106     frMGRS.Visible = False
    
108     Select Case comConversionType.ListIndex
    
            Case 0
110             frmdddddd.Visible = True
112         Case 1
114             frddmmmm.Visible = True
116         Case 2
118             frddmmss.Visible = True
120         Case 3
122             frMGRS.Visible = True
        End Select
        '<EhFooter>
        Exit Sub

comConversionType_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.comConversionType_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub GIS_OnEnter_(translated As Boolean)
        '<EhHeader>
        On Error GoTo GIS_OnEnter__Err
        '</EhHeader>
100     m_bOnTheMap = True
102     translated = True
        '<EhFooter>
        Exit Sub

GIS_OnEnter__Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS_OnEnter_ " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS_OnExit_(translated As Boolean)
        '<EhHeader>
        On Error GoTo GIS_OnExit__Err
        '</EhHeader>
100     m_bOnTheMap = False
102     translated = True
        '<EhFooter>
        Exit Sub

GIS_OnExit__Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS_OnExit_ " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxGISDataGrid_OnEndDragGroupColumn(ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal NewGroupIndex As Long, Allow As Boolean)
        '<EhHeader>
        On Error GoTo dxGISDataGrid_OnEndDragGroupColumn_Err
        '</EhHeader>
100     'DebugPrint "onEndDragGroupColumn"
        '<EhFooter>
        Exit Sub

dxGISDataGrid_OnEndDragGroupColumn_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.dxGISDataGrid_OnEndDragGroupColumn " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxGISDataGrid_OnFilterChanging(FilterText As String)
        '<EhHeader>
        On Error GoTo dxGISDataGrid_OnFilterChanging_Err
        '</EhHeader>
        
        'NOTE: Filter for RS does not works for complex ones with many brackets
        '      for example: ([Province] = 'Kabul') AND (([TYPE] = 'IED Explosion') OR ([TYPE] = 'VBIED Explosion'))
        '      see: http://technet.microsoft.com/en-us/library/ee275540%28BTS.10%29.aspx
        '      where they state: One restriction on these combinations is that OR clauses can only be used at the highest (major) level of the logical operation.
        
        'NOTE: Filter on single quotes is not supported
      ' Commented Out by Petri
      '  FilterText = Replace(FilterText, "[", "")
      '  FilterText = Replace(FilterText, "]", "")

102   '  DoEvents

104     If chkFilterIn.value = vbChecked Then Call chkFilterIn_Click
        
128     If chkSelectIn.value = vbChecked Then Call chkSelectIn_Click
        
        '<EhFooter>
        Exit Sub

dxGISDataGrid_OnFilterChanging_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.dxGISDataGrid_OnFilterChanging " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo Form_KeyDown_Err
        '</EhHeader>
100     If KeyCode = vbKeyEscape And Shift = 1 Then
102         If Not frmLog.Visible Then frmLog.Show vbModeless, Me
        End If
        '<EhFooter>
        Exit Sub

Form_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.Form_KeyDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    On Error Resume Next
    SetErrorMode SEM_NOGPFAULTERRORBOX

End Sub

Public Sub g_clsHotKey_HotKeyPress(ByVal sName As String, _
                                   ByVal eModifiers As EHKModifiers, _
                                   ByVal eKey As KeyCodeConstants)
    Dim f As Form
    Dim i As Integer
    Dim ppic As IPictureDisp
                
    Select Case sName
            
        Case "WMS"
            LoadWMS_WFS
        Case "Debug"
            If Not m_frmDebug.Visible Then m_frmDebug.Show
        Case "Berserk"
                
            bTESTING
            
            '                If Not m_sIncidentIni = "" Then
            '                    If MsgBox("Do You want to Foce INI file Read?", vbYesNo, "Incident Test") = vbYes Then
            '                        CategorizeIncidentsByTypeBerserk True
            '                    Else
            '                        CategorizeIncidentsByTypeBerserk False
            '                    End If
            '                Else
            '                    CategorizeIncidentsByTypeBerserk False
            '                End If

        Case "LyrSearch"

            If m_frmSearch Is Nothing Then
                Set m_frmSearch = New frmSearch
            End If
                
            m_frmSearch.Init GIS10.viewer
                
            m_frmSearch.Show vbModeless, Me

        Case "ForceIni"
            CreateLayerInis

        Case "selector"

        Case "Ticker"
            EmulateTicker

        Case "CompInfo"
            frmComputerInfo.Show vbModal, Me

        Case "INITWORKAROUND" ' DO NOT TAKE THIS AWAY!

            For Each f In Forms
                LoadLanguage f.Name, g_sLanguage, m_Cnn
            Next

        Case "Language"
            frmLanguagePicker.sNewLangugage = g_sLanguage
            frmLanguagePicker.Show vbModal, Me
            
            If frmLanguagePicker.sNewLangugage <> "" Then
                g_sLanguage = frmLanguagePicker.sNewLangugage
                
                For Each f In Forms
                    LoadLanguage f.Name, g_sLanguage, m_Cnn
                Next

            End If
            
        Case "Admin"
            g_sLanguage = InputBox("Type In the Language to Change To:", "OASIS Admin Language Tool", g_sLanguage)

            For Each f In Forms
                LoadLanguage f.Name, g_sLanguage, m_Cnn
            Next

        Case "Dude"

        Case "Script1"

            DOScripts 0

        Case "Script2"

            DOScripts 1

        Case "Script3"

            DOScripts 2

        Case "Script4"

            DOScripts 3
        
        Case "AppState"
            Dim sAppsets As String
                
            sAppsets = "Currrent Connection String: " & m_Cnn.ConnectionString & vbCrLf
            sAppsets = sAppsets & "Client Version: " & App.major & "." & App.minor & "." & App.Revision & " " & App.Comments & vbCrLf
            sAppsets = sAppsets & "Client EXE PATH: " & App.Path & vbCrLf
            sAppsets = sAppsets & "Client DATA PATH: " & g_sAppPath & vbCrLf
            sAppsets = sAppsets & "ADMIN Location: " & m_sAdmLoc & vbCrLf
            sAppsets = sAppsets & "ADMIN 1: " & m_sAdmVal1 & vbCrLf
            sAppsets = sAppsets & "ADMIN 2: " & m_sAdmVal2 & vbCrLf
            sAppsets = sAppsets & "INITIAL MAP NAME: " & m_sInitialMapName & vbCrLf
            sAppsets = sAppsets & "INTERNET CONNECTION: " & m_sInternetCon & vbCrLf

            MsgBox sAppsets

            sAppsets = "SECURITY ANALYSIS FIELD NAME: " & m_sSecAnalysisFieldName & vbCrLf
            sAppsets = sAppsets & "SECURITY ANALYSIS LAYER NAME: " & m_sSecAnalysisLyrName & vbCrLf
            sAppsets = sAppsets & "DEMO MODE: " & g_bDemoLogin & vbCrLf
            sAppsets = sAppsets & "USER ID: " & g_CurrentUserID & vbCrLf
            sAppsets = sAppsets & "APP SERVER PATH: " & g_sAppServerPath & vbCrLf
            sAppsets = sAppsets & "APPSETTING TABLE: " & g_sAppSettingsTable & vbCrLf
            sAppsets = sAppsets & "LANGUAGE: " & g_sLanguage & vbCrLf
            sAppsets = sAppsets & "SERVER TABLE PREFIX: " & g_sRemoteTablePrefix & vbCrLf
            sAppsets = sAppsets & "USER: " & g_sUserName & vbCrLf
            sAppsets = sAppsets & "GIS Project: " & GIS10.ProjectName & vbCrLf
            sAppsets = sAppsets & "SERVER: " & g_sAppServerPath & vbCrLf
            MsgBox sAppsets

        Case "Folders"
            frmFolders.Show vbModal, Me

        Case "GPS"
            frmGPS.Show vbModeless, Me
        Case "OVB"
                
            Dim sProjpath As String
            Dim c As New cCommonDialog
                
            If oVB Is Nothing Then
                Set oVB = CreateObject("OVBScript.ScriptEngine")
    
                If Not oVB Is Nothing Then
    
                    On Error Resume Next
                               
                    If oVB.ScriptProjectPath = "" Then

                        With c
                            .DialogTitle = "Open OASIS VBScript Project File"
                            .CancelError = True
                            .hWnd = Me.hWnd
                            .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
                            .InitDir = g_sAppPath & "\data\user\"
                            .Filter = "OASIS VBScript Project File (*.ovbp)"
                            .FilterIndex = 1
                            .ShowOpen
                            sProjpath = .Filename
                        End With

                    End If
                                
                    With oVB
                            
                        If Not .OpenProject(sProjpath, 0) Then
                            MsgBox "Failed To Open"
                            Set oVB = Nothing
                            Exit Sub
                        End If
                            
                        If .ScriptProjectPath = "" Then
                            Exit Sub
                        End If

                        .ParentForm = Me
                        .Verbose = True
                        .AddScriptingObject "OASISGis", GIS10
                        .AddScriptingObject "OASISGisUtils", GisUtils
                        .AddScriptingObject "RSAppSettings", g_RSAppSettings
                        .AddScriptingObject "RSGISGridTableSettings", g_RSGISGridTableSettings
                        .AddScriptingObject "CN", m_Cnn
                        .AddScriptingObject "OASISGISDataGrid", dxGISDataGrid
                        .AddScriptingObject "OASISToolBar", AB
                        .AddScriptingObject "OASISCharting", frmSecChart

                        .ShowEditor
                            
                    End With

                    Set oVB = Nothing

                End If

            Else
                oVB.ShowEditor
            End If
        
        Case "doPrint"
            'Dim oPrint As New OASISPrint.cMonitor
            'oPrint.ShowPrint "", 0, 0, 0, 0
             'frmMapPrint.Init GIS10.viewer, m_frmMnuOperations.Legend1.Legend, g_bMapTplUpdate, g_sRemoteTablePrefix
             'frmMapPrint.Show vbModeless, Me
        Case "OPSV"
            mnuOPSView_Click
                
        Case "Tracker"
    
            Dim sServer As String
            
            AB.Bands.Item("bMenu").Visible = False
            AB.Bands.Item("sdToolBar").Visible = False
            AB.Bands.Item("tbExtents").Visible = False
            AB.Bands.Item("tbLayer").Visible = False
            AB.Bands.Item("tbUtils").Visible = False
    
            elRSSTool.Visible = False
            elMap.Visible = False
            elOasisProfile.Visible = False
            elAddons.Visible = False
            elDynamicData.Visible = False
            elDynamicReports.Visible = False

           ' AB.Bands.Item("cbSync").Visible = True
            
            sServer = InputBox("ENTER The Server Name:", "OASIS Tracking Utils")
            
            WebBrowser1.Navigate2 sServer
                
            AB.ClientAreaControl = elOasisProfile
            elOasisProfile.Visible = True
        
            AB.RecalcLayout
        End Select
            
End Sub

Private Sub DOScripts(iID As Integer)
        '<EhHeader>
        On Error GoTo DOScripts_Err
        '</EhHeader>
        Dim ofs As FileSystemObject
        Dim oTXT As TextStream
        Dim CurSuffix As String
        Dim sScripts() As String

100     Set ofs = New FileSystemObject
        
102     SafeMoveFirst g_RSAppSettings
104     g_RSAppSettings.Find "SettingName = 'Scripts'"
        
106     If Not g_RSAppSettings.EOF Then
108         If Not g_RSAppSettings.Fields.Item("SettingValue1").value = vbNull Then
110             If Not g_RSAppSettings.Fields.Item("SettingValue1").value = "" Then
                    
112                 sScripts = Split(g_RSAppSettings.Fields.Item("SettingValue1").value, ",")
                                
114                 If UBound(sScripts) < iID Then Exit Sub
                                
116                 With SSC
118                     .AllowUI = True
                                
120                     If ofs.FileExists(g_sAppPath & "\data\user\Sessions\" & sScripts(iID) & ".ovb") Then
122                         .Language = "VBScript"
124                         CurSuffix = ".ovb"
126                     ElseIf ofs.FileExists(g_sAppPath & "\data\user\Sessions\" & sScripts(iID) & ".ojs") Then
128                         .Language = "JScript"
130                         CurSuffix = ".ojs"
                        Else
                            Exit Sub
                        End If
                                
132                     Set oTXT = ofs.OpenTextFile(g_sAppPath & "\data\user\Sessions\" & sScripts(iID) & CurSuffix, ForReading)
134                     .AddCode oTXT.ReadAll
                                
136                     oTXT.Close
138                     Set oTXT = Nothing
140                     Set ofs = Nothing
                                    
                    End With

                End If
            End If
        End If

        '<EhFooter>
        Exit Sub

DOScripts_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.DOScripts " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS_OnEditorChange(translated As Boolean)
  GIS10.UpDate
End Sub

Private Sub GIS_OnKeyUp(translated As Boolean, Key As Integer, ByVal Shift As TatukGIS_XDK10.XShiftState)
        '<EhHeader>
        On Error GoTo GIS_OnKeyUp_Err
        '</EhHeader>
100   If Key = VK_CONTROL Then
102     vkControl = False
      End If
        '<EhFooter>
        Exit Sub

GIS_OnKeyUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS_OnKeyUp " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_fmrAddIncident_ConvPTtoMGRS(x As Double, y As Double)
    m_fmrAddIncident.txtMGRS.Text = ConvPTtoMGRS(x, y)
End Sub

Private Sub m_fmrAddIncident_ConvMGRStoPT(sMGRS As String, x As Double, y As Double)
    
    If ConvMGRStoPT(sMGRS, x, y) Then
        m_fmrAddIncident.txtX.Text = x
        m_fmrAddIncident.txtY.Text = y
    Else
            m_fmrAddIncident.txtX.Text = ""
        m_fmrAddIncident.txtY.Text = ""
    End If
    
End Sub

Private Sub m_fmrAddIncident_SubmitIncident(sGUID As String) '(RSVictims As ADODB.Recordset)
        'g_CurrentTool = m_prevTool
        'SetOASISTool
        '<EhHeader>
        On Error GoTo m_fmrAddIncident_SubmitIncident_Err
        '</EhHeader>
        Dim dID As Long
        Dim varAttachments As Variant
        Dim i As Integer
        Dim RSIncidentForOASISReports As ADODB.Recordset
        Dim iFieldCount As Integer
        Dim m_frmReportsFromRS As frmReportsFromRS ' Dim i As Integer
       
        Dim sPDFPath As String
        Dim sEMHTML As String
        Dim RSUpdater As ADODB.Recordset
        Dim ptg As TatukGIS_XDK10.XGIS_Point
        
'100     sGUID = GetGuid
        
102     g_CurrentTool = oZoom
104     GIS10.Mode = TatukGIS_XDK10.XgisZoomEx
        'idiot
            
106     If m_oIncShpPt Is Nothing Then
108         Set m_oIncShpPt = m_oSQLIncLyr.CreateShape(XgisShapeTypePoint)
            
            Set ptg = New TatukGIS_XDK10.XGIS_Point
            
            With m_fmrAddIncident
            
               ptg.Prepare CDbl(.txtX.Text), CDbl(.txtY.Text)
            
            End With
            
            
110         m_oIncShpPt.Lock TatukGIS_XDK10.XgisLockExtent
112         m_oIncShpPt.AddPart

114         m_oIncShpPt.AddPoint ptg

        End If
            
136     If Not m_oIncShpPt Is Nothing Then
138         m_oIncShpPt.Lock TatukGIS_XDK10.XgisLockExtent
140         Randomize
                
142         dID = CLng(1000 * Rnd(30000))
                
144         m_oIncShpPt.SetField "ID", sGUID
              
146         If Not IsEmpty(m_oIncShpPt.GetField("Name")) Then
148             m_oIncShpPt.SetField "Name", m_fmrAddIncident.txtEnteredBy.Text
            End If
              
150         If Not IsEmpty(m_oIncShpPt.GetField("Type")) Then
152             m_oIncShpPt.SetField "Type", m_fmrAddIncident.ComIncType.List(m_fmrAddIncident.ComIncType.ListIndex)
            End If
              
154         If Not IsEmpty(m_oIncShpPt.GetField("Target")) Then
156             m_oIncShpPt.SetField "Target", m_fmrAddIncident.ComIncTarget.List(m_fmrAddIncident.ComIncTarget.ListIndex)
            End If

158         If Not IsEmpty(m_oIncShpPt.GetField("Incident_DATE")) Then
160             m_oIncShpPt.SetField "Incident_DATE", m_fmrAddIncident.MVIncident.value
            End If
            
162         If Not IsEmpty(m_oIncShpPt.GetField("Incident_DATESERIAL")) Then
164             m_oIncShpPt.SetField "Incident_DATESERIAL", ConvertDateToSerial(m_fmrAddIncident.MVIncident.value)
            End If
              
166         If Not IsEmpty(m_oIncShpPt.GetField("Town")) Then
168             m_oIncShpPt.SetField "Town", Trim(Replace(m_fmrAddIncident.lblNearestTown.caption, "Nearest Town:", ""))
            End If
              
170         If Not IsEmpty(m_oIncShpPt.GetField("District")) Then
172             m_oIncShpPt.SetField "District", Trim(Replace(m_fmrAddIncident.lblDistrict_.caption, "District:", ""))
            End If
              
174         If Not IsEmpty(m_oIncShpPt.GetField("Province")) Then
176             m_oIncShpPt.SetField "Province", Trim(Replace(m_fmrAddIncident.lblProvince_.caption, "Province:", ""))
            End If
              
178         If Not IsEmpty(m_oIncShpPt.GetField("Description")) Then
180             m_oIncShpPt.SetField "Description", m_fmrAddIncident.txtIncidentDescription.Text
            End If
              
182         If Not IsEmpty(m_oIncShpPt.GetField("LocDesc")) Then
184             m_oIncShpPt.SetField "LocDesc", m_fmrAddIncident.txtLocationDescription.Text
            End If
              
186         If Not IsEmpty(m_oIncShpPt.GetField("Source")) Then
188             m_oIncShpPt.SetField "Source", m_fmrAddIncident.ComSource.Text
            End If
              
190         If Not IsEmpty(m_oIncShpPt.GetField("UnKnownAffected")) Then
192             m_oIncShpPt.SetField "UnKnownAffected", "" '0,1
            End If
            
194         If m_fmrAddIncident.OptCasualties(0).value Then
                        
196             If m_fmrAddIncident.chkUnknown.value = vbUnchecked Then
198                 If Not IsEmpty(m_oIncShpPt.GetField("Injured")) Then
200                     m_oIncShpPt.SetField "Injured", m_fmrAddIncident.txtCasualtiesInjured.Text
                    End If
                      
202                 If Not IsEmpty(m_oIncShpPt.GetField("Affected")) Then
204                     m_oIncShpPt.SetField "Affected", m_fmrAddIncident.txtCasualtiesAffected.Text
                    End If
                      
206                 If Not IsEmpty(m_oIncShpPt.GetField("Dead")) Then
208                     m_oIncShpPt.SetField "Dead", m_fmrAddIncident.txtCasualtiesDead.Text
                    End If
                End If

210         ElseIf m_fmrAddIncident.OptCasualties(0).value Then

212             If Not IsEmpty(m_oIncShpPt.GetField("Injured")) Then
214                 m_oIncShpPt.SetField "Injured", 0
                End If
                      
216             If Not IsEmpty(m_oIncShpPt.GetField("Affected")) Then
218                 m_oIncShpPt.SetField "Affected", m_fmrAddIncident.txtCasualtiesAffected.Text
                End If
                      
220             If Not IsEmpty(m_oIncShpPt.GetField("Dead")) Then
222                 m_oIncShpPt.SetField "Dead", 0
                End If
            End If
            
224         If Not IsEmpty(m_oIncShpPt.GetField("Violent")) Then
226             If m_fmrAddIncident.OptIncViolence(0).value Then
228                 m_oIncShpPt.SetField "Violent", 0 '0,1,2
230             ElseIf m_fmrAddIncident.OptIncViolence(1).value = True Then
232                 m_oIncShpPt.SetField "Violent", 1
                Else
234                 m_oIncShpPt.SetField "Violent", 2
                End If
            End If
              
236         If Not IsEmpty(m_oIncShpPt.GetField("Casualties")) Then
238             m_oIncShpPt.SetField "Casualties", "" '0,1,2
            End If
              
240         If Not IsEmpty(m_oIncShpPt.GetField("TimeStamp")) Then
242             m_oIncShpPt.SetField "TimeStamp", Now()
            End If
            
244         If Not IsEmpty(m_oIncShpPt.GetField("GUID")) Then
246             m_oIncShpPt.SetField "GUID", sGUID
            End If
            
248         If Not IsEmpty(m_oIncShpPt.GetField("ReportID")) Then
250             m_oIncShpPt.SetField "ReportID", 10
            End If
              
252         If Not IsEmpty(m_oIncShpPt.GetField("TIME00")) Then
254             If IsNumeric(m_fmrAddIncident.txtHour.Text) Then

256                 Select Case CInt(m_fmrAddIncident.txtHour.Text)

                        Case 0 To 5
258                         m_oIncShpPt.SetField "TIME00", "Night"

260                     Case 6 To 9
262                         m_oIncShpPt.SetField "TIME00", "Morning"

264                     Case 10 To 13
266                         m_oIncShpPt.SetField "TIME00", "Noon"

268                     Case 14 To 17
270                         m_oIncShpPt.SetField "TIME00", "Afternoon"

272                     Case Else
274                         m_oIncShpPt.SetField "TIME00", "Evening"
                    End Select

                Else
276                 m_oIncShpPt.SetField "TIME00", "N/A"
                End If
            End If
          
278         m_fmrAddIncident.SetFocus
          
280         m_oIncShpPt.Unlock
        
282         GIS10.Center = m_oIncShpPt.Centroid
284         m_oSQLIncLyr.SaveData

286         SetNewSynchDBElement GetGuid, sGUID, "OASIS Incidents", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, "oincidents", True
            
288         If m_fmrAddIncident.chkOASISReports = vbChecked Then

290             Set RSIncidentForOASISReports = New ADODB.Recordset

292             iFieldCount = m_oIncShpPt.layer.Fields.Count
294             i = 0

296             Do Until i = iFieldCount
            
298                 If Not m_oIncShpPt.layer.Fields.Item(i).Name = "GUID" And Not m_oIncShpPt.layer.Fields.Item(i).Name = "ID" Then
        
300                     Select Case m_oIncShpPt.layer.Fields.Item(i).FieldType
                
                            Case TatukGIS_XDK10.XgisFieldTypeBoolean
302                             RSIncidentForOASISReports.Fields.Append m_oIncShpPt.layer.Fields.Item(i).Name, adBoolean
                
304                         Case TatukGIS_XDK10.XgisFieldTypeDate
306                             RSIncidentForOASISReports.Fields.Append m_oIncShpPt.layer.Fields.Item(i).Name, adDate

308                         Case TatukGIS_XDK10.XgisFieldTypeFloat
310                             RSIncidentForOASISReports.Fields.Append m_oIncShpPt.layer.Fields.Item(i).Name, adDouble

312                         Case TatukGIS_XDK10.XgisFieldTypeNumber
314                             RSIncidentForOASISReports.Fields.Append m_oIncShpPt.layer.Fields.Item(i).Name, adBigInt

316                         Case TatukGIS_XDK10.XgisFieldTypeString
318                             RSIncidentForOASISReports.Fields.Append m_oIncShpPt.layer.Fields.Item(i).Name, adVarChar, 255
                
320                         Case Else
322                             RSIncidentForOASISReports.Fields.Append m_oIncShpPt.layer.Fields.Item(i).Name, adVarChar, 255
                
                        End Select

                    End If
                
324                 i = i + 1
                Loop

326             frmGridRptProprties.Init RSIncidentForOASISReports
328             frmGridRptProprties.txtTitle.Text = "OASIS Incidents"
330             frmGridRptProprties.Show vbModal, Me
332             RSIncidentForOASISReports.Open
334             RSIncidentForOASISReports.AddNew
            
336             iFieldCount = m_oIncShpPt.layer.Fields.Count
338             i = 0

340             Do Until i = RSIncidentForOASISReports.Fields.Count
                
342                 If RSIncidentForOASISReports.Fields(i).Name = "UID" Then
344                     RSIncidentForOASISReports.Fields("UID").value = m_oIncShpPt.uID
                    Else
346                     RSIncidentForOASISReports.Fields(i).value = m_oIncShpPt.GetField(RSIncidentForOASISReports.Fields(i).Name)
                    End If

348                 i = i + 1
                Loop
            
350             Set m_frmReportsFromRS = New frmReportsFromRS
                        
352             If frmGridRptProprties.chkIncludeMap = vbChecked Then
354                 Clipboard.Clear
356                 GIS10.viewer.PrintClipboard

358                 m_frmReportsFromRS.SetReportRS frmGridRptProprties.txtTitle.Text, RSIncidentForOASISReports, "", Clipboard.GetData(vbCFEMetafile), frmGridRptProprties.txtMapTitle.Text
                Else
360                 m_frmReportsFromRS.SetReportRS frmGridRptProprties.txtTitle.Text, RSIncidentForOASISReports, ""
                End If

362             Unload frmGridRptProprties
364             m_frmReportsFromRS.ShowReport
366             m_frmReportsFromRS.Show vbModal, Me
368             sPDFPath = m_frmReportsFromRS.PDFPath
                
370             Unload m_frmReportsFromRS
372             Set m_frmReportsFromRS = Nothing
374             Set RSIncidentForOASISReports = Nothing
            End If

376         Set m_oIncShpPt = Nothing
        End If
    
378     varAttachments = m_fmrAddIncident.IncAttachments
            
380     SafeMoveFirst g_RSAppSettings
382     g_RSAppSettings.Find "SettingName = 'attachmentsURL'"
            
        '        i = UBound(m_cUploadThreads)
        '
        '        Do While UBound(m_cUploadThreads) >= 0
        '            Set m_cUploadThreads(i) = Nothing
        '            If Not UBound(m_cUploadThreads) = 0 Then
        '                ReDim Preserve m_cUploadThreads(UBound(m_cUploadThreads) - 1)
        '            Else
        '                Exit Do
        '            End If
        '            i = UBound(m_cUploadThreads)
        '        Loop
        '
        
        '        SHCreateThread AddressOf WriteAttachmentHTML, ByVal 0&, CTF_INSIST, ByVal 0&
            
384     If Not UBound(varAttachments) = LBound(varAttachments) Then
            
386         For i = 0 To UBound(varAttachments)
                'm_Cnn.Execute "INSERT INTO Attachments (incidentID, FilePath, DateInserted, AttachmentTable) VALUES (" & dID & ", '" & varAttachments(i) & "', '" & Now() & "', 'oincidents' )"
                                
388             Set RSUpdater = New ADODB.Recordset

390             With RSUpdater

392                 .Open "SELECT * FROM Attachments", m_Cnn, adOpenDynamic, adLockBatchOptimistic
394                 .AddNew
396                 .Fields("incidentID").value = dID
398                 .Fields("FilePath").value = varAttachments(i)
400                 .Fields("DateInserted").value = Now()
402                 .Fields("AttachmentTable").value = "oincidents"
404                 .UpdateBatch adAffectCurrent
406                 .Close

                End With

408             Set RSUpdater = Nothing
                
                '               SHCreateThread AddressOf WriteAttachmentHTML, ByVal 0&, CTF_INSIST, ByVal 0&
                
                '                If Not UBound(m_cUploadThreads) = 0 Then
                '                    ReDim Preserve m_cUploadThreads(UBound(m_cUploadThreads) + 1)
                '                End If
                '
                '                Set m_cUploadThreads(UBound(m_cUploadThreads)) = New clsThreads
                '
                '                m_cUploadThreads(UBound(m_cUploadThreads)).Initialize AddressOf WriteAttachmentHTML
                
410             UploadAttachments g_RSAppSettings.Fields.Item("SettingValue1").value, CStr(varAttachments(i))
            Next
        
        Else

412         If Not varAttachments(0) = "" Then
                
414             Set RSUpdater = New ADODB.Recordset

416             With RSUpdater
            
418                 .Open "SELECT * FROM Attachments", m_Cnn, adOpenDynamic, adLockBatchOptimistic
420                 .AddNew
422                 .Fields("incidentID").value = dID
424                 .Fields("FilePath").value = varAttachments(0)
426                 .Fields("DateInserted").value = Now()
428                 .Fields("AttachmentTable").value = "oincidents"
430                 .UpdateBatch adAffectCurrent
432                 .Close

                End With

434             Set RSUpdater = Nothing

436             UploadAttachments g_RSAppSettings.Fields.Item("SettingValue1").value, CStr(varAttachments(0))
            End If
        End If
        
438     If m_fmrAddIncident.chkHTMLReport = vbChecked Then
        
            On Error Resume Next
        
440         Kill g_sAppPath & "\data\user\Exports\" & "exp.jpg"
442         Kill g_sAppPath & "\data\user\Exports\" & "exp.JGW"
444         Kill g_sAppPath & "\data\user\Exports\" & "exp.JFW"
446         Kill g_sAppPath & "\data\user\Exports\" & "exp.tab"
448         Kill g_sAppPath & "\data\user\Exports\" & "OasisReport.html"
        
450         GIS10.UpDate
            
452         GIS10.viewer.ExportToImage g_sAppPath & "\data\user\Exports\" & "exp.jpg", GIS10.VisibleExtent, ScaleX(GIS10.Width, vbTwips, vbPixels), ScaleY(GIS10.Height, vbTwips, vbPixels), 100, 0, 96

454         sEMHTML = WriteHTMLReport("exp.jpg")

        End If

456     If m_fmrAddIncident.chkCreateEmail = vbChecked Then
458         SendOutLookMail "", "", "", "OASIS Incident Report", "Auto Generated by OASIS", sPDFPath, "", True, sEMHTML
        End If

460     m_fmrAddIncident.Hide
        
        '<EhFooter>
        Exit Sub

m_fmrAddIncident_SubmitIncident_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_fmrAddIncident_SubmitIncident " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetNewSynchDBElement(sNewGUID As String, _
                                sID As String, _
                                sTitle As String, _
                                sDescription As String, _
                                sBy As String, _
                                sRFC3339DateTime As String, _
                                sTableName As String, _
                                bIsGeoLayer As Boolean, _
                                Optional supdates As String = "false")
        '<EhHeader>
        On Error GoTo SetNewSynchDBElement_Err
        '</EhHeader>
    
        Dim RSUpdater As ADODB.Recordset
100     Set RSUpdater = New ADODB.Recordset
102     With RSUpdater
            
104         .Open "SELECT * FROM SynchHistory", m_Cnn, adOpenDynamic, adLockBatchOptimistic

108         If Not .EOF Then
                .AddNew
110             .Fields("sID").value = sID
112             .Fields("sGUID").value = sNewGUID
114             .Fields("sTableName").value = sTableName
116             .Fields("swhen").value = sRFC3339DateTime
118             .Fields("sStatus").value = "pending"
120             .Fields("sequence").value = 1
122             .Fields("sBy").value = sBy
124             .Fields("sdelete").value = "false"
126             .Fields("updates").value = supdates
128             .Fields("noconflict").value = "local"
130             .UpdateBatch adAffectCurrent
132             .Close
            End If

        End With
134     Set RSUpdater = Nothing

        '<EhFooter>
        Exit Sub

SetNewSynchDBElement_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SetNewSynchDBElement " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub InsertNewRSSSyncFeedItem(sNewGUID As String, _
                                     sTitle As String, _
                                     sDescription As String, _
                                     sBy As String, _
                                     sRFC3339DateTime As String, _
                                     sTableName As String, _
                                     bIsGeoLayer As Boolean, _
                                     oChannel As MSXML2.IXMLDOMElement, _
                                     Optional bCheckHistory As Boolean = False, _
                                     Optional CN As ADODB.Connection)
    
    Dim oXMLDoc As New MSXML2.DOMDocument
    Dim oXMLElement As MSXML2.IXMLDOMElement
    Dim oRssItem As MSXML2.IXMLDOMElement
    Dim oSynchElement As MSXML2.IXMLDOMElement
    Dim oHisElements As MSXML2.IXMLDOMElement
    Dim oTitleElements As MSXML2.IXMLDOMElement
    Dim oDescElements As MSXML2.IXMLDOMElement
    Dim Obj As Object
        
    'Create "item" element
    Set oRssItem = oXMLDoc.createElement("item")
    
    'Create "sx:sync" element
    Set oSynchElement = oXMLDoc.createElement("sx:sync")
    
    'Set "id" attribute for "sx:sync" element
    oSynchElement.setAttribute "id", sNewGUID
    
    'Set "updates" attribute for "sx:sync" element
    oSynchElement.setAttribute "updates", "1"
    
    'Set "deleted" attribute for "sx:sync" element
    oSynchElement.setAttribute "deleted", "false"
    
    'Set "noconflicts" attribute for "sx:sync" element
    oSynchElement.setAttribute "noconflicts", "true"
    
    If bCheckHistory Then
        'bCheckHistory = AddHistoryElements(oSynchElement, sNewGUID, cn)
    End If
    
    If Not bCheckHistory Then
    
        'Create "history" element
        Set oHisElements = oXMLDoc.createElement("sx:history")
    
        'Get the current timedate and format it as RFC 3339
        'sRFC3339DateTime = RFC3339DateTime
    
        'Set "sequence" attribute for "sx:history" element
        oHisElements.setAttribute "sequence", "1"
        
        'Set "when" attribute for "sx:history" element
        oHisElements.setAttribute "when", sRFC3339DateTime
        
        'Set "by" attribute for "sx:history" element
        oHisElements.setAttribute "by", sBy
        
        'Append "sx:history" element to "sx:sync" element
        oSynchElement.appendChild oHisElements
        
    End If
        
    'Append "sx:sync" element to "item" element
    oRssItem.appendChild oSynchElement
        
    'Create & populate "title" element
    Set oTitleElements = oXMLDoc.createElement("title")
    oTitleElements.Text = sTitle
        
    'Append "title" element to "item" element
    oRssItem.appendChild oTitleElements
        
    'Create & populate "description" element
    Set oDescElements = oXMLDoc.createElement("description")
    oDescElements.Text = sDescription
        
    'Append "description" element to "item" element
    oRssItem.appendChild oDescElements
        
    'Create & populate "Table" element, Reuse the Description Element
    Set oDescElements = oXMLDoc.createElement("TableName")
    oDescElements.Text = sTableName
    oDescElements.setAttribute "isGeoTable", LCase$(CStr(bIsGeoLayer))
    'Append "tableName" element to "item" element
    oRssItem.appendChild oDescElements
        
    'Append "item" element to "channel" element
    oChannel.appendChild oRssItem

End Sub

Private Function CreateNewRSSSyncFeed(sSourcePath As String, _
                                      sNewGUID As String, _
                                      sTitle As String, _
                                      sDescription As String, _
                                      sBy As String, _
                                      Optional bSave As Boolean = False) As String
        '<EhHeader>
        On Error GoTo CreateNewRSSSyncFeed_Err
        '</EhHeader>

        Dim oXMLDoc As New MSXML2.DOMDocument
        Dim oXMLElement As MSXML2.IXMLDOMElement
        Dim oChannel As MSXML2.IXMLDOMElement
        Dim oRssItem As MSXML2.IXMLDOMElement
        Dim oSynchElement As MSXML2.IXMLDOMElement
        Dim oHisElements As MSXML2.IXMLDOMElement
        Dim oTitleElements As MSXML2.IXMLDOMElement
        Dim oDescElements As MSXML2.IXMLDOMElement
        Dim sRFC3339DateTime As String
        Dim Obj As Object
        
100     oXMLDoc.async = False

102     Set oChannel = oXMLDoc.createElement("channel")
    
        ' Validate that "channel" element exists
104     If oChannel Is Nothing Then
106         MsgBox "Unable to get channel"
            Exit Function
        End If
    
        ' Create "item" element
108     Set oRssItem = oXMLDoc.createElement("item")
    
        ' Create "sx:sync" element
110     Set oSynchElement = oXMLDoc.createElement("sx:sync")
    
        ' Set "id" attribute for "sx:sync" element
112     oSynchElement.setAttribute "id", sNewGUID
    
        ' Set "updates" attribute for "sx:sync" element
114     oSynchElement.setAttribute "updates", "1"
    
        ' Set "deleted" attribute for "sx:sync" element
116     oSynchElement.setAttribute "deleted", "false"
    
        ' Set "noconflicts" attribute for "sx:sync" element
118     oSynchElement.setAttribute "noconflicts", "false"
    
        ' Create "history" element
120     Set oHisElements = oXMLDoc.createElement("sx:history")
    
        ' Get the current timedate and format it as RFC 3339
122     sRFC3339DateTime = RFC3339DateTime
    
        ' Set "sequence" attribute for "sx:history" element
124     oHisElements.setAttribute "sequence", "1"
        
        ' Set "when" attribute for "sx:history" element
126     oHisElements.setAttribute "when", sRFC3339DateTime
        
        ' Set "by" attribute for "sx:history" element
128     oHisElements.setAttribute "by", sBy
        
        ' Append "sx:history" element to "sx:sync" element
130     oSynchElement.appendChild oHisElements
        
        ' Append "sx:sync" element to "item" element
132     oRssItem.appendChild oSynchElement
        
        ' Create & populate "title" element
134     Set oTitleElements = oXMLDoc.createElement("title")
136     oTitleElements.Text = sTitle
        
        ' Append "title" element to "item" element
138     oRssItem.appendChild oTitleElements
        
        ' Create & populate "description" element
140     Set oDescElements = oXMLDoc.createElement("description")
142     oDescElements.Text = sDescription
        
        ' Append "description" element to "item" element
144     oRssItem.appendChild oDescElements
        
        ' Append "item" element to "channel" element
146     oChannel.appendChild oRssItem
        
148     oXMLDoc.loadXML "<rss version=""""2.0"""" xmlns:sx=""""http://feedsync.org/2007/feedsync"""">" & oChannel.xml & "</rss>"
        
150     If bSave Then oXMLDoc.Save sSourcePath
        'oChannel.Normalize
152     CreateNewRSSSyncFeed = "<rss version=""""2.0"""" xmlns:sx=""""http://feedsync.org/2007/feedsync"""">" & oChannel.xml & "</rss>"
        'oXMLDoc.Xml

        '<EhFooter>
        Exit Function

CreateNewRSSSyncFeed_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CreateNewRSSSyncFeed " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function GetNewRSSSyncFeed(sSourcePath As String, _
                                   sNewGUID As String, _
                                   sTitle As String, _
                                   sDescription As String, _
                                   sBy As String, _
                                   sRFC3339DateTime As String, _
                                   sTableName As String, _
                                   bIsGeoLayer As Boolean) As MSXML2.IXMLDOMElement
        '<EhHeader>
        On Error GoTo GetNewRSSSyncFeed_Err
        '</EhHeader>

        Dim oXMLElement As MSXML2.IXMLDOMElement
        Dim oChannel As MSXML2.IXMLDOMElement
        Dim oRssItem As MSXML2.IXMLDOMElement
        Dim oSynchElement As MSXML2.IXMLDOMElement
        Dim oHisElements As MSXML2.IXMLDOMElement
        Dim oTitleElements As MSXML2.IXMLDOMElement
        Dim oDescElements As MSXML2.IXMLDOMElement
        Dim Obj As Object
        Dim oXMLDoc As New MSXML2.DOMDocument

100     Set oChannel = oXMLDoc.createElement("channel")
        
        ' Create "item" element
102     Set oRssItem = oXMLDoc.createElement("item")
    
        ' Create "sx:sync" element
104     Set oSynchElement = oXMLDoc.createElement("sx:sync")
    
        ' Set "id" attribute for "sx:sync" element
106     oSynchElement.setAttribute "id", sNewGUID
    
        ' Set "updates" attribute for "sx:sync" element
108     oSynchElement.setAttribute "updates", "1"
    
        ' Set "deleted" attribute for "sx:sync" element
110     oSynchElement.setAttribute "deleted", "false"
    
        ' Set "noconflicts" attribute for "sx:sync" element
112     oSynchElement.setAttribute "noconflicts", "false"
    
        ' Create "history" element
114     Set oHisElements = oXMLDoc.createElement("sx:history")
    
        ' Set "sequence" attribute for "sx:history" element
116     oHisElements.setAttribute "sequence", "1"
        
        ' Set "when" attribute for "sx:history" element
118     oHisElements.setAttribute "when", sRFC3339DateTime
        
        ' Set "by" attribute for "sx:history" element
120     oHisElements.setAttribute "by", sBy
        
        ' Append "sx:history" element to "sx:sync" element
122     oSynchElement.appendChild oHisElements
        
        ' Append "sx:sync" element to "item" element
124     oRssItem.appendChild oSynchElement
        
        ' Create & populate "title" element
126     Set oTitleElements = oXMLDoc.createElement("title")
128     oTitleElements.Text = sTitle
        
        ' Append "title" element to "item" element
130     oRssItem.appendChild oTitleElements
        
        ' Create & populate "description" element
132     Set oDescElements = oXMLDoc.createElement("description")
134     oDescElements.Text = sDescription
        
        ' Append "description" element to "item" element
136     oRssItem.appendChild oDescElements
        
        ' Create & populate "Table" element, Reuse the Description Element
138     Set oDescElements = oXMLDoc.createElement("TableName")
140     oDescElements.Text = sTableName
        
142     oDescElements.setAttribute "isGeoTable", LCase$(CStr(bIsGeoLayer))
        
        ' Append "tableName" element to "item" element
144     oRssItem.appendChild oDescElements
                
        ' Append "item" element to "channel" element
146     oChannel.appendChild oRssItem
        
148     Set GetNewRSSSyncFeed = oChannel

        '<EhFooter>
        Exit Function

GetNewRSSSyncFeed_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GetNewRSSSyncFeed " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function GetSynchSQL(sSynchTable As String, _
                             sJoinSynchTable, _
                             iYear As Integer, _
                             iMonth As Integer, _
                             iDay As Integer, _
                             iHour As Integer, _
                             iMin As Integer, _
                             iSec As Integer) As String
        '<EhHeader>
        On Error GoTo GetSynchSQL_Err
        '</EhHeader>
        Dim sRFC3339DateTime As String
    
100     sRFC3339DateTime = RFC3339DateTimeEX(iYear, CStr(iMonth), CStr(iDay), CStr(iHour), CStr(iMin), CStr(iSec))
102     GetSynchSQL = "select * From " & sSynchTable & ", " & sJoinSynchTable & " WHERE [" & sSynchTable & "].sGUID = [" & sJoinSynchTable & "].sGUID" & IIf(sRFC3339DateTime <> "", " AND [" & sJoinSynchTable & "].swhen = '" & sRFC3339DateTime & "'", "")
        
        '<EhFooter>
        Exit Function

GetSynchSQL_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GetSynchSQL " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub m_frmChangeTracer_ChangeScope(sScope As String, sLayer As String)
        '<EhHeader>
        On Error GoTo m_frmChangeTracer_ChangeScope_Err
        '</EhHeader>
    Dim oVLayer As TatukGIS_XDK10.XGIS_LayerVector
100         DebugPrint sScope
102     Set oVLayer = GIS10.get(sLayer)
    
104     If Not oVLayer Is Nothing Then
106         oVLayer.scope = sScope
        End If
    
108     GIS10.UpDate
        '<EhFooter>
        Exit Sub

m_frmChangeTracer_ChangeScope_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmChangeTracer_ChangeScope " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmChangeTracer_GetFields(sLayer As String)
        '<EhHeader>
        On Error GoTo m_frmChangeTracer_GetFields_Err
        '</EhHeader>
    Dim oVLayer As TatukGIS_XDK10.XGIS_LayerVector
    Dim i As Integer

100     Set oVLayer = GIS10.get(sLayer)

102     m_frmChangeTracer.ComAnalysisField.Clear
        
        Set m_ColAnlFieldType = New Collection
        
        m_frmChangeTracer.ComAnalysisField.AddItem "--None--"
        
104     If Not oVLayer Is Nothing Then
106         For i = 0 To oVLayer.Fields.Count - 1
108            m_frmChangeTracer.ComAnalysisField.AddItem oVLayer.Fields.Item(i).Name
               m_frmChangeTracer.ComAnalysisField.ItemData(m_frmChangeTracer.ComAnalysisField.ListCount - 1) = oVLayer.FieldInfo(i).FieldType
            Next
        End If

110     If m_frmChangeTracer.ComAnalysisField.ListCount > 0 Then m_frmChangeTracer.ComAnalysisField.ListIndex = 0

        '<EhFooter>
        Exit Sub

m_frmChangeTracer_GetFields_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmChangeTracer_GetFields " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmChangeTracer_GetUniqueValues(sLayer As String, sField As String)
        '<EhHeader>
        On Error GoTo m_frmChangeTracer_GetUniqueValues_Err
        '</EhHeader>
    Dim oVLayer As TatukGIS_XDK10.XGIS_LayerVector
    Dim i As Integer

100     Set oVLayer = GIS10.get(sLayer)
        'oVLayer.Scope = "DISTINCT " & sField
    
104     m_frmChangeTracer.lvUniqueValues.ColumnHeaders.Clear
106     m_frmChangeTracer.lvUniqueValues.ListItems.Clear
    
108     m_frmChangeTracer.lvUniqueValues.ColumnHeaders.Add Text:="Unique Value"
    
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
        
110     For Each oShp9 In oVLayer.Loop(oVLayer.Extent, "", Nothing, "", True)
114         m_frmChangeTracer.lvUniqueValues.ListItems.Add Text:=oShp9.GetField(sField)
        Next
    
118     If Not m_frmChangeTracer.lvUniqueValues.ListItems.Count < 1 Then
120         m_frmChangeTracer.RemoveDuplicates m_frmChangeTracer.lvUniqueValues
        End If
    
        '<EhFooter>
        Exit Sub

m_frmChangeTracer_GetUniqueValues_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmChangeTracer_GetUniqueValues " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmFreqSettings_ApplyAnalysis()
        '<EhHeader>
        On Error GoTo m_frmFreqSettings_ApplyAnalysis_Err
        '</EhHeader>
        Dim sOverLLayer As String
        Dim sAnalysisLyr As String


        'Bugfix Incident animation do not work... it works...
        'Bugfix Coordinate conversion from UTM - > MGRS does not work. we do not support TM in this version
        
100     With m_frmFreqSettings
    
102         sAnalysisLyr = .AnalysisLayer
104         sOverLLayer = .OverLayLayer
        
106         If Not sAnalysisLyr = "" And Not sOverLLayer = "" Then

108             DoFrequencyAnalysis sOverLLayer, sAnalysisLyr
            Else
110             MsgBox "You did not fill in the proper values. Please Retry.", vbInformation, "OASIS Analysis Tools"
            End If
    
        End With

        '<EhFooter>
        Exit Sub

m_frmFreqSettings_ApplyAnalysis_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmFreqSettings_ApplyAnalysis " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmFreqSettings_GetFields(sLayer As String)
        '<EhHeader>
        On Error GoTo m_frmFreqSettings_GetFields_Err
        '</EhHeader>
        Dim oVLayer As TatukGIS_XDK10.XGIS_LayerVector
        Dim i As Integer

100     Set oVLayer = GIS10.get(sLayer)

102     With m_frmFreqSettings
    
104         .ComAnalysisField.Clear

106         If Not oVLayer Is Nothing Then

108             For i = 0 To oVLayer.Fields.Count - 1
110                 .ComAnalysisField.AddItem oVLayer.Fields.Item(i).Name
                Next

            End If

112         If .ComAnalysisField.ListCount > 0 Then .ComAnalysisField.ListIndex = 0

        End With

        '<EhFooter>
        Exit Sub

m_frmFreqSettings_GetFields_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmFreqSettings_GetFields " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmFreqSettings_GetUniqueValues(sLayer As String, _
                                              sField As String)
        '<EhHeader>
        On Error GoTo m_frmFreqSettings_GetUniqueValues_Err
        '</EhHeader>
        Dim oVLayer As TatukGIS_XDK10.XGIS_LayerVector
        Dim i As Integer
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
100     Set oVLayer = GIS10.get(sLayer)

104     With m_frmFreqSettings
106         With .lvUniqueValues
        
108             .ColumnHeaders.Clear
110             .ListItems.Clear
    
112             .ColumnHeaders.Add Text:="Unique Value"
                    
114             For Each oShp9 In oVLayer.Loop(oVLayer.Extent, "", Nothing, "", True)
118                 .ListItems.Add Text:=oShp9.GetField(sField)
                Next
    
            End With
        
122         If Not .lvUniqueValues.ListItems.Count < 1 Then
124             .RemoveDuplicates .lvUniqueValues
            End If

        End With
    
        '<EhFooter>
        Exit Sub

m_frmFreqSettings_GetUniqueValues_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmFreqSettings_GetUniqueValues " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadAvailableThematics()
        '<EhHeader>
        On Error GoTo LoadAvailableThematics_Err
        '</EhHeader>
        Dim rsThGroup As New ADODB.Recordset

        'Set g_RSThemeSettings = New ADODB.Recordset
                
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'ThemeTool'"
                
104     If g_RSAppSettings.Fields.Item("SettingValue1").value = "0" Then
106         m_frmMnuOperations.elLegend.Grid(gsRowHeight, 1) = (0)
            Exit Sub
        End If
    
108     rsThGroup.Open "SELECT * FROM ThemeGroups", m_Cnn, adOpenDynamic, adLockReadOnly
        
110     m_frmMnuOperations.ComThematicsGroup.Clear
112     m_frmMnuOperations.ComThematicsGroup.AddItem "--- No Thematic Group ---"
        
        If Not rsThGroup.Bof And Not rsThGroup.EOF Then
114         SafeMoveFirst rsThGroup
    
116         Do While Not rsThGroup.EOF
                If Not IsNull(rsThGroup.Fields.Item("Name").value) Then
118                 m_frmMnuOperations.ComThematicsGroup.AddItem rsThGroup.Fields.Item("Name").value
                End If
120             rsThGroup.MoveNext
            Loop

        End If
        
122     m_frmMnuOperations.ComThematicsGroup.ListIndex = 0
    
124     m_frmMnuOperations.elLegend.Grid(gsRowHeight, 1) = (765)
        '<EhFooter>
        Exit Sub

LoadAvailableThematics_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.LoadAvailableThematics " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_ActivateTheme(sTheme As String, _
                                             sThemeGR As String)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_ActivateTheme_Err
        '</EhHeader>
        Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
        Dim rsTHGR As New ADODB.Recordset
        Dim rsTh As New ADODB.Recordset
        Dim sFile As String
        Dim oIni As New TatukGIS_XDK10.XGIS_Ini
        
100     rsTHGR.Open "SELECT * FROM ThemeGroups WHERE Name = '" & sThemeGR & "'", m_Cnn, adOpenDynamic, adLockReadOnly
    
102     If (Not rsTHGR.EOF) And (Not rsTHGR.Bof) Then
104         SafeMoveFirst rsTHGR
        Else
            On Error Resume Next
    
106         rsTHGR.Close
108         Set rsTHGR = Nothing
110         Set rsTh = Nothing
        
112         MsgBox "The Theme Group: " & sThemeGR & " is not correct. Please contact your OASIS administrator and try again.", vbInformation, "OASIS Theme Manager"
            Exit Sub
        End If
    
114     rsTh.Open "SELECT * FROM Themes WHERE ThemeGroup = " & rsTHGR.Fields.Item("ID").value, m_Cnn, adOpenDynamic, adLockReadOnly

        If Not IsNull(rsTh.Fields.Item("AnalysisLayer").value) Then
116         Set oLyr = GIS10.get(rsTh.Fields.Item("AnalysisLayer").value)
            
118         If Not oLyr Is Nothing Then
                Dim ofs As New FileSystemObject
                    
                If ofs.FileExists(g_sAppPath & "\data\user\Maps\Layers\" & rsTh.Fields.Item("ThemeConfigName").value) Then

120                 With oLyr
                        .UseConfig = True
                        sFile = g_sAppPath & "\data\user\Maps\Layers\" & Replace$(rsTh.Fields.Item("ThemeConfigName").value, ".ini", "")
122                     .ConfigName = sFile '"C:\OASIS\NFI1"
128                     .RereadConfig
130                     .Params.Visible = True
                        .draw
                    End With
        
134                 ShowThemelegend oLyr
                Else
                    MsgBox "Seems Like File " & g_sAppPath & "\data\user\Maps\Layers\" & rsTh.Fields.Item("ThemeConfigName").value & " is missing or configuration is incorrect.", vbInformation, "OASIS Client"
                End If

            Else
        
                On Error Resume Next
    
136             rsTHGR.Close
138             rsTh.Close
140             Set rsTHGR = Nothing
142             Set rsTh = Nothing
144             MsgBox "It seems like Theme settings are corrupt for theme: " & sTheme & vbCrLf & "Please contact your OASIS administrator to adjust this.", vbInformation, "OASIS Theme Manager"
                Exit Sub
        
            End If
        End If

        On Error Resume Next
    
146     rsTHGR.Close
148     rsTh.Close
150     Set rsTHGR = Nothing
152     Set rsTh = Nothing
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_ActivateTheme_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.m_frmMnuOperations_ActivateTheme " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub



Private Sub m_frmMnuOperations_setScope()
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_setScope_Err
        '</EhHeader>
100     m_frmChangeTracer.Init
102     m_frmChangeTracer.Show vbModeless, Me
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_setScope_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_setScope " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmTextAnnoSettings_DeleteAnnoText()
        '<EhHeader>
        On Error GoTo m_frmTextAnnoSettings_DeleteAnnoText_Err
        '</EhHeader>

100     If m_frmTextAnnoSettings.lstTexts.ListIndex = -1 Then Exit Sub
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
102     For Each oShp9 In m_oDrawLyr.Loop(m_oDrawLyr.Extent, "GIS_UID = " & m_frmTextAnnoSettings.lstTexts.ItemData(m_frmTextAnnoSettings.lstTexts.ListIndex), Nothing, "", True)
        
104         If Not oShp9 Is Nothing Then
106             m_oDrawLyr.Delete oShp9.uID
108             GIS10.UpDate
            End If
        Next
        
110     ListAllAnnotationShps
    
112     If m_frmTextAnnoSettings.lstTexts.ListCount > 0 Then
114         m_frmTextAnnoSettings.lstTexts.ListIndex = 0
        End If

        '<EhFooter>
        Exit Sub

m_frmTextAnnoSettings_DeleteAnnoText_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmTextAnnoSettings_DeleteAnnoText " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmTextAnnoSettings_GetAnnoTextProp()
    Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
    
    For Each oShp9 In m_oDrawLyr.Loop(m_oDrawLyr.Extent, "GIS_UID = " & m_frmTextAnnoSettings.lstTexts.ItemData(m_frmTextAnnoSettings.lstTexts.ListIndex), Nothing, "", True)
          
        If Not oShp9 Is Nothing Then
    
            With oShp9.Params.labels
                m_frmTextAnnoSettings.panColorBack.BackColor = .color
                m_frmTextAnnoSettings.panColorFore.BackColor = .FontColor
                m_frmTextAnnoSettings.txtAnnoText.Text = .value
                m_frmTextAnnoSettings.txtRotation.Text = .rotate
            End With
        
            GIS10.UpDate
        End If
    Next

End Sub

Private Sub m_frmTextAnnoSettings_ResetAll()
    RemoveAllShps m_oDrawLyr
    GIS10.UpDate
    ListAllAnnotationShps
End Sub

Private Sub m_frmTextAnnoSettings_UpdateAnnoText()
        '<EhHeader>
        On Error GoTo m_frmTextAnnoSettings_UpdateAnnoText_Err
        '</EhHeader>
        Dim oshp As TatukGIS_XDK10.XGIS_Shape

100     If m_frmTextAnnoSettings.lstTexts.ListIndex = -1 Then Exit Sub

102     For Each oshp In m_oDrawLyr.Loop(m_oDrawLyr.Extent, "GIS_UID = " & m_frmTextAnnoSettings.lstTexts.ItemData(m_frmTextAnnoSettings.lstTexts.ListIndex), Nothing, "", True)
        
104         If Not oshp Is Nothing Then
    
106             With oshp.Params.labels
108                 .value = m_frmTextAnnoSettings.txtAnnoText.Text
110                 .rotate = m_frmTextAnnoSettings.txtRotation.Text
                    .color = m_frmTextAnnoSettings.panColorBack.BackColor
                    .FontColor = m_frmTextAnnoSettings.panColorFore.BackColor
                End With
        
112             GIS10.UpDate
            End If
        Next
114     ListAllAnnotationShps

116     m_frmTextAnnoSettings.lstTexts.ListIndex = 0

        '<EhFooter>
        Exit Sub

m_frmTextAnnoSettings_UpdateAnnoText_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmTextAnnoSettings_UpdateAnnoText " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmIncidentsV2ControlPanel_LoadIncidents()
    m_frmIncidentsV2DataEntry.Show vbModeless, Me
End Sub

Private Sub m_frmIncidentsV2ControlPanel_LoadSitrep()
    Dim ppic As StdPicture
    m_frmIncidentsV2SitrepGenerator.Show vbModeless, Me
    m_frmIncidentsV2SitrepGenerator.ShowWait True

    DoEvents
                
    Clipboard.Clear
    GIS10.PrintClipboard
    Set ppic = Clipboard.GetData(vbCFEMetafile)
    m_frmIncidentsV2SitrepGenerator.Init g_sUserName, dxGISDataGrid.Filter.FilterText, ppic
    m_frmIncidentsV2SitrepGenerator.ShowWait False
End Sub

Private Sub m_frmIncidentsV2DataEntry_GetLocationOnMap()

    ReDim ptsSel(0)
    ReDim gdpts(0)
    g_CurrentTool = oCreateLocationPoint 'oCreateLocationPoint
    g_CurrentFeatureType = Custom
    m_bPolyL = False
    m_bPolyG = False
    SetOASISTool

End Sub

Private Sub m_InternetCheck_CompleteEX(enmRunnerType As OASIS_SynchNG.RunnerType)

    If enmRunnerType = InternetConnectionCheck Then
        m_bInternetCon = m_InternetCheck.Connected
        m_sInternetCon = m_InternetCheck.InternetConnectionType
        
        If m_bInternetCon Then
            SetIMess "Internet Status: " & m_sInternetCon
        Else
            SetIMess "No Internet Status: " & m_sInternetCon
        End If
        
    End If

End Sub

Private Sub m_InternetCheck_WorkerError(errnum As Long, errDesc As String, errSource As String)
    DebugPrint errDesc
End Sub

Private Sub m_oClipboardViewer_ClipboardChanged()
Dim lCount As Long
Dim lFormat As Long
Dim lID As Long
   
   ' To cope with an problems caused by clipboard applications
   ' that don't behave correctly.  VB Add-ins are the only
   ' example I've found of this because they have to use the
   ' ridiculous PasteFace method to set up a picture on the
   ' button.
   On Error Resume Next
   
   
   ' Show the new clipboard contents:
   'DebugPrint "Clipboard changed"
   With ClipBoard_Ext
      lCount = .GetCurrentFormats(Me.hWnd)
      For lFormat = 1 To lCount
         lID = .GetCurrentFormatID(lFormat)
         DebugPrint "Clipboard Format:" & .GetCurrentFormatName(lFormat)
      Next lFormat
      
   End With

End Sub

Private Sub TranslateClipboard()
    Dim lID As Long
    Dim sText As String
    Dim hBmp As Long
    Dim sOut As String
    Dim sHex As String
    Dim bData() As Byte
    Dim sFiles() As String
    Dim iFileCount As Long
    Dim i As Long, l As Long

    Screen.MousePointer = vbHourglass
    
    lID = 666 'lstFormats.itemData(lstFormats.ListIndex)

    Select Case lID

        Case CF_BITMAP
            ' Here is the simple way!
           ' Set picView = Clipboard.GetData(CF_BITMAP)
           ' picView.ZOrder
          
            ' As an example, here is the hard way to do it
            ' - the clipboard contains a bitmap handle:
            'picView.Cls
            'If (ClipBoard_Ext.ClipboardOpen(Me.hWnd)) Then
            '   dim hBmp as long, hBmpOld as long, hDC as long, tBM as BITMAP
            '    hBmp = ClipBoard_Ext.GetClipboardMemoryHandle(CF_BITMAP)
            '    hDC = CreateCompatibleDC(picView.hDC)
            '    hBmpOld = SelectObject(hDC, hBmp)
            '    GetObjectAPI pic.Handle, Len(tBM), tBM
            '    BitBlt picView.hDC,0,0, tBM.bmWidth,tBM.bmHeight, hDC,0,0,SRCCOPY
            '    SelectObject hDC,hBmpOld
            '    DeleteObject hDC
            '    ClipBoard_Ext.ClipboardClose
            'End If
          
        Case CF_DIB
            ' The simple way!
            'Set picView = Clipboard.GetData(CF_DIB)
            'picView.ZOrder
          
            ' Actually, the clipboard handle points to a BITMAPINFO
            ' structure followed by the bitmap bits.  Left as an
            ' exercise...  See my Image Processing sample for details
            ' of DIB Sections to find out a bit more.
          
        Case CF_ENHMETAFILE
            ' The simple way!
            'Set picView = Clipboard.GetData(CF_ENHMETAFILE)
            'picView.ZOrder
          
            ' The clipboard handle is a handle to an Enhanced Metafile:
            ' Left as an exercise.
          
        Case CF_HDROP
            ' the clipboard handle can be passed to the DragQueryFile
            ' function to get the information:
          
            If ClipBoard_Ext.ClipboardOpen(Me.hWnd) Then
                ClipBoard_Ext.GetFileList sFiles(), iFileCount

                For i = 1 To iFileCount
                    'txtView.Text = txtView.Text & sFiles(i) & vbCrLf
                Next i

            End If
          
        Case CF_METAFILEPICT
            ' The clipboard handle is a handle to an old style Metafile:
            'Set picView = Clipboard.GetData(CF_ENHMETAFILE)
          
        Case Else

            ' Assume Text
            If (ClipBoard_Ext.ClipboardOpen(Me.hWnd)) Then

                ' Get a string and show that:
                If (ClipBoard_Ext.GetTextData(lID, sText)) Then
                    ' txtView.Text = sText
                End If

                ClipBoard_Ext.ClipboardClose
            End If
          
    End Select
    
    Screen.MousePointer = vbNormal

End Sub

Private Sub m_oSQLLyrSynch_CompleteEX(enmRunnerType As OASIS_SynchNG.RunnerType)
    If enmRunnerType = SQLLyrSynch Then
                
    End If
End Sub

Private Sub m_oSQLLyrSynch_WorkerError(errnum As Long, errDesc As String, errSource As String)
    SetMess errDesc & errSource
End Sub


Private Sub mnu16rows_Click()
    mnu4Rows.Checked = False
    mnu8rows.Checked = False
    mnu16rows.Checked = True
    mnuUnlimitedRows.Checked = False
    dxGISDataGrid.MaxRowLineCount = 16
End Sub

Private Sub mnu4Rows_Click()
    mnu4Rows.Checked = True
    mnu8rows.Checked = False
    mnu16rows.Checked = False
    mnuUnlimitedRows.Checked = False
    dxGISDataGrid.MaxRowLineCount = 4
End Sub

Private Sub mnu8rows_Click()
    mnu4Rows.Checked = False
    mnu8rows.Checked = True
    mnu16rows.Checked = False
    mnuUnlimitedRows.Checked = False
    dxGISDataGrid.MaxRowLineCount = 8
End Sub

Private Sub mnuAddToClipboard_Click()
    GIS10.viewer.PrintClipboard
    MsgBox "The map has been added to your clipboard....", vbInformation
End Sub

Private Sub mnuAutoHeight_Click()
    mnuAutoHeight.Checked = Not mnuAutoHeight.Checked
    SetGridOption egoRowAutoHeight, mnuAutoHeight.Checked
End Sub

Private Sub mnuAutoSize_Click()

    mnuAutoSize.Checked = Not mnuAutoSize.Checked
    SetGridOption 18, mnuAutoSize.Checked

End Sub

Private Sub mnuAutoZoom_Click()
    mnuAutoZoom.Checked = Not mnuAutoZoom.Checked
End Sub

Private Sub mnuGetCoords_Click()
   Clipboard.Clear
   Clipboard.SetText AB.Bands.Item("bStatus").Tools.Item("lblCoords").caption
   MsgBox "Coordinates have been added to your clipboard.", vbInformation, "OASIS Coordinate utilities"
'    ConvertMultiCoords CDbl(2), CDbl(2)
End Sub

Private Sub ConvertMultiCoords(x As Double, _
                               y As Double)
    Dim iProjectionType As Integer
    Dim iDatumNumber As Integer
    Dim iUnits As Integer
    Dim dOriginLongitude As Double

End Sub

Private Sub mnuGridSettings_Click()
    'TODO Some Optimization stuff
End Sub

Private Sub mnuHideSelected_Click()
        '<EhHeader>
        On Error GoTo mnuHideSelected_Click_Err
        '</EhHeader>
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
100     Set lL = GIS10.get(GISGridLayerName)
        
102     If Not lL Is Nothing Then
104         HideShowSelected lL, False
106         mnuHideSelected.Visible = False
108         mnuShowSelected.Visible = True
        End If

        '<EhFooter>
        Exit Sub

mnuHideSelected_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.mnuHideSelected_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuLayer_Click(Index As Integer)
    abGridTools.Tools.Item("comLyr").CBListIndex = Index
End Sub

Private Sub mnuLoadAll_Click()
    chkOnlyVisible.value = vbUnchecked
    mnuVisibleExtent.Checked = False
    mnuLoadAll.Checked = True
    LoadLayerAttrDataToGridInit 'True

End Sub

Private Sub fillm_LyrCol()
        '<EhHeader>
        On Error GoTo fillm_LyrCol_Err
        '</EhHeader>
        Dim i As Integer
        
100     Set m_LyrCol = New Collection
            
102     m_LyrCol.Add "--All--", "--All--"
            
104     For i = 0 To GIS10.items.Count - 1
                
106         If GisUtils.IsInherited(GIS10.items.Item(i), "XGIS_LayerVector") Then
108             m_LyrCol.Add GIS10.items.Item(i).Name, GIS10.items.Item(i).caption
110             m_frmMainSettings.ComTipLyr.AddItem GIS10.items.Item(i).caption
            End If
        
        Next

        '<EhFooter>
        Exit Sub

fillm_LyrCol_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.fillm_LyrCol " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuOPSSetting_Click()
        '<EhHeader>
        On Error GoTo mnuOPSSetting_Click_Err
        '</EhHeader>
        Dim i As Integer

100     With m_frmMainSettings
102         .ComTipLyr.Clear
104         .ComTipLyr.AddItem "--All--"

106         fillm_LyrCol
            
108         FindIndexStrEx .ComTipLyr, oMapTipSetting.MapTipLayer
            
110         If .ComTipLyr.ListIndex = -1 Then FindIndexStrEx .ComTipLyr, "--All--"
112         .watermark.GIS_Viewer = GIS10.viewer
114         .Arrow.GIS_Viewer = GIS10.viewer
116         .cpArrow.color = NArrow.Color1
118         .Arrow.Color1 = NArrow.Color1
120         .Arrow.Color2 = NArrow.Color2
122         .cboArrow.ListIndex = NArrow.Symbol
124         .chkArrowTransparent.value = IIf(NArrow.TRANSPARENT = True, vbChecked, vbUnchecked)
126         .chkUseNorth.value = IIf(oMapObjects.UseNorthArrow, vbChecked, vbUnchecked)
            
128         If Len(NArrow.Path) > 1 Then
130             .Arrow.Path = NArrow.Path
132             .chkUseCustom.Tag = NArrow.Path
134             .chkUseCustom.value = vbChecked
            End If
            
136         .Show vbModal, Me
            
138         picToolTip.Visible = False
140         tmrToolTip.Enabled = False
142         tmrToolTip.Interval = 1000 * oMapTipSetting.TipDelay
144         tmrToolTip.Enabled = oMapTipSetting.Enabled
        End With
    
146     With oSelectionStyle
148         GIS10.SelectionColor = .color
150         GIS10.SelectionOutlineOnly = .OutLineOnly
152         GIS10.SelectionTransparency = .Transparency
154         GIS10.SelectionWidth = .Width
        End With
    
156     If Not m_oSQLIncLyr Is Nothing Then

158         With oIncidentLayerSettings
160             m_oSQLIncLyr.CachedPaint = .CachedPaint
                '.ConfigFilePAth
                'm_oSQLIncLyr.HideFromLegend = .HideFromLegend
162             m_oSQLIncLyr.IgnoreShapeParams = False  '.IgnoreShapeParams
164             m_oSQLIncLyr.IncrementalPaint = .IncrementalPaint
                '.UseConfig
                '.UseFileParams
166             'm_oSQLIncLyr.draw
            End With

        End If
    
168     With oMapSettings
            '        .AutoScroll
            '        .iMAPUnits
            '        .MapRotation
            '        .ScrollBars
            '        .StoreLayerParamsInProject
        End With

        '<EhFooter>
        Exit Sub

mnuOPSSetting_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.mnuOPSSetting_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuMapTip_Click()
    mnuMapTip.Checked = Not mnuMapTip.Checked
    oMapTipSetting.Enabled = mnuMapTip.Checked
End Sub

Private Sub mnuOPSView_Click()

    If mnuOPSView.Checked Then
        If mnuOPSView.Tag = True Then NArrow.GIS_Viewer = GIS10.viewer
        If MsgScroll.ListCount > 0 Then
            elScroller.Visible = True
            elScroller.Move 5, elTatukGIS.Height - MsgScroll.Height, elTatukGIS.Width, 255
        End If
        
        If Not m_frmOPSView Is Nothing Then Unload m_frmOPSView
        mnuOPSView.Checked = False
        'Call elTatukGIS_RealignFrame
    Else
        mnuOPSView.Tag = IIf(NArrow.GIS_Viewer Is Nothing, False, True)
         NArrow.GIS_Viewer = Nothing
        Set m_frmOPSView = New frmOPSView
        m_frmOPSView.Show
        elScroller.Visible = False
        mnuOPSView.Checked = True
        SetParent GIS10.hWnd, m_frmOPSView.elHolder.hWnd
        
    End If

End Sub

Private Sub mnuPanTool_Click()
    GIS10.Mode = TatukGIS_XDK10.XgisDrag
End Sub

Private Sub mnuPrevious_Click()
        '<EhHeader>
        On Error GoTo mnuPrevious_Click_Err
        '</EhHeader>
        Dim iDeduct As Integer

100     m_bPrevActionUsed = True
102     GIS10.Lock
    
104     If UBound(m_PrevExt) <> LBound(m_PrevExt) Then
106         If GisUtils.GisIsSameExtent(GIS10.VisibleExtent, m_PrevExt(UBound(m_PrevExt))) Then
108             iDeduct = 1
            End If
        
110         GIS10.VisibleExtent = m_PrevExt(UBound(m_PrevExt) - iDeduct)
112         iDeduct = iDeduct + 1

114         If UBound(m_PrevExt) <> LBound(m_PrevExt) Then ReDim Preserve m_PrevExt(UBound(m_PrevExt) - iDeduct)
        Else
116         GIS10.VisibleExtent = m_PrevExt(UBound(m_PrevExt))
        End If
        
118     GIS10.Unlock

        '<EhFooter>
        Exit Sub

mnuPrevious_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.mnuPrevious_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SelAutoFlash(uID As Long, sLayerName As String)
        '<EhHeader>
        On Error GoTo SelAutoFlash_Err
        '</EhHeader>

        'GIS109 done
        Dim oLyr9 As TatukGIS_XDK10.XGIS_LayerVector
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
    
100     Set oLyr9 = GIS10.get(sLayerName)

102     For Each oShp9 In oLyr9.Loop(oLyr9.Extent, "GIS_UID = " & uID, Nothing, "", True)
104         If Not oShp9 Is Nothing Then oShp9.Flash: Exit For
        Next
    
        '<EhFooter>
        Exit Sub

SelAutoFlash_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SelAutoFlash " & _
               "at line " & Erl
        '</EhFooter>
End Sub


Private Sub mnuRec2Clip_Click()
    DebugPrint "bCopy"
    FindFeatureFromGrid False, False, False, True
End Sub

Private Sub mnuSendToSMS_Click()
    frmSMSMain.Show vbModal
End Sub

Private Sub mnuShowCentroid_Click()
    mnuShowCentroid.Checked = Not mnuShowCentroid.Checked
    With dxGISDataGrid
            .Columns(.Columns.Count - 2).Visible = mnuShowCentroid.Checked
            .Columns(.Columns.Count - 1).Visible = mnuShowCentroid.Checked
    End With
End Sub

Private Sub mnuShowRecordCount_Click()
    Dim sumgroup As DXDBGRIDLibCtl.dxGridSummaryGroup
    Dim sumitem As DXDBGRIDLibCtl.dxGridSummaryItem
    Dim i As Integer
    Dim Col As DXDBGRIDLibCtl.dxGridColumn

    mnuShowRecordCount.Checked = IIf(mnuShowRecordCount.Checked, False, True)
dxGISDataGrid.Option = egoCStyleFormatting
dxGISDataGrid.OptionEnabled = True

        With dxGISDataGrid
    If mnuShowRecordCount.Checked Then
            Set sumgroup = .SummaryGroups.Add
            sumgroup.DefaultGroup = True
            Set sumitem = sumgroup.SummaryItems.Add
            
            .Columns(1).SummaryFooterFormat = "%.0f records"
            sumitem.SummaryType = cstCount
            
            

            For i = 0 To .Columns.Count - 1
                Set Col = .Columns(i)

                If Col.Visible Then
                    Col.SummaryFooterType = cstCount
                    Exit For
                End If

            Next

            .Option = egoShowFooter
            .OptionEnabled = True
            
            
            

        Else
            For i = 0 To .SummaryGroups.Count
                .SummaryGroups.Remove i
            Next
            
            .Option = egoShowFooter
            .OptionEnabled = False
        End If

    End With

End Sub

Private Sub mnuShowSelected_Click()
        '<EhHeader>
        On Error GoTo mnuShowSelected_Click_Err
        '</EhHeader>
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
100     Set lL = GIS10.get(GISGridLayerName)
        
102     If Not lL Is Nothing Then
104         HideShowSelected lL, True
106         mnuHideSelected.Visible = True
108         mnuShowSelected.Visible = False
        End If

        '<EhFooter>
        Exit Sub

mnuShowSelected_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.mnuShowSelected_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuUnlimitedRows_Click()
    mnu4Rows.Checked = False
    mnu8rows.Checked = False
    mnu16rows.Checked = False
    mnuUnlimitedRows.Checked = True
    dxGISDataGrid.MaxRowLineCount = -1
End Sub

Private Sub mnuVisibleExtent_Click()
    chkOnlyVisible.value = vbChecked
    mnuVisibleExtent.Checked = True
    mnuLoadAll.Checked = False
    LoadLayerAttrDataToGridInit ', True
End Sub

Private Sub mnuZoomIN_Click()
'GIS109 Done
    GIS10.zoom = GIS10.zoom * 2
End Sub

Private Sub mnuZoomOUT_Click()
'GIS109 Done
    GIS10.zoom = GIS10.zoom / 2
End Sub

Private Sub mnuZoomRectanletool_Click()
'GIS109 Done
    GIS10.Mode = TatukGIS_XDK10.XgisZoom
End Sub

Sub ShowFormMT(ModalFlag As Variant)
    'This is the multithreaded procedure... We cannot
    'directly load a form here, since doing so will cause
    'VB to crash... We must call the Main Thread and make
    'it perform the Form load operation...
    'The ObjectInThreadContext property returns a reference
    'to the original Form (Form1) running on the original
    'Thread... We call the ShowFormNow() Sub to show the form
    'But this Sub is called in context to the original Thread
        '<EhHeader>
        On Error GoTo ShowFormMT_Err
        '</EhHeader>
100 FormThread.ObjectInThreadContext.ShowFormNow (CLng(ModalFlag))

    'REMEMBER:Here in this case, ModalFlag is a "real"
    'parameter unlike in the FindPrimes() Sub
        '<EhFooter>
        Exit Sub

ShowFormMT_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ShowFormMT " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub MsgScroll_HotSpotClick(ByVal Index As Long)
        '<EhHeader>
        On Error GoTo MsgScroll_HotSpotClick_Err
        '</EhHeader>
 
        Dim lShape As New TatukGIS_XDK10.XGIS_Shape
        Dim lLayer As New TatukGIS_XDK10.XGIS_LayerVector
            
        For Each lShape In m_oSQLIncLyr.Loop(m_oSQLIncLyr.Extent, "ID = '" & MsgScroll.GetKey(Index) & "'", Nothing, "", True)
            GIS10.VisibleExtent = lShape.Extent
112         lShape.Flash 8, 10
        Next
        
116     MsgScroll.ItemColor(Index) = vbBlue
    
        '<EhFooter>
        Exit Sub

MsgScroll_HotSpotClick_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.MsgScroll_HotSpotClick " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetMess(sMess As String)
        '<EhHeader>
        On Error GoTo SetMess_Err
        '</EhHeader>

100     With AB.Bands.Item("bStatus")
    
    
102     If (.Width - (.Tools.Item("lblIConnection").Width + .Tools.Item("lblCoords").Width)) > 0 Then

104         .Tools.Item("lblSynchProgress").Width = .Width - (.Tools.Item("lblIConnection").Width + .Tools.Item("lblCoords").Width)
         End If
106      .Tools.Item("lblSynchProgress").Visible = True
108         With .Tools.Item("lblSynchProgress")
110             .Text = sMess
112             .caption = sMess
            End With
        End With

        '<EhFooter>
        Exit Sub

SetMess_Err:
        '</EhFooter>
End Sub

Public Sub SetIMess(sMess As String)
    With AB.Bands.Item("bStatus").Tools.Item("lblIConnection")
        .Text = sMess
        .caption = sMess
        
    End With
End Sub

Private Function GetTimeFormatted() As String

    GetTimeFormatted = IIf(Len(Hour(Now())) = 1, "0" & Hour(Now()), Hour(Now())) & ":" & IIf(Len(Minute(Now())) = 1, "0" & Minute(Now()), Minute(Now())) & "." & IIf(Len(Second(Now())) = 1, "0" & Second(Now()), Second(Now()))

End Function

Private Sub NArrow_OnMouseUp(translated As Boolean, ByVal Button As TatukGIS_XDK10.XMouseButton, ByVal Shift As TatukGIS_XDK10.XShiftState, ByVal x As Long, ByVal y As Long)
    'NArrow
   ' narrow
End Sub

Private Sub OASSelect_Click()
    ComFeatureLayer.Enabled = False
    
    Select Case OASSelect.Text
    
        Case "Area Select"
            ChangeSelectionType oAreaSelect
        Case "Circle Select"
            ChangeSelectionType oCircleSelect
        Case "Feature Select"
            ChangeSelectionType oFeatureSelect
            ComFeatureLayer.Enabled = True
        Case "Line Select"
            ChangeSelectionType oLineSelect
        Case Else
            MsgBox "You must choose a Selector Tool From the dropdown meny first.", vbInformation, "OASIS Selector Tools"
    End Select

End Sub

Private Sub OASSelect_MouseDownOnDropdown()
    PopupMenu mnuSelToolPopUp
End Sub


Private Sub mnuAreaSelect_Click()
    ChangeSelectionType oAreaSelect
    ComFeatureLayer.Enabled = False
    OASSelect.Text = "Area Select"
    OASSelect.toolTipText = "Current Tool Is Area Select tool"
End Sub

Private Sub mnuCircleSelect_Click()
    ChangeSelectionType oCircleSelect
    ComFeatureLayer.Enabled = False
    OASSelect.Text = "Circle Select"
    OASSelect.toolTipText = "Current Tool Is Circle Select tool"
End Sub

Private Sub mnuFeatureSelect_Click()
    ChangeSelectionType oFeatureSelect
    ComFeatureLayer.Enabled = True
    OASSelect.Text = "Feature Select"
    OASSelect.toolTipText = "Current Tool Is Feature Select tool"
End Sub

Private Sub mnuLineSelect_Click()
    ChangeSelectionType oLineSelect
    ComFeatureLayer.Enabled = False
    OASSelect.Text = "Line Select"
    OASSelect.toolTipText = "Current Tool Is Line Select tool"
End Sub

Private Sub ChangeToolCursor()
    'GIS109 Done
    'Prepare the Cursor.....
    GIS10.CursorPrepare 2048, LoadResPicture(101, vbResCursor) 'LoadCursorFromFile("C:\Documents and Settings\iMMAP\Desktop\Cursores Zelda\immap.ani")
    
    GIS10.Mode = TatukGIS_XDK10.XgisUserDefined
    GIS10.Cursor = 2048
    

End Sub

Private Sub OASSelect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DebugPrint ""
End Sub

Private Sub oClientInterCom_OnDataArrival(bData As String)
    On Error Resume Next
    'DebugPrint bData
    DebugPrint bData
    Select Case left(CStr(bData), 2)
    
        Case "??"
            'DebugPrint "INFO ONLY"
        Case "!!"
            'DebugPrint bData
        Case "**"
            'DebugPrint "Completed Action"
            InterComHandling bData
        Case Else
            
    End Select
    
    If CStr(Trim$(bData)) = "**??Synchronisation complete" And g_bRefreshMapAfterSynch Then
    
        If Not frmMain.GIS10.Mode = TatukGIS_XDK10.XgisEdit Then
            frmMain.GIS10.UpDate
        End If
    
    End If
    
    If left(CStr(Trim$(bData)), 21) = "**?? TABLES CREATED: " Then
    
        Dim sTables() As String
        Dim sTablesString As String
        Dim i As Long
        Dim bReload As Boolean
        Dim oRS As ADODB.Recordset
        
        sTablesString = Replace(CStr(Trim(bData)), "**?? TABLES CREATED: ", "")
        sTables = Split(sTablesString, ";")
        
        Dim sProject As String
        sProject = ReadTextFile(GIS10.ProjectName)
        
        Do Until i > UBound(sTables)

            If right(sTables(i), 4) = "_FEA" Then
        
                sTables(i) = left(sTables(i), Len(sTables(i)) - 4)
                If InStr(sProject, "Name=" & sTables(i)) > 0 Then bReload = True
                oRS.Close
            
            End If

            i = i + 1
        Loop
        
        Set oRS = Nothing

        If bReload Then
            If MsgBox("New layers have been loaded into OASIS.  Do you want to add it to the map?", vbYesNo) = vbYes Then
                'AddSQLLayersToProject9
                Dim sProjectPath As String
                sProjectPath = GIS10.ProjectName
                InitMap sProjectPath, GIS10.VisibleExtent
            End If
        End If
        
    End If
    
End Sub

Private Sub InterComHandling(sComString As String)
    'RT_enmPaddy = "00"
    'RT_IncidentSynch = "01"
    'RT_SQLLyrSynch = "02"
    'RT_GeoMarks = "03"
    'RT_InternetConnectionCheck = "04"
    'RT_IncidentNotifier = "05"
    'RT_IncidentDeleted = "06"
    'RT_IncidentEdited = "07"
    'RT_GeoMarksDeleted = "08"
    'RT_GeoMarksEdited = "09"
    
    If g_bOnlineCheckedAtLogin Then
    
        Dim Col As Collection
        Dim Item As Variant
        Dim oRS As ADODB.Recordset
        Dim i As Long
        Dim j As Long
        Dim sITM() As String

        Select Case Mid$(sComString, 3, 2)
    
            Case "00"
    
            Case "01"
                Set Col = oSync.NewIncidents
            
                SetMess "Synchronised items recieved." & " Time: " & GetTimeFormatted

                sITM = Split(Mid$(sComString, 5), "//")
            
                For j = LBound(sITM) To IIf(UBound(sITM) > 1000, 1000, UBound(sITM))

                    If Len(sITM(j)) > 1 Then
                    
                        Set oRS = New ADODB.Recordset
                        If Not bSQLServerInUse Then
                            oRS.Open "SELECT ID, Province, District, TYPE, TARGET, Incident_DATE FROM oincidents_FEA WHERE (((oincidents_FEA.Incident_DATE)>#" & Format(Now() - 7, "Medium Date") & "#)) AND [ID] = '" & sITM(j) & "'", m_Cnn, adOpenForwardOnly, adLockReadOnly
                        Else
                            oRS.Open "SELECT ID, Province, District, TYPE, TARGET, Incident_DATE FROM oincidents_FEA WHERE (((oincidents_FEA.Incident_DATE)>'" & Format(Now() - 7, "Medium Date") & "')) AND [ID] = '" & sITM(j) & "'", m_Cnn, adOpenForwardOnly, adLockReadOnly
                        End If
                        
                        'oRS.Open "SELECT ID, Province, District, TYPE, TARGET, Incident_DATE FROM oincidents_FEA WHERE (((oincidents_FEA.Incident_DATE)>Now()-7)) AND [ID] = '" & sITM(j) & "'", m_Cnn, adOpenForwardOnly, adLockReadOnly
                        'oRS.Open "SELECT ID, Province, District, TYPE, TARGET, Incident_DATE FROM oincidents_FEA WHERE [ID] = '" & sITM(j) & "'", m_Cnn, adOpenForwardOnly, adLockReadOnly
                        g_ColAlerts.Add sITM(j) ', CStr(Item)
                        i = i + 1

                        If Not oRS.EOF Then

                            With oRS.Fields
                                MsgScroll.AddItem "New Incident #" & g_ColAlerts.Count & " In: " & .Item("Province").value & " - " & .Item("District").value & " - " & .Item("Incident_DATE").value, sITM(j)
                            End With

                            SetMess "Synchronised items recieved. #" & g_ColAlerts.Count & " Time: " & GetTimeFormatted
                        End If
                    End If

                Next

                If MsgScroll.ListCount > 0 Then
            
                    If MsgScroll.ListCount < 10 Then
                    
                        Do Until MsgScroll.ListCount = 10
                            'This is needed to ensure it is rendered correctly
                            'MsgScroll.AddItem " . . . . . . . . . . . . . . . . ."
                            MsgScroll.AddItem "-                                -"
                        Loop

                    End If

                    If Not mnuOPSView.Checked Then GIS10.Height = elTatukGIS.Height + 150

                    If Not elScroller.Visible Then

                        SafeMoveFirst g_RSAppSettings
                        g_RSAppSettings.Find "SettingName = 'Notifier'"

                        If Not g_RSAppSettings.EOF And Not g_RSAppSettings.Bof Then
                            If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").value) Then
                                MsgScroll.BackColor = g_RSAppSettings.Fields.Item("SettingValue1").value
                                MsgScroll.ForeColor = g_RSAppSettings.Fields.Item("SettingValue2").value
                            End If
                        End If

                        If Not mnuOPSView.Checked Then
    
                            elScroller.Visible = True
                            elScroller.Move 5, elTatukGIS.Height - MsgScroll.Height, elTatukGIS.Width, 255

                        End If
                    
                    End If

                    elScroller.ZOrder
                    cmdMSGScroller.ZOrder
                    MsgScroll.AutoScroll = True

                Else
                    elScroller.Visible = False

                    If Not mnuOPSView.Checked Then
                        GIS10.Height = elTatukGIS.Height + 400
                    End If
                End If

                SetMess "Incident Synch Check Complete..." & " Time: " & GetTimeFormatted

            Case "02"

            Case "03"

            Case "04"

                m_bInternetCon = m_InternetCheck.Connected
                '    m_sInternetCon = m_InternetCheck.InternetConnectionType
        
                If m_bInternetCon Then
                    SetIMess "Online" ' "Internet Status: " & m_sInternetCon
                ElseIf InStr(g_sAppServerPath, "localhost") > 0 Or InStr(g_sAppServerPath, "127.0.0.1") > 0 Then
                    SetIMess "Localhost" ' "No Internet Status: " & m_sInternetCon
                Else
                    SetIMess "Offline" ' "No Internet Status: " & m_sInternetCon
                End If

            Case "05"

            Case "06"

            Case "07"

            Case "08"

            Case "09"
        
            Case "??"

                If Len(sComString) > 7 Then SetMess right(sComString, Len(sComString) - 4)
            
        End Select
    
        If g_bRefreshMapAfterSynch Then
            If Trim$(sComString) = "**??Synchronisation complete" Then GIS10.UpDate
        End If

    End If

End Sub

Private Sub EmulateTicker()
        '<EhHeader>
        On Error GoTo EmulateTicker_Err
        '</EhHeader>
        Dim i As Integer
        Dim oRS As ADODB.Recordset

100     SetMess "Synchronised items recieved." & " Time: " & GetTimeFormatted
   
102     Set oRS = New ADODB.Recordset
104     oRS.Open "SELECT UID, ID, Province, District, TYPE, TARGET, Incident_DATE FROM oincidents_FEA", m_Cnn, adOpenForwardOnly, adLockReadOnly

        Do While Not oRS.EOF

106         g_ColAlerts.Add oRS.Fields.Item("ID").value
        
108         i = i + 1

110         With oRS.Fields
112             MsgScroll.AddItem "New Incident #" & i & " In: " & .Item("Province").value & " - " & .Item("District").value & " - " & .Item("Incident_DATE").value, .Item("ID").value
            End With

            oRS.MoveNext

        Loop

114     SetMess "Synchronised items recieved. #" & i & " Time: " & GetTimeFormatted
   
116     If Not mnuOPSView.Checked Then
118         GIS10.Height = elTatukGIS.Height + 150
        End If

120     If Not elScroller.Visible Then

122         SafeMoveFirst g_RSAppSettings
124         g_RSAppSettings.Find "SettingName = 'Notifier'"

126         If Not g_RSAppSettings.EOF And Not g_RSAppSettings.Bof Then
128             If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").value) Then
130                 MsgScroll.BackColor = g_RSAppSettings.Fields.Item("SettingValue1").value
132                 MsgScroll.ForeColor = g_RSAppSettings.Fields.Item("SettingValue2").value
                End If
            End If

134         elScroller.Visible = True
136         elScroller.Move 5, elTatukGIS.Height - MsgScroll.Height, elTatukGIS.Width, 255
        End If

138     elScroller.ZOrder
140     cmdMSGScroller.ZOrder
142     MsgScroll.AutoScroll = True

144     SetMess "Incident Synch Check Complete..." & " Time: " & GetTimeFormatted

        '<EhFooter>
        Exit Sub

EmulateTicker_Err:
        MsgBox Err.Description & vbCrLf & _
           "in OASISClient.frmMain.EmulateTicker " & _
           "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub oSync_Complete()
    SetMess "Incident Synch Complete..." & " Time: " & GetTimeFormatted
End Sub

Private Sub oSync_CompleteEX(enmRunnerType As OASIS_SynchNG.RunnerType)
        'Dim oItem As OASIS_SynchNG.UpdateAlerts
        '<EhHeader>
        On Error GoTo oSync_CompleteEX_Err
        '</EhHeader>
        Dim Item As Variant
        Dim Col As Collection
        Dim i As Integer
        Dim oRS As ADODB.Recordset

100     Select Case enmRunnerType
    
            Case OASIS_SynchNG.enmPaddy
    
102         Case OASIS_SynchNG.GeoMarks
    
104         Case OASIS_SynchNG.IncidentSynch
106             Set Col = oSync.NewIncidents

                SetMess "Synchronised items recieved." & " Time: " & GetTimeFormatted

108             For Each Item In Col
110                 Set oRS = New ADODB.Recordset
112                 oRS.Open "SELECT ID, Province, District, TYPE, TARGET, Incident_DATE FROM oincidents_FEA WHERE [ID] = '" & Item & "'", m_Cnn, adOpenForwardOnly, adLockReadOnly
114                 g_ColAlerts.Add Item ', CStr(Item)
                    i = i + 1

116                 With oRS.Fields
118                     MsgScroll.AddItem "New Incident #" & i & " In: " & .Item("Province").value & " - " & .Item("District").value & " - " & .Item("Incident_DATE").value, CStr(Item)
                    End With

                    SetMess "Synchronised items recieved. #" & i & " Time: " & GetTimeFormatted
                Next
           
120             If MsgScroll.ListCount > 0 Then
                    If Not mnuOPSView.Checked Then
122                     GIS10.Height = elTatukGIS.Height + 150
                    End If
                    
124                 If Not elScroller.Visible Then
                        
                        SafeMoveFirst g_RSAppSettings
                        g_RSAppSettings.Find "SettingName = 'Notifier'"
            
                        If Not g_RSAppSettings.EOF And Not g_RSAppSettings.Bof Then
                            If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").value) Then
                                MsgScroll.BackColor = g_RSAppSettings.Fields.Item("SettingValue1").value
                                MsgScroll.ForeColor = g_RSAppSettings.Fields.Item("SettingValue2").value
                            End If
                        End If
                        
                        elScroller.Visible = True
126                     elScroller.Move 5, elTatukGIS.Height - MsgScroll.Height, elTatukGIS.Width, 255
elScroller.Move 5, elTatukGIS.Height - MsgScroll.Height, elTatukGIS.Width, 255
                    End If

128                 elScroller.ZOrder
130                 cmdMSGScroller.ZOrder
                    MsgScroll.AutoScroll = True
                Else
132                 elScroller.Visible = False

                    If Not mnuOPSView.Checked Then
134                     GIS10.Height = elTatukGIS.Height + 400
                    End If
                End If
                       
                SetMess "Incident Synch Check Complete..." & " Time: " & GetTimeFormatted
                
136         Case OASIS_SynchNG.InternetConnectionCheck
        
138         Case OASIS_SynchNG.SQLLyrSynch
    
        End Select

        '<EhFooter>
        Exit Sub

oSync_CompleteEX_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.oSync_CompleteEX " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub oSync_StatusEX(ByVal i As Long, ByVal lTotal As Long, Cancel As Boolean, sMess As String)
    DebugPrint i & "  " & lTotal & "  " & sMess
    SetMess "Working # " & i & " of " & lTotal & " " & sMess
End Sub

Private Sub oSync_WorkerError(errnum As Long, _
                              errDesc As String, _
                              errSource As String)
    SetMess "Synch Error..." & errDesc & " " & errSource & " " & errnum & " Time: " & GetTimeFormatted
End Sub

Private Sub SelAttributes1_OnClick(Index As Integer, translated As Boolean)
    DebugPrint ""
End Sub

Private Sub SelAttributes1_OnEnter_(Index As Integer, translated As Boolean)
    DebugPrint ""
End Sub

Private Sub SelAttributes1_OnMouseDown(Index As Integer, _
                                       translated As Boolean, _
                                       ByVal Button As TatukGIS_XDK10.XMouseButton, _
                                       ByVal Shift As TatukGIS_XDK10.XShiftState, _
                                       ByVal x As Long, _
                                       ByVal y As Long)

100     If Button = TatukGIS_XDK10.XmbRight Then
    
            ' Get the position of the cursor
102         GetCursorPos ptPopUpPos
            'abGridPop.Bands("popGrid").PopupMenu , ScaleX(Point.X, vbPixels, vbTwips) - Me.left, ScaleY(Point.Y, vbPixels, vbTwips) - Me.top

106         PopupMenu mnuSelector, x:=ScaleX(ptPopUpPos.x, vbPixels, vbTwips) - Me.left, y:=ScaleY(ptPopUpPos.y, vbPixels, vbTwips) - (Me.top + 250)
  
        End If

End Sub

Private Sub SelAttributes1_OnMouseUp(Index As Integer, translated As Boolean, ByVal Button As TatukGIS_XDK10.XMouseButton, ByVal Shift As TatukGIS_XDK10.XShiftState, ByVal x As Long, ByVal y As Long)
    DebugPrint ""
End Sub

Private Sub SSC_Error()
        '<EhHeader>
        On Error GoTo SSC_Error_Err
        '</EhHeader>
100     MsgBox SSC.Error.Description & SSC.Error.Line & SSC.Error.Text
        '<EhFooter>
        Exit Sub

SSC_Error_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SSC_Error " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub ThreadedCheckInternetConnection(dummyVAr As Variant)
   m_bInternetCon = CheckInternConnection(m_sInternetCon)
End Sub

Private Sub tmrInternetCheck_Timer()

    If g_bOnlineCheckedAtLogin Then
    
        Dim lStart As Long
        lStart = GetTickCount
    
        If Not bDebugMode Then
            If Not pCheckInet.IsThreadRunning Then
                pCheckInet.CreateWin32Thread Me, "ThreadedCheckInternetConnection", 0
            End If

        Else
            m_bInternetCon = CheckInternConnection(m_sInternetCon)
        End If
    
        If m_bInternetCon Then
            SetIMess "Online"
        ElseIf InStr(g_sAppServerPath, "localhost") > 0 Or InStr(g_sAppServerPath, "127.0.0.1") > 0 Then
            SetIMess "Localhost" ' "No Internet Status: " & m_sInternetCon
        Else
            SetIMess "Offline"
        End If
    
        If m_bInternetCon Or InStr(g_sAppServerPath, "localhost") > 0 Or InStr(g_sAppServerPath, "127.0.0.1") > 0 Then
    
            
        If InStr(g_sAppServerPath, "localhost") > 0 Or InStr(g_sAppServerPath, "127.0.0.1") > 0 Then
        m_sInternetCon = "Localhost Available for Synchronisation Through: " & m_sInternetCon
        Else
        m_sInternetCon = "Internet Available for Synchronisation Through: " & m_sInternetCon
        End If
        
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'StartSynchWorker
            '
            If g_lHw < 1 Then
        
                DebugPrint "Started Running Synch... " & " Time: " & GetTimeFormatted
                SetMess "Started Running Synch... " & " Time: " & GetTimeFormatted
    
                g_lHw = GetDesktopWindow()
                
                'ShellExecute g_lHw, vbNullString, App.Path & "\bin\OASISCommsMon.exe", "3^" & m_Cnn.ConnectionString & "^" & g_sRemoteTablePrefix & "^" & g_sAppServerPath & "^" & g_bHasEncrypt & "^" & g_sKey & "^" & Me.hWnd & "^" & tmrInternetCheck.Interval & "^" & g_udtSynchUpdateOptions.GeoMarks & "^" & g_udtSynchUpdateOptions.SynchLayersSettings, "C:\", 1
                If g_bOnlineCheckedAtLogin Then
                    If bSQLServerInUse Then
                        ShellExecute g_lHw, vbNullString, App.Path & "\bin\OASISCommsMon.exe", "3^" & g_sGlobalConnectionString & "^" & g_sRemoteTablePrefix & "^" & g_sAppServerPath & "^" & g_bHasEncrypt & "^" & g_sKey & "^" & Me.hWnd & "^" & tmrInternetCheck.Interval & "^" & g_udtSynchUpdateOptions.GeoMarks & "^" & g_udtSynchUpdateOptions.SynchLayersSettings, "C:\", 1
                    Else
                        ShellExecute g_lHw, vbNullString, App.Path & "\bin\OASISCommsMon.exe", "3^" & m_Cnn.ConnectionString & "^" & g_sRemoteTablePrefix & "^" & g_sAppServerPath & "^" & g_bHasEncrypt & "^" & g_sKey & "^" & Me.hWnd & "^" & tmrInternetCheck.Interval & "^" & g_udtSynchUpdateOptions.GeoMarks & "^" & g_udtSynchUpdateOptions.SynchLayersSettings, "C:\", 1
                    End If
                End If
                
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Else
            m_sInternetCon = "Internet Not Available for Synchronisation: " & m_sInternetCon
        End If

        'Me.caption = "OASIS Client Version: " & App.major & "." & App.minor & "." & App.Revision & " " & App.Comments '& " " & m_sInternetCon
        Me.caption = App.Title & " [" & App.major & "." & App.minor & "." & App.Revision & "] powered by iMMAP"
    
    End If

End Sub



Private Sub tmrUpdate_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
     Dim m_Cursor As POINTAPI
     Dim m_hDC As Long
     Dim lRtn As Long
     Dim x As Long
     Dim y As Long
     Dim lScrHt As Long
     Dim lScrWt As Long
     
    'If Not m_bOnTheMap Then Exit Sub
    
     Cls
     
     lScrHt = Screen.Height \ Screen.TwipsPerPixelY
     lScrWt = Screen.Width \ Screen.TwipsPerPixelX
     ' Get the position of the mouse cursor
     GetCursorPos m_Cursor
     ' update coordinates label
'     lblCoord = "  X = " & m_Cursor.X & "  Y = " & m_Cursor.Y
     ' convert x and y positions into twips and add buffer.  Buffer necessary
     ' to create some space from the mouse cursor so that we don't see the corner
     ' of the zoom box in the magnification.
     ' If we are at the right of the screen
     If m_Cursor.x + (elMagnifier.Width \ Screen.TwipsPerPixelX) + (MOUSE_BUFFER \ Screen.TwipsPerPixelX) > lScrWt And _
          m_Cursor.y + (elMagnifier.Height \ Screen.TwipsPerPixelY) + (MOUSE_BUFFER \ Screen.TwipsPerPixelY) > lScrHt Then
          x = (m_Cursor.x * Screen.TwipsPerPixelX) - (elMagnifier.Width + MOUSE_BUFFER)
          y = (m_Cursor.y * Screen.TwipsPerPixelY) - (elMagnifier.Height + MOUSE_BUFFER)
     ElseIf m_Cursor.x + (elMagnifier.Width \ Screen.TwipsPerPixelX) + _
          (MOUSE_BUFFER \ Screen.TwipsPerPixelX) > lScrWt Then
          x = (m_Cursor.x * Screen.TwipsPerPixelX) - (elMagnifier.Width + MOUSE_BUFFER)
          y = m_Cursor.y * Screen.TwipsPerPixelY + MOUSE_BUFFER
     ElseIf m_Cursor.y + (elMagnifier.Height \ Screen.TwipsPerPixelY) + _
          (MOUSE_BUFFER \ Screen.TwipsPerPixelY) > lScrHt Then
          x = m_Cursor.x * Screen.TwipsPerPixelX + MOUSE_BUFFER
          y = (m_Cursor.y * Screen.TwipsPerPixelY) - (elMagnifier.Height + MOUSE_BUFFER)
     Else
          x = m_Cursor.x * Screen.TwipsPerPixelX + MOUSE_BUFFER
          y = m_Cursor.y * Screen.TwipsPerPixelY + MOUSE_BUFFER
     End If
     ' move the form with the cursor
     'Me.Move X, Y, Me.Width, Me.Height
     ' Get the screen device context
     m_hDC = GetWindowDC(0)
     ' Blit the coordinates, passed in the api call, and stretch it into
     ' our form
     StretchBlt GetWindowDC(elMagnifier.hWnd), 0, 0, ScaleX(elMagnifier.Width, vbTwips, vbPixels), ScaleY(elMagnifier.Height, vbTwips, vbPixels), _
        m_hDC, m_Cursor.x - 24, m_Cursor.y - 24, 48, 48, vbSrcCopy
     ' Draw a box to make the form distinguishable from the background.  Set the forms
     ' forecolor to make changes to it.
     'frmZoom.Line (0, 0)-(frmZoom.ScaleWidth - 1, frmZoom.ScaleHeight - 1), , B
     ' Bring the window to the top.
     'Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
     ' release the screen's device context
     lRtn = ReleaseDC(0, m_hDC)
     ' If at coordinate 0, 0 then quit
     'If m_Cursor.X = 0 And m_Cursor.Y = 0 Then
     '     Unload Me
     '     Set frmZoom = Nothing
     'End If
End Sub ' tmrUpdate_Timer


Private Sub me1Lat_LostFocus()
        '<EhHeader>
        On Error GoTo me1Lat_LostFocus_Err
        '</EhHeader>
100     UpdateLat 1
        On Error Resume Next
102     txtMGRS.Text = ConvPTtoMGRS(CDbl(Replace(me2Long.Text, "_", "")), CDbl(Replace(me2Lat.Text, "_", "")))
104     UpdateDisplayValue
        '<EhFooter>
        Exit Sub

me1Lat_LostFocus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.me1Lat_LostFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtFont_Click(Index As Integer)
    frmOASISFonts.Init txtFont(Index).FontName, txtFont(0).Text
    
    cmdZoomTo.SetFocus
    
    frmOASISFonts.Show vbModal, Me
    
    If Len(frmOASISFonts.Tag) = 1 Then
        txtFont(0).FontName = frmOASISFonts.SymFontName
        txtFont(0).Text = frmOASISFonts.SymFontCharacter
    End If
    
    frmOASISFonts.Tag = "KILL"
    
    Unload frmOASISFonts
    
End Sub

Private Sub txtMGRS_LostFocus()
        '<EhHeader>
        On Error GoTo txtMGRS_LostFocus_Err
        '</EhHeader>
        On Error Resume Next
        Dim x  As Double
        Dim y As Double
    
100     Map1.MilitaryGridReferenceToPoint Replace(txtMGRS.Text, " ", ""), x, y
    
102     me2Long.Text = "0" & CStr(left(x, Len(me2Long.Text) - 1))
104     me2Lat.Text = CStr(left(y, Len(me2Long.Text) - 1))
    
106     me2Lat_LostFocus
108     me2Long_LostFocus
110     UpdateDisplayValue
        '<EhFooter>
        Exit Sub

txtMGRS_LostFocus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.txtMGRS_LostFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub me2Lat_LostFocus()
        '<EhHeader>
        On Error GoTo me2Lat_LostFocus_Err
        '</EhHeader>
100     UpdateLat 2
        On Error Resume Next
102     txtMGRS.Text = ConvPTtoMGRS(CDbl(Replace(me2Long.Text, "_", "")), CDbl(Replace(me2Lat.Text, "_", "")))
104     UpdateDisplayValue
        '<EhFooter>
        Exit Sub

me2Lat_LostFocus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.me2Lat_LostFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub me3Lat_LostFocus()
        '<EhHeader>
        On Error GoTo me3Lat_LostFocus_Err
        '</EhHeader>
100     UpdateLat 3
        On Error Resume Next
102     txtMGRS.Text = ConvPTtoMGRS(CDbl(Replace(me2Long.Text, "_", "")), CDbl(Replace(me2Lat.Text, "_", "")))
104     UpdateDisplayValue
        '<EhFooter>
        Exit Sub

me3Lat_LostFocus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.me3Lat_LostFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub me1Long_LostFocus()
        '<EhHeader>
        On Error GoTo me1Long_LostFocus_Err
        '</EhHeader>
100     UpdateLon 1
        On Error Resume Next
102     txtMGRS.Text = ConvPTtoMGRS(CDbl(Replace(me2Long.Text, "_", "")), CDbl(Replace(me2Lat.Text, "_", "")))
104     UpdateDisplayValue
        '<EhFooter>
        Exit Sub

me1Long_LostFocus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.me1Long_LostFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub me2Long_LostFocus()
        '<EhHeader>
        On Error GoTo me2Long_LostFocus_Err
        '</EhHeader>
100     UpdateLon 2
        On Error Resume Next
102     txtMGRS.Text = ConvPTtoMGRS(CDbl(Replace(me2Long.Text, "_", "")), CDbl(Replace(me2Lat.Text, "_", "")))
104     UpdateDisplayValue
        '<EhFooter>
        Exit Sub

me2Long_LostFocus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.me2Long_LostFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub me3Long_LostFocus()
        '<EhHeader>
        On Error GoTo me3Long_LostFocus_Err
        '</EhHeader>
100     UpdateLon 3
        On Error Resume Next
102     txtMGRS.Text = ConvPTtoMGRS(CDbl(Replace(me2Long.Text, "_", "")), CDbl(Replace(me2Lat.Text, "_", "")))
104     UpdateDisplayValue
        '<EhFooter>
        Exit Sub

me3Long_LostFocus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.me3Long_LostFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ClearCoordValues()
        '<EhHeader>
        On Error GoTo ClearCoordValues_Err
        '</EhHeader>
100     me1Lat.Text = "__:__:__.__"
102     me2Lat.Text = "__._____"
104     me3Lat.Text = "__:__.____"
106     me1Long.Text = "___:__:__.__"
108     me2Long.Text = "___._____"
110     me3Long.Text = "___:__.____"
112     txtMGRS.Text = ""
114     txtConversionResutl.Text = ""
116     comConversionType.ListIndex = 0
        '<EhFooter>
        Exit Sub

ClearCoordValues_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ClearCoordValues " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function ConvPTtoMGRS(x As Double, y As Double) As String
    ConvPTtoMGRS = Map1.MilitaryGridReferenceFromPoint(x, y)
End Function

Private Function ConvMGRStoPT(sMGRS As String, x As Double, y As Double) As Boolean
    On Error Resume Next
    ConvMGRStoPT = Map1.MilitaryGridReferenceToPoint(sMGRS, x, y)
End Function

Private Sub UpdateDisplayValue()
        '<EhHeader>
        On Error GoTo UpdateDisplayValue_Err
        '</EhHeader>

100     With txtConversionResutl
102         .Text = ""
104         .Text = .Text & "Lat:" & me1Lat.Text & vbCrLf
106         .Text = .Text & "Lon:" & me1Long.Text & vbCrLf & "-------------" & vbCrLf
108         .Text = .Text & "Lat:" & me2Lat.Text & vbCrLf
110         .Text = .Text & "Lon:" & me2Long.Text & vbCrLf & "-------------" & vbCrLf
112         .Text = .Text & "Lat:" & me3Lat.Text & vbCrLf
114         .Text = .Text & "Lon:" & me3Long.Text & vbCrLf & "-------------" & vbCrLf
116         .Text = .Text & "MGRS:" & txtMGRS.Text
        End With
        '<EhFooter>
        Exit Sub

UpdateDisplayValue_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.UpdateDisplayValue " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub UpdateLat(FrameNumber As Integer)
        '<EhHeader>
        On Error GoTo UpdateLat_Err
        '</EhHeader>
        Dim de As Single, mi As Single, se As Single, lT As String
        Dim Temp1 As Single, temp2 As Single, temp3 As Single
    
100     Select Case FrameNumber

            Case 1 ' dd:mm:ss.ss
102             lT = Replace(me1Lat.Text, "_", "0")
104             de = Val(left$(lT, 2))
106             mi = Val(Mid$(lT, 4, 2))
108             se = Val(right$(lT, 5))

110             If de > 59 Then
112                 MsgBox "Degrees must be from 0 to 59", , "Entry Error"
114                 me1Lat.SetFocus
                    Exit Sub
                End If
    
116             If mi > 59 Then
118                 MsgBox "Minutes must be from 0 to 59", , "Entry Error"
120                 me1Lat.SetFocus
                    Exit Sub
                End If
            
122             If se > 59.99 Then
124                 MsgBox "Seconds must be from 0.00 to 59.99", , "Entry Error"
126                 me1Lat.SetFocus
                    Exit Sub
                End If
    
128         Case 2 ' dd.ddddd
130             lT = Replace(me2Lat.Text, "_", "0")
132             de = Val(lT)
134             mi = 0
136             se = 0
            
138             If de > 59.99999! Then
140                 MsgBox "Degrees must be from 0 to 59.99999", , "Entry Error"
142                 me2Lat.SetFocus
                    Exit Sub
                End If
        
144         Case 3 ' dd:mm.mmmm
146             lT = Replace(me3Lat.Text, "_", "0")

148             de = Val(left$(lT, 2))
150             mi = Val(right$(lT, 7))
152             se = 0

154             If de > 59 Then
156                 MsgBox "Degrees must be from 0 to 59", , "Entry Error"
158                 me3Lat.SetFocus
                    Exit Sub
                End If
    
160             If mi > 59.9999! Then
162                 MsgBox "Minutes must be from 0 to 59.9999", , "Entry Error"
164                 me3Lat.SetFocus
                    Exit Sub
                End If

        End Select

166     Select Case FrameNumber

            Case 1 ' from dd:mm:ss.ss

                '   to dd:mm.mmmm
                '
168             Temp1 = (mi * 60) + se
170             temp2 = (Temp1 / 60)

172             If temp2 + de <> 0 Then me3Lat.Text = Format$(de, "00") & ":" & Format$(temp2, "00.0000")

                '   to dd.ddddd
                '
174             temp3 = de + (temp2 / 60)

176             If temp3 <> 0 Then me2Lat.Text = Format$(temp3, "00.00000")
        
178         Case 2 ' from dd.ddddd

                '   to dd:mm.mmmm
                '
180             Temp1 = (de - Int(de)) * 60

182             If Temp1 + Int(de) <> 0 Then me3Lat.Text = Format$(Int(de), "00") & ":" & Format$(Temp1, "00.0000")

                '   to dd:mm:ss.ss
                '
184             temp2 = Int(Temp1)
186             temp3 = (Temp1 - Int(Temp1)) * 60

188             If temp3 + Int(de) <> 0 Then me1Lat.Text = Format$(Int(de), "00") & ":" & Format$(temp2, "00") & ":" & Format$(temp3, "00.00")

190         Case 3 ' dd:mm.mmmm
        
                '   to dd.ddddd
                '
192             Temp1 = mi / 60
194             temp2 = Temp1 + de

196             If temp2 <> 0 Then me2Lat.Text = Format$(temp2, "00.00000")
            
                '   to dd:mm:ss.ss
                '
198             Temp1 = (mi - Int(mi)) * 60

200             If Temp1 + de <> 0 Then me1Lat.Text = Format$(de, "00") & ":" & Format$(Int(mi), "00") & ":" & Format$(Temp1, "00.00")

        End Select

        '<EhFooter>
        Exit Sub

UpdateLat_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.UpdateLat " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub UpdateLon(FrameNumber As Integer)
        '<EhHeader>
        On Error GoTo UpdateLon_Err
        '</EhHeader>
        Dim de As Single, mi As Single, se As Single, lT As String
        Dim Temp1 As Single, temp2 As Single, temp3 As Single
    
100     Select Case FrameNumber
            Case 1 ' ddd:mm:ss.ss
102             lT = Replace(me1Long.Text, "_", "0")
104             de = Val(left$(lT, 3))
106             mi = Val(Mid$(lT, 5, 2))
108             se = Val(right$(lT, 5))

110             If de > 179 Then
112                 MsgBox "Degrees must be from 0 to 179", , "Entry Error"
114                 me1Long.SetFocus
                    Exit Sub
                End If
    
116             If mi > 59 Then
118                 MsgBox "Minutes must be from 0 to 59", , "Entry Error"
120                 me1Long.SetFocus
                    Exit Sub
                End If
            
122             If se > 59.99 Then
124                 MsgBox "Seconds must be from 0.00 to 59.99", , "Entry Error"
126                 me1Long.SetFocus
                    Exit Sub
                End If
    
128         Case 2 ' ddd.ddddd
130             lT = Replace(me2Long.Text, "_", "0")
132             de = Val(lT)
134             mi = 0
136             se = 0
            
138             If de > 180! Then
140                 MsgBox "Degrees must be from 0 to 179.99999", , "Entry Error"
142                 me2Long.SetFocus
                    Exit Sub
                End If
        
144         Case 3 ' ddd:mm.mmmm
146             lT = Replace(me3Long.Text, "_", "0")

148             de = Val(left$(lT, 3))
150             mi = Val(right$(lT, 7))
152             se = 0

154             If de > 179 Then
156                 MsgBox "Degrees must be from 0 to 179", , "Entry Error"
158                 me3Long.SetFocus
                    Exit Sub
                End If
    
160             If mi > 59.9999! Then
162                 MsgBox "Minutes must be from 0 to 59.9999", , "Entry Error"
164                 me3Long.SetFocus
                    Exit Sub
                End If

        End Select

166     Select Case FrameNumber
            Case 1 ' from dd:mm:ss.ss

                '   to ddd:mm.mmmm
                '
168             Temp1 = (mi * 60) + se
170             temp2 = (Temp1 / 60)
172             If temp2 + de <> 0 Then me3Long.Text = Format$(de, "000") & ":" & Format$(temp2, "00.0000")

                '   to ddd.ddddd
                '
174             temp3 = de + (temp2 / 60)
176             If temp3 <> 0 Then me2Long.Text = Format$(temp3, "000.00000")
        
178         Case 2 ' from ddd.ddddd

                '   to ddd:mm.mmmm
                '
180             Temp1 = (de - Int(de)) * 60
182             If Temp1 + Int(de) <> 0 Then me3Long.Text = Format$(Int(de), "000") & ":" & Format$(Temp1, "00.0000")

                '   to ddd:mm:ss.ss
                '
184             temp2 = Int(Temp1)
186             temp3 = (Temp1 - Int(Temp1)) * 60
188             If temp3 + Int(de) <> 0 Then me1Long.Text = Format$(Int(de), "000") & ":" & Format$(temp2, "00") & ":" & Format$(temp3, "00.00")

190         Case 3 ' ddd:mm.mmmm
        
                '   to ddd.ddddd
                '
192             Temp1 = mi / 60
194             temp2 = Temp1 + de
196             If temp2 <> 0 Then me2Long.Text = Format$(temp2, "000.00000")
            
                '   to ddd:mm:ss.ss
                '
198             Temp1 = (mi - Int(mi)) * 60
200             If Temp1 + de <> 0 Then me1Long.Text = Format$(de, "000") & ":" & Format$(Int(mi), "00") & ":" & Format$(Temp1, "00.00")

        End Select

    

        '<EhFooter>
        Exit Sub

UpdateLon_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.UpdateLon " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckAvailableTools()
                'AvailableMapTools
        '<EhHeader>
        On Error GoTo CheckAvailableTools_Err
        '</EhHeader>
100             g_RSAppSettings.Requery
                
102             SafeMoveFirst g_RSAppSettings
104             g_RSAppSettings.Find "SettingName = 'AvailableMapTools'"

                Dim sTools() As String

                If IsNull(g_RSAppSettings.Fields.Item("SettingValue1").value) Then Exit Sub

106             sTools = Split(g_RSAppSettings.Fields.Item("SettingValue1").value, ",")


                Dim tl As ActiveBar3LibraryCtl.Tool
                
                'Activate Again In version 2.5
                AB.Bands.Item("tbLayer").Tools.Remove "btnPrint"
                         
108             For Each tl In AB.Bands.Item("tbLayer").Tools
110                 If InStr(g_RSAppSettings.Fields.Item("SettingValue1").value, tl.Name) Then
112                     'DebugPrint "add: " & tl.Name
                    Else
114                     'DebugPrint "remove: " & tl.Name
116                     AB.Bands.Item("tbLayer").Tools.Remove tl.Name
                    End If
                Next

118             For Each tl In AB.Bands.Item("tbExtents").Tools
120                 If InStr(g_RSAppSettings.Fields.Item("SettingValue2").value, tl.Name) Then
122                     DebugPrint "add: " & tl.Name
                    Else
124                     DebugPrint "remove: " & tl.Name
126                     AB.Bands.Item("tbExtents").Tools.Remove tl.Name
                    End If
                Next

128             For Each tl In AB.Bands.Item("sdToolBar").Tools
130                 If InStr(g_RSAppSettings.Fields.Item("SettingValue3").value, tl.Name) Then
132                     'DebugPrint "add: " & tl.Name
                    Else
134                     'DebugPrint "remove: " & tl.Name
136                     AB.Bands.Item("sdToolBar").Tools.Remove tl.Name
                    End If
                Next
                
138             For Each tl In AB.Bands.Item("tbUtils").Tools
140                 If InStr(g_RSAppSettings.Fields.Item("SettingValue4").value, tl.Name) Then
142                     'DebugPrint "add: " & tl.Name
                    Else
144                     'DebugPrint "remove: " & tl.Name
146                     AB.Bands.Item("tbUtils").Tools.Remove tl.Name
                    End If
                Next

148             AB.Refresh
       
        '<EhFooter>
        Exit Sub

CheckAvailableTools_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CheckAvailableTools " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub RealaignTB()
        '<EhHeader>
        On Error GoTo RealaignTB_Err
        '</EhHeader>
Dim iCurrOffset As Integer

100        With AB.Bands
102             .Item("tbrHelp").DockingOffset = Me.Width - 600
104             If .Item("sdToolBar").Visible Then
        
106                 .Item("sdToolBar").DockingOffset = 3995
108                 iCurrOffset = .Item("sdToolBar").Width + .Item("sdToolBar").DockingOffset
            
110                 If .Item("tbExtents").Visible Then
112                     .Item("tbExtents").DockingOffset = iCurrOffset + 23
114                     iCurrOffset = .Item("tbExtents").Width + .Item("tbExtents").DockingOffset
                    End If

116                 If .Item("tbLayer").Visible Then
118                     .Item("tbLayer").DockingOffset = iCurrOffset + 23
120                     iCurrOffset = .Item("tbLayer").Width + .Item("tbLayer").DockingOffset
                    End If
                
122                 If .Item("tbUtils").Visible Then
124                     .Item("tbUtils").DockingOffset = iCurrOffset + 23
126                     iCurrOffset = .Item("tbUtils").Width + .Item("tbUtils").DockingOffset
                    End If
                End If
        
            End With
 
        '<EhFooter>
        Exit Sub

RealaignTB_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.RealaignTB " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ActivateOperations()
        '<EhHeader>
        On Error GoTo ActivateOperations_Err
        '</EhHeader>
        Dim i As Integer
        Dim iCurrOffset As Integer
        
100     AB.ClientAreaControl = elMap
102     elMap.Visible = True
            
        'Tool Panels
104     SafeMoveFirst g_RSAppSettings
106     g_RSAppSettings.Find "SettingName = 'MainMapTools'"
        
108     AB.Bands.Item("sdToolBar").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue1").value) = 1, True, False)
110     AB.Bands.Item("tbExtents").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue2").value) = 1, True, False)
112     AB.Bands.Item("tbLayer").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue3").value) = 1, True, False)
        '114     AB.Bands.Item("bToolbarStyle").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue4").Value) = 1, True, False)
116     AB.Bands.Item("tbUtils").Visible = IIf(Trim$(g_RSAppSettings.Fields("SettingValue5").value) = "1", True, False)
        AB.Bands.Item("tbrHelp").Visible = True
        AB.Bands.Item("tbrHelp").DockingOffset = Me.Width - 600
        
        If Not IsNull(g_RSAppSettings.Fields("SettingValue10").value) Then
            AB.Bands.Item("tbrHelp").Tools.Item("btnWebHelp").Text = g_RSAppSettings.Fields("SettingValue10").value
        Else
            AB.Bands.Item("tbrHelp").Tools.Item("btnWebHelp").Text = "http://www.immap.org"
        End If
        
        RealaignTB
        
118     If Not m_bMapInitialized Then
            'If Not bDebugMode Then
            '    If Not pGetIncThread.IsThreadRunning Then
            '        pGetIncThread.CreateWin32Thread Me, "GetNewIncidents", 1
            '    End If
            'End If
                    
120         If m_bInternetCon And g_bOnlineCheckedAtLogin Then

122             With oSync
            
124                 If Not .IsRunning Then
        
126                     .Paddy = False
128                     .RType = IncidentSynch
130                     .LocalConnectionString = m_Cnn.ConnectionString
132                     .RemoteTablePrefix = g_sRemoteTablePrefix
134                     .WebsiteURL = g_sAppServerPath
136                     .HasEncrypt = g_bHasEncrypt
138                     .EncryptKey = g_sKey
140                     .Start

                    End If

                End With

            End If
                    
            'GetNewIncidents 1
            InitMapFromMapLibraryDefault
142         'InitMap
        End If

        'AB.RecalcLayout

        '<EhFooter>
        Exit Sub

ActivateOperations_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.ActivateOperations " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub AB_ChildBandChange(ByVal Band As ActiveBar3LibraryCtl.Band)
        '<EhHeader>
        On Error GoTo AB_ChildBandChange_Err
        '</EhHeader>
    
        Dim i As Integer
    
100     AB.Bands.Item("bMenu").Visible = False
102     AB.Bands.Item("sdToolBar").Visible = False
104     AB.Bands.Item("tbExtents").Visible = False
106     AB.Bands.Item("tbLayer").Visible = False
108     AB.Bands.Item("tbUtils").Visible = False
        AB.Bands.Item("tbrHelp").Visible = False

    
110     elRSSTool.Visible = False
112     elMap.Visible = False
114     elOasisProfile.Visible = False
118     elAddons.Visible = False
120     elDynamicData.Visible = False
122     elDynamicReports.Visible = False
    
124     Select Case Band.Name

            Case "cbFolderList", "cbShortcuts"

                'AB.ClientAreaControl = GIS
                'MailControl.Visible = True
                'CalendarControl.Visible = False
126         Case "cbOperations"
        
128             If Not m_bLOADING Then ActivateOperations

130         Case "cbProfile"

                If Not m_bAlreadyinitedProfile Then
132                 SafeMoveFirst g_RSAppSettings
134                 g_RSAppSettings.Find "SettingName = 'InitURL'"

136                 If Not g_RSAppSettings.EOF Then
138                     If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").value) Then
140                         WebBrowser1.Navigate2 g_RSAppSettings.Fields.Item("SettingValue1").value
                        End If
                    End If

                    m_bAlreadyinitedProfile = True
                End If
                
                AB.ClientAreaControl = elOasisProfile
                elOasisProfile.Visible = True
                
                'TODO Remove Below
                
142             If lLOG = 0 Then
144                 'lLOG = ShellExecute(Me.hWnd, vbNullString, "C:\OASIS\Client\Common_Pipeline_DB_v3.04.03\Common_Pipeline_DB.mdb", vbNullString, "C:\", 1)
                
146                 'AB.ClientAreaControl = elOasisProfile
148                 'elOasisProfile.Visible = True
                    'TODO Remove Below
                    'If lLOG > 0 Then
                    '    SetParent lLOG, elOasisProfile.Hwnd
                    'End If
                End If
                
150         Case "cbContent"
152             AB.ClientAreaControl = elRSSTool
154             elRSSTool.Visible = True
                
                m_frmDynamicContent.Init g_bFeedUpdate
                '   m_frmDynamicContent.Visible = True
                '   DebugPrint m_frmDynamicContent.Visible
                
156             If Len(RSSBrowser1.ServerURL) = 0 Then
                    RSSBrowser1.Visible = True
158                 RSSBrowser1.ServerURL = g_sAppServerPath & "/oasis4.asp"
160                 RSSBrowser1.UserGroupPrefix = g_sRemoteTablePrefix
162                 RSSBrowser1.Init2 m_Cnn.ConnectionString, g_bFeedUpdate
                End If
                               
164         Case "cbDynamicData"
166             AB.ClientAreaControl = elDynamicData
168             elDynamicData.Visible = True
             
170             If Not DynamicDataModule1.IsEditing Then
172                 DynamicDataModule1.Init m_frmMnuDynamicDataModule.listDatabases, m_frmMnuDynamicDataModule.lstDataElements
        
                End If
                
174         Case "cbReports"
176             AB.ClientAreaControl = elDynamicReports
178             elDynamicReports.Visible = True
             
180             If Not m_frmMnuDynamicReportsModule.Combo1.ListCount > 0 Then
             
182                 DynamicDataReports1.Init m_frmMnuDynamicReportsModule.listQueries, m_frmMnuDynamicReportsModule.Combo1, m_frmMnuDynamicReportsModule.ComFilter, m_frmMnuDynamicReportsModule.listGroup
         
                End If
             
184         Case "cbW3"

186         Case "cbSync", "cbJournal"
        
188             MsgBox "This functionality is currently not available for your profile." & vbCrLf & "Please contact your OASIS system administrator if you want to activate these modules.", vbInformation, "OASIS Client Support"

190         Case "cbAddOns"
192             AB.ClientAreaControl = elAddons
                '  WebBrowser2.Navigate2 "http://www.immap.org"
194             elAddons.Visible = True
        End Select
        
196     MsgScroll.ZOrder
        
198     AB.RecalcLayout
        '<EhFooter>
        Exit Sub

AB_ChildBandChange_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.AB_ChildBandChange " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadMenuFrames()
        '<EhHeader>
        On Error GoTo LoadMenuFrames_Err
        '</EhHeader>

100     Set m_frmMnuOperations = New frmMnuOperations

        Set m_frmMnuDynamicDataModule = New frmMnuDynamicDataModule
        Set m_frmMnuDynamicReportsModule = New frmMnuDynamicReportsModule

106     Set m_frmMnuOASISProfile = New frmMnuOASISProfile
108     Set m_frmAddOns = New frmAddons
110     Set m_frmChangeTracer = New frmChangeTracer
112     Set m_frmFreqSettings = New frmFreqSettings
        Set m_frmDynamicContent = New frmDynamicContent
        '<EhFooter>
        Exit Sub

LoadMenuFrames_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.LoadMenuFrames " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetMenuSystem()
        '<EhHeader>
        On Error GoTo SetMenuSystem_Err
        '</EhHeader>

100     LoadMenuFrames

        Dim common As New AB3Common
102     common.Init AB
    
        'NavigationBar
104     Set m_navBar = New NavigationBarExtension
106     m_navBar.Init AB, common
108     m_navBar.Extend
    
110     elRSSTool.BorderWidth = 0
112     elOasisProfile.BorderWidth = 0
114     elMap.BorderWidth = 0
    
        'Dockable Forms
116     AB.Bands("bNavPane").ChildBands("cbProfile").Tools("frmMail").Custom = m_frmMnuOASISProfile
118     AB.Bands("bNavPane").ChildBands("cbOperations").Tools("frmCalendar").Custom = m_frmMnuOperations
120     AB.Bands("bNavPane").ChildBands("cbContent").Tools("frmContacts").Custom = m_frmDynamicContent
        'AB.Bands("bNavPane").ChildBands("cbSync").Tools("frmTasks").Custom = frmAddIncident
    AB.Bands("bNavPane").ChildBands("cbDynamicData").Tools("frmDynamicData").Custom = m_frmMnuDynamicDataModule
    AB.Bands("bNavPane").ChildBands("cbReports").Tools("frmReports").Custom = m_frmMnuDynamicReportsModule
    
122     'AB.Bands("bNavPane").ChildBands("cbW3").Tools("frmW3").Custom = m_frmW3Wizard
    
124     AB.Bands("bNavPane").ChildBands("cbAddOns").Tools("frmNotes").Custom = m_frmAddOns
   
126     AB.Bands("bNavPane").ChildBands.CurrentChildBand = AB.Bands("bNavPane").ChildBands("cbOperations")
    
        'ToolbarStyle Combo
128     Set m_styleCombo = New ToolbarStyleABCombo
130     m_styleCombo.Init AB, common
132     Set m_styleCombo.Combo = AB.Bands("bToolbarStyle").Tools("tToolbarStyle")
134     m_styleCombo.Extend
        '<EhFooter>
        Exit Sub

SetMenuSystem_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SetMenuSystem " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub OpenSpatialAnalysis()
        '<EhHeader>
        On Error GoTo OpenSpatialAnalysis_Err
        '</EhHeader>
 
100     Set m_frmSpatialAnalysis = New frmSpatialAnalysis
102     m_frmSpatialAnalysis.SetGISViewer GIS10.viewer
104     m_frmSpatialAnalysis.InitForm
106     m_frmSpatialAnalysis.Show vbModeless, Me

        '<EhFooter>
        Exit Sub

OpenSpatialAnalysis_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.OpenSpatialAnalysis " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub OpenDynamicDataEntry()

End Sub

Private Sub AB_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)
        '<EhHeader>
        On Error GoTo AB_ToolClick_Err
        '</EhHeader>
        Dim xportLyr As Object
        Dim c As New cCommonDialog
        Dim udtOASISChartO As OASISChartObj
        Dim sAdoString As String
        Dim sLayerName As String
        Dim sDelyr As String
        Dim oSQLLyr As New TatukGIS_XDK10.XGIS_LayerSqlAdo
        Dim oRS As ADODB.Recordset
        Dim oExtent As TatukGIS_XDK10.XGIS_Extent
        Dim oLyr As New TatukGIS_XDK10.XGIS_LayerVector
        Dim oAbsLayer As XGIS_LayerAbstract
        'GIS109 Todo
100     m_bPolyG = False
102     m_bPolyL = False

104     If m_SelUID > 0 Then
106         m_oDrawLyr.Delete m_SelUID
108         m_SelUID = 0
        End If
        
110     g_CurrentTool = oZoom

112     Select Case Tool.Name

            Case Is = "btnSearchLayers"

114             If m_frmSearch Is Nothing Then
116                 Set m_frmSearch = New frmSearch
                End If
                
118             m_frmSearch.Init GIS10.viewer
120             m_frmSearch.Show vbModeless, Me

122         Case Is = "btnAddAnnotation"

124             If m_frmTextAnnoSettings Is Nothing Then Set m_frmTextAnnoSettings = New frmTextAnnoSettings
                
126             ListAllAnnotationShps
128             m_frmTextAnnoSettings.Show vbModeless, Me
130             g_CurrentTool = oCreateLocationText
132             GIS10.Mode = TatukGIS_XDK10.XgisSelect

134         Case Is = "btnCharting"
            
136             frmOASISChartOCTfiles.Show vbModeless, Me
                '104             With udtOASISChartO
                'MsgBox "A slight hickup occured... OASIS Devteam is working on this...", vbInformation, "OASIS Beta Hickups"
   
                '106              With .udtChartTemplate
                '108                  .enmFormat = tplBin
                '110                  .sDecription = "Some Potato Junkie"
                '112                  .sName = "C:\OASIS\Client\data\db\demo.oct"
                '                    End With
                '
                '                End With
                '
                '114             With udtOASISChartO
                '116                 .bAnnoTBR = True
                '118                 .bChartTBR = True
                '                End With
                '
                '120             m_frmOASISCharts.SetChart udtOASISChartO
                '122             m_frmOASISCharts.Show vbModal, Me
138         Case Is = "btnAdminLocator"
140             m_frmLocator.Show vbModeless, Me
142             m_frmLocator.Init GIS10.viewer

144         Case Is = "mnuResetMap"
146             GIS10.FullExtent

148         Case Is = "btnZoomin"

                If ctlZoomSlider1.GetZoom > 1 Then
                    ctlZoomSlider1.ZoomIn
150             Else
                    GIS10.zoom = GIS10.zoom * 2
                End If

152         Case Is = "btnZoomout"

                If ctlZoomSlider1.GetZoom < 19 Then
                    ctlZoomSlider1.ZoomOut
154             Else
                    GIS10.zoom = GIS10.zoom / 2
                End If

156         Case Is = "btnZoom"
158             GIS10.Mode = TatukGIS_XDK10.XgisZoomEx
                ctlSelector1.RenewSelection

160         Case Is = "btnZoomRect"
162             GIS10.Mode = TatukGIS_XDK10.XgisZoom
                ctlSelector1.RenewSelection

164         Case Is = "btnPan"
166             GIS10.Mode = TatukGIS_XDK10.XgisDrag
                ctlSelector1.RenewSelection

168         Case Is = "btnSelect"
170             GIS10.Mode = TatukGIS_XDK10.XgisSelect
                ctlSelector1.RenewSelection

172         Case Is = "mnuAddLayer"
174             AddLayer

176         Case "mnuFile"

178         Case "mnuEdit"

180         Case "mnuView"

182         Case "mnuTools"

184         Case "mnuHelp"

186         Case "mnuMaps"

188         Case "mnuabout"
            
190         Case "btnCannedReports"
                'Case "btnDynReports"
192             Set m_frmCannedReports = New frmCannedReports
194             m_frmCannedReports.Init True
196             m_frmCannedReports.Show vbModal, Me
198             Set m_frmCannedReports = Nothing

200         Case "btnDynamicDataEntry"
                '                Dim rsRemote As ADODB.Recordset
                '                Dim RS As ADODB.Recordset
                '
                '178             If g_bUpdateDynamicDataDefs Then
                '180                 If m_bInternetCon Then
                '182                     Set rsRemote = New ADODB.Recordset
                '184                     DynamDataDefsUpdate m_Cnn, rsRemote, RS
                '186                     g_bUpdateDynamicDataDefs = False
                '                    End If
                '                End If
                '
                '188             Call OpenDynamicDataEntry

202         Case "btnSpatialAnalysis"
204             Call OpenSpatialAnalysis

206         Case "btnOASISv1Charts"
                'MsgBox "to be implemented"
208             frmChartSettings.Show vbModeless, Me

210         Case "mnuCoorpWebSite"
212
                ShellExecute Me.hWnd, vbNullString, "mailto:support@immap.org", vbNullString, "C:\", SW_SHOWNORMAL

            Case "btnWebHelp"
                ShellExecute Me.hWnd, vbNullString, Tool.Text, vbNullString, "C:\", SW_SHOWNORMAL

214         Case "mnuOASISWebSite"

            Case "btnMapPrinttemplate"
    
               ' frmMainPrint.Show vbModeless, Me
                Dim sTmpMap As String
                sTmpMap = g_sAppPath & "\data\user\Maps\" & GUIDGen & ".TTKGP"
                
                GIS10.SaveProjectAs sTmpMap
                
                With GIS10.VisibleExtent
                    frmMainPrint.InitPrint sTmpMap, .xmin, .xmax, .ymin, .ymax
                End With
    
                frmMainPrint.Show vbModal, Me
   
                Unload frmMainPrint
                
                On Error Resume Next
                
                Kill sTmpMap

216         Case "mnuHelpGo"

218         Case "mnuOnlineSupport"

220         Case "mnuRemoveLayer"

222         Case "mnuSaveSession"

224         Case "mnuSaveMap"

226         Case "mnuOpenMap"

228         Case "mnuOpenSession"

230         Case "btnZoom"

232         Case "btnSelect"

234         Case "btnDeselect"

236         Case "btnAddLyr"
238             AddLayer
240             FillCOPValues

244         Case "btnRemoveLyr"
246             frmLayerSelection.Init GIS10, True
248             frmLayerSelection.Show vbModal, Me

250             If frmLayerSelection.GetItem <> "" Then
252                 If MsgBox("Do you really want to remove map layer: " & frmLayerSelection.GetItem, vbYesNo, "OASIS Client") = vbYes Then

                        sDelyr = frmLayerSelection.GetItem
254                     GIS10.Delete sDelyr

                        Set oAbsLayer = GIS10.get(sDelyr)
                        
                        If oAbsLayer Is Nothing Then
                            
                            abGridTools.Tools.Item("comLyr").CBRemoveItem abGridTools.Tools.Item("comLyr").CBListIndex
1108                        dxGISDataGrid.Visible = False
1110                        dxGISDataGrid.Columns.DestroyColumns
1112                        Set dxGISDataGrid.DataSource = Nothing
1114                        dxGISDataGrid.Visible = True
                            '-Nothing
                        End If
                        
256                     FillCOPValues
260                     GIS10.UpDate
                    End If
                End If

262         Case "btnLyrSettings"
                
264         Case "btnLegend"
                
                '            abCOP.Bands("bnMapLegend").Visible = Not abCOP.Bands("bnMapLegend").Visible
                ''            abCOP.Bands("bnMapLegend").Tools("frmLegend").Visible = Not abCOP.Bands("bnMapLegend").Tools("frmLegend").Visible
                '            frmLegend.Update
                
266         Case "btnInfo"

268             If m_frmAttributes Is Nothing Then
270                 Set m_frmAttributes = New frmAttributes
                End If
                
272             m_frmAttributes.Init GIS10.viewer
274             m_frmAttributes.Show vbModeless, Me
276             g_CurrentTool = oInfo
278             GIS10.Mode = TatukGIS_XDK10.XgisSelect
                ctlSelector1.RenewSelection

280         Case "btnFullExtent"

282             SafeMoveFirst g_RSAppSettings
284             g_RSAppSettings.Find "[SettingName] = 'PCodeAdminLevel0'"

286             Set oRS = New ADODB.Recordset
288             oRS.Open "SELECT * FROM [ttkGISLayerSQLInProject] WHERE [LayerCaption] = '" & g_RSAppSettings.Fields("SettingValue1").value & "'", m_Cnn, adOpenDynamic, adLockBatchOptimistic
                
290             If Not oRS.State = 0 Then
                
292                 If Not oRS.EOF Then
                        
294                     Set oExtent = New TatukGIS_XDK10.XGIS_Extent

296                     With oRS
                            
298                         oExtent.Prepare .Fields("XMIN").value, .Fields("YMIN").value, .Fields("XMAX").value, .Fields("YMAX").value
300                         GIS10.VisibleExtent = oExtent
302                         Set oExtent = Nothing
                        
                        End With
                    
                    Else
304                     GIS10.FullExtent
                    End If
                    
306                 oRS.Close
                
                Else
                
308                 GIS10.FullExtent
                End If
                
310             Set oRS = Nothing

312         Case "btnLayerExtent"
314             frmLayerSelection.Init GIS10, True
316             frmLayerSelection.Show vbModal, Me
                
318             If Not frmLayerSelection.GetItem = "" Then
320                 GIS10.VisibleExtent = GIS10.get((frmLayerSelection.GetItem)).Extent
                End If
                
322             frmLayerSelection.ClearItem

324         Case "btnSelectionExtent"

326         Case "btnAddClippBoard"
328             GIS10.viewer.PrintClipboard

330         Case "btnPrevExtent"
332             SetMapExtent g_PrevExt

334         Case "btnRefreshMap"
336             GIS10.UpDate

338         Case "btnPrint"
                
340             frmMapPrint.Init GIS10.viewer, m_frmMnuOperations.Legend1.Legend, g_bMapTplUpdate, g_sRemoteTablePrefix
                'frmMapPrint.Show vbModal, Me
342             frmMapPrint.Show vbModeless, Me
        
344         Case "btnOpenMap"

346             With c
348                 .DialogTitle = "Open Map Definition File"
                    '.CancelError = True
350                 .hWnd = Me.hWnd
352                 .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
354                 .InitDir = g_sAppPath & "\data\user\Maps\"
356                 .Filter = "Map Definition files (*.TTKGP)|*.TTKGP"
358                 .FilterIndex = 1
360                 .ShowOpen
        
                    ctlSelector1.RenewSelection
                End With
                
362             If Len(c.Filename) > 0 Then
364                 InitMap c.Filename
366                 LoadLayerAttrDataToGridInit
                End If
                
368         Case "btnRepairDB"
                Dim sConnString As String
                'TODO m_cnn Must be Closed and opened!!!
370             sConnString = m_Cnn.ConnectionString
372             m_Cnn.Close
374             Set m_Cnn = Nothing
        
376             FixDB True
        
378             Set m_Cnn = New ADODB.Connection
380             m_Cnn.ConnectionString = sConnString
382             m_Cnn.Open
        
384         Case "btnCreateDBLyr"
        
386             frmLayerSelection.Init GIS10
        
388             frmLayerSelection.Show vbModal, Me

390             If frmLayerSelection.GetItem <> "" Then
                    
392                 Set oLyr = GIS10.get((frmLayerSelection.GetItem))

394                 With c
                
396                     .Filter = "Microsoft Access (*.mdb)|*.mdb"

398                     .DialogTitle = "Create OASIS SQL Layers"
400                     .InitDir = g_sAppPath & "\data\db\"
402                     .ShowOpen
                    End With
                    
404                 If Len(c.Filename) > 0 Then
406                     Set xportLyr = New TatukGIS_XDK10.XGIS_LayerSqlAdo
                
408                     sLayerName = InputBox("Insert the name of the layer", "OASIS GIS SQL Layers", "My Layer")
                
410                     If sLayerName = "" Then
412                         MsgBox "It seems like you have not entered a proper name of the table to be created. please try again!", vbInformation, "OASIS Data Creator"
414                         frmLayerSelection.ClearItem
                            Exit Sub
                        End If
                        
416                     If MsgBox("Is this the OASIS database?", vbYesNo, "Confirm if this is the OASIS database") = vbYes Then
                        
418                         sAdoString = GetConnectionString(c.Filename)
                        Else
                        
420                         sAdoString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & c.Filename & ";" 'Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                        End If
                
422                     sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = Native" & vbCrLf & "layer=" & sLayerName & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString
                    
424                     xportLyr.Path = sAdoString
426                     oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
428                     xportLyr.SaveAll
                
                        Dim mFileSysObj As New FileSystemObject
                        Dim sPath As String
                
430                     sPath = Mid$(c.Filename, 1, InStrRev(c.Filename, "\")) & sLayerName & ".ttkls"
                
432                     mFileSysObj.CreateTextFile sPath, True
            
434                     Open sPath For Output As #1
436                     Print #1, sAdoString
438                     Close #1

                        '                        Dim iwidth As Integer
                        '                        Dim oPixLayer As New TatukGIS_XDK10.XGIS_LayerPixel
                        '                        Set oPixLayer = GIS10.Get("1m_herat_geo")
                        '                        iwidth = Round((GIS10.Extent.XMax - GIS10.Extent.XMin) / (oPixLayer.Extent.XMax - oPixLayer.Extent.XMin) * oPixLayer.BitWidth)
                        '                        GIS10.ExportToImage "C:\OASIS\Client\data\db\PixelStore2_test.ttkps", GIS10.Extent, iwidth, 0
                    
                    End If
                End If
                
440             frmLayerSelection.ClearItem
                
                'm_oIncidentLyr.ExportLayer
442         Case "btnExportMapDefFile"

444             With c
                
446                 .Filter = "Map Definition files (*.TTKGP)|*.TTKGP"

448                 .DialogTitle = "Save Map Definition File"
450                 .InitDir = g_sAppPath & "\data\user\Maps\"
452                 .ShowSave

454                 If Len(.Filename) > 0 Then
456                     GIS10.SaveProjectAs .Filename & ".TTKGP"
                    End If

                End With

458         Case "btnExportToShape"

460             With c
                
462                 frmLayerSelection.Init GIS10
        
464                 frmLayerSelection.Show vbModal, Me

                    'Dim olyr As TatukGIS_XDK10.XGIS_LayerVector

466                 If frmLayerSelection.GetItem <> "" Then
468                     Set oLyr = GIS10.get((frmLayerSelection.GetItem))

470                     If Not oLyr Is Nothing Then
472                         frmExportFormats.FraExportFormats.Visible = True
474                         frmExportFormats.Show vbModal, Me
                
476                         If frmExportFormats.bExport Then
                        
478                             If frmExportFormats.chkShapeFile.value = vbChecked Then
                                    
480                                 .Filter = "ESRI Shape files (*.shp)|*.shp"
482                                 .CancelError = False
484                                 .DialogTitle = "Export to ESRI shape file"
486                                 .InitDir = g_sAppPath & "\data\user\Maps\"
488                                 .ShowSave
                    
490                                 Set xportLyr = New TatukGIS_XDK10.XGIS_LayerSHP
                                    
492                                 If Len(.Filename) < 1 Then
494                                     frmLayerSelection.ClearItem
                                        Exit Sub
                                    End If
                                    
496                                 xportLyr.Path = .Filename & ".shp"
498                                 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
500                                 xportLyr.SaveAll
                                End If
                
502                             If frmExportFormats.chkAutocadDWG.value = vbChecked Then
504                                 .Filter = "Autocad DXF files (*.dxf)|*.dxf"
506                                 .CancelError = False
508                                 .DialogTitle = "Export to Autocad DXF file"
510                                 .InitDir = g_sAppPath & "\data\user\Maps\"
512                                 .ShowSave
                    
514                                 If Len(.Filename) < 1 Then
516                                     frmLayerSelection.ClearItem
                                        Exit Sub
                                    End If
                    
518                                 Set xportLyr = New TatukGIS_XDK10.XGIS_LayerDXF
                    
520                                 xportLyr.Path = .Filename & ".dxf"
522                                 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
524                                 xportLyr.SaveAll
                                End If
                        
526                             If frmExportFormats.chkGoogleKML.value = vbChecked Then
528                                 .Filter = "GOOGLE KML files (*.kml)|*.kml"
530                                 .CancelError = False
532                                 .DialogTitle = "Export to Google KML file"
534                                 .InitDir = g_sAppPath & "\data\user\Maps\"
536                                 .ShowSave
                    
538                                 If Len(.Filename) < 1 Then
540                                     frmLayerSelection.ClearItem
                                        Exit Sub
                                    End If
                                        
542                                 Set xportLyr = New TatukGIS_XDK10.XGIS_LayerKML
                    
544                                 xportLyr.Path = .Filename & ".kml"
546                                 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
548                                 xportLyr.SaveAll
                                End If
                        
550                             If frmExportFormats.chkMapInfoTAB.value = vbChecked Then
552                                 .Filter = "OGIS GML files (*.GML)|*.GML"
554                                 .CancelError = False
556                                 .DialogTitle = "Export to OGIS GML  file"
558                                 .InitDir = g_sAppPath & "\data\user\Maps\"
560                                 .ShowSave
                                    
562                                 If Len(.Filename) < 1 Then
564                                     frmLayerSelection.ClearItem
                                        Exit Sub
                                    End If
                    
566                                 Set xportLyr = New TatukGIS_XDK10.XGIS_LayerGML
                    
568                                 xportLyr.Path = .Filename & ".gml"
                            
570                                 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
572                                 xportLyr.SaveAll
                                End If
                
574                             If frmExportFormats.chkGPSExport.value = vbChecked Then
576                                 .Filter = "GPS Export files (*.gpx)|*.gpx"
                                    .CancelError = False
578                                 .DialogTitle = "Export to GPS Export file"
580                                 .InitDir = g_sAppPath & "\data\user\Maps\"
582                                 .ShowSave
                    
584                                 If Len(.Filename) < 1 Then
586                                     frmLayerSelection.ClearItem
                                        Exit Sub
                                    End If
                                        
588                                 Set xportLyr = New TatukGIS_XDK10.XGIS_LayerGPX
                    
590                                 xportLyr.Path = .Filename & ".gpx"
592                                 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", True
594                                 xportLyr.SaveAll
                                End If
                
                            End If
                        End If
                    End If

                End With

596             frmLayerSelection.ClearItem
                
598         Case "btnLoadSQLLyr"

600             If MsgBox("Would You like to use existing SQL layers from the OASIS Databases?", vbYesNo, "Add OASIS Layers") = vbNo Then

602                 With c
                
604                     .Filter = "OASIS SQL Layer files (*.ttkls)|*.ttkls"
606                     .CancelError = True
608                     .DialogTitle = "Open OASIS SQL Layer file"
610                     .InitDir = g_sAppPath & "\data\db\"
612                     .ShowOpen

614                     If Len(.Filename) < 1 Then
                            Exit Sub
                        End If
            
616                     oSQLLyr.Path = .Filename
618                     oSQLLyr.Open
            
620                     GIS10.Add oSQLLyr
                    End With

                Else
                                    
622                 frmSQLLayers.Show vbModal, Me
                    
624                 If frmSQLLayers.m_bOK Then
626                     If Not frmSQLLayers.comSQLLayers.ListIndex < 1 Then

628                         frmSQLLayers.comPath.ListIndex = frmSQLLayers.comSQLLayers.ListIndex
                            
630                         If frmSQLLayers.comPath.Text = "ClientDB" Then
                                'sAdoString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\Data\db\OasisClient.mdb" & ";"
632                             sAdoString = GetConnectionString(g_sAppPath & "\Data\db\OasisClient.mdb")
634                             sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = Native" & vbCrLf & "layer=" & frmSQLLayers.comSQLLayers.List(frmSQLLayers.comSQLLayers.ListIndex) & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString

                            Else
636                             sAdoString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & frmSQLLayers.comPath.Text & ";"
638                             sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = Native" & vbCrLf & "layer=dd_" & frmSQLLayers.comSQLLayers.List(frmSQLLayers.comSQLLayers.ListIndex) & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString
640                             oSQLLyr.caption = frmSQLLayers.comSQLLayers.List(frmSQLLayers.comSQLLayers.ListIndex)
                            End If
                            
642                         oSQLLyr.Path = sAdoString
644                         oSQLLyr.Open
646                         GIS10.Add oSQLLyr
                        End If
                    End If
                End If
 
648         Case "btnAdvancedDBManagement"
650             frmAccessTest.Show vbModal, Me

652         Case "btnEmergency"
                Dim ActLyrs() As String
                Dim DeActLyrs() As String
                Dim oActLyr As TatukGIS_XDK10.XGIS_LayerAbstract
                Dim i As Integer
                
654             g_RSAppSettings.MoveFirst
656             g_RSAppSettings.Find "SettingName = 'EmergencyLayers'"
                
658             m_frmMnuOperations.Legend1.AllowParams = True
                
660             If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").value) Then
662                 If Len(g_RSAppSettings.Fields.Item("SettingValue1").value) > 0 Then
664                     ActLyrs = Split(g_RSAppSettings.Fields.Item("SettingValue1").value, ",")

666                     For i = LBound(ActLyrs) To UBound(ActLyrs)
668                         Set oActLyr = GIS10.get(ActLyrs(i))

670                         If Not oActLyr Is Nothing Then
672                             oActLyr.Params.Visible = True
674                             oActLyr.draw
                                
                            End If

                        Next

                    End If
                End If
                
676             If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue2").value) Then
678                 DeActLyrs = Split(g_RSAppSettings.Fields.Item("SettingValue2").value, ",")
                End If
                
680         Case "btnScript0"

682             DOScripts 0

684         Case "btnScript1"

686             DOScripts 1

688         Case "btnScript2"

690             DOScripts 2

692         Case "btnScript3"

694             DOScripts 3

696         Case "btnScript4"

698             DOScripts 4

700         Case Else
                
702             DOScripts 100
        
        End Select

        On Error Resume Next
        
        Dim sName As String
704     sName = SSC.Procedures.Item("OASISToolBar_ToolClick").Name
        
706     If Len(sName) > 0 Then
708         SSC.Run "OASISToolBar_ToolClick", Tool
        End If
        
        '<EhFooter>
        Exit Sub

AB_ToolClick_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.AB_ToolClick " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub abGridTools_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)
        '<EhHeader>
        On Error GoTo abGridTools_ToolClick_Err
        '</EhHeader>
        Dim c As New cCommonDialog
        Dim oRS As ADODB.Recordset
        Dim sSummaryGroups As String
        Dim m_frmReportsFromRS As frmReportsFromRS
        Dim i As Integer
        
100     Set m_frmReportsFromRS = New frmReportsFromRS
        
102     Select Case Tool.Name
    
            Case Is = "btnReport"
            
                If GISGridRS Is Nothing Then
                    MsgBox "There is no data to export!"
                Else
104                 frmExportFormats.FraExportFormats.Visible = False
106                 frmExportFormats.Show vbModal, Me
                End If
108             If frmExportFormats.bExport And Not GISGridRS Is Nothing Then
                    
110                 If frmExportFormats.chkFormats(0).value = vbChecked Or frmExportFormats.chkFormats(1).value = vbChecked Or frmExportFormats.chkFormats(2).value = vbChecked Or frmExportFormats.chkFormats(3).value = vbChecked Then

112                     With c
114                         .CancelError = False
116                         .DialogTitle = "Export Grid Data to..."
118                         .InitDir = g_sAppPath & "\data\gis\"
                        
120                         If frmExportFormats.chkFormats(0).value = vbChecked Then
122                             .Filter = "Microsoft Excel (.xls)|*.xls"
124                             .DefaultExt = ".xls"
126                             .ShowSave

                                'g_sAppPath & "\data\user\Exports\OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xls"
128                             If Not .Filename = "" Then dxGISDataGrid.M.ExportToXLS .Filename
                            End If

130                         If frmExportFormats.chkFormats(1).value = vbChecked Then
132                             .Filter = "XML data format (.xml)|*.xml"
134                             .DefaultExt = ".xml"
136                             .ShowSave

138                             If Not .Filename = "" Then dxGISDataGrid.M.ExportToXML .Filename 'g_sAppPath & "\data\user\Exports\OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xml"
                            End If

140                         If frmExportFormats.chkFormats(2).value = vbChecked Then
142                             .Filter = "HTML Web page (.html)|*.html"
144                             .DefaultExt = ".html"
146                             .ShowSave

148                             If Not .Filename = "" Then dxGISDataGrid.M.ExportToHTML .Filename 'g_sAppPath & "\data\user\Exports\OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".htm"
                            End If

150                         If frmExportFormats.chkFormats(3).value = vbChecked Then
152                             .Filter = "Tab Limited Text Format (.txt)|*.txt"
154                             .DefaultExt = ".txt"
156                             .ShowSave

158                             If Not .Filename = "" Then dxGISDataGrid.M.SaveAllToTextFile .Filename 'g_sAppPath & "\data\user\Exports\OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".txt"
                            End If

                        End With

                    End If
                    
160                 If frmExportFormats.chkFormats(4).value = vbChecked Then

162                     With dxGISDataGrid.Dataset

164                         Set oRS = New ADODB.Recordset

166                         .DisableControls
168                         .First
        
170                         For i = 1 To GISGridRS.Fields.Count - 1

172                             oRS.Fields.Append GISGridRS.Fields(i).Name, GISGridRS.Fields(GISGridRS.Fields(i).Name).Type, GISGridRS.Fields(GISGridRS.Fields(i).Name).DefinedSize

174                             If dxGISDataGrid.Columns(i).GroupIndex = -1 Then
176                                 If dxGISDataGrid.Columns(i).Sorted = csUp Then
178                                     GISGridRS.Sort = "[" & dxGISDataGrid.Columns(i).FieldName & "]"
180                                 ElseIf dxGISDataGrid.Columns(i).Sorted = csDown Then
182                                     GISGridRS.Sort = "[" & dxGISDataGrid.Columns(i).FieldName & "] DESC"
                                    End If
                                End If

184                         Next i
        
186                         i = 0

188                         Do Until i = dxGISDataGrid.Columns.GroupColumnCount
190                             sSummaryGroups = sSummaryGroups & dxGISDataGrid.Columns.GroupColumn(i).FieldName & ":::"
192                             i = i + 1
                            Loop

194                         If oRS.Fields.Count = 0 Then
196                             MsgBox "No data!"
                                Exit Sub
                            Else
198                             frmGridRptProprties.Init oRS, sSummaryGroups
200                             frmGridRptProprties.txtTitle.Text = abGridTools.Tools.Item("comLyr").Text
202                             frmGridRptProprties.Show vbModal, Me
                            End If

204                         If frmGridRptProprties.bOkClicked Then
        
206                             If Len(dxGISDataGrid.Filter.FilterText) > 2 Then GISGridRS.Filter = Me.dxGISDataGrid.Filter.FilterText
        
208                             If oRS.Fields.Count > 0 Then
            
210                                 oRS.Open
212                                 While Not .EOF

214                                     oRS.AddNew

216                                     For i = 0 To oRS.Fields.Count - 1
218                                         oRS.Fields.Item(i).value = .FieldValues(oRS.Fields.Item(i).Name)
220                                     Next i

222                                     .Next
                                    Wend
        
                                End If
        
                            End If

224                         .EnableControls

                        End With

226                     If frmGridRptProprties.bOkClicked And Not oRS.State = adStateClosed Then

228                         If frmGridRptProprties.chkIncludeMap = vbChecked Then
230                             Clipboard.Clear
232                             GIS10.viewer.PrintClipboard
234                             m_frmReportsFromRS.SetReportRS frmGridRptProprties.txtTitle.Text, oRS, sSummaryGroups, Clipboard.GetData(vbCFEMetafile), frmGridRptProprties.txtMapTitle.Text, oRS.Sort
                            Else
236                             m_frmReportsFromRS.SetReportRS frmGridRptProprties.txtTitle.Text, oRS, sSummaryGroups, , , , oRS.Sort
                            End If

238                         m_frmReportsFromRS.ShowReport
240                         m_frmReportsFromRS.Show vbModal, Me
        
                        End If

242                     Unload frmGridRptProprties
244                     Set m_frmReportsFromRS = Nothing
246                     Set oRS = Nothing

                    End If
                End If
                    
248         Case Is = "btnTheme"

                Dim lL As TatukGIS_XDK10.XGIS_LayerVector
262             Set lL = GIS10.get(GISGridLayerName)
               
                Set m_oXgisThemeWiz = New TatukGIS_XDK10.XGIS_ControlLegendVectorWiz
266             m_oXgisThemeWiz.Execute lL, lL.Shape.ShapeType, lL.ParamsList
            
268             lL.Params.Visible = True
270             GIS10.UpDate

272         Case Is = "btnSpatialAnalysis"

274             With m_frmFreqSettings

276                 If Not .Visible Then
278                     .Init
280                     .Show vbModeless, Me
                    End If

                End With
            Case Is = "btnReset"
            
                GISComboGridChange "btnRefresh"

        End Select

        '<EhFooter>
        Exit Sub

abGridTools_ToolClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.abGridTools_ToolClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdLoadSQLLyr_Click()
        '<EhHeader>
        On Error GoTo cmdLoadSQLLyr_Click_Err
        '</EhHeader>
100     LoadSQLLyr '1
102     FillCOPValues
        '<EhFooter>
        Exit Sub

cmdLoadSQLLyr_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdLoadSQLLyr_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub bTESTING()
    Dim oSQLIncLyr As New TatukGIS_XDK10.XGIS_LayerSqlAdo
   
    SafeMoveFirst g_RSAppSettings
    g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"

    oSQLIncLyr.Name = g_RSAppSettings.Fields.Item("SettingValue1").value & Now()
    oSQLIncLyr.SQLParameter("LAYER") = "oincidents"
    oSQLIncLyr.SQLParameter("DIALECT") = g_sGlobalDialect
    oSQLIncLyr.SQLParameter("ADO") = g_sGlobalConnectionString

    With oIncidentLayerSettings
        oSQLIncLyr.CachedPaint = .CachedPaint
        oSQLIncLyr.HideFromLegend = False
        oSQLIncLyr.IgnoreShapeParams = .IgnoreShapeParams
        oSQLIncLyr.IncrementalPaint = .IncrementalPaint
        '.UseConfig
        '.UseFileParams
    End With

    oSQLIncLyr.Params.Visible = True
        
    GIS10.Add oSQLIncLyr

    CategorizeIncidentsByType oSQLIncLyr.Name, True

    aTESTING
'118     oSQLIncLyr.Active = False

End Sub

Private Sub aTESTING()
        '<EhHeader>
        On Error GoTo aTESTING_Err
        '</EhHeader>
        Dim oSQLIncLyr As New TatukGIS_XDK10.XGIS_LayerSqlAdo
    
100     oSQLIncLyr.Name = "TestCities" & Now()
102     oSQLIncLyr.SQLParameter("LAYER") = "TestCities"
104     oSQLIncLyr.SQLParameter("DIALECT") = g_sGlobalDialect
106     oSQLIncLyr.SQLParameter("ADO") = g_sGlobalConnectionString

108     With oIncidentLayerSettings
110         oSQLIncLyr.CachedPaint = .CachedPaint
112         oSQLIncLyr.HideFromLegend = False
114         oSQLIncLyr.IgnoreShapeParams = .IgnoreShapeParams
116         oSQLIncLyr.IncrementalPaint = .IncrementalPaint
        End With

118     oSQLIncLyr.Params.Visible = True
        
120     GIS10.Add oSQLIncLyr

122     aCategorizeIncidentsByType oSQLIncLyr.Name, True
        
        DumpAllGISProperties
        
        '<EhFooter>
        Exit Sub

aTESTING_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.aTESTING " & _
               "at line " & Erl
    
        '</EhFooter>
End Sub

Public Sub aCategorizeIncidentsByType(Optional sLyrName As String, _
                                      Optional bIgnoreINI As Boolean)
        '<EhHeader>
        On Error GoTo aCategorizeIncidentsByType_Err
        '</EhHeader>
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
        Dim SymbolList As New TatukGIS_XDK10.XGIS_SymbolList
        Dim RS As New ADODB.Recordset
        Dim sFont As String
        Dim sDBFieldName As String
        Dim sGISFieldName As String
   
100     Set lL = GIS10.get(sLyrName)
        
102     lL.ParamsList.Clear

104     lL.ParamsList.Add

106     sFont = "ERS v2 Incidents"
108     sFont = sFont & ":" & 35 & ":NORMAL"
110     lL.Params.Marker.color = 16777215
112     lL.Params.Marker.OutlineColor = 16711680
114     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
116     lL.Params.Marker.Size = 600
118     lL.Params.Marker.ShowLegend = 1
120     lL.Params.Legend = "Uncategorized"

122     lL.ParamsList.Add
124     lL.Params.Query = "STATUS = 'Provincial capital'"
126     sFont = "ESRI Crime Analysis"
128     sFont = sFont & ":41:NORMAL"
130     lL.Params.Marker.color = vbRed
132     lL.Params.Marker.OutlineColor = vbGreen
134     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
136     lL.Params.Marker.Size = 800
138     lL.Params.Marker.ShowLegend = 1
140     lL.Params.Legend = "Provincial Capital"
    
    
142     lL.ParamsList.Add
144     lL.Params.Query = "STATUS = 'National and provincial capital'"
146     sFont = "ERS v2 Incidents"
148     sFont = sFont & ":37:NORMAL"
150     lL.Params.Marker.color = vbGreen
152     lL.Params.Marker.OutlineColor = vbBlue
154     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
156     lL.Params.Marker.Size = 500
158     lL.Params.Marker.ShowLegend = 1
160     lL.Params.Legend = "National & Provincial Capital"
    
    
162     lL.ParamsList.Add
164     lL.Params.Query = "STATUS = 'National capital'"
166     sFont = "ERS v2 Incidents"
168     sFont = sFont & ":21:NORMAL"
170     lL.Params.Marker.color = vbBlack
172     lL.Params.Marker.OutlineColor = vbWhite
174     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
176     lL.Params.Marker.Size = 600
178     lL.Params.Marker.ShowLegend = 1
180     lL.Params.Legend = "National Capital"
            
        '<EhFooter>
        Exit Sub

aCategorizeIncidentsByType_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.aCategorizeIncidentsByType " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub LoadSQLLyr()
        '<EhHeader>
        On Error GoTo LoadSQLLyr_Err
        '</EhHeader>
    
        SafeMoveFirst g_RSAppSettings
        g_RSAppSettings.Find "SettingName = 'IncidentsV2'"

        If g_RSAppSettings.EOF Then
            g_sIncidentsV2DDName = "Sec"
        Else
            g_sIncidentsV2DDName = g_RSAppSettings.Fields.Item("SettingValue1").value
        End If
        
            
            g_bIncidentsV2 = IIf(bSQLServerInUse, DoesTableExist(g_sGlobalConnectionString, "dd_" & g_sIncidentsV2DDName & "_mastertable"), DoesTableExist("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\DynamicData\IncidentsV2.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False", "dd_" & g_sIncidentsV2DDName & "_mastertable"))
            g_sIncidentsV2ConnectionString = IIf(bSQLServerInUse, g_sGlobalConnectionString, "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\DynamicData\IncidentsV2.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False")
        
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'stdLyrs'"
    
104     If g_RSAppSettings.EOF Or g_RSAppSettings.Bof Then Exit Sub
    
106     If g_RSAppSettings.Fields.Item("SettingValue1").value = "1" Then

108         Set m_oSQLIncLyr = New TatukGIS_XDK10.XGIS_LayerSqlAdo
    
110         SafeMoveFirst g_RSAppSettings
112         g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"
114         m_oSQLIncLyr.Name = g_RSAppSettings.Fields.Item("SettingValue1").value

            If Not g_bIncidentsV2 Then
116             m_oSQLIncLyr.SQLParameter("LAYER") = "oincidents"
118             m_oSQLIncLyr.SQLParameter("DIALECT") = g_sGlobalDialect
                m_oSQLIncLyr.SQLParameter("ADO") = g_sGlobalConnectionString
            Else
                m_oSQLIncLyr.SQLParameter("LAYER") = "dd_" & g_sIncidentsV2DDName & "_qryIncidents"
                m_oSQLIncLyr.SQLParameter("DIALECT") = g_sGlobalDialect
                m_oSQLIncLyr.SQLParameter("ADO") = g_sIncidentsV2ConnectionString
            End If

            'Config issue here - if i uncomment out the following the incidents layer does no appear
            m_oSQLIncLyr.HideFromLegend = True
            '            With oIncidentLayerSettings
            '                m_oSQLIncLyr.HideFromLegend = .HideFromLegend
            '                m_oSQLIncLyr.CachedPaint = .CachedPaint
            '                m_oSQLIncLyr.IgnoreShapeParams = .IgnoreShapeParams
            '                m_oSQLIncLyr.IncrementalPaint = .IncrementalPaint
            '                m_oSQLIncLyr.UseConfig = .UseConfig
            '                m_oSQLIncLyr.UseFileParams = .UseFileParams
            '                m_oSQLIncLyr.Params.Visible = .VisibleFromStart
            '            End With

            'CreateTTKGPReference "oincidents_FEA"
124         GIS10.Add m_oSQLIncLyr
126         CategorizeIncidentsByType '1
130         m_oSQLIncLyr.Active = False

        End If

        ' The correct Dialect and Storage type configuration settings must be considered. The Dialect setting specifies the database product while the Storage type setting specifies the format of the data within the SQL database file.
        '
        'As of the DK v. 8.x, the Dialect (database) options are:
        '
        '  GIS_SQL_DIALECT_NAME_MSJET = 'MSJET' ;
        '  GIS_SQL_DIALECT_NAME_MSSQL = 'MSSQL' ;
        '  GIS_SQL_DIALECT_NAME_INTERBASE = 'INTERBASE' ;
        '  GIS_SQL_DIALECT_NAME_MYSQL = 'MYSQL' ;
        '  GIS_SQL_DIALECT_NAME_DB2 = 'DB2' ;
        '  GIS_SQL_DIALECT_NAME_SYBASE = 'SYBASE' ;
        '  GIS_SQL_DIALECT_NAME_ORACLE = 'ORACLE' ;
        '  GIS_SQL_DIALECT_NAME_PROGRESS = 'PROGRESS' ;
        '  GIS_SQL_DIALECT_NAME_INFORMIX = 'INFORMIX' ;
        '  GIS_SQL_DIALECT_NAME_ADVANTAGE = 'ADVANTAGE' ;
        '  GIS_SQL_DIALECT_NAME_SAPDB = 'SAPDB' ;
        '  GIS_SQL_DIALECT_NAME_POSTGRESQL = 'POSTGRESQL' ;
        '  GIS_SQL_DIALECT_NAME_FLASHFILER = 'FLASHFILER' ;
        '  GIS_SQL_DIALECT_NAME_NEXUSDB = 'NEXUSDB' ;
        '
        'As for the DK v. 8.x, the storage options are:
        '
        '  GIS_INI_LAYERSQL_OPENGIS = 'OpenGIS' ;
        '  GIS_INI_LAYERSQL_NATIVE = 'Native' ;
        '  GIS_INI_LAYERSQL_OPENGISBLOB = 'OpenGisBlob' ;
        '  GIS_INI_LAYERSQL_OPENGISBLOB2 = 'OpenGisBlob2' ;
        '  GIS_INI_LAYERSQL_OPENGISNORMALIZED = 'OpenGisNormalized' ;
        '  GIS_INI_LAYERSQL_OPENGISNORMALIZED2 = 'OpenGisNormalized2' ;
        '  GIS_INI_LAYERSQL_GEOMEDIA = 'Geomedia' ;
        '  GIS_INI_LAYERSQL_PIXELSTORE2 = 'PixelStore2' ;
        '
        '‘Native” corresponds to the TatukGIS proprietary format. OpenGisBlob corresponds to the OpenGIS (OGC) WKB format. OpenGisNormalized corresponds to the OpenGIS WKT format (but with the x and y coordinates organized into separate database columns ).

        '<EhFooter>
        Exit Sub

LoadSQLLyr_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.LoadSQLLyr " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateTTKGPReference(sName As String)
        '<EhHeader>
        On Error GoTo CreateTTKGPReference_Err
        '</EhHeader>
    
        Dim RSttkGISLayerSQL As New ADODB.Recordset
        Dim sShortName As String
        Dim oLocalCnn As ADODB.Connection
        Set oLocalCnn = New ADODB.Connection
        oLocalCnn.Open g_sGlobalConnectionString

100     If Not DoesTableExist(g_sGlobalConnectionString, "ttkGISLayerSQL") Then
            
102         oLocalCnn.Execute "CREATE TABLE [ttkGISLayerSQL]  ([NAME] TEXT(255), [XMIN] INTEGER, [XMAX] INTEGER, [YMIN] INTEGER, [YMAX] INTEGER, [SHAPETYPE] INTEGER)"
104         DebugPrint "   --- Creating table [ttkGISLayerSQL]"

        End If
            
106     sShortName = left$(sName, Len(sName) - 4)
            
108     With RSttkGISLayerSQL
            
110         .Open "SELECT * FROM [ttkGISLayerSQL] WHERE [Name] = '" & sShortName & "'", oLocalCnn, adOpenDynamic, adLockBatchOptimistic
                
112         If .EOF Then
           
114             .AddNew
116             .Fields("Name") = sShortName
118             .Fields("XMIN") = -180
120             .Fields("XMAX") = 180
122             .Fields("YMIN") = -180
124             .Fields("YMAX") = 180
126             .Fields("SHAPETYPE") = 0
128             .UpdateBatch adAffectCurrent
130             DebugPrint "   --- Populating table [ttkGISLayerSQL] for GEOTABLE: [" & sName & "]"
           
            End If
           
132         .Close
            
        End With
            
        oLocalCnn.Close
        Set oLocalCnn = Nothing
134     Set RSttkGISLayerSQL = Nothing

        '<EhFooter>
        Exit Sub

CreateTTKGPReference_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CreateTTKGPReference " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateIncidentsTTKGPReference()
        '<EhHeader>
        On Error GoTo CreateIncidentsTTKGPReference_Err
        '</EhHeader>
    
        Dim RSttkGISLayerSQL As New ADODB.Recordset
        Dim sShortName As String
        Dim oLocalCnn As ADODB.Connection
100     Set oLocalCnn = New ADODB.Connection
102     oLocalCnn.Open g_sGlobalConnectionString

104     If Not DoesTableExist(g_sGlobalConnectionString, "ttkGISLayerSQL") Then
            
106         oLocalCnn.Execute "CREATE TABLE [ttkGISLayerSQL]  ([NAME] TEXT(255), [XMIN] INTEGER, [XMAX] INTEGER, [YMIN] INTEGER, [YMAX] INTEGER, [SHAPETYPE] INTEGER)"
108         DebugPrint "   --- Creating table [ttkGISLayerSQL]"

        End If
            
    '110     sShortName = Left$(sName, Len(sName) - 4)
            
110     With RSttkGISLayerSQL
            
112         .Open "SELECT * FROM [ttkGISLayerSQL] WHERE [Name] = 'oincidents'", oLocalCnn, adOpenDynamic, adLockBatchOptimistic
                
114         If .EOF Then
           
116             If g_DatabaseSpecExtent Is Nothing Then
118                 Set g_DatabaseSpecExtent = GIS10.Extent
                End If
           
120             .AddNew
122             .Fields("Name") = "oincidents"
124             .Fields("XMIN") = g_DatabaseSpecExtent.xmin
126             .Fields("XMAX") = g_DatabaseSpecExtent.xmax
128             .Fields("YMIN") = g_DatabaseSpecExtent.ymin
130             .Fields("YMAX") = g_DatabaseSpecExtent.ymax
132             .Fields("SHAPETYPE") = 2
134             .UpdateBatch adAffectCurrent
136             DebugPrint "   --- Populating table [ttkGISLayerSQL] for GEOTABLE: [oincidents]"
           
            End If
           
138         .Close
            
        End With
            
140     oLocalCnn.Close
142     Set oLocalCnn = Nothing
144     Set RSttkGISLayerSQL = Nothing

        '<EhFooter>
        Exit Sub

CreateIncidentsTTKGPReference_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CreateIncidentsTTKGPReference " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub CategorizeIncidentsByType(Optional sLyrName As String, _
                                     Optional bIgnoreINI As Boolean)
        '<EhHeader>
        On Error GoTo CategorizeIncidentsByType_Err
        '</EhHeader>
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
        Dim SymbolList As New TatukGIS_XDK10.XGIS_SymbolList
        Dim RS As New ADODB.Recordset
        Dim sFont As String
        Dim sDBFieldName As String
        Dim sGISFieldName As String
        Dim sTable As String
        Dim oCn As ADODB.Connection
  
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"

104     If sLyrName = "" Then sLyrName = g_RSAppSettings.Fields.Item("SettingValue1").value
        
106     Set lL = GIS10.get(sLyrName)
        
108     If Not bIgnoreINI Then
110         If Not m_sIncidentIni = "" Then

112             With lL
114                 .ConfigName = m_sIncidentIni
116                 .StoreParamsInProject = False
118                 .UseConfig = True
120                 .RereadConfig
122                 .draw
                End With

                Exit Sub
            End If
        End If
        
124     Set RS = New ADODB.Recordset
                
126     Set RS.ActiveConnection = g_RSAppSettings.ActiveConnection

128     If g_bIncidentsV2 Then
                
130         Set oCn = New ADODB.Connection
132         oCn.Open g_sIncidentsV2ConnectionString
134         Set RS.ActiveConnection = oCn
        Else
136         Set RS.ActiveConnection = g_RSAppSettings.ActiveConnection
        End If
                
138     RS.CursorType = adOpenDynamic
140     RS.CursorLocation = g_sGlobalCursorLocation
142     RS.LockType = adLockReadOnly
                
144     Select Case m_frmMnuOperations.abGridTools.Tools.Item("comIncCategory").Text  'shp.GetField("Type")

            Case "Target"
146             sDBFieldName = IIf(g_bIncidentsV2, "option", "Name")
148             sGISFieldName = IIf(g_bIncidentsV2, "Target", "Target")
150             sTable = IIf(g_bIncidentsV2, "dd_" & g_sIncidentsV2DDName & "_ddEventVictim", "IncTarget")
152             RS.Open "SELECT * FROM " & sTable & " ORDER BY [" & sDBFieldName & "]"

154         Case "Time"
156             sDBFieldName = IIf(g_bIncidentsV2, "Incident_Time_Name", "Incident_Time_Name")
158             sGISFieldName = IIf(g_bIncidentsV2, "TIME00", "TIME00")
160             sTable = IIf(g_bIncidentsV2, "dd_" & g_sIncidentsV2DDName & "_ddIncTimeCategory", "IncTimeCategory")
162             RS.Open "SELECT * FROM " & sTable & " ORDER BY [" & sDBFieldName & "]"

164         Case "Type"
166             sDBFieldName = IIf(g_bIncidentsV2, "option", "Incident_Type_Name")
168             sGISFieldName = IIf(g_bIncidentsV2, "Type", "Type")
170             sTable = IIf(g_bIncidentsV2, "dd_" & g_sIncidentsV2DDName & "_ddEventType", "IncTypeCategory")
172             RS.Open "SELECT * FROM " & sTable & " ORDER BY [" & sDBFieldName & "]"

174         Case Else
176             sDBFieldName = IIf(g_bIncidentsV2, "option", "Incident_Type_Name")
178             sGISFieldName = IIf(g_bIncidentsV2, "Type", "Type")
180             sTable = IIf(g_bIncidentsV2, "dd_" & g_sIncidentsV2DDName & "_ddEventType", "IncTypeCategory")
182             RS.Open "SELECT * FROM " & sTable & " ORDER BY [" & sDBFieldName & "]"

        End Select
                
        '    sDBFieldName = "Incident_Type_Name"
        '    sGISFieldName = "Type"
        '    RS.Open "SELECT * FROM IncTypeCategory"
                
184     SafeMoveFirst RS
186     lL.ParamsList.Clear

        'll.ParamsList.Add
        'll.Params.Query = sGISFieldName & " <> ''"
        'sFont = RS.Fields.Item("Font_Name").value
        'sFont = sFont & ":72:NORMAL"
            
        'll.Params.Marker.Color = RS.Fields.Item("bgColor").value
        'll.Params.Marker.OutlineColor = RS.Fields.Item("color").value
        'll.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
        'll.Params.Marker.Size = RS.Fields.Item("size").value
        'll.Params.Marker.ShowLegend = 1
        'll.Params.Legend = "Other"

188     lL.ParamsList.Add
        'll.Params.AreaColor=RGB(102:102:102)
        'DebugPrint  .Item("Incident_Type_Name").value
190     sFont = "ERS v2 Incidents"
192     sFont = sFont & ":" & 35 & ":NORMAL"
    
194     lL.Params.Marker.color = 16777215
196     lL.Params.Marker.OutlineColor = 16711680
198     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
200     lL.Params.Marker.Size = 600
202     lL.Params.Marker.ShowLegend = 1
204     lL.Params.Legend = "Uncategorized"

206     'DebugPrint "Current Font:" & sFont
        
208     Do While Not RS.EOF

210         With RS.Fields
        
212             lL.ParamsList.Add
214             lL.Params.Query = sGISFieldName & " = '" & .Item(sDBFieldName).value & "'"
                '''DebugPrint "THIS IS QUERY:" & sGISFieldName & " = '" & .Item(sDBFieldName).Value & "'"
                '''DebugPrint "FieldName" & sGISFieldName & " The DB Value:" & sDBFieldName
                'll.Params.AreaColor=RGB(102:102:102)
                'DebugPrint  .Item("Incident_Type_Name").value
216             sFont = .Item("Font_Name").value
218             sFont = sFont & ":" & .Item("Ascii").value & ":NORMAL"
                '''DebugPrint "FONT=" & sFont
220             lL.Params.Marker.color = .Item("bgColor").value
222             lL.Params.Marker.OutlineColor = .Item("color").value
224             lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
226             lL.Params.Marker.Size = .Item("size").value
228             lL.Params.Marker.ShowLegend = 1
230             lL.Params.Legend = .Item(sDBFieldName).value
                '''DebugPrint "Legend Source" & .Item(sDBFieldName).Value
232             RS.MoveNext
            End With

        Loop
        
234     If sLyrName = abGridTools.Tools.Item("comLyr").Text Then
236         LoadLayerAttrDataToGridInit
        End If
            
        'll.UseFileParams
        
        '<EhFooter>
        Exit Sub

CategorizeIncidentsByType_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CategorizeIncidentsByType " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub CategorizeIncidentsByTypeBerserk(bUseConfig As Boolean)
        '<EhHeader>
        On Error GoTo CategorizeIncidentsByType_Err
        '</EhHeader>
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
        Dim SymbolList As New TatukGIS_XDK10.XGIS_SymbolList
        Dim RS As New ADODB.Recordset
        Dim sFont As String
        Dim sDBFieldName As String
        Dim sGISFieldName As String
        Dim sLyrName As String
  
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"

        sLyrName = g_RSAppSettings.Fields.Item("SettingValue1").value
        
104     Set lL = GIS10.get(sLyrName)

        If bUseConfig Then ' Not m_sIncidentIni = "" Then

            With lL
                .ConfigName = m_sIncidentIni
                .StoreParamsInProject = False
                .UseConfig = True
                .RereadConfig
                .draw
            End With

            Exit Sub
        End If
    
106     Set RS = New ADODB.Recordset
                
108     Set RS.ActiveConnection = m_Cnn
                
110     RS.CursorType = adOpenDynamic
112     RS.CursorLocation = g_sGlobalCursorLocation
114     RS.LockType = adLockReadOnly

118     sDBFieldName = "Name"
120     sGISFieldName = "TARGET"
122     RS.Open "SELECT * FROM IncTarget ORDER BY [Name]"
        DebugPrint "Target Choosen"
                
150     lL.ParamsList.Clear
152     SafeMoveFirst RS

        'll.ParamsList.Add
        'll.Params.Query = sGISFieldName & " <> ''"
        'sFont = RS.Fields.Item("Font_Name").value
        'sFont = sFont & ":72:NORMAL"
            
        'll.Params.Marker.Color = RS.Fields.Item("bgColor").value
        'll.Params.Marker.OutlineColor = RS.Fields.Item("color").value
        'll.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
        'll.Params.Marker.Size = RS.Fields.Item("size").value
        'll.Params.Marker.ShowLegend = 1
        'll.Params.Legend = "Other"

154     lL.ParamsList.Add
156     sFont = "ERS v2 Incidents"
158     sFont = sFont & ":" & 35 & ":NORMAL"
    
160     lL.Params.Marker.color = 16777215
162     lL.Params.Marker.OutlineColor = 16711680
164     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
166     lL.Params.Marker.Size = 600
168     lL.Params.Marker.ShowLegend = 1
170     lL.Params.Legend = "Uncategorized"

        DebugPrint "Current Font:" & sFont
        
172     Do While Not RS.EOF

174         With RS.Fields
        
176             lL.ParamsList.Add
178             lL.Params.Query = sGISFieldName & " = '" & .Item(sDBFieldName).value & "'"
                DebugPrint "THIS IS QUERY:" & sGISFieldName & " = '" & .Item(sDBFieldName).value & "'"
                DebugPrint "FieldName" & sGISFieldName & " The DB Value:" & sDBFieldName

180             sFont = .Item("Font_Name").value
182             sFont = sFont & ":" & .Item("Ascii").value & ":NORMAL"
                DebugPrint "FONT=" & sFont
184             lL.Params.Marker.color = .Item("bgColor").value
186             lL.Params.Marker.OutlineColor = .Item("color").value
188             lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
190             lL.Params.Marker.Size = .Item("size").value
192             lL.Params.Marker.ShowLegend = 1
194             lL.Params.Legend = .Item(sDBFieldName).value
                DebugPrint "Legend Source" & .Item(sDBFieldName).value
196             RS.MoveNext
            End With

        Loop
        
        If sLyrName = abGridTools.Tools.Item("comLyr").Text Then
            LoadLayerAttrDataToGridInit
        End If
            
        lL.draw
            
        '<EhFooter>
        Exit Sub

CategorizeIncidentsByType_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.CategorizeIncidentsByType " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateLayerInis()
    Dim oLyrAbs As TatukGIS_XDK10.XGIS_LayerAbstract
    Dim i As Integer
    
    With oLyrAbs
 
        For i = 0 To GIS10.items.Count - 1
            Set oLyrAbs = GIS10.get(GIS10.items.Item(i).Name)
            oLyrAbs.StoreParamsInProject = False
            oLyrAbs.ConfigName = g_sAppPath & "\Data\gis\temp\" & oLyrAbs.Name
            oLyrAbs.WriteConfig
            oLyrAbs.SaveAll
        Next
    
    End With

End Sub

Public Sub ReadLayerStyleFromDB(oLayer As TatukGIS_XDK10.XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo ReadLayerStyleFromDB_Err
        '</EhHeader>

100     With oLayer
        
102         With .Params
104             With .area
            
106                 DebugPrint .color
                    'DebugPrint  .Symbol.FontName
                    'DebugPrint  .SymbolSize
                    'DebugPrint  .SymbolGap
                    'DebugPrint  .SymbolRotate
                    'DebugPrint  .Bitmap
                    'DebugPrint  .Pattern
                    'DebugPrint  .OutlineColor
                    'DebugPrint  .OutlineWidth
                    'DebugPrint  .OutlineStyle
                    'DebugPrint  .OutlineSymbol
                    'DebugPrint  .OutlineSymbolGap
                    'DebugPrint  .OutlineSymbolRotate
                    'DebugPrint  .OutlineBitmap
                    'DebugPrint  .OutlinePattern
                    'DebugPrint  .SmartSize
                    'DebugPrint  .SmartSizeField
            
108                 .color = RGB(0, 0, 255)
                    '                .Symbol
                    '                .SymbolSize
                    '                .SymbolGap
                    '                .SymbolRotate
                    '                .Bitmap
                    '                .Pattern
                    '                .OutlineColor = RGB(0, 0, 255)
                    '                .OutlineWidth
                    '                .OutlineStyle
                    '                .OutlineSymbol
                    '                .OutlineSymbolGap
                    '                .OutlineSymbolRotate
                    '                .OutlineBitmap
                    '                .OutlinePattern
                    '                .SmartSize
                    '                .SmartSizeField
            
                End With
           
110             With .Marker
           
112                 DebugPrint .color
                    'DebugPrint  .Marker.Symbol
                    '.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
                    'DebugPrint  .SymbolSize
                    'DebugPrint  .SymbolGap
                    'DebugPrint  .SymbolRotate
                    'DebugPrint  .Bitmap
                    'DebugPrint  .Pattern
                    'DebugPrint  .OutlineColor
                    'DebugPrint  .OutlineWidth
                    'DebugPrint  .OutlineStyle
                    'DebugPrint  .OutlineSymbol
                    'DebugPrint  .OutlineSymbolGap
                    'DebugPrint  .OutlineSymbolRotate
                    'DebugPrint  .OutlineBitmap
                    'DebugPrint  .OutlinePattern
                    'DebugPrint  .SmartSize
                    'DebugPrint  .SmartSizeField
                End With
           
114             With .Line
           
116                 DebugPrint .color
                    'DebugPrint  .Marker.Symbol
                    '.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
                    'DebugPrint  .SymbolSize
                    'DebugPrint  .SymbolGap
                    'DebugPrint  .SymbolRotate
                    'DebugPrint  .Bitmap
                    'DebugPrint  .Pattern
                    'DebugPrint  .OutlineColor
                    'DebugPrint  .OutlineWidth
                    'DebugPrint  .OutlineStyle
                    'DebugPrint  .OutlineSymbol
                    'DebugPrint  .OutlineSymbolGap
                    'DebugPrint  .OutlineSymbolRotate
                    'DebugPrint  .OutlineBitmap
                    'DebugPrint  .OutlinePattern
                    'DebugPrint  .SmartSize
                    'DebugPrint  .SmartSizeField
                End With
           
            End With

        End With

        '<EhFooter>
        Exit Sub

ReadLayerStyleFromDB_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ReadLayerStyleFromDB " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub comActiveLyr_Click()
        '<EhHeader>
        On Error GoTo comActiveLyr_Click_Err
        '</EhHeader>
  'GIS109 Done
        Dim i As Integer
        Dim lL As TatukGIS_XDK10.XGIS_LayerAbstract
  
100     If Not m_bLOADING Then

102         For i = 1 To comActiveLyr.ListCount - 1
104             Set lL = GIS10.get(comActiveLyr.List(i))

106             If comActiveLyr.ListIndex = 0 Then
108                 lL.Active = True
                Else

110                 If i = comActiveLyr.ListIndex Then
112                     lL.Active = True
                    Else
114                     lL.Active = False
                    End If
                End If

            Next
  
116         GIS10.Invalidate
        End If
    
        '<EhFooter>
        Exit Sub

comActiveLyr_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.comActiveLyr_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ctTreeBookmrks_Change(ByVal nIndex As Long)
        ' DebugPrint  "ctTreeBookmrks_Change" & nIndex
        '<EhHeader>
        On Error GoTo ctTreeBookmrks_Change_Err
        '</EhHeader>
        'GIS109 Done
100     DebugPrint ctTreeBookmrks.NodeText(nIndex) & " Change Level:" & ctTreeBookmrks.NodeLevel(nIndex)
   
        ' Exit Sub
    
        Dim RS As New ADODB.Recordset
   
102     If ctTreeBookmrks.NodeText(nIndex) = "Start View" Then 'm_DefaultViewName Then
            If Not g_DatabaseSpecExtent Is Nothing Then
                GIS10.VisibleExtent = g_DatabaseSpecExtent
            Else
                GIS10.VisibleExtent = GIS10.Extent
            End If
104         'GIS10.viewer.SetViewport m_DefaultViewx, m_DefaultViewy
106         'GIS10.viewer.zoom = m_DefaultViewz
108         GIS10.UpDate
            'Map1.ZoomTo m_DefaultViewz, m_DefaultViewx, m_DefaultViewy
        Else

110         If ctTreeBookmrks.NodeLevel(nIndex) = 2 Then
112             RS.Open "SELECT X,Y,Z FROM GeoBookMarks WHERE Name ='" & ctTreeBookmrks.NodeText(nIndex) & "'", m_Cnn, adOpenDynamic, adLockReadOnly

114             If Not RS.Bof Then SafeMoveFirst RS
            
                Dim ptg As New TatukGIS_XDK10.XGIS_Point

                If IsNull(RS.Fields.Item("X")) Or IsNull(RS.Fields.Item("Y")) Or IsNull(RS.Fields.Item("Z")) Then
                
                    MsgBox "XY coordinates or zoom factor for this bookmark are invalid!"
                
                Else
                
116                 ptg.Prepare CDbl(RS.Fields.Item("X")), CDbl(RS.Fields.Item("Y"))

118                 GIS10.zoom = RS.Fields.Item("Z")
120                 GIS10.CenterViewport ptg
122                 GIS10.UpDate
                End If
            End If
        End If

        'DebugPrint  ctTreeBookmrks.NodeLevel(nIndex)

        '<EhFooter>
        Exit Sub

ctTreeBookmrks_Change_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.ctTreeBookmrks_Change " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ctTreeBookmrks_HeaderClick(ByVal nColumn As Integer)
        '<EhHeader>
        On Error GoTo ctTreeBookmrks_HeaderClick_Err
        '</EhHeader>
100     DebugPrint "TreeHeaderClicked" & nColumn
        '<EhFooter>
        Exit Sub

ctTreeBookmrks_HeaderClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ctTreeBookmrks_HeaderClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub EventLayer_OnPaintShape(translated As Boolean, _
                                    ByVal Shape As Object)
        '<EhHeader>
        On Error GoTo EventLayer_OnPaintShape_Err
        '</EhHeader>
        Dim population As Double
        Dim area As Double
        Dim factor As Double
        Dim fl As TatukGIS_XDK10.XGIS_ShapePolygon
        Dim i As Integer

100     translated = True
        
12      If Not Shape.ShapeType = XgisShapeTypePolygon Then Exit Sub
        
13      Set fl = GIS10.get(m_sSecAnalysisLyrName).GetShape(Shape.uID).MakeEditable
        
15      If fl Is Nothing Then Exit Sub
'        fl.Flash 4, 10
'67       fl.MakeEditable = True
        
        With fl
        
102         factor = .GetField(m_sSecAnalysisFieldName) - 100

            With .Params.area
                        
                If factor < arRangeVals(0) Then
                   .color = &H8C07&
                   .Pattern = XBrushStyle.XbsClear
                Else
                    For i = 1 To UBound(arRangeColors)
                        If factor < arRangeVals(i) Then
                            .color = arRangeColors(i)
                            .Pattern = XBrushStyle.XbsSolid
                            Exit For
                        End If
                    Next
                End If
                
            End With
            .MakeEditable = False
138         .draw
            'GIS10.UpDate
        End With
        
        '<EhFooter>
        Exit Sub

EventLayer_OnPaintShape_Err:
        DebugPrint Err.Description & vbCrLf & "in OASISClient.frmMain.EventLayer_OnPaintShape " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Public Function SaveAsShape(sLayerName As String, sNewLAyerName As String) As TatukGIS_XDK10.XGIS_LayerSHP
        '<EhHeader>
        On Error GoTo SaveAsShape_Err
        '</EhHeader>
    Dim oTargetLayer As New TatukGIS_XDK10.XGIS_LayerSHP
    Dim oSourceLyr As TatukGIS_XDK10.XGIS_LayerVector

100     Set oSourceLyr = GIS10.get(sLayerName)
    
102     oTargetLayer.Path = g_sAppPath & "\data\gis\temp\" & sNewLAyerName & ".shp"
104     oTargetLayer.ExportLayer oSourceLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XgisShapeTypeUnknown, "", False
106     GIS10.Add oTargetLayer
108     Set SaveAsShape = oTargetLayer
    
        '<EhFooter>
        Exit Function

SaveAsShape_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SaveAsShape " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub m_frmSpatialAnalysis_GetSimpleSpatialAnalysis(sContraint As String, _
                                                         sOverLayLayer As String, _
                                                         sTargetLayer As String, _
                                                         sTargetID As String, _
                                                         sScope As String, _
                                                         sDE_9IM As String, _
                                                         ByRef sSpatialString As String)
        '<EhHeader>
        On Error GoTo m_frmSpatialAnalysis_GetSimpleSpatialAnalysis_Err
        '</EhHeader>
        Dim oTestShp As TatukGIS_XDK10.XGIS_Shape
        Dim oOverLayLayer As TatukGIS_XDK10.XGIS_LayerVector
        Dim oOverlayLayer2 As TatukGIS_XDK10.XGIS_LayerVector
        Dim oTargetLyr As TatukGIS_XDK10.XGIS_LayerVector
        Dim lTot As Long
        Dim lnow As Long
        Dim sResult() As String
    
100     ReDim sResult(0)
        'GIS109 Done
        Set oTargetLyr = GIS10.get(sTargetLayer)

104     oTargetLyr.scope = sContraint
    
110     Set oOverLayLayer = GIS10.get(sOverLayLayer)
        oOverLayLayer.scope = sScope
        'This workaround fixes the date issue (tatuk limitation)
112     Set oOverlayLayer2 = New TatukGIS_XDK10.XGIS_LayerVector
114     oOverlayLayer2.ImportLayer oOverLayLayer, oOverLayLayer.Extent, TatukGIS_XDK10.XgisShapeTypeUnknown, sScope, True

        For Each oTestShp In oTargetLyr.Loop(GisUtils.GisWholeWorld, "", Nothing, "", True)

116         If Not oTestShp Is Nothing Then
118             lnow = GetNumWithinGeometry(oTestShp, oOverlayLayer2, "", sDE_9IM)
120             lTot = lTot + lnow
122             sResult(UBound(sResult)) = oTestShp.GetField(sTargetID) & ",,," & lnow
126             If Not oTestShp Is Nothing Then ReDim Preserve sResult(UBound(sResult) + 1) 'GIS109 Todo IS this array really true
            End If

        Next

128     oTargetLyr.scope = ""
130     sSpatialString = Join(sResult, ":::")

        '<EhFooter>
        Exit Sub

m_frmSpatialAnalysis_GetSimpleSpatialAnalysis_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.m_frmSpatialAnalysis_GetSimpleSpatialAnalysis " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetNumWithinGeometry(oShpToSearch As TatukGIS_XDK10.XGIS_Shape, _
                                      oLyr As TatukGIS_XDK10.XGIS_LayerVector, _
                                      sScope As String, _
                                      sDE_9IM As String) As Long
        '<EhHeader>
        On Error GoTo GetNumWithinGeometry_Err
        '</EhHeader>
        Dim i As Long
        Dim aShape As TatukGIS_XDK10.XGIS_Shape
        'GIS109 Done
100     If oLyr Is Nothing Then Exit Function
    
102     With oLyr
    
104         .scope = sScope
'106         .MoveFirst GisUtils.GisWholeWorld, "", Nothing, "", True
     
            For Each aShape In .Loop(oShpToSearch.Extent, "", oShpToSearch, sDE_9IM, True)
                i = i + 1
            Next
     
116         .scope = ""
    
        End With
    
118     GetNumWithinGeometry = i
    
        '<EhFooter>
        Exit Function

GetNumWithinGeometry_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GetNumWithinGeometry " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub CheckIncidentFrequency(oOverLayLayer As TatukGIS_XDK10.XGIS_LayerVector, _
                                  oTargetLayer As TatukGIS_XDK10.XGIS_LayerVector, _
                                  Optional bResetPreviousScoring As Boolean = True)
        '<EhHeader>
        On Error GoTo CheckIncidentFrequency_Err
        '</EhHeader>
        Dim Shh As Shape
        Dim admShape As Shape
        Dim i As Integer
        Dim shp As TatukGIS_XDK10.XGIS_ShapePoint
        Dim aShape As Object
        Dim oPolyShape As TatukGIS_XDK10.XGIS_ShapePolygon
        Dim iVal As Integer
        Dim iTimeVal As Integer
        Dim iTargetVal As Integer
        Dim iTypeVal As Integer
        Dim oRSType As ADODB.Recordset
        Dim oRSTarget As ADODB.Recordset
        Dim oRSTime As ADODB.Recordset
        'GIS109 Done
        
100     If oTargetLayer Is Nothing Then
102         MsgBox "No Target layer selected. Choose a new target layer and try again", vbInformation, "OASIS Security Analysis"
            Exit Sub
        End If
        
104     If oOverLayLayer Is Nothing Then
106         MsgBox "No Overlay layer selected. Choose a new Overlay layer and try again", vbInformation, "OASIS Security Analysis"
            Exit Sub
        End If
        
108     oTargetLayer.IgnoreShapeParams = True
110     oTargetLayer.CachedPaint = False

112     If frmScoring.m_bApply Then
        
114         If frmScoring.chkChkActivate(0).value = vbChecked Then
116             Set oRSType = New ADODB.Recordset
        
118             With oRSType
120                 .Open "SELECT Incident_Type_Name, Scoring FROM IncTypeCategory ORDER BY [Incident_Type_Name]", m_Cnn, adOpenDynamic, adLockReadOnly
                End With

            End If
        
122         If frmScoring.chkChkActivate(1).value = vbChecked Then
        
124             Set oRSTarget = New ADODB.Recordset

126             With oRSTarget
128                 .Open "SELECT Name, Scoring FROM IncTarget ORDER BY [NAME]", m_Cnn, adOpenDynamic, adLockReadOnly
                End With

            End If
        
130         If frmScoring.chkChkActivate(2).value = vbChecked Then

132             Set oRSTime = New ADODB.Recordset

134             With oRSTime
136                 .Open "SELECT Incident_Time_Name, Scoring FROM IncTimeCategory ORDER BY [Incident_Time_Name]", m_Cnn, adOpenDynamic, adLockReadOnly
                End With

            End If
        
        End If

138     If bResetPreviousScoring Then
140         ResetScoring oTargetLayer
       End If
        
142     For Each shp In oOverLayLayer.Loop(GIS10.viewer.VisibleExtent, "", Nothing, "", True)
        
144         For Each aShape In oTargetLayer.Loop(shp.Extent, "", shp, GisUtils.GIS_RELATE_WITHIN, True)

146             If Not aShape.GetField(m_sSecAnalysisFieldName) = vbNull Then
148                 Set oPolyShape = aShape.MakeEditable
                
150                 If frmScoring.m_bApply Then
                    
152                     iTargetVal = 0
154                     iTimeVal = 0
156                     iTypeVal = 0

158                     If frmScoring.chkChkActivate(0).value = vbChecked Then

160                         With oRSType
162                             SafeMoveFirst oRSType
164                             .Find "Incident_Type_Name = '" & shp.GetField("TYPE") & "'"
166                             If Not .EOF Then iTypeVal = CInt(.Fields.Item("Scoring").value)
                            End With

                        End If
                    
168                     If frmScoring.chkChkActivate(1).value = vbChecked Then

170                         With oRSTarget
172                             SafeMoveFirst oRSTarget
174                             .Find "Name = '" & shp.GetField("TARGET") & "'"
176                             If Not .EOF Then iTargetVal = CInt(.Fields.Item("Scoring").value)
                            End With

                        End If
                    
178                     If frmScoring.chkChkActivate(2).value = vbChecked Then

180                         With oRSTime
182                             SafeMoveFirst oRSTime
184                             .Find "Incident_Time_Name = '" & shp.GetField("TIME00") & "'"
186                             If Not .EOF Then iTimeVal = CInt(.Fields.Item("Scoring").value)
                            End With

                        End If
                    End If
                
188                 iVal = aShape.GetField(m_sSecAnalysisFieldName) + 1
190                 oPolyShape.SetField m_sSecAnalysisFieldName, CLng(iVal + iTargetVal + iTimeVal + iTypeVal)  'CInt(aShape.GetField("Scoring") + 1)
                Else
                    'TODO CHECK IF THIS IS NEEDED!
                    'Set oPolyShape = aShape.MakeEditable
                    'oPolyShape.SetField "Scoring", 100
                End If
                    
            Next
                
192         i = i + 1

        Next
      '  oTargetLayer.SaveData
194     oTargetLayer.SaveAll
196     oTargetLayer.Active = True 'Params.Visible = True
        '<EhFooter>
        Exit Sub

CheckIncidentFrequency_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CheckIncidentFrequency " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub ResetScoring(oLayer As TatukGIS_XDK10.XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo ResetScoring_Err
        '</EhHeader>
        On Error Resume Next
        
    
        
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim oshp As TatukGIS_XDK10.XGIS_Shape
        'GIS109 Done
       GIS10.get(m_sSecAnalysisLyrName).RevertAll
     GIS10.get(m_sSecAnalysisLyrName).SaveAll
     
     
102     For Each shp In oLayer.Loop(oLayer.Extent, "", Nothing, "", True)
104         Set oshp = shp.MakeEditable
106         oshp.SetField m_sSecAnalysisFieldName, 100
            Set oshp = Nothing
        Next
        
        
        '<EhFooter>
        Exit Sub

ResetScoring_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ResetScoring " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub AddLayer(Optional oLayer As Object)
        '<EhHeader>
        On Error GoTo AddLayer_Err
        '</EhHeader>
        Dim oAbsLayer As TatukGIS_XDK10.XGIS_LayerAbstract
        Dim c As New cCommonDialog
        Dim sName As String
100     With c
    
102         If oLayer Is Nothing Then
        
104             .Filter = GisUtils.GisSupportedFiles(XGIS_FileType.XgisFileTypeAll, False)
106             .CancelError = True
108             .DialogTitle = "Add Map Layers"
110             .InitDir = g_sAppPath & "\data\gis\"
112             .ShowOpen
            
                'Check the type of layer
            
                If Not Len(.Filename) < 3 Then
            
114                 If InStr(UCase(GisUtils.GisSupportedFiles(XGIS_FileType.XgisFileTypePixel, False)), UCase(Mid(.Filename, InStr(.Filename, ".")))) = 0 Then
                
                        'Check For dublicates
                        sName = Mid$(.FileTitle, 1, Len(.FileTitle) - 4)
113                     Set oLayer = GIS10.get(sName)
                        
115                     If Not oLayer Is Nothing Then 'We Found a dublicate
                            sName = sName & Day(Now()) & Month(Now()) & Minute(Now()) & Second(Now())
                        End If
                        
                        'IS VECTOR FILES
116                     If InStr(UCase(.Filename), ".SHP") Then
118                         Set oLayer = New TatukGIS_XDK10.XGIS_LayerSHP
120                     ElseIf InStr(UCase(.Filename), ".KML") Then
122                         Set oLayer = New TatukGIS_XDK10.XGIS_LayerKML
124                     ElseIf InStr(UCase(.Filename), ".TAB") Then
126                         Set oLayer = New TatukGIS_XDK10.XGIS_LayerTAB
128                     ElseIf InStr(UCase(.Filename), ".MIF") Then
130                         Set oLayer = New TatukGIS_XDK10.XGIS_LayerMIF
132                     ElseIf InStr(UCase(.Filename), ".GML") Then
134                         Set oLayer = New TatukGIS_XDK10.XGIS_LayerGML
136                     ElseIf InStr(UCase(.Filename), ".DXF") Then
138                         Set oLayer = New TatukGIS_XDK10.XGIS_LayerDXF
                        ElseIf InStr(UCase(.Filename), ".GPX") Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerGPX
                        Else
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerVector
                        End If
                
140                     oLayer.Path = .Filename
                        
141                     Set oLayer = GisUtils.GisCreateLayer(sName, .Filename)

142                     'oLayer.Open
144                     GIS10.Add oLayer
146                     m_oColUserLayers.Add oLayer.Name
148                     FillCOPValues
150                 ElseIf InStr(UCase(GisUtils.GisSupportedFiles(XGIS_FileType.XgisFileTypePixel, False)), UCase(Mid(.Filename, InStr(.Filename, ".")))) > 0 Then
            
                        Dim sExt As String
                        'Check For dublicates
                        sName = Mid$(.FileTitle, 1, Len(.FileTitle) - 4)
313                     Set oLayer = GIS10.get(sName)
                        
315                     If Not oLayer Is Nothing Then 'We Found a dublicate
                            sName = sName & Day(Now()) & Month(Now()) & Minute(Now()) & Second(Now())
                        End If
                        
                        sExt = UCase$(right$(.Filename, Len(.Filename) - InStr(.Filename, ".")))
                
                        If InStr("*.toc)|*.TOC|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerCADRG
                        ElseIf InStr("(*.tif;*.tiff)|*.TIF;*.TIFF|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerTIFF
                        ElseIf InStr("(*.bil)|*.BIL|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerBIL
                        ElseIf InStr("(*.ddf)|*.DDF|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerSDTS_RPE
                        ElseIf InStr("(*.png)|*.PNG|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerPNG
                        ElseIf InStr("(*.psi)|*.PSI|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerPixel
                        ElseIf InStr("(*.sid)|*.SID|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerMrSID
                        ElseIf InStr("(*.jpg;*.jpeg)|*.JPG;*.JPEG|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerJPG
                        ElseIf InStr("(*.jp2;*.j2k;*.jpf;*.jpx;*.jpc)|*.JP2;*.J2K;*.JPF;*.JPX;*.JPC|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerJPG
                        ElseIf InStr("(*.gif)|*.GIF|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerGIF
                        ElseIf InStr("(*.img)|*.IMG|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerIMG
                        ElseIf InStr("(*.ecw)|*.ECW|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerECW
                        ElseIf InStr("(*.bmp)|*.BMP|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerBMP
                        ElseIf InStr("(*.ttkps)|*.TTKPS|", sExt) Then
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerPixelStoreAdo2
                        Else
                            Set oLayer = New TatukGIS_XDK10.XGIS_LayerPixel
                        End If
                        
180                     oLayer.Path = .Filename
182                    Set oLayer = GisUtils.GisCreateLayer(sName, .Filename) ' oLayer.Open
                
184                     GIS10.Add oLayer
186                     FillCOPValues
188                 ElseIf InStr(UCase(GisUtils.GisSupportedFiles(XGIS_FileType.XgisFileTypeProject, False)), UCase(Mid(.Filename, InStr(.Filename, ".")))) > 0 Then
            
190                     GIS10.Open .Filename
            
                    End If
                End If

            Else
192             GIS10.Add oLayer
194             GIS10.UpDate
            End If
            
196         GIS10.UpDate
        End With
        
        
        
        '<EhFooter>
        Exit Sub

AddLayer_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.AddLayer " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub ClearOASISIncidentLyr()
    'm_oIncidentLyr
End Sub

Private Sub JoinAttributeDataToGeometries(sLayer As String, _
                                          oRS As ADODB.Recordset, _
                                          sSpatialJoinKey As String, _
                                          sAttributeJoinKey As String)
        '<EhHeader>
        On Error GoTo JoinAttributeDataToGeometries_Err
        '</EhHeader>
        Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector

100     Set oLyr = GIS10.get(sLayer)

102     If Not oLyr Is Nothing Then

104         With oLyr
106             .JoinADO = oRS
108             .JoinPrimary = sSpatialJoinKey
110             .JoinForeign = sAttributeJoinKey
112             .Paint
            End With

        End If
    
        '<EhFooter>
        Exit Sub

JoinAttributeDataToGeometries_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.JoinAttributeDataToGeometries " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadOASISIncidents()
        '<EhHeader>
        On Error GoTo LoadOASISIncidents_Err
        '</EhHeader>
    
        Dim shpInc As TatukGIS_XDK10.XGIS_ShapePoint
        Dim sWhereClause As String
        
100     ClearOASISIncidentLyr
        
102     Set m_oIncRS = New ADODB.Recordset
    
        'TODO MAKE SURE TO CHECK THE VALUE OF DATE Incident_Date
    
104     Select Case m_frmMnuOperations.abGridTools.Tools.Item("comDateFrom").Text
       
            Case "1 week"

106             If MonthView1.Week = 1 Then
108                 MonthView1.Week = 52
                Else
110                 MonthView1.Week = MonthView1.Week - 1
                End If

112             sWhereClause = "Incident_DATE > #" & MonthView1.value & "#"
                SecurityLayerDateFrom = MonthView1.value
                SecurityLayerDateTill = Now()

                'DebugPrint  1 & "/" & Month( & "/" & Year
114         Case "2 weeks"

116             If MonthView1.Week = 1 Then
118                 MonthView1.Week = 51
120             ElseIf MonthView1.Week = 2 Then
122                 MonthView1.Week = 52
                Else
124                 MonthView1.Week = MonthView1.Week - 2
                End If

126             sWhereClause = "Incident_DATE > #" & MonthView1.value & "#"
                SecurityLayerDateFrom = MonthView1.value
                SecurityLayerDateTill = Now()

128         Case "3 weeks"

130             If MonthView1.Week = 1 Then
132                 MonthView1.Week = 50
134             ElseIf MonthView1.Week = 2 Then
136                 MonthView1.Week = 51
138             ElseIf MonthView1.Week = 3 Then
140                 MonthView1.Week = 52
                Else
142                 MonthView1.Week = MonthView1.Week - 3
                End If

                SecurityLayerDateFrom = MonthView1.value
                SecurityLayerDateTill = Now()
144             sWhereClause = "Incident_DATE > #" & MonthView1.value & "#"

146         Case "1 month"

148             If MonthView1.Month = 1 Then
150                 MonthView1.Month = 12
                Else
152                 MonthView1.Month = MonthView1.Month - 1
                End If
               
                SecurityLayerDateFrom = MonthView1.value
                SecurityLayerDateTill = Now()
154             sWhereClause = "Incident_DATE > #" & MonthView1.value & "#"

156         Case "Custom"
158             frmDatePicker.Show vbModal, Me
               
                SecurityLayerDateFrom = MonthView1.value
                SecurityLayerDateTill = Now()
160             sWhereClause = "Incident_DATE BETWEEN #" & frmDatePicker.dtFrom.value & "# AND #" & frmDatePicker.dtTo.value & "#  "
               
162         Case Else
164             sWhereClause = ""
                SecurityLayerDateFrom = Format("1-1-1900", "Medium Date")
                SecurityLayerDateTill = Now()
166             MonthView1.value = Format(Now(), "dd/mm/yyyy")
        End Select
    
168     With m_oIncRS
170         .Open "SELECT * FROM incReport " & sWhereClause, m_Cnn

172         SafeMoveFirst m_oIncRS

174         If Not .Bof Then

176             Do While Not .EOF

178                 If Not .Fields.Item("Longitude (E)").value = vbNull Or Not .Fields.Item("Latitude (N)").value = vbNull Then
180                     Set shpInc = m_oIncidentLyr.CreateShape(XgisShapeTypePoint)
            
182                     shpInc.Lock TatukGIS_XDK10.XgisLockExtent
184                     shpInc.AddPart
            
186                     With .Fields
188                         shpInc.AddPoint GisUtils.GISPoint(.Item("Longitude (E)").value, .Item("Latitude (N)").value)

                            '"ID"
                            '"Name"
                            '"Type"
                            '"Target"
                            '"Time"
                            '"Decription"
190                         shpInc.SetField "ID", .Item("ID").value
192                         shpInc.SetField "Name", .Item("IncidentID").value
194                         shpInc.SetField "Type", .Item("Incident Type").value
196                         shpInc.SetField "Target", .Item("Incident Target").value
198                         shpInc.SetField "Time", .Item("Incident Time").value
                        End With
                
200                     shpInc.Unlock

202                     m_oIncidentLyr.AddShape shpInc
                    End If

204                 .MoveNext
                Loop

            End If

        End With

        '<EhFooter>
        Exit Sub

LoadOASISIncidents_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.LoadOASISIncidents " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadOASISIncidentsEX(Optional bCommitFilter As Boolean = True, Optional bShowCustomDialog As Boolean = True)
        '<EhHeader>
        On Error GoTo LoadOASISIncidentsEX_Err
        '</EhHeader>
    
        Dim shpInc As TatukGIS_XDK10.XGIS_ShapePoint
        Dim sWhereClause As String
               
100     Set m_oIncRS = New ADODB.Recordset
    
        'TODO MAKE SURE TO CHECK THE VALUE OF DATE Incident_Date
    
102     MonthView1.value = Format(Now(), "yyyy/mm/dd")
    
104     Select Case m_frmMnuOperations.abGridTools.Tools.Item("comDateFrom").Text
       
            Case "1 week"

106             If MonthView1.Week = 1 Then
108                 MonthView1.Week = 52
                    MonthView1.Year = MonthView1.Year - 1
                Else
110                 MonthView1.Week = MonthView1.Week - 1
                End If

                SecurityLayerDateFrom = MonthView1.value
                SecurityLayerDateTill = Now()
112             sWhereClause = "Incident_DATESERIAL > " & ConvertDateToSerial(MonthView1.value)
114             'sWhereClause = "Incident_DATE > #" & MonthView1.Value & "#"

                'DebugPrint  1 & "/" & Month( & "/" & Year
116         Case "2 weeks"

118             If MonthView1.Week = 1 Then
120                 MonthView1.Week = 51
                    MonthView1.Year = MonthView1.Year - 1
122             ElseIf MonthView1.Week = 2 Then
124                 MonthView1.Week = 52
                    MonthView1.Year = MonthView1.Year - 1
                Else
126                 MonthView1.Week = MonthView1.Week - 2
                End If
                
                SecurityLayerDateFrom = MonthView1.value
                SecurityLayerDateTill = Now()
128             sWhereClause = "Incident_DATESERIAL > " & ConvertDateToSerial(MonthView1.value)
                'sWhereClause = "Incident_DATE > #" & MonthView1.Value & "#"

130         Case "3 weeks"

132             If MonthView1.Week = 1 Then
134                 MonthView1.Week = 50
                    MonthView1.Year = MonthView1.Year - 1
136             ElseIf MonthView1.Week = 2 Then
138                 MonthView1.Week = 51
                    MonthView1.Year = MonthView1.Year - 1
140             ElseIf MonthView1.Week = 3 Then
142                 MonthView1.Week = 52
                    MonthView1.Year = MonthView1.Year - 1
                Else
144                 MonthView1.Week = MonthView1.Week - 3
                End If

                SecurityLayerDateFrom = MonthView1.value
                SecurityLayerDateTill = Now()
146             sWhereClause = "Incident_DATESERIAL > " & ConvertDateToSerial(MonthView1.value)
                'sWhereClause = "Incident_DATE > #" & MonthView1.Value & "#"

148         Case "1 month"

150             If MonthView1.Month = 1 Then
152                 MonthView1.Month = 12
                    MonthView1.Year = MonthView1.Year - 1
                Else
                
                    If GetCountOfDaysInMonth(MonthView1.Year, MonthView1.Month - 1) < MonthView1.Day Then
                    
                        MonthView1.Day = GetCountOfDaysInMonth(MonthView1.Year, MonthView1.Month - 1)
                    
                    End If

154                 MonthView1.Month = MonthView1.Month - 1
                    
                End If
               
                SecurityLayerDateFrom = MonthView1.value
                SecurityLayerDateTill = Now()
156             sWhereClause = "Incident_DATESERIAL > " & ConvertDateToSerial(MonthView1.value)
                'sWhereClause = "Incident_DATE > #" & MonthView1.Value & "#"

158         Case "Custom"
160             If bShowCustomDialog Then frmDatePicker.Show vbModal, Me
               
                SecurityLayerDateFrom = frmDatePicker.dtFrom.value
                SecurityLayerDateTill = frmDatePicker.dtTo.value
162             sWhereClause = "Incident_DATESERIAL >= " & ConvertDateToSerial(frmDatePicker.dtFrom.value) & " AND Incident_DATESERIAL <= " & ConvertDateToSerial(frmDatePicker.dtTo.value)
                'sWhereClause = "Incident_DATE BETWEEN #" & frmDatePicker.dtFrom.Value & "# AND #" & frmDatePicker.dtTo.Value & "#  "
               
164         Case Else
166             sWhereClause = ""
                SecurityLayerDateFrom = Format("1-1-1900", "Medium Date")
                SecurityLayerDateTill = Now()
168             MonthView1.value = Format(Now(), "dd/mm/yyyy")
        End Select
        
        'this code was not used since it would make transparent layers not transparent any more!
        '170     m_oSQLIncLyr.scope = sWhereClause  '"Incident_DATE BETWEEN #2007-07-18# AND #2007-07-19#"  '"DATE00 > #2007-07-18#"  sWhereClause

170     m_oSQLIncLyr.scope = sWhereClause  '"Incident_DATE BETWEEN #2007-07-18# AND #2007-07-19#"  '"DATE00 > #2007-07-18#"  sWhereClause

        If m_oSQLIncLyr.Active And bCommitFilter Then
            'had to change to update since it would not always remove the old layer filter
172         'm_oSQLIncLyr.draw
            GIS10.UpDate
        End If

        '<EhFooter>
        Exit Sub

LoadOASISIncidentsEX_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.LoadOASISIncidentsEX " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub CreateW3Layers()
        '<EhHeader>
        On Error GoTo CreateW3Layers_Err
        '</EhHeader>
        Dim RS As New ADODB.Recordset
        Dim WRS As New ADODB.Recordset
        'Dim cn As New Connection
        Dim shpInc As TatukGIS_XDK10.XGIS_Shape
        Dim SymbolList As New TatukGIS_XDK10.XGIS_SymbolList

        'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
        '.RecordSource = "1organisation"
        '.Refresh

100     If m_oW3Lyr Is Nothing Then
            Exit Sub
        End If

102     RS.Open "SELECT * FROM 1placeType", m_Cnn, adOpenDynamic, adLockReadOnly

104     With WRS
106         .Open "SELECT * FROM qryWhere", m_Cnn, adOpenDynamic, adLockOptimistic

108         SafeMoveFirst WRS

110         If Not .Bof Then

112             Do While Not .EOF

114                 If Not .Fields.Item("latX").value = vbNull Or Not .Fields.Item("longY").value = vbNull Then
116                     Set shpInc = m_oW3Lyr.CreateShape(XgisShapeTypePoint)
            
118                     shpInc.Lock TatukGIS_XDK10.XgisLockExtent
120                     shpInc.AddPart
            
122                     With .Fields
124                         shpInc.AddPoint GisUtils.GISPoint(CDbl(Replace(.Item("longY").value, ",", ".")), CDbl(Replace(.Item("latX").value, ",", ".")))

                            '"ID"
                            '"Name"
                            '"Type"
                            '"Target"
                            '"Time"
                            '"Decription"
                            
126                         shpInc.SetField "Name", .Item("Name").value
                            
                            'Add the lookUp type
128                         SafeMoveFirst RS
130                         RS.Find "id = '" & .Item("placeTypeId").value & "'"
                            
132                         shpInc.SetField "Type", RS.Fields.Item("name").value
134                         shpInc.SetField "Description", .Item("Description").value
136                         shpInc.SetField "PCode", "IQ200700" & .Item("pCode").value
                        End With
                
138                     shpInc.Unlock

    '140                     m_oW3Lyr.AddShape shpInc
                    End If

140                 .MoveNext
                Loop

            End If

        End With

        
142     m_oW3Lyr.ParamsList.Add
144     m_oW3Lyr.Params.Query = "Type <> 'home'"
        'sFont = .Item("Font_Name").value
        'sFont = sFont & ":" & .Item("Ascii").value & ":NORMAL"
            
146     m_oW3Lyr.Params.Marker.color = vbWhite
148     m_oW3Lyr.Params.Marker.OutlineColor = vbBlue
        '
150     m_oW3Lyr.Params.Marker.Symbol = SymbolList.Prepare(Replace(g_sAppPath, "\", "\\\\") & "\\\\Data\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE") '"..\\\\..\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE") 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
152     m_oW3Lyr.Params.Marker.Size = 440
154     m_oW3Lyr.Params.Marker.ShowLegend = 1
156     m_oW3Lyr.Params.Legend = "My Activities"
    '        m_oW3Lyr.Params.LabelField = "name"
    '        m_oW3Lyr.Params.LabelVisible = True

158     m_oW3Lyr.Paint
        '<EhFooter>
        Exit Sub

CreateW3Layers_Err:
       ' MsgBox Err.Description & vbCrLf & _
       '        "in OASISClient.frmMain.CreateW3Layers " & _
       '        "at line " & Erl
       ' Resume Next
        Err.Clear
        '</EhFooter>
End Sub

Public Sub CreateW3WhoLayers()
        '<EhHeader>
        On Error GoTo CreateW3WhoLayers_Err
        '</EhHeader>
        Dim WRS As New ADODB.Recordset
        Dim shpInc As TatukGIS_XDK10.XGIS_Shape
        Dim SymbolList As New TatukGIS_XDK10.XGIS_SymbolList
        
100     If m_oW3WHOLyr Is Nothing Then
            Exit Sub
        End If

102     With WRS
104         .Open "SELECT * FROM qryWho", m_Cnn, adOpenDynamic, adLockOptimistic

106         SafeMoveFirst WRS

108         If Not .Bof Then

110             Do While Not .EOF

112                 If Not .Fields.Item("latX").value = vbNull Or Not .Fields.Item("longY").value = vbNull Then
114                     Set shpInc = m_oW3WHOLyr.CreateShape(XgisShapeTypePoint)
            
116                     shpInc.Lock TatukGIS_XDK10.XgisLockExtent
118                     shpInc.AddPart
            
120                     With .Fields
122                         shpInc.AddPoint GisUtils.GISPoint(CDbl(Replace(.Item("longY").value, ",", ".")), CDbl(Replace(.Item("latX").value, ",", ".")))

                            '"PlaceType"
                            '"Cluster"
                            '"offname"

124                         shpInc.SetField "Organisation", .Item("name").value
                            
126                         shpInc.SetField "PlaceType", .Item("PlaceType").value
128                         shpInc.SetField "Cluster", .Item("Cluster").value
130                         shpInc.SetField "offname", .Item("offname").value
                        End With
                
132                     shpInc.Unlock

                    End If

134                 .MoveNext
                Loop

            End If

        End With

136     With m_oW3WHOLyr
138         .ParamsList.Add
140         .Params.Query = "PlaceType <> 'home'"
142         .Params.Marker.color = vbWhite
144         .Params.Marker.OutlineColor = vbBlue
146         .Params.Marker.Symbol = SymbolList.Prepare(Replace(g_sAppPath, "\", "\\\\") & "\\\\Data\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN2-32.BMP?TRUE") '"..\\\\..\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE") 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
148         .Params.Marker.Size = 440
150         .Params.Marker.ShowLegend = 1
152         .Params.Legend = "W3 Who"
154         .Paint
        End With

        '<EhFooter>
        Exit Sub

CreateW3WhoLayers_Err:
        'MsgBox Err.Description & vbCrLf & _
        '       "in OASISClient.frmMain.CreateW3WhoLayers " & _
        '       "at line " & Erl
        'Resume Next
        Err.Clear
        '</EhFooter>
End Sub

Public Sub CreateOASISLayers()
        '<EhHeader>
        On Error GoTo CreateOASISLayers_Err
        '</EhHeader>
        
        On Error Resume Next
        
100     Kill g_sAppPath & "\data\gis\temp\incidents.shp"
102     Kill g_sAppPath & "\data\gis\temp\incidents.dbf"
104     Kill g_sAppPath & "\data\gis\temp\incidents.shx"
        
106     Kill g_sAppPath & "\data\gis\temp\operations.shp"
108     Kill g_sAppPath & "\data\gis\temp\operations.dbf"
110     Kill g_sAppPath & "\data\gis\temp\operations.shx"
    
112     SafeMoveFirst g_RSAppSettings
114     g_RSAppSettings.Find "SettingName = 'stdLyrs'"
    
116     If g_RSAppSettings.Fields.Item("SettingValue1").value = "1" Then
    
118         Set m_oIncidentLyr = GisUtils.GisCreateLayer("OASIS Incidents", g_sAppPath & "\data\gis\temp\incidents.shp")
        
120         m_oIncidentLyr.AddField "ID", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
122         m_oIncidentLyr.AddField "Name", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
124         m_oIncidentLyr.AddField "Type", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
126         m_oIncidentLyr.AddField "Target", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
128         m_oIncidentLyr.AddField "Time", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
130         m_oIncidentLyr.AddField "Description", TatukGIS_XDK10.XgisFieldTypeString, 255, 0
        End If
        
132     If g_RSAppSettings.Fields.Item("SettingValue2").value = "1" Then
134         Set m_oW3Lyr = New TatukGIS_XDK10.XGIS_LayerVector
  
136         m_oW3Lyr.AddField "Name", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
138         m_oW3Lyr.AddField "Type", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
140         m_oW3Lyr.AddField "Description", TatukGIS_XDK10.XgisFieldTypeString, 255, 0
142         m_oW3Lyr.AddField "PCode", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
    
144         m_oW3Lyr.Params.area.color = RGB(0, 0, 255)
146         m_oW3Lyr.Params.Label.Field = "Name"
148         m_oW3Lyr.Params.Label.Visible = "YES"
            'Params.MarkerSymbol := SynbolList.Prepare( 'mysymbol.bmp?FALSE' );
            'm_oW3Lyr.Params.MarkerSymbol = "..\\\\..\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE"

            'm_oW3Lyr.Transparency = 50
150         m_oW3Lyr.Name = "Who What Where"
152         m_oW3Lyr.HideFromLegend = False

        End If

154     If g_RSAppSettings.Fields.Item("SettingValue3").value = "1" Then

156         Set m_oW3WHOLyr = New TatukGIS_XDK10.XGIS_LayerVector
158         m_oW3WHOLyr.AddField "Organisation", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
160         m_oW3WHOLyr.AddField "PlaceType", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
162         m_oW3WHOLyr.AddField "Cluster", TatukGIS_XDK10.XgisFieldTypeString, 255, 0
164         m_oW3WHOLyr.AddField "offname", TatukGIS_XDK10.XgisFieldTypeString, 100, 0
    
166         m_oW3WHOLyr.Params.area.color = RGB(0, 0, 255)
168         m_oW3WHOLyr.Params.Label.Field = "Organisation"
170         m_oW3WHOLyr.Params.Label.Visible = "YES"
            'Params.MarkerSymbol := SynbolList.Prepare( 'mysymbol.bmp?FALSE' );
            'm_oW3Lyr.Params.MarkerSymbol = "..\\\\..\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE"

            'm_oW3Lyr.Transparency = 50
172         m_oW3WHOLyr.Name = "W3-Who"
174         m_oW3WHOLyr.HideFromLegend = False
        End If

        CreateAddedScribbles

'176     Set m_oDrawLyr = New TatukGIS_XDK10.XGIS_LayerVector
'178     m_oDrawLyr.Params.area.Color = RGB(0, 0, 255)
'180     m_oDrawLyr.Transparency = 90
'182     m_oDrawLyr.Name = "Draw_Layer"
'184     m_oDrawLyr.HideFromLegend = True
    
186     Set m_oCosmeticLayer = New TatukGIS_XDK10.XGIS_LayerVector

188     With m_oCosmeticLayer
190         With .Params
    
192             .area.color = RGB(0, 0, 255)
    
            End With
    
194         .Transparency = 50
196         .Name = "Cosmetic"
198         .HideFromLegend = True
200         .AddField "cosLabel", TatukGIS_XDK10.XgisFieldTypeString, 255, 0
202         .AddField "cosValue", TatukGIS_XDK10.XgisFieldTypeNumber, 16, 0
204         .AddField "cosText", TatukGIS_XDK10.XgisFieldTypeString, 255, 0
206         .AddField "cosDescription", TatukGIS_XDK10.XgisFieldTypeString, 100, 0

        End With
  
208     Set m_oBufferLyr = New TatukGIS_XDK10.XGIS_LayerVector
210     m_oBufferLyr.Params.area.color = RGB(0, 0, 255)
212     m_oBufferLyr.Params.area.OutlineColor = RGB(0, 0, 255)
214     m_oBufferLyr.Transparency = 60
216     m_oBufferLyr.Name = "Buffers"
218     m_oBufferLyr.HideFromLegend = True

        On Error Resume Next
220     If g_RSAppSettings.Fields.Item("SettingValue1").value = "1" Then m_oIncidentLyr.SaveAll
    
        'shape.SetField( "name", "some text" )
    
        '
        '      shp := lv.CreateShaoe( gisShapeTypePoint );
        '  shp.Lock( gisLockExtent ) ;
        '  shp.AddPart() ;
        '  shp.AddPoint( GisPoint( x, y ) ) ;
        '  shp.Unlock ;
        '  shp.SetField( 'name', name ) ;
        '  shp.SetField( 'address', address ) ;
    
        '  shp = Layer.CreateShape(XgisShapeTypePoint)
        'shp.Lock (XgisLockExtent)
        'shp.AddPart()
        'shp.AddPoint (GisPoint(3, 3))
        'shp.Unlock()
        '<EhFooter>
        Exit Sub

CreateOASISLayers_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CreateOASISLayers " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadMapThread(sMAP As Variant)
    On Error Resume Next
    GIS10.Open sMAP, False
End Sub

Public Sub LoadMapProducts(Optional sInitialMapName As String)
        '<EhHeader>
        On Error GoTo LoadMapProducts_Err
        '</EhHeader>

100     If Not sInitialMapName = "" Then
            On Error Resume Next
            'GIS109 Done
            GIS10.Open sInitialMapName, True 'False
        Else
            Exit Sub
        End If
        '<EhFooter>
        Exit Sub

LoadMapProducts_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.LoadMapProducts " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetLayerFilter(oLayer As TatukGIS_XDK10.XGIS_LayerVector, sFieldName As String, sVal As String)
        '<EhHeader>
        On Error GoTo SetLayerFilter_Err
        '</EhHeader>
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
        
100     For Each oShp9 In oLayer.Loop(oLayer.Extent, "", Nothing, "", True)
104         If oShp9.GetField(sFieldName) = sVal Then
106             oShp9.IsHidden = True
            End If
        Next
        '<EhFooter>
        Exit Sub

SetLayerFilter_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SetLayerFilter " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub miniMapRefresh()
        '<EhHeader>
        On Error GoTo miniMapRefresh_Err
        '</EhHeader>
            Dim ptg1 As TatukGIS_XDK10.XGIS_Point
            Dim ptg2 As TatukGIS_XDK10.XGIS_Point
            Dim ptg3 As TatukGIS_XDK10.XGIS_Point
            Dim ptg4 As TatukGIS_XDK10.XGIS_Point
            Dim ex As TatukGIS_XDK10.XGIS_Extent

    Exit Sub

100         If GIS10.IsEmpty Then Exit Sub

102         Set ex = GIS10.VisibleExtent
104         Set ptg1 = GisUtils.GISPoint(ex.xmin, ex.ymin)
106         Set ptg2 = GisUtils.GISPoint(ex.xmax, ex.ymin)
108         Set ptg3 = GisUtils.GISPoint(ex.xmax, ex.ymax)
110         Set ptg4 = GisUtils.GISPoint(ex.xmin, ex.ymax)

112         minishp.Reset
114         minishp.Lock TatukGIS_XDK10.XgisLockExtent
116         minishp.AddPart
118         minishp.AddPoint ptg1
120         minishp.AddPoint ptg2
122         minishp.AddPoint ptg3
124         minishp.AddPoint ptg4
126         minishp.Unlock

128         minishpo.Reset
130         minishpo.Lock TatukGIS_XDK10.XgisLockExtent
132         minishpo.AddPart
134         minishpo.AddPoint ptg1
136         minishpo.AddPoint ptg2
138         minishpo.AddPoint ptg3
140         minishpo.AddPoint ptg4
142         minishpo.AddPoint ptg1
144         minishpo.Unlock

146         m_frmOvMap.GISm.UpDate
        '<EhFooter>
        Exit Sub

miniMapRefresh_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.miniMapRefresh " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
    End Sub



Private Sub MouseUp(ByVal x As Integer, ByVal y As Integer)
        '<EhHeader>
        On Error GoTo MouseUp_Err
        '</EhHeader>
            Dim ptg As TatukGIS_XDK10.XGIS_Point
            Dim P1, p2, p3 As TatukGIS_XDK10.XGIS_Point
            Dim p4 As New TatukGIS_XDK10.XGIS_Point

100         Set ptg = m_frmOvMap.GISm.ScreenToMap(GisUtils.POINT(x, y))
102         minishp.SetPosition miniRecalc(ptg), m_frmOvMap.GISm.get(MINIMAP_R_NAME), 5
104         m_frmOvMap.GISm.UpDate
106         fminiMove = False
108         Set P1 = minishp.GetPoint(0, 0)
110         Set p2 = minishp.GetPoint(0, 1)
112         Set p3 = minishp.GetPoint(0, 2)
114         p4.x = P1.x + (p2.x - P1.x) / 2
116         p4.y = P1.y + (p3.y - p2.y) / 2
118         GIS10.Center = p4
        '<EhFooter>
        Exit Sub

MouseUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.MouseUp " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub OpenSQLDBLayer()
'Q10701
'Is there a way to open an SQL layer other than using the .TTKLS file? I want to access features directly from the database, avoiding the .ttkls file.


'Embed into the path parameter the content of the TTKLS with a line separated by "\n" ad with last line "\n.ttkls".

'So you can use:

' GIS10.Open ("Storage=Native\nLAYER=lwaters\nDIALECT=MSJET\nADO=Provider=Microsoft.Jet.OLEDB.4.0;Data Source=gistest.mdb\n.ttkls")

'Or just:

' ll = GisCreateLayer("somename", "Storage=Native\nLAYER=lwaters\nDIALECT=MSJET\nADO=Provider=Microsoft.Jet.OLEDB.4.0;Data Source=gistest.mdb\n.ttkls")

'So just embed the content of TTKLS file with line separated by "\n" ad with last line "\n.ttkls".

'You can use same mechanism for the PixelStore TTKPS file.

End Sub

Private Sub abCOP_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)
        '<EhHeader>
        On Error GoTo abCOP_ToolClick_Err
        '</EhHeader>

100     Select Case Tool.Name

        End Select

        '<EhFooter>
        Exit Sub

abCOP_ToolClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.abCOP_ToolClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub CloneLyrSettings(oLyrToCopyFrom As TatukGIS_XDK10.XGIS_LayerVector, oLyrToAssign As TatukGIS_XDK10.XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo CloneLyrSettings_Err
        '</EhHeader>
100     oLyrToAssign.ParamsList.Assign_ oLyrToCopyFrom.ParamsList
102     oLyrToAssign.draw
        '<EhFooter>
        Exit Sub

CloneLyrSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CloneLyrSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetMapExtent(oXtent As TatukGIS_XDK10.XGIS_Extent)
        '<EhHeader>
        On Error GoTo SetMapExtent_Err
        '</EhHeader>
100     GIS10.VisibleExtent = oXtent
        '<EhFooter>
        Exit Sub

SetMapExtent_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SetMapExtent " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadLayerAttrDataToGridInit()
        '<EhHeader>
        On Error GoTo LoadLayerAttrDataToGridInit_Err
        '</EhHeader>

        Dim GlobalGISGridLayerOriginal As TatukGIS_XDK10.XGIS_LayerVector
        Dim lMilliSecsTimed As Long
        Dim iTot As Long
        Dim sExcludedFldsString As String
        Dim sLayerName As String
    
100     lMilliSecsTimed = GetTickCount

102     If m_bLOADING Then Exit Sub
104     SafeMoveFirst g_RSGISGridTableSettings

        'Clear the guy
106     If abGridTools.Tools.Item("comLyr").Text = "---Nothing---" Then
108         dxGISDataGrid.Visible = False
110         dxGISDataGrid.Columns.DestroyColumns
112         Set dxGISDataGrid.DataSource = Nothing
114         dxGISDataGrid.Visible = True
            Exit Sub
        End If
    
116     g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"
    
        'Get the layer data
118     If Not g_RSGISGridTableSettings.EOF And Not g_RSGISGridTableSettings.Bof Then
120         sLayerName = g_RSGISGridTableSettings.Fields("Name").value
122         Set GlobalGISGridLayerOriginal = GIS10.get(g_RSGISGridTableSettings.Fields("Name").value)
        Else
124         sLayerName = abGridTools.Tools.Item("comLyr").Text
126         Set GlobalGISGridLayerOriginal = GIS10.get(abGridTools.Tools.Item("comLyr").Text)
        End If
                
        GISGridLayerName = sLayerName
    
        'Layer not found?
128     If GlobalGISGridLayerOriginal Is Nothing Then
130         MsgBox "The Data You are trying to browse does not anymore exist in the map." & vbCrLf & "Either it has been removed or the data is corrupt." & vbCrLf & "Contact your OASIS adminstrator if this problem remains."
132         Set GlobalGISGridLayerOriginal = Nothing
            Exit Sub
        Else
            If Not GlobalGISGridLayerOriginal.Params.Visible Then
                If MsgBox("The layer is not visible. Would you like you to make it visible?", vbYesNo, "OASIS Client") = vbYes Then
                    GlobalGISGridLayerOriginal.Params.Visible = True
                End If
            End If
        End If
    
134     GlobalGISGridLayerOriginal.Lock
136     Set GlobalGISGridLayer = New TatukGIS_XDK10.XGIS_LayerVector
138     dxProgressBar1.Visible = True
140     dxGISDataGrid.Visible = False
142     dxProgressBar1.Pos = 0
144     elMap.Refresh
        
146     If chkOnlyVisible.value = vbChecked Then
156         Set GISGridExtent = GIS10.viewer.VisibleExtent
        'ElseIf GIS10.viewer.VisibleExtent.XMin > GlobalGISGridLayer.Extent.XMin Or GIS10.viewer.VisibleExtent.XMax < GlobalGISGridLayer.Extent.XMax Or GIS10.viewer.VisibleExtent.YMin > GlobalGISGridLayer.Extent.YMin Or GIS10.viewer.VisibleExtent.YMax < GlobalGISGridLayer.Extent.YMax Then
            'oExtent = GlobalGISGridLayerOriginal.Extent
        Else
158         Set GISGridExtent = GlobalGISGridLayerOriginal.Extent
        End If
        
160     'GlobalGISGridLayerOriginal.ExportLayer GlobalGISGridLayer, GISGridExtent, TatukGIS_XDK10.XgisShapeTypeUnknown, GlobalGISGridLayerOriginal.scope, False
162     'GlobalGISGridLayerOriginal.Unlock

Set GlobalGISGridLayer = GlobalGISGridLayerOriginal
        GlobalGISGridLayer.ScopeExtent.Assign_ GISGridExtent
        GlobalGISGridLayer.ScopeExtent = GISGridExtent

        'GIS10
        'WE need to check this line when upgrading to tatuk10 - if should catch sql layers
164     If GlobalGISGridLayerOriginal.ConfigName = "_DUMMY.ini" Then 'And Not GlobalGISGridLayerOriginal.Name = "Incidents" Then

166         GlobalGISGridLayer.caption = GlobalGISGridLayerOriginal.caption
168         GlobalGISGridLayer.Name = GlobalGISGridLayerOriginal.Table
170         sLayerName = GlobalGISGridLayerOriginal.Name

        End If
        
172    ' Set GlobalGISGridLayerOriginal = Nothing
    
174     With g_RSGISGridTableSettings
    
176         SafeMoveFirst g_RSGISGridTableSettings
178         .Find "Name = '" & sLayerName & "'"
            
180         If Not .Bof And Not .EOF Then
        
182             DebugPrint GlobalGISGridLayer.items.Count
184             iTot = GlobalGISGridLayer.GetLastUid
        If GlobalGISGridLayer.ConfigName = "_DUMMY.ini" Then
            Dim oRS As ADODB.Recordset
            Set oRS = New ADODB.Recordset
            Dim oCn As ADODB.Connection
            Set oCn = New ADODB.Connection
            oCn.Open GlobalGISGridLayer.SQLParameter("ADO")
            oRS.Open "SELECT count(*) from [" & GlobalGISGridLayer.Name & "_FEA]", oCn, adOpenDynamic, adLockBatchOptimistic
            
            If Not oRS.EOF Then
            iTot = oRS.Fields(0).value
            End If
            Set oCn = Nothing
            Set oRS = Nothing
            
        End If
                'Get list of excluded fields
            
186             If Not .EOF Then
188                 If Not .Fields.Item("excludedFlds").value = vbNull Then
190                     sExcludedFldsString = .Fields.Item("excludedFlds").value
                    End If
                End If
        
192             If Not .EOF Then
            
                    'Check count of records
194                 If .Fields.Item("datasetwarning").value Then
                        
196                     DebugPrint "datasetwarning -- iTot: " & iTot
198                     DebugPrint "datasetwarning -- .Fields.Item('warninglevel').Value: " & .Fields.Item("warninglevel").value
                
                        'Too many records - ABORT!
200                     If iTot > .Fields.Item("MaxRec").value Then
202                         MsgBox "Note! According to performace settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
204                         Set GlobalGISGridLayer = Nothing
                            Exit Sub

                        End If

                        'Many records - ask the question
206                     If iTot > .Fields.Item("warninglevel").value Then
208                         If MsgBox("The data you are about to browse contains: " & iTot & " records." & vbCrLf & "This might take very long time. Would you like to continue?", vbYesNo) = vbNo Then
210                             Set GlobalGISGridLayer = Nothing
                                Exit Sub
                            End If
                        End If
                    
                    Else
212                     DebugPrint "no datasetwarning"

                        'Too many records - ABORT!
214                     If iTot > .Fields.Item("MaxRec").value Then
216                         MsgBox "Note! According to performance settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
218                         Set GlobalGISGridLayer = Nothing
                            Exit Sub
                        End If
                    End If

                Else
                    'UserDefGridRecords
                    'the Layer is userdrawn, Use default values from AppSettings
220                 SafeMoveFirst g_RSAppSettings
222                 g_RSAppSettings.Find "SettingName = 'UserDefGridRecords'"
                
224                 If g_RSAppSettings.Fields.Item("SettingValue1").value = "1" Then
                    
226                     If iTot > CLng(g_RSAppSettings.Fields.Item("SettingValue2").value) Then
                        
                            'Too many records - ABORT!
228                         MsgBox "Note! According to performace settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
230                         Set GlobalGISGridLayer = Nothing
                            Exit Sub
                        
                        Else

                            'Too many records - ABORT!
232                         If iTot > CLng(g_RSAppSettings.Fields.Item("SettingValue3").value) Then
234                             If MsgBox("The data you are about to browse contains:" & iTot & " records." & vbCrLf & "This might take very long time. Would you like to continue?", vbYesNo) = vbNo Then
236                                 Set GlobalGISGridLayer = Nothing
                                    Exit Sub
                                End If
                            End If
                        
                        End If
                    
                    End If
                
                End If
            
            End If
        
        End With
    
238     If bDebugMode Then
240         LoadLayerAttrDataToGrid sExcludedFldsString
        Else

242         If Not pLoadLyrAttrToGrdThread.IsThreadRunning Then
244             pLoadLyrAttrToGrdThread.CreateWin32Thread Me, "LoadLayerAttrDataToGrid", sExcludedFldsString
            End If
        End If

        On Error Resume Next
246     DebugPrint "GIS load operation took " & Round(((GetTickCount - lMilliSecsTimed) / 1000), 2) & " seconds"
     
        '<EhFooter>
        Exit Sub

LoadLayerAttrDataToGridInit_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.LoadLayerAttrDataToGridInit " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CreateCustomRS(oRS As ADODB.Recordset, Optional bCopyData As Boolean = True) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo CreateCustomRS_Err
        '</EhHeader>

        Dim sVisibleFields As String
        Dim NewRS As New ADODB.Recordset
        Dim i As Long
    
100     i = 0
                
102     Do Until i = oRS.Fields.Count
            
104        ' If InStr(UCase(sVisibleFields), UCase(oRS.Fields(i).Name)) > 0 Then

106             With oRS.Fields(i)
108                 NewRS.Fields.Append .Name, IIf(.Type = adLongVarWChar, adVarWChar, .Type), IIf(.DefinedSize > 8100, 8100, .DefinedSize)
                End With
            
           ' End If

110         i = i + 1
        
        Loop
        
                'add x and y
        NewRS.Fields.Append "X_", adDouble
        NewRS.Fields.Append "Y_", adDouble

        If Not bCopyData Then Set CreateCustomRS = NewRS
112     If Not NewRS.Fields.Count = 0 And bCopyData Then
    
114         NewRS.Open
116         'NewRS.AddNew
118         CopyRSValues oRS, NewRS
120         Set CreateCustomRS = NewRS
    
        End If

        '<EhFooter>
        Exit Function

CreateCustomRS_Err:
        MsgBox "frmMain.CreateCustomRS_Err (" & Erl & " ) " & Err.Description
        
        '</EhFooter>
End Function

Private Sub CopyRSValues(SourceRS As ADODB.Recordset, _
                         DestRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo CopyRSValues_Err
        '</EhHeader>

        On Error Resume Next
        Dim i As Long
        
        

102     If Not SourceRS Is Nothing Then

104         If Not SourceRS.EOF And Not SourceRS.Bof Then

                Do Until SourceRS.EOF
                
                    i = 0
                    DestRS.AddNew
                    
106                 Do Until i >= SourceRS.Fields.Count
    
108                     With DestRS.Fields(SourceRS.Fields(i).Name)
                            
112                         .value = Empty
114                         .value = SourceRS.Fields(SourceRS.Fields(i).Name).value
    
                        End With
    
116                     i = i + 1
                    Loop
                    
                    SourceRS.MoveNext
                Loop

            End If

        End If

        '<EhFooter>
        Exit Sub

CopyRSValues_Err:
        MsgBox "frmMain.CopyRSValues_Err (" & Erl & " ) " & Err.Description
        
    '</EhFooter>
End Sub

Public Sub LoadLayerAttrDataToGrid(sExcludedFieldsPassed As Variant)
        '<EhHeader>
        On Error GoTo LoadLayerAttrDataToGrid_Err
        '</EhHeader>
        
        Dim oShapes As XGIS_LayerVectorEnumeratorFactory
        Dim lngFldType As Long
        Dim lngDataLength As Long
        Dim flds() As String
        Dim Col As DXDBGRIDLibCtl.dxGridColumn ' Variant
        Dim k As Long
        Dim i As Long
        Dim j As Long
        Dim iFieldIndex As Integer
        Dim iFieldCount As Integer
        Dim sCnn As String
        Dim strVals As Variant
        Dim arVals As Variant
        Dim arVal As Variant
        Dim arFieldNames As Variant
        Dim varVal As Variant
        Dim fldPos() As Integer
        Dim bExclude As Boolean
        Dim iCountOfFields As Integer
        Dim lUID As Long
        Dim bSQLLayer As Boolean
        
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
        Dim oASHP9 As TatukGIS_XDK10.XGIS_Shape
        Dim lLayer As TatukGIS_XDK10.XGIS_LayerVector
        Dim lLayerTemp As TatukGIS_XDK10.XGIS_LayerVector
        
        Dim oCn As New ADODB.Connection
        Dim oRS As New ADODB.Recordset
        Dim oRSNew As New ADODB.Recordset
                
100     bSQLLayer = False
        'this breaks support for old sql layers
102     If GlobalGISGridLayer.ConfigName = "_DUMMY.ini" Then
104         bSQLLayer = True
        End If
    
        Dim sExcludedFlds() As String

106     If ((Not sExcludedFlds) = -1&) Then
108         sExcludedFlds = Split(sExcludedFieldsPassed, ",")
        End If

110     With dxGISDataGrid
        
112         .Visible = False
114         .Columns.DestroyColumns
116         Set .DataSource = Nothing
118         Set GISGridRS = New ADODB.Recordset

120         Set dxProgressBar1.Container = dxGISDataGrid.Container
122         dxProgressBar1.left = dxGISDataGrid.left
124         dxProgressBar1.top = dxGISDataGrid.top
126         dxProgressBar1.Width = dxGISDataGrid.Width
128         dxProgressBar1.Height = dxGISDataGrid.Height
130         dxProgressBar1.MinPos = 0
132         dxProgressBar1.MaxPos = dxProgressBar1.MinPos + 1 + GlobalGISGridLayer.GetLastUid
134         dxProgressBar1.Pos = 0
136         dxProgressBar1.Step = 1
138         dxProgressBar1.Visible = True
140         dxProgressBar1.DoStep
142         elMap.Refresh
            
            'LockWindowUpdate elGisAttr.Hwnd
144         If mnuAutoSize.Checked Then
146             .Options.Set (egoAutoWidth)
            Else
148             .Options.Unset (egoAutoWidth)
            End If
            
            If mnuAutoHeight.Checked Then
                .Options.Set (egoRowAutoHeight)
            Else
                .Options.Unset (egoRowAutoHeight)
            End If
            
150         .Options.Set (egoShowGroupPanel)
152         .Options.Set (egoBandMoving)
154         .Options.Set (egoColumnMoving)
156         .Options.Set (egoMultiSort)
158         .Options.Set (egoShowFooter)
160         .Options.Set (egoAutoSort)
162         .Options.Set (egoShowButtons)
164         .Options.Set (egoShowRowFooter)
166         .Options.Set (egoAutoSearch)
168         .Options.Set (egoAutoExpandOnSearch)
170         .Options.Set (egoAnsiSort)
172         .Options.Set (egoLoadAllRecords)
174         .Options.Set (egoAutoSearch)
176         .Options.Unset (egoCanNavigation)
178         .Options.Unset (egoDynamicLoad)
180         .DatasetType = dtADODataset
182         .Filter.FilterActive = True
184         .Filter.FilterStatus = fsAlways
    
186         ReDim flds(0)

188         If bSQLLayer Then
                
190             Set lLayer = GIS10.get(GISGridLayerName)
192             'oRS.Open "SELECT * FROM [ttkGISLayerSQLInProject] WHERE [LayerCaption] = '" & GlobalGISGridLayer.caption & "'", m_Cnn, adOpenDynamic, adLockReadOnly
                 
194             'If Not oRS.EOF Or lLayer.Name = "Incidents" Then
                    
196                 If lLayer.Name = "Incidents" Then
198                     sCnn = g_sGlobalConnectionString 'IIf(bSQLServerInUse, g_sGlobalConnectionString, Replace(oRS.Fields("ADO").Value, "CLIENTDBPATH", ))

                    Else
                    
200                     sCnn = lLayer.SQLParameter("ADO") ' IIf(bSQLServerInUse, g_sGlobalConnectionString, Replace(oRS.Fields("ADO").value, "CLIENTDBPATH", g_sAppPath))

                    End If

202                 'oRS.Close

                    If g_bIncidentsV2 And GlobalGISGridLayer.Name = "dd_" & g_sIncidentsV2DDName & "_qryIncidents" Then
                        oCn.Open g_sIncidentsV2ConnectionString
                    Else
204                     If bSQLServerInUse Then
206                         oCn.Open GetConnectionString("")
                        Else
208                         oCn.Open sCnn
                        End If
                    End If
                    
                    'oCn.CursorLocation = adUseClient
210                 Set oRS = New ADODB.Recordset
212                 oRS.Open "SELECT top 1 UID as GIS_UID, * FROM [" & GlobalGISGridLayer.Name & "_FEA]", oCn, adOpenForwardOnly, adLockBatchOptimistic
                    
214                 If Not oRS.EOF Then
                    
216                     Set oRSNew = CreateCustomRS(oRS, False)
218                     oRSNew.Open

220                     iFieldCount = oRSNew.Fields.Count
222                     j = 0
                            
224                     dxProgressBar1.MaxPos = 1
226                     dxProgressBar1.Step = 1
228                     dxProgressBar1.Pos = 0
                        
                        Set lLayerTemp = New TatukGIS_XDK10.XGIS_LayerVector
                        lLayer.ExportLayer lLayerTemp, GIS10.VisibleExtent, XgisShapeTypeUnknown, "", False
                        If lLayerTemp.items.Count > 0 Then
                            dxProgressBar1.MaxPos = lLayerTemp.items.Count
                        Else
                            dxProgressBar1.DoStep
                        End If
                        
230                     Set oShapes = lLayer.Loop(GISGridExtent, "", Nothing, "", False)
                        
232                     'For Each oShp9 In oShapes ' lLayer.Loop(oExtent, "", Nothing, "", False)
234                         'dxProgressBar1.MaxPos = dxProgressBar1.MaxPos + 1
                        'Next
                        
236                     For Each oShp9 In oShapes 'lLayer.Loop(oExtent, "", Nothing, "", False)
238                         oRSNew.AddNew

240                         iFieldIndex = 0
                            
242                         Do Until iFieldIndex = iFieldCount - 2
244                             oRSNew.Fields(iFieldIndex).value = oShp9.GetField(oRSNew.Fields(iFieldIndex).Name)
246                             iFieldIndex = iFieldIndex + 1
                            Loop
                            
                            oRSNew.Fields(iFieldCount - 2).value = oShp9.Centroid.x
                            oRSNew.Fields(iFieldCount - 1).value = oShp9.Centroid.y

                            'oRSNew.Fields("GIS_UID").Value = oShp9.uID
248                         dxProgressBar1.DoStep
250                         elMap.Refresh
                            
252                         j = j + 1
                        Next
                        
254                     Set oShapes = Nothing
                        
256                     If j > 0 Then
                        
258                         Set GISGridRS = oRSNew
260                         Set dxGISDataGrid.DataSource = oRSNew
262                         Set GlobalGISGridLayer = lLayer
264                         dxGISDataGrid.Columns.RetrieveFields
                    
266                         For i = 0 To dxGISDataGrid.Columns.Count - 1

                                If dxGISDataGrid.Columns(i).ColumnType = gedMemoEdit Then
                                    dxGISDataGrid.Columns(i).MemoColumn.WordWrap = True
                                    dxGISDataGrid.Columns(i).MemoColumn.ScrollBars = ssHorizontal
                                End If
                                
268                             If Not ((Not sExcludedFlds) = -1&) Then
270                                 If SequentialSearchStringArray(sExcludedFlds, dxGISDataGrid.Columns(i).FieldName) > -1 Then
272                                     dxGISDataGrid.Columns(i).Visible = False
                                    End If
                                End If
                
                            Next
                            .Columns(.Columns.Count - 2).Visible = mnuShowCentroid.Checked
                            .Columns(.Columns.Count - 1).Visible = mnuShowCentroid.Checked
                        Else
274                         MsgBox "There are no shapes in the map!", vbInformation
                        End If
                    
                    Else
276                     Set dxGISDataGrid.DataSource = Nothing
                    End If
                 
                    'Else

278                 If Not oRS.State = 0 Then oRS.Close
280                 Set oRS = Nothing
                'End If
                
            Else

282             GISGridRS.Fields.Append "GIS_UID", adBigInt, 100
284             iCountOfFields = 1
            
286             For i = 0 To GlobalGISGridLayer.Fields.Count - 1

288                 bExclude = False
                
290                 If Not ((Not sExcludedFlds) = -1&) Then
292                     If SequentialSearchStringArray(sExcludedFlds, GlobalGISGridLayer.Fields.Item(i).Name) > -1 Then
294                         bExclude = True
                        End If
                    End If
                
296                 If Not bExclude Then

298                     iCountOfFields = iCountOfFields + 1
                
300                     ReDim Preserve flds(UBound(flds) + 1)
302                     ReDim Preserve fldPos(UBound(flds))
                    
304                     flds(UBound(flds)) = GlobalGISGridLayer.Fields.Item(i).Name '(i + 1) to Compensate for the GIS_UID
                        'Stop
306                     fldPos(UBound(flds)) = i
                    
308                     Select Case GlobalGISGridLayer.Fields.Item(i).FieldType
        
                            Case Is = TatukGIS_XDK10.XgisFieldTypeString '= 0,
310                             lngFldType = adVarWChar  'xftString
312                             lngDataLength = GlobalGISGridLayer.Fields.Item(i).Width
        
314                         Case Is = TatukGIS_XDK10.XgisFieldTypeNumber ' = 1,
316                             lngFldType = adDouble 'xftFloat
318                             lngDataLength = GlobalGISGridLayer.Fields.Item(i).Width
        
320                         Case Is = TatukGIS_XDK10.XgisFieldTypeFloat '= 2,
322                             lngFldType = adDouble 'xftFloat
324                             lngDataLength = GlobalGISGridLayer.Fields.Item(i).Width
        
326                         Case Is = TatukGIS_XDK10.XgisFieldTypeBoolean '= 3,
328                             lngFldType = adBoolean 'xftBoolean
330                             lngDataLength = GlobalGISGridLayer.Fields.Item(i).Width
        
332                         Case Is = TatukGIS_XDK10.XgisFieldTypeDate '= 4
334                             lngFldType = adDate 'xftDate
336                             lngDataLength = GlobalGISGridLayer.Fields.Item(i).Width
                        End Select
                    
338                     GISGridRS.Fields.Append GlobalGISGridLayer.Fields.Item(i).Name, lngFldType, lngDataLength

                    End If
                
                Next
                
340             arVals = Split(strVals, ",")
342             arFieldNames = Split(strVals, ",")
344             strVals = ""
346             ReDim arVal(UBound(flds))

                'GIS109 TODO Check if this is needed
                '326             If Not GlobalGISGridLayer.Shape Is Nothing Then arVal(0) = GlobalGISGridLayer.Shape.uID
348             j = 0
                GISGridRS.Fields.Append "X_", adDouble
                GISGridRS.Fields.Append "Y_", adDouble
350             GISGridRS.Open
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Populate the RS
                '

352             For Each oASHP9 In GlobalGISGridLayer.Loop(GISGridExtent, "", Nothing, "", True)
                    'LockWindowUpdate 0
354                 dxProgressBar1.DoStep
356                 elMap.Refresh
358                 GISGridRS.AddNew
            
360                 For i = 0 To iCountOfFields - 3 'UBound(flds)

362                     Select Case GlobalGISGridLayer.Fields.Item(fldPos(i)).FieldType 'i to compensate the GIS_UID on 0 position
    
                            Case Is = TatukGIS_XDK10.XgisFieldTypeString '= 0,
364                             varVal = CStr(oASHP9.GetField(flds(i)))
366                             arVal(i) = varVal
    
368                         Case Is = TatukGIS_XDK10.XgisFieldTypeNumber ' = 1,
370                             arVal(i) = CDbl(oASHP9.GetField(flds(i)))
    
372                         Case Is = TatukGIS_XDK10.XgisFieldTypeFloat '= 2,
374                             arVal(i) = CDbl(oASHP9.GetField(flds(i)))
    
376                         Case Is = TatukGIS_XDK10.XgisFieldTypeBoolean '= 3,
378                             arVal(i) = CBool(oASHP9.GetField(flds(i)))
    
380                         Case Is = TatukGIS_XDK10.XgisFieldTypeDate '= 4
382                             arVal(i) = CDate(oASHP9.GetField(flds(i)))
                        End Select
                    
384                     GISGridRS.Fields(i).value = arVal(i)
                
                    Next

386                 j = j + 1

388                 If Not oASHP9 Is Nothing Then
390                     lUID = oASHP9.uID
'Add Centroid to the grid
392                     GISGridRS.Fields(0).value = lUID
                        GISGridRS.Fields(i + 2).value = oASHP9.Centroid.x
                        GISGridRS.Fields(i + 3).value = oASHP9.Centroid.y

                    End If
                    
                Next

394             If j > 0 Then Set .DataSource = GISGridRS

            End If

396         If SafeMoveFirst(GISGridRS) And Not .DataSource Is Nothing Then

398             If Not bSQLLayer Then

400                 Set .DataSource = GISGridRS
402                 .Columns.RetrieveFields

                End If
                
404             .Columns(0).Visible = False

                .Columns(.Columns.Count - 2).Visible = mnuShowCentroid.Checked
                .Columns(.Columns.Count - 1).Visible = mnuShowCentroid.Checked

                        
406             If .Columns.Count > 1 Then
408                 If GISGridRS.Fields(1).Name = "GIS_UID" Or GISGridRS.Fields(1).Name = "UID" Or GISGridRS.Fields(1).Name = "ID" Then
410                     .Columns(1).Visible = False
                    End If
                End If
                
412             .KeyField = GISGridRS.Fields(0).Name
            End If
        
        End With
        
414     If j > 0 Then mnuShowRecordCount_Click
416     dxGISDataGrid.Visible = True
418     dxProgressBar1.Visible = False

        '<EhFooter>
        Exit Sub

LoadLayerAttrDataToGrid_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.LoadLayerAttrDataToGrid " & "at line " & Erl
        'Resume Next
        '</EhFooter>
End Sub

Private Sub LoadLimitedAttr2Grid(oLayer As TatukGIS_XDK10.XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo LoadLimitedAttr2Grid_Err
        '</EhHeader>
        Dim vcm As Variant
        Dim i As Integer
        Dim lngFldType As Long
        Dim lngDataLength As Long
        Dim flds() As String
        Dim Col As DXDBGRIDLibCtl.dxGridColumn ' Variant
        Dim k As Integer
        Dim iTot As Long
        Dim sExcludedFlds() As String
                    
100     If oLayer Is Nothing Then
102         MsgBox "The Data You are trying to browse does not anymore exist in the map." & vbCrLf & "Either it has been removed or the data is corrupt." & vbCrLf & "Contact your OASIS adminstrator if this problem remains."
            Exit Sub
        End If
    
104     With g_RSGISGridTableSettings
        
106         SafeMoveFirst g_RSGISGridTableSettings
        
108         .Find "Name = '" & oLayer.Name & "'"
        
110         iTot = oLayer.GetLastUid
        
112         If ((Not sExcludedFlds) = -1&) Then
114             If Not .EOF Then
116                 If Not .Fields.Item("excludedFlds").value = vbNull Then
118                     sExcludedFlds = Split(.Fields.Item("excludedFlds").value, ",")
                    End If
                End If
            End If
        
120         If Not .EOF Then
122             If .Fields.Item("datasetwarning").value Then
124                 If iTot > .Fields.Item("MaxRec").value Then
126                     MsgBox "Note! According to performace settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
                        Exit Sub
                    Else

128                     If iTot > .Fields.Item("warninglevel").value Then
130                         If MsgBox("The data you are about to browse contains:" & iTot & " records." & vbCrLf & "This might take very long time. Would you like to continue?", vbYesNo) = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                    
                Else

132                 If iTot > .Fields.Item("MaxRec").value Then
134                     MsgBox "Note! According to performace settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
                        Exit Sub
                    End If
                End If

            Else
                'UserDefGridRecords
                'the Layer is userdrawn, Use default values from AppSettings
136             SafeMoveFirst g_RSAppSettings
138             g_RSAppSettings.Find "SettingName = 'UserDefGridRecords'"
                
140             If g_RSAppSettings.Fields.Item("SettingValue1").value = "1" Then
142                 If iTot > CLng(g_RSAppSettings.Fields.Item("SettingValue2").value) Then
144                     MsgBox "Note! According to performace settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
                        Exit Sub
                    Else

146                     If iTot > CLng(g_RSAppSettings.Fields.Item("SettingValue3").value) Then
148                         If MsgBox("The data you are about to browse contains:" & iTot & " records." & vbCrLf & "This might take very long time. Would you like to continue?", vbYesNo) = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        
        End With
        
152     With dxGISDataGrid
154         .Dataset.Close

156         If .Dataset.FieldCount > 0 Then

158             For k = 0 To .Dataset.FieldCount - 1
160                 .Dataset.MemoryDataset.DeleteField .Dataset.FieldByNo(k)
                Next

            End If
            
162         .Dataset.Refresh
            
164         .Columns.DestroyColumns
        
            If mnuAutoSize.Checked Then
              .Options.Set (egoAutoWidth)
            Else
                .Options.Unset (egoAutoWidth)
            End If
            
            If mnuAutoHeight.Checked Then
              .Options.Set (egoRowAutoHeight)
            Else
                .Options.Unset (egoRowAutoHeight)
            End If
            
168         .Options.Set (egoShowGroupPanel)
170         .Options.Set (egoBandMoving)
172         .Options.Set (egoColumnMoving)
174         .Options.Set (egoMultiSort)
176         .Options.Set (egoShowFooter)
178         .Options.Set (egoAutoSort)
180         .Options.Set (egoShowButtons)
182         .Options.Set (egoShowRowFooter)
184         .Options.Set (egoAutoSearch)
186         .Options.Set (egoAutoExpandOnSearch)
188         .Options.Set (egoAnsiSort)
190         .Options.Set (egoLoadAllRecords)
192         .Options.Set (egoAutoSearch)
            
194         .Options.Unset (egoCanNavigation)
    
196         .DatasetType = dtMemoryDataset
198         .Dataset.MemoryDataset.ClearData
200         .Filter.FilterActive = True
202         .Filter.FilterStatus = fsAlways
    
204         ReDim flds(0)
            
206         .Dataset.MemoryDataset.AddField "GIS_UID", xftInteger, 100
208         Set Col = .Columns.Add(gedTextEdit)
210         Col.caption = "GIS_UID"
212         Col.FieldName = "GIS_UID"
214         Col.Visible = False
            
216         flds(0) = "GIS_UID"
    
            Dim bExclude As Boolean
                
218         For i = 0 To oLayer.Fields.Count - 1
220             bExclude = False
                
222             If Not ((Not sExcludedFlds) = -1&) Then
224                 If SequentialSearchStringArray(sExcludedFlds, oLayer.Fields.Item(i).Name) > -1 Then
226                     bExclude = True
                    End If
                End If
                
228             If Not bExclude Then
                
230                 ReDim Preserve flds(UBound(flds) + 1)
                    
232                 flds(UBound(flds)) = oLayer.Fields.Item(i).Name '(i + 1) to Compensate for the GIS_UID
        
234                 Select Case oLayer.Fields.Item(i).FieldType
        
                        Case Is = TatukGIS_XDK10.XgisFieldTypeString '= 0,
236                         lngFldType = xftString
238                         lngDataLength = oLayer.Fields.Item(i).Width
        
240                     Case Is = TatukGIS_XDK10.XgisFieldTypeNumber ' = 1,
242                         lngFldType = xftInteger
244                         lngDataLength = oLayer.Fields.Item(i).Width
        
246                     Case Is = TatukGIS_XDK10.XgisFieldTypeFloat '= 2,
248                         lngFldType = xftFloat
250                         lngDataLength = oLayer.Fields.Item(i).Width
        
252                     Case Is = TatukGIS_XDK10.XgisFieldTypeBoolean '= 3,
254                         lngFldType = xftBoolean
256                         lngDataLength = oLayer.Fields.Item(i).Width
        
258                     Case Is = TatukGIS_XDK10.XgisFieldTypeDate '= 4
260                         lngFldType = xftDate
262                         lngDataLength = oLayer.Fields.Item(i).Width
                    End Select
                            
264                 .Dataset.MemoryDataset.AddField oLayer.Fields.Item(i).Name, lngFldType, lngDataLength
                    
266                 Set Col = .Columns.Add(gedTextEdit)
268                 Col.caption = oLayer.Fields.Item(i).Name
270                 Col.FieldName = oLayer.Fields.Item(i).Name
                End If
                
            Next
    

274         .KeyField = "GIS_UID"
276         .Dataset.Open
    
            Dim j As Integer
            Dim strVals As Variant
            Dim arVals As Variant
            Dim arVal As Variant
            Dim varVal As Variant
            Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
            
278         arVals = Split(strVals, ",")
    
280         For Each oShp9 In oLayer.Loop(GIS10.viewer.VisibleExtent, "", Nothing, "", True)
282             strVals = ""
284             ReDim arVal(UBound(flds))
286             arVal(0) = oShp9.uID

288             j = 0

290             For i = 1 To oLayer.Fields.Count - 1

292                 bExclude = False
                
294                 If Not ((Not sExcludedFlds) = -1&) Then
296                     If SequentialSearchStringArray(sExcludedFlds, oLayer.Fields.Item(i - 1).Name) > -1 Then
298                         bExclude = True
                        End If
                    End If
    
300                 If Not bExclude Then
                    
304                     Select Case oLayer.Fields.Item(i - 1).FieldType 'i to compensate the GIS_UID on 0 position
    
                            Case Is = TatukGIS_XDK10.XgisFieldTypeString '= 0,
306                             varVal = CStr(oShp9.GetField(flds(j + 1)))
308                             arVal(j + 1) = varVal
    
310                         Case Is = TatukGIS_XDK10.XgisFieldTypeNumber ' = 1,
312                             arVal(j + 1) = CDbl(oShp9.GetField(flds(j + 1)))
    
314                         Case Is = TatukGIS_XDK10.XgisFieldTypeFloat '= 2,
316                             arVal(j + 1) = CDbl(oShp9.GetField(flds(j + 1)))
    
318                         Case Is = TatukGIS_XDK10.XgisFieldTypeBoolean '= 3,
320                             arVal(j + 1) = CBool(oShp9.GetField(flds(j + 1)))
    
322                         Case Is = TatukGIS_XDK10.XgisFieldTypeDate '= 4
324                             arVal(j + 1) = CDate(oShp9.GetField(flds(j + 1)))
                        End Select
    
326                     strVals = IIf(strVals <> "", strVals & ",", "") & oShp9.GetField(flds(j + 1))
                    
328                     j = j + 1
                    
                    End If
                
                Next
    
330             arVals = arVal
332             .Dataset.AppendRecord arVals
            Next
        
        End With
    
        Exit Sub

        '<EhFooter>
        Exit Sub

LoadLimitedAttr2Grid_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.LoadLimitedAttr2Grid " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Private Sub AddFieldsToMemoryDataset(ByVal vFld As Variant, MDS As IdxMemoryDataset)
        '<EhHeader>
        On Error GoTo AddFieldsToMemoryDataset_Err
        '</EhHeader>
    Dim i As Long
100  With MDS
102   For i = 0 To UBound(vFld)
104    .AddField vFld(i)(0), vFld(i)(1), vFld(i)(2)
      Next
     End With
        '<EhFooter>
        Exit Sub

AddFieldsToMemoryDataset_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.AddFieldsToMemoryDataset " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub abGridPop_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)
        '<EhHeader>
        On Error GoTo abGridPop_ToolClick_Err
        '</EhHeader>
100     Select Case Tool.Name
    
            Case ""
        
102         Case ""
        
104         Case ""
    
        End Select
        '<EhFooter>
        Exit Sub

abGridPop_ToolClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.abGridPop_ToolClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GISComboGridChange(toolName As String)
        '<EhHeader>
        On Error GoTo abGridTools_ComboSelChange_Err

        Dim l As Long
        Dim lMAX As Long

        '</EhHeader>
100     Select Case toolName

            Case Is = "comIncCategory"
102             CategorizeIncidents

104         Case Is = "comLyr", "btnRefresh"
                'Buggfix Nothing was getting things stuck
                
                If Not pLoadLyrAttrToGrdThread.IsThreadRunning Then
                
                    If abGridTools.Tools.Item("comLyr").Text <> "---Nothing---" Then
                        LoadLayerAttrDataToGridInit
                        CreateSummaryGroup
                    Else
                    'GISComboGridChange "btnRefresh"
                    'we should add to clen up the bugger here
    '108         dxGISDataGrid.Visible = False
    '110         dxGISDataGrid.Columns.DestroyColumns
    '112         Set dxGISDataGrid.DataSource = Nothing
    '114         dxGISDataGrid.Visible = True
    
                    
                    End If
                
                Else
                    MsgBox "Please wait until the prior data in the grid has completed loading", vbInformation
                End If
116         Case Is = "comDateFrom"
                'FORMAT "dd.mm.yyyyy" Incident Date (ddmmyyyy)
118             LoadOASISIncidents
                'abGridTools.Tools.Item("comDateFrom").Text

120         Case Is = "s"
            
        End Select
        
        dxGISDataGrid.Option = egoAutoWidth
        mnuAutoSize.Checked = True
        
        'dxGISDataGrid.Option = egoRowAutoHeight
        'mnuAutoHeight.Checked = True
        
        dxGISDataGrid.OptionEnabled = True
        
        
        l = dxGISDataGrid.Columns.Count - 1
        Do Until l < 0
            
            'If dxGISDataGrid.Columns.Item(l).FieldType = xftDateTime Then
            
            dxGISDataGrid.Columns.Item(l).Width = (dxGISDataGrid.Width / Screen.TwipsPerPixelX) / dxGISDataGrid.Columns.Count
            'End If
        
            l = l - 1
        Loop
        
        '<EhFooter>
        Exit Sub

abGridTools_ComboSelChange_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.abGridTools_ComboSelChange " & "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub abGridTools_ComboSelChange(ByVal Tool As ActiveBar3LibraryCtl.Tool)
    GISComboGridChange Tool.Name
End Sub

Private Sub CreateSummaryGroup()
    Dim sumgroup As DXDBGRIDLibCtl.dxGridSummaryGroup
    Dim sumitem As DXDBGRIDLibCtl.dxGridSummaryItem
    Dim i As Integer
    Dim Col As DXDBGRIDLibCtl.dxGridColumn

    Set sumgroup = dxGISDataGrid.SummaryGroups.Add
    sumgroup.DefaultGroup = True
    Set sumitem = sumgroup.SummaryItems.Add
    sumitem.SummaryType = cstCount
    
    dxGISDataGrid.Option = egoCStyleFormatting
    dxGISDataGrid.OptionEnabled = True
    
    sumitem.SummaryFormat = "%.0f records"
    If dxGISDataGrid.Columns.Count > 0 Then
    dxGISDataGrid.Columns(1).SummaryFooterFormat = "%.0f records"
    End If
    
    With dxGISDataGrid
    
    
    


        For i = 0 To .Columns.Count - 1
            Set Col = .Columns(i)

            If Col.Visible Then
                Col.SummaryFooterType = cstCount
                Exit For
            End If

        Next
        
        If mnuShowRecordCount.Checked Then
            .Option = egoShowFooter
            .OptionEnabled = True
        Else
            .Option = egoShowFooter
            .OptionEnabled = True
        End If
    End With

End Sub
Public Sub FilterIncidentsByDate()

End Sub

Public Sub PrepareFontLists()


End Sub

Public Sub CategorizeIncidents()
        '<EhHeader>
        On Error GoTo CategorizeIncidents_Err
        '</EhHeader>
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
        Dim SymbolList As New TatukGIS_XDK10.XGIS_SymbolList
        Dim RS As New ADODB.Recordset
        Dim sFont As String
        Dim sDBFieldName As String
        Dim sGISFieldName As String
        'GIS109 Done
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"

        DebugPrint "OASIS Incident Layer Name:" & g_RSAppSettings.Fields.Item("SettingValue1").value

104     Set lL = GIS10.get(g_RSAppSettings.Fields.Item("SettingValue1").value)

        If Not m_sIncidentIni = "" Then

            With lL
                .ConfigName = m_sIncidentIni
                .StoreParamsInProject = False
                .UseConfig = True
                .RereadConfig
                .draw
            End With

            Exit Sub
        End If
    
106     Set RS = New ADODB.Recordset
                
108     Set RS.ActiveConnection = m_Cnn
                
110     RS.CursorType = adOpenDynamic
112     RS.CursorLocation = g_sGlobalCursorLocation
114     RS.LockType = adLockReadOnly
                
116     Select Case m_frmMnuOperations.abGridTools.Tools.Item("comIncCategory").Text  'shp.GetField("Type")

            Case "--None--"
118             sDBFieldName = "Incident_Type_Name"
120             sGISFieldName = "Type"
122             RS.Open "SELECT * FROM IncTypeCategory ORDER BY [Incident_Type_Name]"
                DebugPrint " None Category"

124         Case "Target"
126             sDBFieldName = "Name"
128             sGISFieldName = "Type"
130             RS.Open "SELECT * FROM IncTarget ORDER BY [NAME]"
                DebugPrint " Incident Target Category"

132         Case "Time"
134             sDBFieldName = "Incident_Time_Name"
136             sGISFieldName = "TIME00"
138             RS.Open "SELECT * FROM IncTimeCategory ORDER BY [Incident_Time_Name]"
                DebugPrint " Time Category"

140         Case "Type"
142             sDBFieldName = "Incident_Type_Name ORDER BY [Incident_Type_Name]"
144             sGISFieldName = "Type"
146             RS.Open "SELECT * FROM IncTypeCategory ORDER BY [Incident_Type_Name]"
                DebugPrint " Type Category"
        End Select

150     'While Not shp Is Nothing
        For Each shp In lL.Loop(GIS10.viewer.Extent, "", Nothing, "", True)
            DebugPrint " INCIDENT Type:" & shp.GetField("Type")

152         SafeMoveFirst RS
154         RS.Find sDBFieldName & " = " & "'" & shp.GetField(sGISFieldName) & "'"
            DebugPrint sGISFieldName & " Value = " & shp.GetField(sGISFieldName)
            
156         If Not RS.Bof And Not RS.EOF Then
158             sFont = RS.Fields("Font_Name").value
160             sFont = sFont & ":" & RS.Fields("Ascii").value & ":NORMAL"
                DebugPrint " Font:" & sFont
            End If
                
162         With shp.Params.Marker
                
164             .color = vbWhite
166             .OutlineColor = vbBlue
168             .Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
170             .Size = 640
                
            End With
        
172         shp.draw
        
        Next
    
176     lL.Paint
            '<EhFooter>
        Exit Sub

CategorizeIncidents_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.CategorizeIncidents " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetApplicationSettings(dumVar As Variant)
        '<EhHeader>
        On Error GoTo SetApplicationSettings_Err
        '</EhHeader>
        Dim sGridOpt() As String
        

100     ClearCoordValues
102     frddmmss.Move frmdddddd.left, frmdddddd.top
104     frddmmmm.Move frmdddddd.left, frmdddddd.top
106     frMGRS.Move frmdddddd.left, frmdddddd.top
        dxGISDataGrid.MaxRowLineCount = 4
        
108     With g_RSAppSettings
        
110         .Requery
112         SafeMoveFirst g_RSAppSettings
114         .Find "SettingName = 'AdminLevel0'"

116         If Not .Fields.Item("SettingValue1").value = vbNullString Then m_sAdmVal1 = .Fields.Item("SettingValue1").value
        
118         SafeMoveFirst g_RSAppSettings
120         .Find "SettingName = 'AdminLevel1'"

122         If Not .Fields.Item("SettingValue1").value = vbNullString Then m_sAdmVal2 = .Fields.Item("SettingValue1").value
        
124         SafeMoveFirst g_RSAppSettings
126         .Find "SettingName = 'AdminLocation'"

128         If Not .Fields.Item("SettingValue1").value = vbNullString Then m_sAdmLoc = .Fields.Item("SettingValue1").value
        
            Dim sModMenus() As String
            Dim i As Integer
    
130         SafeMoveFirst g_RSAppSettings
132         .Find "SettingName = 'VisibleMainModuleMenus'"
    
134         sModMenus = Split(.Fields.Item("SettingValue1").value, ",")
        
136         For i = 0 To AB.Bands("bNavPane").ChildBands.Count - 1
138             AB.Bands("bNavPane").ChildBands(i).Visible = False
            Next
        
140         For i = 0 To UBound(sModMenus)
150             AB.Bands("bNavPane").ChildBands(sModMenus(i)).Visible = True
152         Next i
        
154         CheckAvailableTools
        
156         SafeMoveFirst g_RSAppSettings
158         .Find "SettingName = 'UserDefGridRecords'"
        
160         If Not IsNull(.Fields.Item("SettingValue4").value) Then
                
162             If Len(.Fields.Item("SettingValue4").value) > 2 Then
164                 sGridOpt = Split(.Fields.Item("SettingValue4").value, ",")
166                 chkFilterIn.Visible = IIf(sGridOpt(0) = "1", True, False)

                    If sGridOpt(0) <> "1" Then chkSelectIn.left = chkFilterIn.left
168                 chkSelectIn.Visible = IIf(sGridOpt(1) = "1", True, False)
                End If
        
            End If
        
170         SafeMoveFirst g_RSAppSettings
172         .Find "SettingName = 'CurrentActiveMeny'"
        
174         AB.Bands("bNavPane").ChildBands.CurrentChildBand = AB.Bands("bNavPane").ChildBands(.Fields.Item("SettingValue1").value)

            If AB.Bands("bNavPane").ChildBands.CurrentChildBand.Name = "cbOperations" Then
                m_bMapInitialized = False
                ActivateOperations
                '(keith) Added this to fix non rendering of map when proile was set as default view
            ElseIf AB.Bands("bNavPane").ChildBands.Item("cbProfile").Visible Then
                InitMapFromMapLibraryDefault
                ' InitMap
            End If

176         AB.RecalcLayout

            SafeMoveFirst g_RSAppSettings
            g_RSAppSettings.Find "SettingName = 'ShowSelector'"
            C1TTab1Tab2.TabHeight = 1

            If Not g_RSAppSettings.EOF Then
                C1TTab1Tab2.TabHeight = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", 0, 1)
            End If
            
            SafeMoveFirst g_RSAppSettings
            g_RSAppSettings.Find "SettingName = 'ShowScalebar'"
            Scale1.Visible = False

            If Not g_RSAppSettings.EOF Then
                Scale1.Visible = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", True, False)
            End If
            
            SafeMoveFirst g_RSAppSettings
            g_RSAppSettings.Find "SettingName = 'ShowZoomBar'"
            ctlZoomSlider1.Visible = False

            If Not g_RSAppSettings.EOF Then
                ctlZoomSlider1.Visible = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = "1", True, False)

                If g_RSAppSettings.Fields.Item("SettingValue1").value = "1" Then
                    m_lZoomSelectColour = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue2").value) > 0, (g_RSAppSettings.Fields.Item("SettingValue2").value), vbWhite)
                    m_lZoomBackColour = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue3").value) > 0, (g_RSAppSettings.Fields.Item("SettingValue3").value), &HC0&)
                    m_lZoomCrosshairColour = IIf(Len(g_RSAppSettings.Fields.Item("SettingValue4").value) > 0, (g_RSAppSettings.Fields.Item("SettingValue4").value), vbWhite)
                    
            
                    'm_eZoomberMaxExtent
                    With g_RSAppSettings.Fields
            
                        If Not IsNull(.Item("SettingValue5")) And Not IsNull(.Item("SettingValue6")) And Not IsNull(.Item("SettingValue7")) And Not IsNull(.Item("SettingValue8")) Then
            
                            Set m_eZoomberMaxExtent = New XGIS_Extent
                            m_eZoomberMaxExtent.Prepare .Item("SettingValue5"), .Item("SettingValue7"), .Item("SettingValue6"), .Item("SettingValue8")
            ctlZoomSlider1.Init m_eZoomberMaxExtent, GIS10.VisibleExtent, vbWhite, &HC0&, vbWhite
                        Else
                            Set m_eZoomberMaxExtent = Nothing
                            '  End If
                        End If

                    End With
                    
                    ctlZoomSlider1.SetColours m_lZoomSelectColour, m_lZoomBackColour, m_lZoomCrosshairColour

                End If
            End If

180         SafeMoveFirst g_RSAppSettings
182         .Find "SettingName = 'ShowExpandedToolBoxInGISWin'"
        
184         If CInt(.Fields("SettingValue1").value) = 1 Then
186             C1TabFastFunction.Visible = True
188             C1TabFastFunction.TabsPerPage = CInt(.Fields("SettingValue2").value)

                'GeoMarks
                If CInt(.Fields("SettingValue3").value) = 1 Then
190                 C1TabFastFunction.TabVisible(1) = True
                    LoadGeoBookMarks
                Else
                    Themes.Visible = False
                    C1TabFastFunction.TabVisible(1) = False
                End If
                
                'Geo Convert
                If CInt(.Fields("SettingValue4").value) = 1 Then
194                 C1TabFastFunction.TabVisible(2) = True
                Else
                    elGeoCalc.Visible = False
                    C1TabFastFunction.TabVisible(2) = False
                End If
                
                'Goto
                If CInt(.Fields("SettingValue5").value) = 1 Then
196                 C1TabFastFunction.TabVisible(3) = True
                Else
                    C1TabFastFunction.TabVisible(3) = False
                End If

                'Settings
                If CInt(.Fields("SettingValue6").value) = 1 Then
198                 C1TabFastFunction.TabVisible(4) = True
                Else
                    C1Elastic1.Visible = False
                    C1TabFastFunction.TabVisible(4) = False
                End If
                
                'Magnifier
                If CInt(.Fields("SettingValue7").value) = 1 Then
200                 C1TabFastFunction.TabVisible(5) = True
                Else
                    elMagnifier.Visible = False
                    C1TabFastFunction.TabVisible(5) = False
                End If

            Else
202             C1TabFastFunction.Visible = False
            End If
        
204         SafeMoveFirst g_RSAppSettings
206         .Find "SettingName = 'SecurityTools'"
208         m_frmMnuOperations.cmdSecurityAnalysis.Visible = IIf(.Fields.Item("SettingValue1").value = "1", True, False)
210         m_frmMnuOperations.FraAnalysisLevel.Visible = IIf(.Fields.Item("SettingValue1").value = "1", True, False)
212         m_frmMnuOperations.cmdCreateCharts.Visible = IIf(.Fields.Item("SettingValue2").value = "1", True, False)
214         m_frmMnuOperations.cmdInsertIncident.Visible = IIf(.Fields.Item("SettingValue3").value = "1", True, False)
216         m_frmMnuOperations.cmdSetScope.Visible = IIf(.Fields.Item("SettingValue4").value = "1", True, False)
        
218         SafeMoveFirst g_RSAppSettings
220         .Find "SettingName = 'MineActionTools'"
222         m_frmMnuOperations.DxHUpdateData.Visible = IIf(.Fields.Item("SettingValue1").value = "1", True, False)
        
224         SafeMoveFirst g_RSAppSettings
226         .Find "SettingName = 'InetConnectionSettings'"

228         tmrInternetCheck.Interval = IIf(1000 * CLng(.Fields.Item("SettingValue2").value) > 65000, 65000, 1000 * CLng(g_RSAppSettings.Fields.Item("SettingValue2").value))
230         tmrInternetCheck.Enabled = IIf(.Fields.Item("SettingValue1").value = "1", True, False)
            
            If Not IsNull(.Fields.Item("SettingValue3").value) Then
                g_bRefreshMapAfterSynch = IIf(CStr(Trim$(.Fields.Item("SettingValue3").value)) = "1", True, False)
            End If
            
232         SafeMoveFirst g_RSAppSettings
234         .Find "SettingName = 'OpsTools'"
        
236         If Not .EOF Then
                If Not IsNull(.Fields.Item("SettingValue2").value) Then
                    m_frmMnuOperations.Legend1.Mode = IIf(.Fields.Item("SettingValue2").value = "1", XgisControlLegendModeGroups, XgisControlLegendModeLayers)
                End If

238             cmdCommand2.Visible = IIf(.Fields.Item("SettingValue1").value = "1", True, False)
            Else
240             cmdCommand2.Visible = False
            End If

            SafeMoveFirst g_RSAppSettings
            .Find "SettingName = 'RangeSettingsColours'"
                        
            If Not .EOF Then
                                
                For i = 1 To 10
                    
                    If Not IsNull(.Fields.Item("SettingValue" & i).value) Then
                        If IsNumeric(.Fields.Item("SettingValue" & i).value) Then
                            ReDim Preserve arRangeColors(i - 1)
                            arRangeColors(i - 1) = CLng(.Fields.Item("SettingValue" & i).value)
                        Else
                            Exit For
                        End If

                    Else
                        Exit For
                    End If

                Next
                
                SafeMoveFirst g_RSAppSettings
                .Find "SettingName = 'RangeSettingsMaxVal'"
                
                If Not .EOF Then

                    For i = 1 To UBound(arRangeColors) + 1
                        ReDim Preserve arRangeVals(i - 1)
                        arRangeVals(i - 1) = CLng(.Fields.Item("SettingValue" & i).value)
                    Next

                End If
                
            Else
                ReDim arRangeColors(6)
                ReDim arRangeVals(6)
                
                arRangeColors(0) = 12
                arRangeVals(0) = 101
                arRangeColors(1) = 124124
                arRangeVals(1) = 105
                arRangeColors(2) = 12412
                arRangeVals(2) = 110
                arRangeColors(3) = 255
                arRangeVals(3) = 115
                arRangeColors(4) = 4124
                arRangeVals(4) = 130
                arRangeColors(5) = 1244
                arRangeVals(5) = 150
            
            End If

242         SafeMoveFirst g_RSAppSettings
244         .Find "SettingName = 'Scripts'"

246         If Not .EOF Then
248             If Not IsNull(.Fields.Item("SettingValue2").value) Then
250                 If Len(.Fields.Item("SettingValue2").value) > 4 Then
                
                        Dim sScripts() As String
                        Dim k As Integer
                    
252                     sScripts = Split(.Fields.Item("SettingValue2").value, ",")
                    
                        'TODO Enable to add to different toolbars, tooltips as well as icons and different styles of buttons
                    
254                     For k = LBound(sScripts) To UBound(sScripts)
256                         AddBtn k, ddSIconText, "Script" & k, "btnScript" & k, CLng(Rnd(11000) & k), "tbUtils"
                            'AddBtn 1, ddSIconText, "Script1", "btnScript1", 11111, "tbUtils"
                            'AddBtn 2, ddSIconText, "Script2", "btnScript2", 11112, "tbUtils"
                            'AddBtn 3, ddSIconText, "Script3", "btnScript3", 11113, "tbUtils"
                            'AddBtn 4, ddSIconText, "Script4", "btnScript4", 11114, "tbUtils"
                        Next

                    End If
                End If
            End If

            '  End If

        End With

        '
        'sdToolBar,tbExtents,tbLayer,bToolbarStyle
        
        '<EhFooter>
        Exit Sub

SetApplicationSettings_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.SetApplicationSettings " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadGeoBookMarks()
        '<EhHeader>
        On Error GoTo LoadGeoBookMarks_Err
        '</EhHeader>
        Dim CatRS As New ADODB.Recordset
        Dim bkrRS As New ADODB.Recordset

100     With ctTreeBookmrks
    
102         .ClearNodes 'Clear the ctTree control of any existing nodes
    
104         .MultiSelect = False 'Do not allow for multi select
106         .Multiline = False 'Do not allow for multi line nodes
108         .HorzScroll = True 'Do not display a horizontal scroll bar
110         .TipsOnScroll = True 'Display tips on vertical scroll
112         .TipsType = TipsBoth 'Display tips when the mouse moves over the node and on the vertical thumb
114         .TipsDelay = 1 'Delay the tips display for 5 seconds
116         .TreeLines = True 'Display the tree lines
118         .PictureType = 3 'Display a picture, text, and plus/minus
120         .ClearColumns 'Clear any existing columns from the control
122         .ScrollOnVThumb = False 'Do not allow the nodes to scroll when the vertical thumb is being moved
124         .HorzAutoSize = False 'Do not automatically adjust the width to fit the node
126         .SelectedStyle = 2 'Highlight only the first column of the selected node
    
128      CatRS.Open "SELECT * FROM GeoBookMarksCategories ORDER BY Name", m_Cnn, adOpenStatic, adLockReadOnly
    
130     If Not CatRS.Bof Then SafeMoveFirst CatRS
    
    
'132     .AddNode (m_DefaultViewName), 0, 1
        .AddNode ("Start View"), 0, 1
        
134     Do While Not CatRS.EOF
136         bkrRS.Open "SELECT * FROM GeoBookMarks WHERE BmkrID = " & CatRS.Fields.Item("ID").value & " ORDER BY Name", m_Cnn
        
        
        
138         If Not bkrRS.Bof Then
140             SafeMoveFirst bkrRS
142             .AddNode (CatRS.Fields.Item("Name").value), 0, 1
            End If
        
144         Do While Not bkrRS.EOF
                '.AddPictureNode bkrRS.Fields.item("Name").Value, 0, 2, 3, 0, 12
146             .AddNode (bkrRS.Fields.Item("Name").value), 0, 2
                '.AddNode ("X:" & bkrRS.Fields.item("X").Value), 0, 3
                '.AddNode ("Y:" & bkrRS.Fields.item("Y").Value), 0, 3
                '.AddNode (IIf(IsNull(bkrRS.Fields.item("Description").Value), "", bkrRS.Fields.item("Description").Value)), 0, 3
148             bkrRS.MoveNext
            Loop
150         bkrRS.Close
152         CatRS.MoveNext
        Loop
    
    '        'The following nodes will illustrate how the tips are displayed.  The tips will
    '        'appear for those nodes that are too long to be fully displayed
    '        .AddNode ("This is a very long message"), 0, 1
    '        .AddNode ("This is a child node of the really long message"), 0, 2
    '        .AddNode ("This is another child node of the really long message"), 0, 2
    '
    '        .AddNode ("This is a another very long message"), 0, 1
    '        .AddNode ("This is a child node of the really long message"), 0, 2
    '        .AddNode ("This is another child node of the really long message"), 0, 2
    '
    '        .AddNode ("Short Message"), 0, 1
    '        .AddNode ("This is short"), 0, 2
    '        .AddNode ("This is also short"), 0, 2
    
        End With


    
        '<EhFooter>
        Exit Sub

LoadGeoBookMarks_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.LoadGeoBookMarks " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub EndEdit(Optional oTool As OASIS_TOOLS = -1)
        '<EhHeader>
        On Error GoTo EndEdit_Err
        '</EhHeader>

If oTool = -1 Then
   g_CurrentTool = m_prevTool
   SetOASISTool
Else
   g_CurrentTool = oTool
   GIS10.Mode = TatukGIS_XDK10.XgisEdit
End If



    '  If g_oEditLayer Is Nothing Then Exit Sub
    'On Error Resume Next
104     If Not m_oSQLOpsLyr Is Nothing Then
106            m_oSQLOpsLyr.SaveAll
        End If
      
      On Error Resume Next
      
108   GIS10.Editor.EndEdit
110   Set g_oEditLayer = Nothing
  
  
        '<EhFooter>
        Exit Sub

EndEdit_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.endEdit " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCommand2_Click()
        '<EhHeader>
        On Error GoTo cmdCommand2_Click_Err
        '</EhHeader>
100      m_oSQLOpsLyr.caption = "MY OPERATIONS"
    
102      m_oSQLOpsLyr.HideFromLegend = False
      
104       Set g_oEditLayer = m_oSQLOpsLyr
106       GIS10.Mode = TatukGIS_XDK10.XgisEdit
108       g_CurrentFeatureType = Location
110       m_prevTool = g_CurrentTool

        '<EhFooter>
        Exit Sub

cmdCommand2_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdCommand2_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>
    Dim sPath As String
    Dim sName As String
    Dim lL As New TatukGIS_XDK10.XGIS_LayerSHP
    Dim OASISLyr As String
    Dim AnalysisLyr As String
    Dim q As Integer

    'Dim sCriteria As String
    'Dim oRS As New ADODB.Recordset
    '
    'sCriteria = InputBox("Enter The Name of the Area you would like to filter on", Default:="Colombia City")
    '
    'oRS.Open "SELECT & FROM oIncidents WHERE Admin3 = '" & sCriteria & "'", cnn
    '
    'Set dxGISDataGrid.DataSource = oRS

    'dxGISDataGrid.Dataset.ADODataset

    


100     GIS10.Lock

        If Not EventLayer Is Nothing Then
102         sPath = EventLayer.Path
104         sName = EventLayer.Name
        End If
        
106     SafeMoveFirst g_RSAppSettings
108     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"
110     OASISLyr = g_RSAppSettings.Fields.Item("SettingValue1").value
    
    Set EventLayer = Nothing

'     GIS10.get(m_sSecAnalysisLyrName).draw

        If GIS10.get(m_sSecAnalysisLyrName) Is Nothing Then
            MsgBox "Layer " & m_sSecAnalysisLyrName & " does not exist. OASIS Cannot do security analysis", vbInformation, "Confuguration failure"
            GIS10.Unlock
            Exit Sub
        End If

        If GIS10.get(OASISLyr) Is Nothing Then
            MsgBox "Layer " & OASISLyr & " does not exist. OASIS Cannot do security analysis", vbInformation, "Confuguration failure"
            GIS10.Unlock
            Exit Sub
        End If
        
        For q = 0 To m_frmMnuOperations.OptLevel.UBound
            m_frmMnuOperations.OptLevel(q).Enabled = False
        Next
112     CheckIncidentFrequency GIS10.get(OASISLyr), GIS10.get(m_sSecAnalysisLyrName)

    
114
                
        If Not GIS10.get(m_sSecAnalysisLyrName).FileInfo = "TatukGIS SQL Vector Coverage (TTKLS)" Then
116         Set EventLayer = GIS10.get(m_sSecAnalysisLyrName)
118         EventLayer.Params.Visible = True
        End If
        GoTo Hell
120     If Not sName = "" Then
122         If m_bThematicsDone Then
124             GIS10.get(sName).RereadConfig
            Else
126             m_bThematicsDone = True
            End If
        End If
Hell:
128     GIS10.Unlock
130     GIS10.UpDate
        DoEvents
        Exit Sub


132     GIS10.Delete sName
    
    
134     Set EventLayer = Nothing
    
        'GIS10.Update
    
136     lL.Path = sPath
138     lL.Name = sName
140     lL.Open
    
142     GIS10.Add lL
    
144     Set EventLayer = GIS10.get(sName)
        
        
        SafeMoveFirst g_RSAppSettings
        g_RSAppSettings.Find "SettingName = 'SecurityTools'"
     
        If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue10").value) Then
            GIS10.get(m_sSecAnalysisLyrName).HideFromLegend = CBool(g_RSAppSettings.Fields.Item("SettingValue1").value)
        Else
            GIS10.get(m_sSecAnalysisLyrName).HideFromLegend = True
        End If
        m_frmMnuOperations.Legend1.UpDate
        
        For q = 0 To m_frmMnuOperations.OptLevel.UBound
            m_frmMnuOperations.OptLevel(q).Enabled = True
        Next
146     GIS10.UpDate
        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.Command1_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>

End Sub

Private Sub DoFrequencyAnalysis(sFreqOverlayLyr As String, sAnalysisLyr As String)
        '<EhHeader>
        On Error GoTo DoFrequencyAnalysis_Err
        '</EhHeader>
    Dim lL As New TatukGIS_XDK10.XGIS_LayerSHP
    Dim OASISLyr As String
    Dim AnalysisLyr As String
    
    

100     GIS10.Lock
    
102     m_sSecAnalysisFieldName = "ScoringFld"
    
104     If sFreqOverlayLyr = sAnalysisLyr Then Exit Sub
    
106     CheckFrequency GIS10.get(sFreqOverlayLyr), GIS10.get(sAnalysisLyr), "ScoringFld"

108     Set EventLayer = Nothing
110     m_sSecAnalysisFieldName = "ScoringFld"
112     Set EventLayer = GIS10.get(sAnalysisLyr)
    
114     GIS10.Unlock
116     GIS10.UpDate

        '<EhFooter>
        Exit Sub

DoFrequencyAnalysis_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.DoFrequencyAnalysis " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateScoringField(oTargetLayer As TatukGIS_XDK10.XGIS_LayerVector, sFieldAnalysisName As String)
        '<EhHeader>
        On Error GoTo CreateScoringField_Err
        '</EhHeader>
        On Error GoTo ErrCode

        On Error Resume Next
100      oTargetLayer.AddField sFieldAnalysisName, TatukGIS_XDK10.XgisFieldTypeNumber, 16, 0
        'End If
        Exit Sub
ErrCode:
        'oTargetLayer.AddField sFieldAnalysisName, TatukGIS_XDK10.XGISFieldTypeNumber, 16, 0
        '<EhFooter>
        Exit Sub

CreateScoringField_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CreateScoringField " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckFrequency(oOverLayLayer As TatukGIS_XDK10.XGIS_LayerVector, _
                                  oTargetLayer As TatukGIS_XDK10.XGIS_LayerVector, sFieldAnalysisName As String, _
                                  Optional bResetPreviousScoring As Boolean = True)
        '<EhHeader>
        On Error GoTo CheckFrequency_Err
        '</EhHeader>
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim aShape As Object
        Dim oPolyShape As TatukGIS_XDK10.XGIS_ShapePolygon
        Dim iVal As Integer

100     CreateScoringField oTargetLayer, sFieldAnalysisName

102     If bResetPreviousScoring Then
104         ResetScoring oTargetLayer
        End If
    
        For Each shp In oOverLayLayer.Loop(GIS10.viewer.VisibleExtent, "", Nothing, "", True)

112         For Each aShape In oTargetLayer.Loop(shp.Extent, "", shp, GisUtils.GIS_RELATE_WITHIN, True)
        
116             If Not aShape.GetField(sFieldAnalysisName) = vbNull Then
118                 Set oPolyShape = aShape.MakeEditable
120                 iVal = aShape.GetField(sFieldAnalysisName) + 1
122                 oPolyShape.SetField sFieldAnalysisName, CLng(iVal)
                End If
            
            Next
        Next
    
        '<EhFooter>
        Exit Sub

CheckFrequency_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CheckFrequency " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxGISDataGrid_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, _
                                       ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
    
    If mnuAutoZoom.Checked Then
        'AutoZoom Node
        FindFeatureFromGrid False, False, True
    End If

End Sub

Private Sub dxGISDataGrid_OnFilterRecord(accept As Boolean)
        '<EhHeader>
        On Error GoTo dxGISDataGrid_OnFilterRecord_Err
        '</EhHeader>
100     accept = True
102     'DebugPrint "OnFilterRecord"
        '<EhFooter>
        Exit Sub

dxGISDataGrid_OnFilterRecord_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.dxGISDataGrid_OnFilterRecord " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxGISDataGrid_OnMouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal x As Single, ByVal y As Single)
        '<EhHeader>
        On Error GoTo dxGISDataGrid_OnMouseDown_Err
        '</EhHeader>


100 If Button = vbRightButton Then
    
            ' Get the position of the cursor
102     GetCursorPos ptPopUpPos
        'abGridPop.Bands("popGrid").PopupMenu , ScaleX(Point.X, vbPixels, vbTwips) - Me.left, ScaleY(Point.Y, vbPixels, vbTwips) - Me.top
104     If dxGISDataGrid.ex.SelectedCount > 0 Then
106         PopupMenu mnuGridAction, x:=ScaleX(ptPopUpPos.x, vbPixels, vbTwips) - Me.left, y:=ScaleY(ptPopUpPos.y, vbPixels, vbTwips) - (Me.top + 250)
        End If
    End If

        '<EhFooter>
        Exit Sub

dxGISDataGrid_OnMouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.dxGISDataGrid_OnMouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxGISDataGrid_OnMouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal x As Single, ByVal y As Single)
'    DebugPrint  "MOUSEMOVE: " & X
End Sub

Private Sub dxGISDataGrid_OnSelectedCountChange()
        '<EhHeader>
        On Error GoTo dxGISDataGrid_OnSelectedCountChange_Err
        '</EhHeader>
        Dim i As Integer
        Dim Node As dxGridNode
        Dim bookmark As Variant

        'DebugPrint  "SEL COUNT CHANGE " & Now & dxGISDataGrid.ex.SelectedCount
    
100     If Not dxGISDataGrid.M.IsGridMode Then

102         For i = 0 To dxGISDataGrid.ex.SelectedCount - 1
104             Set Node = dxGISDataGrid.ex.SelectedNodes(i)
                'DebugPrint  node.Values(0) ' SelectedNodes(i)
            Next

        Else

106         For i = 0 To dxGISDataGrid.ex.SelectedCount - 1
108             bookmark = dxGISDataGrid.ex.SelectedRows(i)

110             If dxGISDataGrid.Dataset.BookmarkValid(bookmark) Then
                    'DebugPrint  dxGISDataGrid.Dataset.FieldValues(dxGISDataGrid.Dataset.FieldNameByNo(0))
112                 dxGISDataGrid.Dataset.GotoBookmark bookmark
                    'DebugPrint  dxGISDataGrid.Dataset.FieldValues(dxGISDataGrid.Dataset.FieldNameByNo(0))
                
                    '...
                End If

                'DebugPrint  node.Values(0) ' SelectedNodes(i)
            Next

        End If

        'dxGISDataGrid.m.CopySelectedToClipboard
    
114     'If mnuAutoZoom.Checked Then
        '    AutoZoom
        'End If

        '<EhFooter>
        Exit Sub

dxGISDataGrid_OnSelectedCountChange_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.dxGISDataGrid_OnSelectedCountChange " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub AutoZoom(Node As DXDBGRIDLibCtl.IdxGridNode)
        '<EhHeader>
        On Error GoTo AutoZoom_Err
        '</EhHeader>

        

        Dim lShape As TatukGIS_XDK10.XGIS_Shape
     
100     With dxGISDataGrid.ex
        
102         Set lShape = GlobalGISGridLayer.GetShape(Node.values(0))

104         If Not lShape Is Nothing And Not IsNull(lShape) Then
106             'GIS10.Lock
108             GIS10.VisibleExtent = lShape.Extent
110             GIS10.Unlock

            End If
    
        End With
        
112     Set lShape = Nothing

        '<EhFooter>
        Exit Sub

AutoZoom_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.AutoZoom " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub HideShowSelected(oLyr As TatukGIS_XDK10.XGIS_LayerVector, _
                             bShow As Boolean)
        '<EhHeader>
        On Error GoTo HideShowSelected_Err
        '</EhHeader>
        'GIS109 Done
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim oshp As TatukGIS_XDK10.XGIS_Shape
        Dim xt As TatukGIS_XDK10.XGIS_Extent
            
102     GIS10.Lock

        For Each shp In oLyr.Loop(oLyr.Extent, "GIS_Selected=True", Nothing, "", True)

104         If Not shp Is Nothing Then
106             Set oshp = shp.MakeEditable
108             oshp.IsHidden = Not bShow
                oshp.Invalidate True
                Set oshp = Nothing
            End If

        Next
            
112     GIS10.Unlock
        
        '<EhFooter>
        Exit Sub

HideShowSelected_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.HideShowSelected " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub elTatukGIS_RealignFrame()
        'GIS10.Width = elTatukGIS.Width '+ 1000
        '<EhHeader>
        On Error GoTo elTatukGIS_RealignFrame_Err
        '</EhHeader>
    
100     GIS10.Move 0, 0, elTatukGIS.Width, elTatukGIS.Height
        
102     If MsgScroll.ListCount > 0 Then
            'GIS10.Height = elTatukGIS.Height + 150
104         elScroller.Visible = True
106         elScroller.Move 5, elTatukGIS.Height - MsgScroll.Height, elTatukGIS.Width
108         elScroller.Move 5, elTatukGIS.Height - MsgScroll.Height, elTatukGIS.Width, 255
110         elScroller.ZOrder
112         cmdMSGScroller.ZOrder
        Else
114         elScroller.Visible = False
            'GIS10.Height = elTatukGIS.Height + 400
        End If

116     C1TabFastFunction.ZOrder 0
118     NArrow.left = elTatukGIS.Width - 1330
120     NArrow.ZOrder 0

122     With ctlZoomSlider1
            NArrow.top = 0
            NArrow.Height = 1095
            NArrow.Width = 1575
            .Width = 250
124         .left = (elTatukGIS.Width + (NArrow.Width / 2) - (.Width / 2)) - 1330
126         .top = NArrow.top + NArrow.Height + 20
128         .Height = elTatukGIS.Height - (NArrow.Height + NArrow.top + 150 + IIf(watermark.Path <> "", watermark.Height + 150, 0))
130         .ZOrder 0
            
        End With

132     With watermark
134         .Move elTatukGIS.Width - 1330, elTatukGIS.Height - (MsgScroll.Height + .Height + 20), .Width, .Height
136         .ZOrder 0
        End With

        '<EhFooter>
        Exit Sub

elTatukGIS_RealignFrame_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.elTatukGIS_RealignFrame " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateAddedScribbles()
        '<EhHeader>
        On Error GoTo CreateAddedScribbles_Err
        '</EhHeader>
        Dim xportLyr As TatukGIS_XDK10.XGIS_LayerSqlAdo
        Dim sAdoString As String
        Dim sLayerName As String
        Dim RS As ADODB.Recordset
        Dim Col As ADOX.Column

100     Set m_oDrawLyr = New TatukGIS_XDK10.XGIS_LayerSqlAdo
102     m_oDrawLyr.Params.area.color = RGB(0, 0, 255)
104     m_oDrawLyr.Transparency = 70
106     m_oDrawLyr.Name = "Draw_Layer"
108     m_oDrawLyr.HideFromLegend = True

110     sAdoString = GetConnectionString(g_sAppPath & "\Data\Db\OASISClient.mdb")

112     sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = Native" & vbCrLf & "layer=" & "Draw_Layer" & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString

114     If Not DoTableExists("Draw_Layer_GEO", m_Cnn) Then
116         Set RS = New ADODB.Recordset

118         RS.Open "SELECT * FROM oincidents_GEO WHERE UID = 1000000", m_Cnn

120         CreateTable "Draw_Layer_GEO", RS, m_Cnn

124         Set RS = New ADODB.Recordset
126         RS.Open "SELECT UID, NAME FROM oincidents_FEA WHERE UID = 1000000", m_Cnn
            '128         RS.Fields.Append "StringStyle", adVarChar, 255
128         CreateTable "Draw_Layer_FEA", RS, m_Cnn
               
            Set Col = New ADOX.Column

130         With Col
132             .Name = "StringStyle"
134             .Type = adVarChar
136             .DefinedSize = 255
            End With
                
            AddFieldToTable m_Cnn, "Draw_Layer_FEA", Col
                
            Set Col = Nothing
        End If

138     m_oDrawLyr.Path = sAdoString
        '116 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, TatukGIS_XDK10.XGISShapeTypeUnknown, "", True
140     m_oDrawLyr.SaveAll
142     m_oDrawLyr.Open

        '<EhFooter>
        Exit Sub

CreateAddedScribbles_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.CreateAddedScribbles " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
        Dim bPrompt As Boolean
        Dim OASISSynchFolderImporter As New clSynchFolderImporter
        Dim RSUpdater As ADODB.Recordset
        
        SavePrivateAppSettings
        Me.WindowState = vbMinimized
        
        If pLoadLyrAttrToGrdThread.IsThreadRunning Then
            pLoadLyrAttrToGrdThread.TerminateWin32Thread
        End If
        
        If Not m_oDrawLyr Is Nothing Then
        
            On Error Resume Next

            If g_ZoomToSettings.SaveOnExit Then m_oDrawLyr.SaveAll
        
        End If
        
100     OASISSynchFolderImporter.ScanAndProcessSynchFolder g_sAppPath & "\Data\Sync\import", g_sAppPath & "\Data\Db\OASISClient.mdb"
102     Set OASISSynchFolderImporter = Nothing

    '    If Not SilentHttpComms Is Nothing Then Set SilentHttpComms = Nothing
        If Not GISGridRS Is Nothing Then Set GISGridRS = Nothing
104     EndEdit
    
106     If Not m_Cnn Is Nothing Then
108         If Not m_Cnn.State = adStateClosed Then
                      
                Set RSUpdater = New ADODB.Recordset

                With RSUpdater

                    .Open "SELECT * FROM Personnell", m_Cnn, adOpenDynamic, adLockBatchOptimistic
                    .Find "Personnell_ID = " & g_CurrentUserID

                    If Not .EOF Then
                        .Fields("LatestViewX").value = Replace(GIS10.viewer.CenterPtg.x, ",", ".")
                        .Fields("LatestViewY").value = Replace(GIS10.viewer.CenterPtg.y, ",", ".")
                        .Fields("LatestViewZ").value = Replace(GIS10.viewer.zoom, ",", ".")
                        .Fields("LatestMapName").value = GIS10.Name
                        .UpdateBatch adAffectCurrent
                        .Close
                    End If

                End With

                Set RSUpdater = Nothing
            
            End If
        End If
        
        If g_bDefaultMapChanged Then
            If MsgBox("Do you want the current map to load as the default map?", vbYesNo, "OASIS Client") = vbYes Then
                m_Cnn.Execute "UPDATE [ttkGISProjectDef] SET [InUse]=false"
                m_Cnn.Execute "UPDATE [ttkGISProjectDef] SET [InUse]=true WHERE [sGUID] = '" & g_bDefaultMapChangedGUID & "'"
            End If
        End If

'112     SafeMoveFirst g_RSAppSettings
'114     g_RSAppSettings.Find "SettingName = 'InitialOperationsTabNumber'"
'
'116     If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue2").Value) Then
'118         If g_RSAppSettings.Fields.Item("SettingValue2").Value = "1" Then bPrompt = True
'        End If
'
'120     If bPrompt And g_bMapLoadedCorrectly Then
'122         If MsgBox("Do you want to save changes to the map?", vbYesNo, "OASIS Client") = vbYes Then
'                'm_oSQLIncLyr.SaveAll
'                On Error Resume Next
'
'124             'Without these lines of code the incidents layer will not appear next load
'                GIS10.Delete m_oSQLIncLyr.Name
'                GIS10.Delete "Draw_Layer"
'                GIS10.SaveAll
'                UpdateProjectFileToDB
'            End If
'        End If
    
        On Error Resume Next
        Shell "taskkill /F /IM OASIS_SynchNG.exe"
        'OASISFolderMonitorImporter.StopMonitoring
        
        'Set OASISFolderMonitorImporter = Nothing
        
126     SaveStartUpParams

        g_clsHotKey.UnregisterKey "Language"
        g_clsHotKey.UnregisterKey "OPSV"
        'g_clsHotKey.UnregisterKey "doPrint"
        g_clsHotKey.UnregisterKey "Admin"
        g_clsHotKey.UnregisterKey "Dude"
        g_clsHotKey.UnregisterKey "Script1"
        g_clsHotKey.UnregisterKey "Script2"
        g_clsHotKey.UnregisterKey "Script3"
        g_clsHotKey.UnregisterKey "Script4"
        g_clsHotKey.UnregisterKey "AppState"
        g_clsHotKey.UnregisterKey "Folders"
        g_clsHotKey.UnregisterKey "GPS"
        g_clsHotKey.UnregisterKey "OVB"
        g_clsHotKey.UnregisterKey "CompInfo"
        g_clsHotKey.UnregisterKey "Ticker"
        g_clsHotKey.UnregisterKey "ForceIni"
        g_clsHotKey.UnregisterKey "Berserk"
        g_clsHotKey.UnregisterKey "Tracker"
        g_clsHotKey.UnregisterKey "LyrSearch"
        g_clsHotKey.UnregisterKey "WMS"
        g_clsHotKey.UnregisterKey "Debug"
        
128     Set g_clsHotKey = Nothing
130     Set oVB = Nothing
132     Set m_oAES = Nothing
    
134     SetErrorMode SEM_NOGPFAULTERRORBOX
    
        '    Dim i As Integer
        '
        '    i = UBound(m_cUploadThreads)
        '
        '    Do While UBound(m_cUploadThreads) >= 0
        '        Set m_cUploadThreads(i) = Nothing
        '
        '        If Not UBound(m_cUploadThreads) = 0 Then
        '            ReDim Preserve m_cUploadThreads(UBound(m_cUploadThreads) - 1)
        '        Else
        '            Exit Do
        '        End If
        '
        '        i = UBound(m_cUploadThreads)
        '    Loop

136     If pThread.IsThreadRunning = True Then
138         pThread.TerminateWin32Thread
        End If
    
140     If FormThread.IsThreadRunning = True Then

142         FormThread.TerminateWin32Thread
        End If
    
144     If pInitThread.IsThreadRunning Then
146         pInitThread.TerminateWin32Thread
        End If
    
148     If pSubmitGeoMarksThread.IsThreadRunning Then
150         pSubmitGeoMarksThread.TerminateWin32Thread
        End If
    
156     If pUpdateCheckerThread.IsThreadRunning Then
158         pUpdateCheckerThread.TerminateWin32Thread
        End If
    
160     If pSynchThread.IsThreadRunning Then
162         pSynchThread.TerminateWin32Thread
        End If
    
164     If pCheckSynchThread.IsThreadRunning Then
166         pCheckSynchThread.TerminateWin32Thread
        End If
    
168     If incThread.IsThreadRunning Then
170         incThread.TerminateWin32Thread
        End If
    
172     If pGetIncThread.IsThreadRunning Then
174         pGetIncThread.TerminateWin32Thread
        End If
        
        If pLoadMap.IsThreadRunning Then
            pLoadMap.TerminateWin32Thread
        End If
        
        If pCheckInet.IsThreadRunning Then
            pCheckInet.TerminateWin32Thread
        End If
        
        
        
176     ' If Not m_frmOASISProgress Is Nothing Then
178     '    Unload m_frmOASISProgress
        ' End If
        
180     Set pCheckSynchThread = Nothing
182     Set pSynchThread = Nothing
184     Set pUpdateCheckerThread = Nothing
186     Set pLoadLyrAttrToGrdThread = Nothing
188     Set pSubmitGeoMarksThread = Nothing
190     Set pInitThread = Nothing
192     Set pThread = Nothing
194     Set FormThread = Nothing
196     Set incThread = Nothing
198     Set pGetIncThread = Nothing
        Set pLoadMap = Nothing '
        Set pCheckInet = Nothing
        'VERY IMPORTANT - MUST EXPLICITLY CALL END TO TERMINATE APP
        
202     Set mRSUGSettings = Nothing
    
204     UnloadMenuFrames
206     UnloadAllForms
    
208     m_Cnn.Close
    
210     Set m_Cnn = Nothing
    
'212     Set g_PictureDialogLarge = Nothing
'214     Set g_PictureDialogSmall = Nothing
'216     Set g_PictureDialogLogo = Nothing
        
        If g_bOnlineCheckedAtLogin Then oClientInterCom.UnregisterChannel
        If g_bOnlineCheckedAtLogin Then Set oClientInterCom = Nothing
218     If g_bOnlineCheckedAtLogin Then Set oSync = Nothing
220     If g_bOnlineCheckedAtLogin Then Set m_InternetCheck = Nothing
222     If g_bOnlineCheckedAtLogin Then Set m_oSQLLyrSynch = Nothing
        
        Dim Process As Variant

224     For Each Process In GetObject("winmgmts:").ExecQuery("select name from Win32_Process where name='OASIS_SynchNG.exe'")
226         Process.Terminate (0)
        Next
        
        For Each Process In GetObject("winmgmts:").ExecQuery("select name from Win32_Process where name='OASISCommsMon.exe'")
227         Process.Terminate (0)
        Next
        
        For Each Process In GetObject("winmgmts:").ExecQuery("select name from Win32_Process where name='OASIS_Inter_Comms.exe'")
            Process.Terminate (0)
        Next
        
228     KillALL
230     KILLEMHARD
232     Unload Me
    
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.Form_Unload " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub KillALL()

End Sub

Private Sub ShowThemelegend(oLayer As TatukGIS_XDK10.XGIS_LayerVector)
    
End Sub

Public Sub UnloadAllForms()
        '<EhHeader>
        On Error GoTo UnloadAllForms_Err
        '</EhHeader>
    On Error Resume Next
    Dim oFRM As Form

100 For Each oFRM In Forms
102     If oFRM.Name <> Me.Name Then
104         Unload oFRM
106         Set oFRM = Nothing
        End If
    Next

        '<EhFooter>
        Exit Sub

UnloadAllForms_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.UnloadAllForms " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Public Sub UnloadMenuFrames()
        '<EhHeader>
        On Error GoTo UnloadMenuFrames_Err
        '</EhHeader>
        On Error Resume Next
        
100     Unload m_frmOvMap
104     Unload m_fmrAddIncident
108     Unload m_frmMnuOperations
114     Unload m_frmMnuOASISProfile
116     Unload m_frmAddOns
118     Unload m_frmChangeTracer
120     Unload m_frmFreqSettings
        Unload m_frmAttributes
        Unload m_frmSearch
        Unload m_frmTextAnnoSettings
        Unload m_frmLocator
        Unload m_frmUpdateSettings
        Unload m_frmMainSettings
        Unload m_frmSelectorReports
        Unload m_frmSelectorSettings
        Unload m_frmDynamicContent
        
        Set m_frmDynamicContent = Nothing
        Set m_frmSelectorSettings = Nothing
        Set m_frmSelectorReports = Nothing
        Set m_frmMainSettings = Nothing
        Set m_frmAttributes = Nothing
        Set m_frmSearch = Nothing
126     Set m_frmFreqSettings = Nothing
128     Set m_frmChangeTracer = Nothing
130     Set m_frmAddOns = Nothing
134     Set m_frmOvMap = Nothing
138     Set m_fmrAddIncident = Nothing
140     Set m_frmMnuOperations = Nothing
146     Set m_frmMnuOASISProfile = Nothing
148     Set m_frmLocator = Nothing
        Set m_frmUpdateSettings = Nothing
150     Set m_frmOASISCharts = Nothing
        Set m_frmTextAnnoSettings = Nothing
        
        Set ClipBoard_Ext = Nothing
        
        '<EhFooter>
        Exit Sub

UnloadMenuFrames_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.UnloadMenuFrames " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS10_OnAfterPaint(translated As Boolean)
        '<EhHeader>
        On Error GoTo GIS_OnAfterPaint_Err
        '</EhHeader>
            Dim ex As TatukGIS_XDK10.XGIS_Extent
            
            If Not m_bPrevActionUsed Then
                ReDim Preserve m_PrevExt(UBound(m_PrevExt) + 1)
                Set m_PrevExt(UBound(m_PrevExt)) = GIS10.VisibleExtent
            Else
                m_bPrevActionUsed = False
            End If
            
100         translated = True
102         Set ex = GIS10.VisibleExtent
104         Set lP1 = GisUtils.GISPoint(ex.xmin, ex.ymin)
106         Set lP2 = GisUtils.GISPoint(ex.xmax, ex.ymin)
108         Set lP3 = GisUtils.GISPoint(ex.xmax, ex.ymax)
110         Set lP4 = GisUtils.GISPoint(ex.xmin, ex.ymax)
    '        lblP1.Caption = "P1 : x: " + Str(lP1.X) + "   y: " + Str(lP1.Y)
    '        lblP2.Caption = "P2 : x: " + Str(lP2.X) + "   y: " + Str(lP2.Y)
    '        lblP3.Caption = "P3 : x: " + Str(lP3.X) + "   y: " + Str(lP3.Y)
    '        lblP4.Caption = "P4 : x: " + Str(lP4.X) + "   y: " + Str(lP4.Y)
112         miniMapRefresh
            ctlZoomSlider1.SetZoomPointerFromExtent GIS10.VisibleExtent

        '<EhFooter>
        Exit Sub

GIS_OnAfterPaint_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS_OnAfterPaint " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS_OnExtentChange(translated As Boolean)
        '<EhHeader>
        On Error GoTo GIS_OnExtentChange_Err
        '</EhHeader>
        On Error Resume Next
100     Set g_PrevExt = GIS10.VisibleExtent
        '<EhFooter>
        Exit Sub

GIS_OnExtentChange_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS_OnExtentChange " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS10_OnKeyDown(translated As Boolean, Key As Integer, ByVal Shift As TatukGIS_XDK10.XShiftState)
        '<EhHeader>
        On Error GoTo GIS_OnKeyDown_Err
        '</EhHeader>
100   If Key = VK_CONTROL Then
102     If Not vkControl Then ' avoid multiple call on key repeat
104       GIS10.Mode = TatukGIS_XDK10.XgisSelect
106       vkControl = True
        End If
      End If
  
108   If Key = VK_DELETE Then
110     If GIS10.Mode = TatukGIS_XDK10.XgisEdit Then
112       GIS10.Editor.DeleteShape
114       GIS10.Mode = TatukGIS_XDK10.XgisSelect
        End If
      End If

        '<EhFooter>
        Exit Sub

GIS_OnKeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS_OnKeyDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub GetAdmCode(ptg As TatukGIS_XDK10.XGIS_Point, _
                      sAdm1 As String, _
                      sAdm2 As String, _
                      sAdmloc As String)
        '<EhHeader>
        On Error GoTo GetAdmCode_Err
        '</EhHeader>

        Dim oVecLyr As TatukGIS_XDK10.XGIS_LayerVector
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim oshp As TatukGIS_XDK10.XGIS_Shape
        Dim j As Integer
    
100     Set oVecLyr = GIS10.get(m_sAdmVal1)
                        
102     If Not oVecLyr Is Nothing Then
        
104         Set oshp = oVecLyr.Locate(ptg, 5 / GIS10.zoom, True)
                        
106         If Not oshp Is Nothing Then
        
108             SafeMoveFirst g_RSAppSettings
110             g_RSAppSettings.Find "SettingName = 'AdminLevel0'"
112             sAdm1 = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value)
    
                'For j = 0 To oVecLyr.Fields.Count - 1
                '    DebugPrint  oVecLyr.Name & "_____" & oVecLyr.Fields.Item(j).Name & ":" & oShp.GetField(oVecLyr.Fields.Item(j).Name)
                'Next

            End If
        End If

114     Set oVecLyr = GIS10.get(m_sAdmVal2)
                        
116     If Not oVecLyr Is Nothing Then
118         Set oshp = oVecLyr.Locate(ptg, 5 / GIS10.zoom, True)
                        
120         If Not oshp Is Nothing Then
            
122             SafeMoveFirst g_RSAppSettings
124             g_RSAppSettings.Find "SettingName = 'AdminLevel1'"
126             sAdm2 = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value)

                'For j = 0 To oVecLyr.Fields.Count - 1
                '    DebugPrint  oVecLyr.Name & "_____" & oVecLyr.Fields.Item(j).Name & ":" & oShp.GetField(oVecLyr.Fields.Item(j).Name)
                'Next

            End If
        End If

        If m_bUseDistrictOnly Then
        Dim x As Double
        Dim y As Double
            'x = oShp.Centroid.x
            'y = oShp.Centroid.y
            
            If Not oshp Is Nothing Then
                Set ptg = oshp.Centroid
                sAdmloc = "N/A"
            End If
        Else

128         SafeMoveFirst g_RSAppSettings
130         g_RSAppSettings.Find "SettingName = 'AdminLocation'"
            
132     Set oVecLyr = GIS10.get(g_RSAppSettings.Fields.Item("SettingValue1").value)
                        
134     If Not oVecLyr Is Nothing Then
136         Set oshp = oVecLyr.Locate(ptg, 20 / GIS10.zoom, True)
                        
138         If Not oshp Is Nothing Then

140             sAdmloc = oshp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value)
    
                '            For j = 0 To oVecLyr.Fields.Count - 1
                '                DebugPrint  oVecLyr.Name & "_____" & oVecLyr.Fields.Item(j).Name & ":" & oShp.GetField(oVecLyr.Fields.Item(j).Name)
                '            Next

                End If
            End If
        End If

        '<EhFooter>
        Exit Sub

GetAdmCode_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GetAdmCode " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function GetAdminNameFromPoint(ptg As TatukGIS_XDK10.XGIS_Point, _
                                      sAdmLayer As String, _
                                      sAdmField As String, _
                                      dPrecision As Double) As String

    Dim oVecLyr As TatukGIS_XDK10.XGIS_LayerVector
    Dim shp As TatukGIS_XDK10.XGIS_Shape
    Dim oshp As TatukGIS_XDK10.XGIS_Shape
    Dim j As Integer
    
    Set oVecLyr = GIS10.get(sAdmLayer)
    GetAdminNameFromPoint = ""

    If Not oVecLyr Is Nothing Then
        
        Set oshp = oVecLyr.Locate(ptg, dPrecision, False)

        If Not oshp Is Nothing Then GetAdminNameFromPoint = oshp.GetField(sAdmField)

    End If
    
    Set oVecLyr = Nothing
    Set oshp = Nothing

End Function

Private Sub GIS10_OnMouseDown(translated As Boolean, _
                              ByVal Button As TatukGIS_XDK10.XMouseButton, _
                              ByVal Shift As TatukGIS_XDK10.XShiftState, _
                              ByVal x As Long, _
                              ByVal y As Long)
        '<EhHeader>
        On Error GoTo GIS10_OnMouseDown_Err
        '</EhHeader>
100     translated = True

        Dim sAdm1 As String
        Dim sAdm2 As String
        Dim sAdm3 As String
        Dim sAdm4 As String
        Dim sAdm5 As String
        Dim sAdmloc As String
        Dim sAdmins() As String
        Dim sAdminNames() As String
        Dim lUID As Long
        Dim i As Integer
        Dim j As Integer
        Dim fdesc As String
        Dim oVecLyr As TatukGIS_XDK10.XGIS_LayerVector
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim oshp As TatukGIS_XDK10.XGIS_Shape
        Dim bRunUrl As Boolean
  
102     DebugPrint "----- **MouseDown Start** ----"
104     DebugPrint "gdpts:" & UBound(gdpts)
106     DebugPrint "ptsSel:" & UBound(ptsSel)
  
108     If Button = XmbRight Then
            Dim Pt As POINTAPI
110         GetCursorPos Pt
112         PopupMenu mnuMAPPopUp, , ScaleX(Pt.x - 2, vbPixels, vbTwips), ScaleY(Pt.y - 8, vbPixels, vbTwips)
            Exit Sub
        End If
  
114     Set m_dPTG = GIS10.ScreenToMap(GisUtils.POINT(x, y))

        'ProjTest m_dPTG.X, m_dPTG.y

116     Select Case g_CurrentTool

            Case oPointBuffer
            m_oBufferLyr.RevertAll
                Set shp = m_oBufferLyr.CreateShape(XgisShapeTypePoint)  ' GIS10.Locate(m_dPTG, 5 / GIS10.zoom)
                shp.Lock XgisLockExtent
                shp.AddPart
                shp.AddPoint m_dPTG
                shp.Unlock
                AddShapeWithBufferAndHighlightIntersect shp, ctlSelector1.GetBuffer

            Case oSingleSelect
            
118             Set shp = GIS10.Locate(m_dPTG, 5 / GIS10.zoom) ' 5 pixels precision
            
120             If Not shp Is Nothing Then
122                 shp.MakeEditable
124                 shp.Flash
126                 m_frmAttributes.ShowShape shp
128                 shp.IsSelected = True
130                 m_frmAttributes.ShowSelected shp.layer
                End If
    
132         Case oInfo
            
134             If m_frmAttributes Is Nothing Then
136                 Set m_frmAttributes = New frmAttributes
                End If
                
138             If Not m_frmAttributes.Visible Then
140                 m_frmAttributes.Init GIS10.viewer
                End If
                
142             m_frmAttributes.Show vbModeless, Me
            
144             Set oVecLyr = GIS10.get(m_frmAttributes.LyrCol.Item(m_frmAttributes.ComLayer.List(m_frmAttributes.ComLayer.ListIndex)))
            
146             If Not oVecLyr Is Nothing Then
                 
148                 Set oshp = oVecLyr.Locate(m_dPTG, 15 / GIS10.zoom, True)
150                 m_frmAttributes.ShowShape oshp
                    
152                 m_frmAttributes.caption = "Active layer: " & oVecLyr.caption

154                 If Not oshp Is Nothing Then

156                     If Not g_RSGISGridTableSettings.EOF Or Not g_RSGISGridTableSettings.Bof Then
                        
158                         SafeMoveFirst g_RSGISGridTableSettings
160                         g_RSGISGridTableSettings.Find "name = '" & oVecLyr.Name & "'"
                    
162                         If Not g_RSGISGridTableSettings.EOF Then
164                             If g_RSGISGridTableSettings.Fields.Item("isURLLayer").value Then
166                                 If g_RSGISGridTableSettings.Fields.Item("autoRunUrls").value Then
168                                     bRunUrl = True
                                    End If
                                End If
                            End If
    
170                         If bRunUrl Then
172                             If oshp.GetField(g_RSGISGridTableSettings.Fields.Item("URLLayerField").value) <> "" Then
174                                 ShellExecute Me.hWnd, vbNullString, oshp.GetField(g_RSGISGridTableSettings.Fields.Item("URLLayerField").value), vbNullString, "C:\", 1
                                End If
                            End If
                            
                        End If

176                     If Shift = XssShift Then
178                         frmSendMess.lvwAttributes.ListItems.Add , , "X"
180                         frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oshp.Centroid.x
                            
182                         frmSendMess.lvwAttributes.ListItems.Add , , "Y"
184                         frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oshp.Centroid.y
                            
186                         For j = 0 To oVecLyr.Fields.Count - 1
188                             frmSendMess.lvwAttributes.ListItems.Add , , oVecLyr.Fields.Item(j).Name
190                             frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oshp.GetField(oVecLyr.Fields.Item(j).Name)
                            Next
                            
192                         frmSendMess.Show vbModal, Me
                        End If
                        
                        'New Stuff
                        'frmAttributes.ShowSelected oVecLyr
                        
                    End If
                    
                Else
                    
194                 Set oshp = GIS10.Locate(m_dPTG, 5 / GIS10.zoom) ' 5 pixels precision
            
196                 If Not oshp Is Nothing Then m_frmAttributes.caption = "Active layer: " & oshp.layer.caption
                    
198                 m_frmAttributes.ShowShape oshp

200                 If Shift = 17 Then
202                     frmSendMess.lvwAttributes.ListItems.Add , , "X"
204                     frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oshp.Centroid.x
                            
206                     frmSendMess.lvwAttributes.ListItems.Add , , "Y"
208                     frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oshp.Centroid.y
                            
210                     Set oVecLyr = oshp.layer
                            
212                     For j = 0 To oVecLyr.Fields.Count - 1
214                         frmSendMess.lvwAttributes.ListItems.Add , , oVecLyr.Fields.Item(j).Name
216                         frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oshp.GetField(oVecLyr.Fields.Item(j).Name)
                        Next
                            
218                     frmSendMess.Show vbModal, Me
                    End If
                End If

                'frmAttributes.Show vbModeless, Me
            
220         Case OASIS_TOOLS.oCreateLocationArea, oCreateLocationPolyline, oCreateLocationPoint, oCreateLocationLine, oCreateLocationMultipoint, oCreateLocationPoint

222             Select Case g_CurrentFeatureType
                
                    Case OASISFeatureTypes.Custom

224                     If m_frmSpatialiseDD.Visible Then

                            If Not m_ShapeTemp Is Nothing Then m_ShapeTemp.Delete
226                         Set oshp = mDDLayer.CreateShape(XgisShapeTypePoint)
                            Set m_ShapeTemp = oshp
228                         oshp.Lock TatukGIS_XDK10.XgisLockExtent
230                         oshp.AddPart
232                         oshp.AddPoint m_dPTG
234                         UpdateSpatialiseDDList oshp

                        End If
                        
236                     If g_bIncidentsV2 Then
                        
238                         If m_frmIncidentsV2DataEntry.Visible Then
                                
240                             With m_frmIncidentsV2DataEntry

                                    Dim iMaxAdminPolygons As Integer
                                    Dim iCount As Integer
242                                 iMaxAdminPolygons = 4
                                    
244                                 SafeMoveFirst g_RSAppSettings
246                                 g_RSAppSettings.Find "SettingName = 'AdminLevel3'"

248                                 If IsNull(g_RSAppSettings.Fields("SettingValue2").value) Then
250                                     .txtLocationAdmin5.Visible = False
252                                     iMaxAdminPolygons = 3
                                    End If
                                        
254                                 SafeMoveFirst g_RSAppSettings
256                                 g_RSAppSettings.Find "SettingName = 'AdminLevel2'"

258                                 If IsNull(g_RSAppSettings.Fields("SettingValue2").value) Then
260                                     .txtLocationAdmin4.Visible = False
262                                     iMaxAdminPolygons = 2
                                    End If
                                    
264                                 SafeMoveFirst g_RSAppSettings
266                                 g_RSAppSettings.Find "SettingName = 'AdminLevel1'"

268                                 If IsNull(g_RSAppSettings.Fields("SettingValue2").value) Then
270                                     .txtLocationAdmin3.Visible = False
272                                     iMaxAdminPolygons = 1
                                    End If
                                    
274                                 SafeMoveFirst g_RSAppSettings
276                                 g_RSAppSettings.Find "SettingName = 'AdminLevel0'"

278                                 If IsNull(g_RSAppSettings.Fields("SettingValue2").value) Then
280                                     .txtLocationAdmin2.Visible = False
282                                     iMaxAdminPolygons = 0
                                    End If
                                    
284                                 ReDim sAdmins(6)
            
286                                 iCount = 0

288                                 Do Until iCount = iMaxAdminPolygons
290                                     SafeMoveFirst g_RSAppSettings
292                                     g_RSAppSettings.Find "SettingName = 'AdminLevel" & iCount & "'"
294                                     sAdmins(iCount) = GetAdminNameFromPoint(m_dPTG, g_RSAppSettings.Fields("SettingValue1").value, g_RSAppSettings.Fields("SettingValue2").value, 200 / GIS10.zoom)
296                                     iCount = iCount + 1
                                    Loop
                                         
298                                 SafeMoveFirst g_RSAppSettings
300                                 g_RSAppSettings.Find "SettingName = 'AdminLocation'"
302                                 sAdmins(iCount) = GetAdminNameFromPoint(m_dPTG, g_RSAppSettings.Fields("SettingValue1").value, g_RSAppSettings.Fields("SettingValue2").value, 200 / GIS10.zoom)
                                  
                                End With

304                             m_frmIncidentsV2DataEntry.SetLocationParams sAdmins(0), sAdmins(1), sAdmins(2), sAdmins(3), sAdmins(4), CStr(m_dPTG.x), CStr(m_dPTG.y)

                                Dim oCNIncidentsV2 As New ADODB.Connection
                                Dim oRSIncidentsV2 As New ADODB.Recordset
                                Dim dDist As Double
306                             oCNIncidentsV2.Open g_sIncidentsV2ConnectionString

308                             With oRSIncidentsV2
                                
310                                 .Open "SELECT [option] from [dd_" & g_sIncidentsV2DDName & "_ddLayersToFind]", oCNIncidentsV2, adOpenDynamic, adLockReadOnly

312                                 If .State = adStateOpen Then
314                                     m_frmIncidentsV2DataEntry.lstLocations.Clear

316                                     Do Until .EOF

                                            If Not GIS10.get(.Fields("option").value) Is Nothing Then
                                            
317                                         dDist = GetNearestShape(m_dPTG, .Fields("option").value, False)
318                                         m_frmIncidentsV2DataEntry.lstLocations.AddItem GIS10.get(.Fields("option").value).caption & ": " & IIf(dDist = 0.00006666666, "?", Round(dDist, 2)) & " km"
                                            
                                            End If
                                            
320                                         .MoveNext
                                        Loop

322                                     .Close
                                    
                                    End If
    
                                End With
                                
324                             oCNIncidentsV2.Close
326                             Set oRSIncidentsV2 = Nothing
328                             Set oCNIncidentsV2 = Nothing

330                             ReDim ptsSel(0)
332                             ReDim gdpts(0)
334                             g_CurrentTool = oPan
336                             g_CurrentFeatureType = Custom
338                             m_bPolyL = False
340                             m_bPolyG = False
342                             SetOASISTool
    
                                ',
                                'Round(GetNearestShape(m_dPTG, "WV Warehouses", False), 2) & " km",
                                'Round(GetNearestShape(m_dPTG, "WV Op Areas", False), 2) & " km"
                            
                            End If
                        
                        End If
                    
344                 Case OASISFeatureTypes.GeoMark
                    
                       
426                 Case OASISFeatureTypes.Incident
428                     GetAdmCode m_dPTG, sAdm1, sAdm2, sAdmloc
                        'MsgBox GetNearestShape(m_dPTG, "Office Locations")
                    
430                     If Not m_oIncShpPt Is Nothing Then
432                         m_oSQLIncLyr.Delete m_oIncShpPt.uID
                        End If
                    
434                     Set m_oIncShpPt = m_oSQLIncLyr.CreateShape(XgisShapeTypePoint)
                        
436                     m_oIncShpPt.Lock TatukGIS_XDK10.XgisLockExtent
438                     m_oIncShpPt.AddPart

440                     m_oIncShpPt.AddPoint m_dPTG
442                     m_oIncShpPt.SetField "ID", Rnd(300000)
444                     m_oIncShpPt.SetField "Name", m_fmrAddIncident.txtEnteredBy.Text
446                     m_oIncShpPt.SetField "Type", m_fmrAddIncident.ComIncType.List(m_fmrAddIncident.ComIncType.ListIndex)
448                     m_oIncShpPt.SetField "Target", m_fmrAddIncident.ComIncTarget.List(m_fmrAddIncident.ComIncTarget.ListIndex)
450                     m_oIncShpPt.SetField "Incident_DATE", m_fmrAddIncident.MVIncident.value
452                     m_oIncShpPt.SetField "Town", sAdmloc
454                     m_oIncShpPt.SetField "District", sAdm2
456                     m_oIncShpPt.SetField "Province", sAdm1
458                     m_oIncShpPt.SetField "Description", m_fmrAddIncident.txtIncidentDescription.Text & vbCrLf & m_fmrAddIncident.txtLocationDescription.Text
                        
                        'm_fmrAddIncident.SetNearestValues GetNearestShape(m_dPTG, "Office Locations"), GetNearestShape(m_dPTG, "Operational Areas")

460                     m_fmrAddIncident.SetCoordinateValue CStr(m_dPTG.x), "Long:", CStr(m_dPTG.y), "Lat:", "Lat/Long WGS84", Map1.MilitaryGridReferenceFromPoint(m_dPTG.x, m_dPTG.y)
462                     m_fmrAddIncident.SetFocus
                    
464                     m_oIncShpPt.Unlock

                        'lUID = m_oIncShpPt.Uid
                    
                        'Set m_oIncShpPt = m_oSQLIncLyr.AddShape(m_oIncShpPt)
                        
                        'm_oSQLIncLyr.Delete lUID
                    
466                     m_fmrAddIncident.SetAdmValues sAdm1, sAdm2, sAdmloc
                                        
468                     GIS10.UpDate

470                 Case OASISFeatureTypes.Location
                
472                     If Button = XmbLeft Then
474                         If g_oEditLayer Is Nothing Then
                                Exit Sub
                            Else
 
476                             Select Case g_CurrentTool
 
                                    Case OASIS_TOOLS.oCreateLocationArea
                                    
478                                     GIS10.Editor.CreateShape g_oEditLayer, m_dPTG, TatukGIS_XDK10.XgisShapeTypePolygon, XgisDimensionTypeUnknown
480                                     Set g_oEditLayer = Nothing
                                        
482                                 Case OASIS_TOOLS.oCreateLocationPolyline

484                                     GIS10.Editor.CreateShape g_oEditLayer, m_dPTG, TatukGIS_XDK10.XgisShapeTypeArc, XgisDimensionTypeUnknown
486                                     Set g_oEditLayer = Nothing

488                                 Case OASIS_TOOLS.oCreateLocationPoint
490                                     GIS10.Editor.CreateShape g_oEditLayer, m_dPTG, TatukGIS_XDK10.XgisShapeTypePoint, XgisDimensionTypeUnknown
492                                     Set g_oEditLayer = Nothing

494                                 Case oCreateLocationMultipoint
496                                     GIS10.Editor.CreateShape g_oEditLayer, m_dPTG, TatukGIS_XDK10.XgisShapeTypeMultiPoint, XgisDimensionTypeUnknown
498                                     Set g_oEditLayer = Nothing
                                End Select
                    
                            End If

                            '                End If

                        Else

500                         Select Case g_CurrentTool
            
                                Case OASIS_TOOLS.oCreateLocationArea, oCreateLocationPolyline, oCreateLocationPoint, oCreateLocationLine, oCreateLocationMultipoint, oCreateLocationPoint

502                                 EndEdit
                                    'Set m_oShpPolygon = m_oShpPolygon.AsPolygon
                                    'frmRegionStyle.Init m_oShpPolygon
                                    'frmRegionStyle.Show vbModal, Me
                                    'cmdCommand2_Click
504                                 GIS10.UpDate

                                    ' Case OASIS_TOOLS.oCreateLocationPolyline
                                    '     Set m_oShpArc = m_oShpArc.AsArc
                            End Select

                        End If

                End Select

506         Case OASIS_TOOLS.oRadiusSelect, OASIS_TOOLS.oCircleSelect
            
508             Set oldPos = GisUtils.POINT(x, y)
510             oldRadius = 0

512         Case oLineSelect, oPolyLineSelect, oPolySelect, oRectSelect, oMultiSelect, oAreaSelect
            
514             If g_CurrentTool = oLineSelect Then m_bPolyL = True
516             If g_CurrentTool = oAreaSelect Then m_bPolyG = True
                
                '        If m_bPolyG Or m_bPolyL Then
                
518             GIS10.PaintMinimum
                
                ' Private gdpts() As POINTAPI
            
520             Set ptsSel(UBound(ptsSel)) = GIS10.ScreenToMap(GisUtils.POINT(x, y))
522             gdpts(UBound(gdpts)).x = x: gdpts(UBound(gdpts)).y = y
524             GDI_Polyline GIS10.hdc, gdpts
                
526             If Not m_bDrawFinished Then
528                 ReDim Preserve gdpts(UBound(gdpts) + 1)
530                 ReDim Preserve ptsSel(UBound(ptsSel) + 1)
                Else
532                 m_bDrawFinished = False
                End If
            
534             Set oldPos = GisUtils.POINT(x, y)

536         Case OASIS_TOOLS.oFeatureSelect
            
538             Set oVecLyr = GIS10.get(m_SelLyrCol.Item(ComFeatureLayer.List(ComFeatureLayer.ListIndex)))
            
540             If Not oVecLyr Is Nothing Then
                 
542                 Set oshp = oVecLyr.Locate(m_dPTG, 15 / GIS10.zoom, True)

544                 If Not oshp Is Nothing Then

546                     doFeatureSelect oshp
                    End If
                    
                End If
            
548         Case OASIS_TOOLS.oCreateLocationText
550             AddAnnoText m_dPTG
        End Select

552     DebugPrint "----- **MouseDown End** ----"
554     DebugPrint "gdpts:" & UBound(gdpts)
556     DebugPrint "ptsSel:" & UBound(ptsSel)
        
        'Buggfix Works as designed is in kilometers Description:I have tried to select settlements with the line selector tool, with setting of 2.0 and contains as the overlay function. I assume 2.0 refers to the buffer size (in km??)?
        'I however seem to get way too many settlements selected, and not only the ones in the selection area.
        
        '<EhFooter>
        Exit Sub

GIS10_OnMouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS10_OnMouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function doFeatureSelect(oshp As TatukGIS_XDK10.XGIS_Shape) As Long
        '<EhHeader>
        On Error GoTo doFeatureSelect_Err
        '</EhHeader>
        Dim i As Integer
        Dim ptg As TatukGIS_XDK10.XGIS_Point
'GIS109 Done
100     RemoveAlltabs
    
102     With m_oDrawLyr
104         .Lock
        
106         doFeatureSelect = oshp.uID

108         .draw
    
            Dim lL As TatukGIS_XDK10.XGIS_LayerVector
            Dim tmp As TatukGIS_XDK10.XGIS_Shape
            Dim otmp As TatukGIS_XDK10.XGIS_Shape
            Dim buf1 As New TatukGIS_XDK10.XGIS_Shape
            Dim buf2 As New TatukGIS_XDK10.XGIS_Shape
        
            Dim sVals As String
            Dim tpl As New TatukGIS_XDK10.XGIS_Topology

110         m_oBufferLyr.RevertAll
            ' 1 degree in km 111.2
112         oshp.MakeEditable
666         buf2.MakeEditable
667         Set buf2 = m_oBufferLyr.AddShape(tpl.MakeBuffer(oshp, m_udtSelectorSettings.dBuffeLevel / 111.2, 36, True))

120         m_lBufUID = buf2.uID
        
122         Set buf1 = Nothing
124         Set tpl = Nothing
    
126         Set lL = GIS10.get(m_SelLyrCol.Item(ComSelLayer.List(ComSelLayer.ListIndex)))

128         If lL Is Nothing Then
130             .Unlock
132             GIS10.UpDate
                Exit Function
            End If
    
134         lL.DeselectAll
    
            ' check all shapes
138         For Each tmp In lL.Loop(buf2.Extent, "", buf2, m_udtSelectorSettings.sSpatialOperator, True)
140             Set otmp = tmp.MakeEditable
142             otmp.IsSelected = True
144             AddSelection otmp
                Set otmp = Nothing
            Next
148         .Unlock
150         GIS10.PaintMinimum

        End With

        '<EhFooter>
        Exit Function

doFeatureSelect_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.doFeatureSelect " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function CreatePolygon() As Long
        '<EhHeader>
        On Error GoTo CreatePolygon_Err
        '</EhHeader>
        Dim i As Integer
        Dim ptg As TatukGIS_XDK10.XGIS_Point
        Dim oshp As TatukGIS_XDK10.XGIS_Shape
    
100     GIS10.Mode = TatukGIS_XDK10.XgisEdit
    
102     GIS10.Editor.CreateShape m_oDrawLyr, ptsSel(0), TatukGIS_XDK10.XgisShapeTypePolygon, XgisDimensionTypeUnknown
    
104     For i = LBound(ptsSel) + 1 To UBound(ptsSel)
106         GIS10.Editor.AddPoint ptsSel(i)
        Next

108     CreatePolygon = GIS10.Editor.uID

110     m_oDrawLyr.draw
    
        '<EhFooter>
        Exit Function

CreatePolygon_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CreatePolygon " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub UpdateSpatialiseDDList(sShp As TatukGIS_XDK10.XGIS_Shape)

    Dim i  As Integer
    Dim sPoints As String
    i = 0
    sPoints = ""
    
    Do Until i = sShp.GetNumPoints
            
        If Len(sPoints) > 1 Then
            sPoints = sPoints & ";" & sShp.GetPoint(0, i).x & ";" & sShp.GetPoint(0, i).y
        Else
            sPoints = sShp.GetPoint(0, i).x & ";" & sShp.GetPoint(0, i).y
        End If

        i = i + 1
                
    Loop

    Set sShp = m_oSpatialiseLayer.AddShape(sShp, True)
    'sShp.draw
    sShp.Flash
    m_oSpatialiseLayer.SaveAll
    m_frmSpatialiseDD.SetListValues sPoints
    EndEdit
    
End Sub

Private Sub DynamicDataModule1_GetCurrentExtentCentroid(oPoint As TatukGIS_XDK10.XGIS_Point)

    oPoint.Prepare GIS10.VisibleExtent.xmin + ((GIS10.VisibleExtent.xmax - GIS10.VisibleExtent.xmin) / 2), GIS10.VisibleExtent.ymin + ((GIS10.VisibleExtent.ymax - GIS10.VisibleExtent.ymin) / 2)

End Sub

Private Function CreatePolygonEX() As Long
        '<EhHeader>
        On Error GoTo CreatePolygonEX_Err
        '</EhHeader>
        Dim i As Integer
        Dim ptg As TatukGIS_XDK10.XGIS_Point
        Dim oshp As TatukGIS_XDK10.XGIS_Shape
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
        Dim tmp As TatukGIS_XDK10.XGIS_Shape
        Dim otmp As TatukGIS_XDK10.XGIS_Shape
        Dim buf1 As TatukGIS_XDK10.IXGIS_Shape
        Dim buf2 As TatukGIS_XDK10.IXGIS_Shape
        Dim sVals As String
        Dim tpl As TatukGIS_XDK10.XGIS_Topology
        Dim ipoints As Integer
        Dim iparts As Integer
        Dim oLayer As TatukGIS_XDK10.XGIS_LayerAbstract
        
        Dim dDistLen As Double
        Dim dLat1 As Double
        Dim dLon1 As Double
  
100     If Not m_frmSpatialiseDD.Visible Then
102         m_oBufferLyr.RevertAll
104         Set oLayer = m_oBufferLyr 'GIS10.get("Buffers") 'm_oDrawLyr '  '  ' m_oBufferLyr 'm_oDrawLyr
        Else
106         Set oLayer = mDDLayer ' m_oSpatialiseLayer
        End If
    
108     With oLayer
110         .Lock
    
112         If m_frmSpatialiseDD.Visible Then
114             Set oshp = mDDLayer.CreateShape(mDDLayer.GetShape(mDDLayer.GetLastUid).ShapeType)
            Else
116             Set oshp = .CreateShape(XgisShapeTypePolygon)
            End If
    
118         oshp.Lock TatukGIS_XDK10.XgisLockExtent
120         oshp.AddPart
        
            dLat1 = ptsSel(0).x
            dLon1 = ptsSel(0).y
        
122         For i = LBound(ptsSel) To UBound(ptsSel)
124             oshp.AddPoint ptsSel(i)
                If i = 1 Then dDistLen = dDistLen + HaversineDistance(dLon1, ptsSel(i).x, dLat1, ptsSel(i).y)
                If i > 1 Then dDistLen = dDistLen + HaversineDistance(ptsSel(i - 1).x, ptsSel(i).x, ptsSel(i - 1).y, ptsSel(i).y)
            Next

126         oshp.Unlock
128         EndEdit

130         If m_frmSpatialiseDD.Visible Then UpdateSpatialiseDDList oshp
132         CreatePolygonEX = oshp.uID

            If Not m_frmSpatialiseDD.Visible And ctlSelector1.DistanceEnabled Then
                'MsgBox "Circumference length: " & Round(dDistLen, 2) & " km" ' & vbCrLf & "Area of coverage: " & CStr(Round(CDbl(oshp.area * 111.3199), 2)) & " km2"
            End If

        End With
    
        '<EhFooter>
        Exit Function

CreatePolygonEX_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.CreatePolygonEX " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function CreatePolyLine() As Long
        '<EhHeader>
        On Error GoTo CreatePolyLine_Err
        '</EhHeader>
        Dim i As Integer
        Dim ptg As TatukGIS_XDK10.XGIS_Point
        Dim oshp As TatukGIS_XDK10.XGIS_Shape
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
        Dim tmp As TatukGIS_XDK10.XGIS_Shape
        Dim otmp As TatukGIS_XDK10.XGIS_Shape
        Dim buf1 As TatukGIS_XDK10.IXGIS_Shape
        Dim buf2 As TatukGIS_XDK10.IXGIS_Shape
        Dim sVals As String
        Dim tpl As TatukGIS_XDK10.XGIS_Topology
        Dim oLayer As TatukGIS_XDK10.XGIS_LayerAbstract
        
        Dim dDistLen As Double
        Dim dLat1 As Double
        Dim dLon1 As Double
        
100     If Not m_frmSpatialiseDD.Visible Then
102         m_oBufferLyr.RevertAll
104         Set oLayer = m_oBufferLyr ' m_oDrawLyr
        Else
106         Set oLayer = mDDLayer
        End If
        
108     With oLayer
110         .Lock
    
112         If m_frmSpatialiseDD.Visible Then
114             Set oshp = mDDLayer.GetShape(mDDLayer.GetLastUid)   ' mDDShape 'm_frmSpatialiseDD.GetShape
            Else
116             Set oshp = .CreateShape(XgisShapeTypeArc)
            End If
    
118         oshp.Lock TatukGIS_XDK10.XgisLockExtent
120         oshp.AddPart

            dLat1 = ptsSel(0).y
            dLon1 = ptsSel(0).x
        
122         For i = LBound(ptsSel) To UBound(ptsSel)
124             oshp.AddPoint ptsSel(i)
                If i = 1 Then dDistLen = dDistLen + HaversineDistance(dLon1, ptsSel(i).x, dLat1, ptsSel(i).y)
                If i > 1 Then dDistLen = dDistLen + HaversineDistance(ptsSel(i - 1).x, ptsSel(i).x, ptsSel(i - 1).y, ptsSel(i).y)
            Next
            
126         oshp.Unlock
128         EndEdit

            If Not m_frmSpatialiseDD.Visible And ctlSelector1.DistanceEnabled Then
                MsgBox "Line length: " & Round(dDistLen, 2) & " km"
            End If

130         If m_frmSpatialiseDD.Visible Then UpdateSpatialiseDDList oshp
134         CreatePolyLine = oshp.uID

        End With

        '<EhFooter>
        Exit Function

CreatePolyLine_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.CreatePolyLine " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub SetSelectionStyle()
Dim i As Integer
    Dim oLyr As New TatukGIS_XDK10.XGIS_LayerVector
    Dim oLine As TatukGIS_XDK10.XGIS_Shape
    Dim oPol As TatukGIS_XDK10.XGIS_Shape
    Dim oPt As TatukGIS_XDK10.XGIS_Shape
    Dim Pt As TatukGIS_XDK10.XGIS_Point
        
    oLyr.Open
    
    oLyr.Params.area.Pattern = GIS10.SelectionPattern
    oLyr.Params.Transparancy = GIS10.SelectionTransparency
    oLyr.Params.area.Pattern = GIS10.SelectionColor
    oLyr.Params.area.Pattern = GIS10.SelectionOutlineOnly
    oLyr.Params.area.Pattern = GIS10.SelectionWidth

    Set oLine = oLyr.CreateShape(XgisShapeTypeArc)
        
    Dim ptg As TatukGIS_XDK10.XGIS_Point
    Dim oshp As TatukGIS_XDK10.XGIS_Shape
    
    oLyr.Lock
    
    oLine.Lock TatukGIS_XDK10.XgisLockExtent
    oLine.AddPart
        
    For i = 0 To 2
        Set Pt = New TatukGIS_XDK10.XGIS_Point
        Pt.x = i
        Pt.y = i
        oLine.AddPoint Pt
    Next
        
    oLine.Unlock
    oLyr.Unlock
        
   ' SetParent GisUtils.GisShowLayerProperties(oLyr, Nothing), elMap.hWnd
    
    Dim WinhWnd As Long
    Dim oldhwnd As Long  ' receives handle of button's former parent
    'WinhWnd = FindWindow(vbNullString, "Vector:")
    
    'oldhwnd = SetParent(WinhWnd, elMap.hWnd)
    'MoveWindow WinhWnd, 0, -20, 1100, 750, 1
End Sub

Private Sub GIS10_OnMouseMove(translated As Boolean, _
                              ByVal Shift As TatukGIS_XDK10.XShiftState, _
                              ByVal x As Long, _
                              ByVal y As Long)
        '<EhHeader>
        On Error GoTo GIS10_OnMouseMove_Err
        '</EhHeader>
        Dim ptg As TatukGIS_XDK10.XGIS_Point
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim i As Integer
        Dim sMGRS As String
            
        Dim j As Integer
        Dim oVecLyr As TatukGIS_XDK10.XGIS_LayerVector
        Dim oshp As TatukGIS_XDK10.XGIS_Shape
        Dim bRunUrl As Boolean
  
100     If m_bScrib Then
102         GIS10.PaintMinimum
            '        pts(UBound(pts)).X = ScaleX(X, vbPixels, vbTwips): pts(UBound(pts)).Y = ScaleY(Y, vbPixels, vbTwips)
104         gdpts(UBound(gdpts)).x = x: gdpts(UBound(gdpts)).y = y
        
106         GDI_Polyline GIS10.hdc, gdpts
108         ReDim Preserve gdpts(UBound(gdpts) + 1)
        
        End If

110     If m_bPolyG Or m_bPolyL Then
112         GIS10.PaintMinimum
114         Set ptsSel(UBound(ptsSel)) = GIS10.ScreenToMap(GisUtils.POINT(x, y))
116         gdpts(UBound(gdpts)).x = x: gdpts(UBound(gdpts)).y = y
118         GDI_Polyline GIS10.hdc, gdpts
        End If

120     If m_bRECT Then
122         GIS10.PaintMinimum

124         If UBound(gdpts) > 0 Then
126             gdpts(1).x = x: gdpts(1).y = y
128             GDI_Rectangle GIS10.hdc, gdpts(0).x, gdpts(0).y, gdpts(1).y, gdpts(1).y
            End If
        
        End If
  
130     Set ptg = GIS10.ScreenToMap(GisUtils.POINT(x, y))

132     If oMapTipSetting.Enabled And Not oMapTipSetting.MapTipLayer = "" Then
134         If oMapTipSetting.MapTipLayer = "--All--" Then
136             Set m_oToolTipSHP = GIS10.Locate(ptg, 5 / GIS10.zoom) ' 5 pixels precision
            Else

138             If Not m_LyrCol Is Nothing Then
140                 Set m_oToolTipSHP = GIS10.get(m_LyrCol.Item(oMapTipSetting.MapTipLayer)).Locate(ptg, 5 / GIS10.zoom)
                Else
142                 fillm_LyrCol
                End If
            End If
        End If
        
144     If g_CurrentTool = oInfo Then
        
146         If Not m_frmAttributes Is Nothing Then
                
148             If m_frmAttributes.Visible Then
                    
150                 If m_frmAttributes.chkDynamicInfo = vbChecked Then
            
152                     Set oVecLyr = GIS10.get(m_frmAttributes.LyrCol.Item(m_frmAttributes.ComLayer.List(m_frmAttributes.ComLayer.ListIndex)))
            
154                     If Not oVecLyr Is Nothing Then
                 
156                         Set oshp = oVecLyr.Locate(ptg, 15 / GIS10.zoom, True)
158                         m_frmAttributes.ShowShape oshp
                    
160                         m_frmAttributes.caption = "Active layer: " & oVecLyr.caption

162                         If Not oshp Is Nothing Then

164                             If Not g_RSGISGridTableSettings.EOF Or Not g_RSGISGridTableSettings.Bof Then
                        
166                                 SafeMoveFirst g_RSGISGridTableSettings
168                                 g_RSGISGridTableSettings.Find "name = '" & oVecLyr.Name & "'"
                    
170                                 If Not g_RSGISGridTableSettings.EOF Then
172                                     If g_RSGISGridTableSettings.Fields.Item("isURLLayer").value Then
174                                         If g_RSGISGridTableSettings.Fields.Item("autoRunUrls").value Then
176                                             bRunUrl = True
                                            End If
                                        End If
                                    End If
    
178                                 If bRunUrl Then
180                                     If oshp.GetField(g_RSGISGridTableSettings.Fields.Item("URLLayerField").value) <> "" Then
182                                         ShellExecute Me.hWnd, vbNullString, oshp.GetField(g_RSGISGridTableSettings.Fields.Item("URLLayerField").value), vbNullString, "C:\", 1
                                        End If
                                    End If
                                End If
                        
                            End If
                    
                        Else

184                         Set oshp = GIS10.Locate(ptg, 5 / GIS10.zoom) ' 5 pixels precision
            
186                         If Not oshp Is Nothing Then m_frmAttributes.caption = "Active layer: " & oshp.layer.caption
                    
188                         m_frmAttributes.ShowShape oshp

                        End If
                    End If
                    
                End If
            End If
        
        End If
        
        'sMGRS = " MGRS:" & Map1.MilitaryGridReferenceFromPoint(ptg.x, ptg.y) & " Zoom:" & Round(GIS10.zoom, 2) & " Scale:" & GIS10.ScaleAsText
190     sMGRS = " MGRS:" & Map1.MilitaryGridReferenceFromPoint(ptg.x, ptg.y) & " Scale:" & GIS10.ScaleAsText
        
192     AB.Bands.Item("bStatus").Tools.Item("lblCoords").Text = "X:" & Round(ptg.x, 8) & " Y:" & Round(ptg.y, 8) & sMGRS
194     AB.Bands.Item("bStatus").Tools.Item("lblCoords").caption = "X:" & Round(ptg.x, 8) & " Y:" & Round(ptg.y, 8) & sMGRS
        
196     If m_oToolTipSHP Is Nothing Then
198         ptTip.x = 0
200         ptTip.y = 0
202         picToolTip.Visible = False
        Else
204         GetCursorPos ptTip
            'StatusBar.SimpleText = m_oToolTipSHP.GetField("name")
        End If

206     If g_CurrentTool = oRadiusSelect Or g_CurrentTool = oCircleSelect Then
208         If Not (Shift = XssLeft) Then Exit Sub

210         SetROP2 GIS10.hdc, R2_XORPEN

212         If oldRadius <> 0 Then
214             Ellipse GIS10.hdc, oldPos.x - oldRadius, oldPos.y - oldRadius, oldPos.x + oldRadius, oldPos.y + oldRadius
            End If

216         oldRadius = Round(Sqr(((oldPos.x - x) * (oldPos.x - x)) + ((oldPos.y - y) * (oldPos.y - y))))

218         Ellipse GIS10.hdc, oldPos.x - oldRadius, oldPos.y - oldRadius, oldPos.x + oldRadius, oldPos.y + oldRadius

        End If

220     If g_CurrentTool = oLineSelect Then
        
        End If
        
222     If g_CurrentTool = oPolyLineSelect Then
        
        End If
        
224     If g_CurrentTool = oPolySelect Then
        
        End If
        
226     If g_CurrentTool = oRectSelect Then
        
        End If

228     If g_CurrentTool = oSingleSelect Then
        
        End If
        
230     If g_CurrentTool = oMultiSelect Then
        
        End If
    
        On Error Resume Next
232     SSC.Run "OASISGis_MouseMove", sMGRS, ptg.x, ptg.y
    
234     translated = True
        '<EhFooter>
        Exit Sub

GIS10_OnMouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS10_OnMouseMove " & _
               "at line " & Erl
        'Resume Next
    '</EhFooter>
End Sub

Private Sub DumpAllGISProperties()
        Dim oLayer As TatukGIS_XDK10.XGIS_LayerAbstract
        Dim i As Integer

        On Error Resume Next

        With GIS10
        
            DebugPrint "gis properties"
            DebugPrint .Align
            DebugPrint .AutoCenter
        
            With .BigExtent
                DebugPrint "bIG eXTENT"
                DebugPrint .xmax
                DebugPrint .xmin
                DebugPrint .ymax
                DebugPrint .ymin
        
            End With

            DebugPrint .BigExtentMargin
            DebugPrint .BorderStyle
            DebugPrint .BusyText
            DebugPrint .CachedPaint
            DebugPrint .CausesValidation
        
            With .Center
                DebugPrint "Center"
                DebugPrint .x
                DebugPrint .y
        
            End With
        
            With .CenterPtg
                DebugPrint .x
                DebugPrint .y
        
            End With
        
            DebugPrint "Charset"
            DebugPrint .Charset
            DebugPrint .ChartLabelCache
            DebugPrint .CodePage
            DebugPrint .color
            DebugPrint .Ctl3D
            DebugPrint .Cursor
            
            DebugPrint "DrawMode"
            DebugPrint .DrawMode
            DebugPrint .FullPaint
            DebugPrint .IncrementalPaint
            DebugPrint .InDirectPaint
            DebugPrint .MinZoomSize
            DebugPrint .OutCodePage
            DebugPrint .ProjectFile
        
        End With
        
        DebugPrint "Starting With ALyers"
        
        For Each oLayer In GIS10.items
            
100         With oLayer

                DebugPrint "|||||||||||||||||||||||||||"
                DebugPrint .Name
                DebugPrint .Active
                DebugPrint .Age
                DebugPrint .CachedPaint
                DebugPrint .caption
                DebugPrint .Charset
                DebugPrint .CodePage
                DebugPrint .Collapsed
                DebugPrint .Comments
                
                DebugPrint "CONFIG NAME"
                DebugPrint .ConfigFile
                DebugPrint .ConfigName
                DebugPrint .DirectMode
                DebugPrint .DormantMode
                DebugPrint .FileInfo
                DebugPrint .HideFromLegend
                DebugPrint .IncrementalPaint
                DebugPrint .IsLocked
                DebugPrint .IsOpened
                
                DebugPrint "IS OPEN"
                DebugPrint .LabelsOnTop
                DebugPrint .OutCodePage
                DebugPrint .Path
                DebugPrint .StoreParamsInProject
                DebugPrint .SubType
                DebugPrint .Transparency
                DebugPrint .UseConfig
                DebugPrint .UseFileParams
                DebugPrint .ZOrder
                DebugPrint .ZOrderEx
                
                
                
102             With .Params
                    DebugPrint "Strating With Params"
                    DebugPrint .IsAssigned
                    DebugPrint .MaxScale
                    DebugPrint .MaxZoom
                    DebugPrint .MinScale
                    DebugPrint .MinZoom
                    
                    DebugPrint .Serial
                    DebugPrint .Style
                    DebugPrint .Visible

104                 With .area
                        DebugPrint "AREA PARAMS"
106                     DebugPrint .color
                        DebugPrint .Symbol.FontName
                        DebugPrint .SymbolSize
                        DebugPrint .SymbolGap
                        DebugPrint .SymbolRotate
                        DebugPrint .BITMAP
                        DebugPrint .Pattern
                        DebugPrint .OutlineColor
                        DebugPrint .OutlineWidth
                        DebugPrint .OutlineStyle
                        DebugPrint .OutlineSymbol
                        DebugPrint .OutlineSymbolGap
                        DebugPrint .OutlineSymbolRotate
                        DebugPrint .OutlineBitmap
                        DebugPrint .OutlinePattern
                        DebugPrint .SmartSize
                        DebugPrint .SmartSizeField
            
                    End With
           
110                 With .Marker
                        DebugPrint "MARKER PARAMS"
112                     DebugPrint .color
                        DebugPrint .Marker.Symbol
                        DebugPrint .SymbolSize
                        DebugPrint .SymbolGap
                        DebugPrint .SymbolRotate
                        DebugPrint .BITMAP
                        DebugPrint .Pattern
                        DebugPrint .OutlineColor
                        DebugPrint .OutlineWidth
                        DebugPrint .OutlineStyle
                        DebugPrint .OutlineSymbol
                        DebugPrint .OutlineSymbolGap
                        DebugPrint .OutlineSymbolRotate
                        DebugPrint .OutlineBitmap
                        DebugPrint .OutlinePattern
                        DebugPrint .SmartSize
                        DebugPrint .SmartSizeField
                    End With
                    
114                 With .Line
                        DebugPrint "LINE:nnnn"
116                     DebugPrint .color
                        DebugPrint .Marker.Symbol
                
                        DebugPrint .SymbolSize
                        DebugPrint .SymbolGap
                        DebugPrint .SymbolRotate
                        DebugPrint .BITMAP
                        DebugPrint .Pattern
                        DebugPrint .OutlineColor
                        'DebugPrint  .OutlineWidth
                        'DebugPrint  .OutlineStyle
                        'DebugPrint  .OutlineSymbol
                        'DebugPrint  .OutlineSymbolGap
                        'DebugPrint  .OutlineSymbolRotate
                        'DebugPrint  .OutlineBitmap
                        'DebugPrint  .OutlinePattern
                        'DebugPrint  .SmartSize
                        'DebugPrint  .SmartSizeField
                    End With
           
                End With

            End With

        Next

End Sub

Private Sub GIS10_OnMouseUp(translated As Boolean, _
                            ByVal Button As TatukGIS_XDK10.XMouseButton, _
                            ByVal Shift As TatukGIS_XDK10.XShiftState, _
                            ByVal x As Long, _
                            ByVal y As Long)
        '<EhHeader>
        On Error GoTo GIS10_OnMouseUp_Err
        '</EhHeader>
                          
        Dim tpl As TatukGIS_XDK10.IXGIS_Topology
        
        Dim tmp As TatukGIS_XDK10.IXGIS_Shape
        Dim otmp As TatukGIS_XDK10.IXGIS_Shape
        Dim buf1 As TatukGIS_XDK10.IXGIS_Shape
        Dim buf2 As TatukGIS_XDK10.IXGIS_Shape
        Dim ptg As TatukGIS_XDK10.IXGIS_Point
        Dim ptg1 As TatukGIS_XDK10.IXGIS_Point
        Dim distance As Double
        Dim sVals As String
        Dim colPlaces As New Collection
        
100     DebugPrint "----- **MouseUp Start** ----"
102     DebugPrint "gdpts:" & UBound(gdpts)
104     DebugPrint "ptsSel:" & UBound(ptsSel)
        
106     If g_CurrentTool = oRadiusSelect Or g_CurrentTool = oCircleSelect Then

108         If oldRadius = 0 Then
                Exit Sub
            End If
    
110         Set ptg = GIS10.ScreenToMap(oldPos)
112         m_oDrawLyr.Lock
114         Set tmp = m_oDrawLyr.CreateShape(XgisShapeTypePoint)
116         tmp.Params.Marker.Size = 0
118         tmp.Lock TatukGIS_XDK10.XgisLockExtent
120         tmp.AddPart
122         tmp.AddPoint ptg
124         tmp.Unlock
126         GIS10.get("Buffers").RevertAll
    
128         m_oDrawLyr.Unlock
130
            'distance recalc
132         Set ptg1 = GIS10.ScreenToMap(GisUtils.POINT(oldPos.x + oldRadius, y))
134         distance = ptg1.x - ptg.x
    
            AddShapeWithBufferAndHighlightIntersect tmp, distance

178         m_oDrawLyr.RevertAll
180         'GIS10.UpDate
182         translated = True
        End If

184     'miniMapRefresh

186     DebugPrint "----- **MouseUP End** ----"
188     DebugPrint "gdpts:" & UBound(gdpts)
190     DebugPrint "ptsSel:" & UBound(ptsSel)

        '<EhFooter>
        Exit Sub

GIS10_OnMouseUp_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GIS_OnMouseUp " & "at line " & Erl
        '</EhFooter>
End Sub

Private Sub AddShapeWithBufferAndHighlightIntersect(oShapeNew As TatukGIS_XDK10.IXGIS_Shape, _
                                                    dBuffer As Double)
        '<EhHeader>
        On Error GoTo AddShapeWithBufferAndHighlightIntersect_Err
        '</EhHeader>

100     If Len(ctlSelector1.GetActiveLayer) > 0 Then

            Dim lL As TatukGIS_XDK10.IXGIS_LayerVector
            Dim oLayerExport As New XGIS_LayerVector
            Dim oRS As New ADODB.Recordset
            Dim oPoint As New XGIS_Point
            Dim tpl As New TatukGIS_XDK10.XGIS_Topology
            Dim buf1 As TatukGIS_XDK10.IXGIS_Shape
            Dim buf2 As TatukGIS_XDK10.IXGIS_Shape
            Dim tmp As TatukGIS_XDK10.IXGIS_Shape
            Dim i As Long
            Dim j As Long
            Dim lCountOfItems As Long
            Dim sFields() As String
            Dim lngFldType As Long
            Dim lngDataLength As Long
            
            ctlSelector1.ToggleGridVisible False

            DoEvents
102         Set lL = GIS10.get(ctlSelector1.GetActiveLayer)
104         lL.DeselectAll

106         oRS.Fields.Append "GIS_UID", adBigInt
        
            sFields = Split(ctlSelector1.GetFields, ";")
            i = 0
            j = 0

            Do Until i = UBound(sFields) Or UBound(sFields) = -1

                Do Until j = lL.Fields.Count

                    If lL.Fields.Item(j).Name = sFields(i + 1) Then
                
308                     Select Case lL.Fields.Item(j).FieldType
        
                            Case Is = TatukGIS_XDK10.XgisFieldTypeString '= 0,
310                             lngFldType = adVarWChar  'xftString
312                             lngDataLength = 255 'lL.Fields.Item(i).Width
        
314                         Case Is = TatukGIS_XDK10.XgisFieldTypeNumber ' = 1,
316                             lngFldType = adDouble 'xftFloat
318                             lngDataLength = lL.Fields.Item(i).Width
        
320                         Case Is = TatukGIS_XDK10.XgisFieldTypeFloat '= 2,
322                             lngFldType = adDouble 'xftFloat
324                             lngDataLength = lL.Fields.Item(i).Width
        
326                         Case Is = TatukGIS_XDK10.XgisFieldTypeBoolean '= 3,
328                             lngFldType = adBoolean 'xftBoolean
330                             lngDataLength = lL.Fields.Item(i).Width
        
332                         Case Is = TatukGIS_XDK10.XgisFieldTypeDate '= 4
334                             lngFldType = adDate 'xftDate
336                             lngDataLength = lL.Fields.Item(i).Width
                        End Select
                        
                        oRS.Fields.Append sFields(i + 1), lngFldType, lngDataLength
                        Exit Do
                
                    End If

                    j = j + 1
                
                Loop
                
                'oRS.Fields.Append sFields(i + 1), adVarChar, 150
                i = i + 1
            Loop

112         If ctlSelector1.DistanceEnabled Then
114             oRS.Fields.Append "Distance_From_Centre", adDouble
            End If
                
116         oRS.Open

118         If ctlSelector1.DistanceEnabled Then
                oRS.Sort = "Distance_From_Centre"
            ElseIf oRS.Fields.Count > 1 Then
                oRS.Sort = oRS.Fields(1).Name
            End If

120         If dBuffer > 0 Then
122             Set buf1 = tpl.MakeBuffer(oShapeNew, dBuffer, 360, True)
124             Set buf2 = GIS10.get("Buffers").AddShape(buf1)
            Else
126             Set buf2 = oShapeNew
            End If

128         Set buf1 = Nothing
130         Set tpl = Nothing
132         Set oPoint = oShapeNew.Centroid
           
            lL.Lock
            lL.ExportLayer oLayerExport, buf2.Extent, XgisShapeTypeUnknown, "", False
            lL.Unlock
            lCountOfItems = oLayerExport.items.Count
            Set oLayerExport = Nothing
            ctlSelector1.InitProgressBar lCountOfItems
                
134         For Each tmp In lL.Loop(buf2.Extent, "", buf2, RELATE_INTERSECT, True)
                
                ctlSelector1.ProgressStep
136             Set tmp = tmp.MakeEditable
138             tmp.IsSelected = True
        
140             oRS.AddNew
142             oRS.Fields("GIS_UID").value = tmp.uID
                i = 0
           ' On Error Resume Next
                Do Until i = UBound(sFields) Or UBound(sFields) = -1
                    
                    oRS.Fields(i + 1).value = IIf(IsNull(tmp.GetField(sFields(i + 1))), "", tmp.GetField(sFields(i + 1)))
                    i = i + 1
                Loop
                    On Error GoTo AddShapeWithBufferAndHighlightIntersect_Err
150             If ctlSelector1.DistanceEnabled Then
152                 oRS.Fields("Distance_From_Centre").value = Round(DistanceBetween2Points(oPoint, tmp), 2)
                End If
                
154             Set tmp = Nothing

            Next
                
156         EndEdit ctlSelector1.GetTool
158         GIS10.UpDate
160         ctlSelector1.SetRS oRS
            
            Set GISGridRS = oRS
            GISGridLayerName = ctlSelector1.GetActiveLayer
            Set dxGISDataGrid.DataSource = oRS
            dxGISDataGrid.KeyField = "GIS_UID"
            dxGISDataGrid.Columns.DestroyColumns
            dxGISDataGrid.Columns.RetrieveFields
            abGridTools.Tools.Item("comLyr").Text = GISGridLayerName
            dxGISDataGrid.Columns(0).Visible = False
        Else
162         MsgBox "Layer not read"
        End If

        ctlSelector1.ToggleGridVisible True
        '<EhFooter>
        Exit Sub

AddShapeWithBufferAndHighlightIntersect_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.AddShapeWithBufferAndHighlightIntersect " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmbSnap_Click()
        '<EhHeader>
        On Error GoTo cmbSnap_Click_Err
        '</EhHeader>
100   If cmbSnap.ListIndex > 0 Then
102     GIS10.Editor.SnapLayer = GIS10.get(cmbSnap.List(cmbSnap.ListIndex))
      Else
104     GIS10.Editor.SnapLayer = Nothing
      End If
  
        '<EhFooter>
        Exit Sub

cmbSnap_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmbSnap_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS10_OnVisibleExtentChange(translated As Boolean)
        '<EhHeader>
        On Error GoTo GIS_OnVisibleExtentChange_Err
        '</EhHeader>
100     Set g_PrevExt = GIS10.VisibleExtent
    
        On Error Resume Next
        Unload frmSplash
    

    
102     translated = True
        '<EhFooter>
        Exit Sub

GIS_OnVisibleExtentChange_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS_OnVisibleExtentChange " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS_OnZoomChange(translated As Boolean)
        '<EhHeader>
        On Error GoTo GIS_OnZoomChange_Err
        '</EhHeader>
        
        'Not sure why but for loading of some maps the next line screwed up the GIS control
        'only a restart would allow some map def fines to load the legend:
        
102     translated = True
104     miniMapRefresh
106     Set g_PrevExt = GIS10.VisibleExtent
        '<EhFooter>
        Exit Sub

GIS_OnZoomChange_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GIS_OnZoomChange " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_fmrAddIncident_CancelIncident()
        ' g_CurrentTool = m_prevTool
        ' SetOASISTool
        '<EhHeader>
        On Error GoTo m_fmrAddIncident_CancelIncident_Err
        '</EhHeader>
100     g_CurrentTool = oZoom
102     GIS10.Mode = TatukGIS_XDK10.XgisZoomEx

        '    m_oIncShpPt.MakeEditable
        '    m_oIncShpPt.Delete
    
104     If Not m_oIncShpPt Is Nothing Then
106         m_oSQLIncLyr.Delete m_oIncShpPt.uID
        End If

108     Set m_oIncShpPt = Nothing
    
110     m_oSQLIncLyr.SaveData
    
112     GIS10.UpDate
    
114     m_fmrAddIncident.Hide
    
        '<EhFooter>
        Exit Sub

m_fmrAddIncident_CancelIncident_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.m_fmrAddIncident_CancelIncident " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_fmrAddIncident_CheckCoordinates(sAdmVal1 As String, _
                                              sAdmVal2 As String, _
                                              sAdmloc As String)
        '<EhHeader>
        On Error GoTo m_fmrAddIncident_CheckCoordinates_Err
        '</EhHeader>
        Dim ptg As New TatukGIS_XDK10.XGIS_Point
        Dim x As Double
        Dim y As Double

100     If Len(m_fmrAddIncident.txtX.Text) > 0 And Len(m_fmrAddIncident.txtY.Text) > 0 Then
102         ptg.Prepare CDbl(m_fmrAddIncident.txtX.Text), CDbl(m_fmrAddIncident.txtY.Text)
        Else

104         If Len(m_fmrAddIncident.txtMGRS.Text) > 0 Then
106             ConvMGRStoPT m_fmrAddIncident.txtMGRS.Text, x, y
108             ptg.Prepare x, y
            Else
                Exit Sub
            End If
        End If

110     GetAdmCode ptg, sAdmVal1, sAdmVal2, sAdmloc
112     m_fmrAddIncident.SetAdmValues sAdmVal1, sAdmVal2, sAdmloc
     
114     GIS10.CenterViewport ptg
        '<EhFooter>
        Exit Sub

m_fmrAddIncident_CheckCoordinates_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_fmrAddIncident_CheckCoordinates " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_fmrAddIncident_GetCoordinatesFromMap(sAdmVal1 As String, _
                                                   sAdmVal2 As String, _
                                                   sAdmloc As String)
        '<EhHeader>
        On Error GoTo m_fmrAddIncident_GetCoordinatesFromMap_Err
        '</EhHeader>
        
        m_bUseDistrictOnly = IIf(m_fmrAddIncident.chkUseDistrict.value = vbChecked, True, False)
        
100     m_prevTool = g_CurrentTool
102     g_CurrentTool = oCreateLocationPoint
104     g_CurrentFeatureType = Incident


        '<EhFooter>
        Exit Sub

m_fmrAddIncident_GetCoordinatesFromMap_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_fmrAddIncident_GetCoordinatesFromMap " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function DoesFieldExist(RS As Recordset, sFldName As String) As Boolean
        '<EhHeader>
        On Error GoTo DoesFieldExist_Err
        '</EhHeader>
        Dim s As String
    
100     s = RS.Fields.Item(sFldName).Name
102     DoesFieldExist = True
        '<EhFooter>
        Exit Function

DoesFieldExist_Err:
        Err.Clear
        '</EhFooter>
End Function

Private Sub m_frmAddOns_MenuPressed(i As Integer)
        '<EhHeader>
        On Error GoTo m_frmAddOns_MenuPressed_Err
        '</EhHeader>
100     Select Case i
    '
    '        Case 0
    '            WebBrowser2.Navigate2 "http://www.unhcr.org"
    '            ShellExecute Me.hwnd, vbNullString, "C:\Program Files\HIS\HIS.mdb", vbNullString, "C:\", 1
    '
    '        Case 1 'NCCI
    '            WebBrowser2.Navigate2 "http://www.mop-iraq.org/dad/"  '"file:///C:/My%20Web%20Sites/NCCI/www.ncciraq.org/index.html"
    '
    '        Case 2 'UNAMI
    '            WebBrowser2.Navigate2 "file:///C:/My%20Web%20Sites/UNAMI/www.uniraq.org/index.html"
    '        Case Else
    '            WebBrowser2.Navigate2 "file:///C:/OASIS/AddOns/addons.html"
    '
                '
        End Select
        '<EhFooter>
        Exit Sub

m_frmAddOns_MenuPressed_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmAddOns_MenuPressed " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Public Sub SetOASISTool()
        '<EhHeader>
        On Error GoTo SetOASISTool_Err
        '</EhHeader>

100     Select Case g_CurrentTool
    
            Case oCreateLocationArea
102             GIS10.Mode = TatukGIS_XDK10.XgisEdit



104         Case oCreateLocationPoint
106             GIS10.Mode = TatukGIS_XDK10.XgisEdit

108         Case oCreateLocationPolyline
110             GIS10.Mode = TatukGIS_XDK10.XgisEdit

112         Case OASIS_TOOLS.oZoomEx
114             GIS10.Mode = TatukGIS_XDK10.XgisZoomEx

116         Case OASIS_TOOLS.oZoom
118             GIS10.Mode = TatukGIS_XDK10.XgisZoom

120         Case OASIS_TOOLS.oSingleSelect
122             GIS10.Mode = TatukGIS_XDK10.XgisSelect

124         Case OASIS_TOOLS.oPan
126             GIS10.Mode = TatukGIS_XDK10.XgisDrag
    
        End Select
    
        '<EhFooter>
        Exit Sub

SetOASISTool_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SetOASISTool " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub m_frmMnuOASISProfile_OASISIntranetClicked()
        '<EhHeader>
        On Error GoTo m_frmMnuOASISProfile_OASISIntranetClicked_Err
        '</EhHeader>
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'IntranetURL'"
        
        If Not g_RSAppSettings.Bof And Not g_RSAppSettings.EOF Then
            If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").value) Then
104             WebBrowser1.Navigate2 g_RSAppSettings.Fields.Item("SettingValue1").value
            Else
                MsgBox "This functionality has not been configured by your OASIS Administrator", vbInformation, "OASIS Configuration message"
            End If
        End If
        
        '<EhFooter>
        Exit Sub

m_frmMnuOASISProfile_OASISIntranetClicked_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOASISProfile_OASISIntranetClicked " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOASISProfile_OASISProfileClicked()
        '<EhHeader>
        On Error GoTo m_frmMnuOASISProfile_OASISProfileClicked_Err
        '</EhHeader>
    
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'InitURL'"
104     WebBrowser1.Navigate2 g_RSAppSettings.Fields.Item("SettingValue1").value

        '<EhFooter>
        Exit Sub

m_frmMnuOASISProfile_OASISProfileClicked_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOASISProfile_OASISProfileClicked " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOASISProfile_OASISSettings()
        '<EhHeader>
        On Error GoTo m_frmMnuOASISProfile_OASISSettings_Err
        '</EhHeader>
100     MsgBox "You don't have enough Administrative Privilages." & vbCrLf & "Please check your administrative rights or contact OASIS support.", vbInformation, "OASIS Client Support"
        '<EhFooter>
        Exit Sub

m_frmMnuOASISProfile_OASISSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOASISProfile_OASISSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOASISProfile_OASISSupportCentre()
        '<EhHeader>
        On Error GoTo m_frmMnuOASISProfile_OASISSupportCentre_Err
        '</EhHeader>
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'OASISSupportCenterURL'"
    
104     If g_RSAppSettings.Fields.Item("SettingValue1").value <> "" Then
106         WebBrowser1.Navigate2 g_RSAppSettings.Fields.Item("SettingValue1").value
        Else
108         MsgBox "It seems like your profile has not access to the OASIS support centre." & vbCrLf & "Please check contact your OASIS administrator for support.", vbInformation, "OASIS Client Support"
        End If

        '<EhFooter>
        Exit Sub

m_frmMnuOASISProfile_OASISSupportCentre_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOASISProfile_OASISSupportCentre " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_ActivateSecurityIncidents(bActivated As Boolean)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_ActivateSecurityIncidents_Err
        '</EhHeader>

100     If m_oSQLIncLyr Is Nothing Then
            Exit Sub
        End If

102     m_oSQLIncLyr.Active = bActivated
        'If bActivated Then
104         GIS10.UpDate
        'End If
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_ActivateSecurityIncidents_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_ActivateSecurityIncidents " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_CategorizeIncidents()
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_CategorizeIncidents_Err
        '</EhHeader>
100     CategorizeIncidentsByType '2
102     GIS10.UpDate
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_CategorizeIncidents_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_CategorizeIncidents " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_createChart()
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_createChart_Err
        '</EhHeader>
100     If Not frmChartProviderSettings.Visible Then frmChartProviderSettings.Show vbModeless, Me
    
    '    If Not frmSecChart.Visible Then frmSecChart.Show vbModeless, Me
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_createChart_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_createChart " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_DoSecurityAnalysis()
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_DoSecurityAnalysis_Err
        '</EhHeader>
    On Error Resume Next
    
    '    frmSecurityCharts.Show
    
    '    Exit Sub
    
100     If Not m_sSecAnalysisLyrName = "" Then
102         GIS10.get(m_sSecAnalysisLyrName).Params.Visible = False
        End If
    
104     SafeMoveFirst g_RSAppSettings
    
106     Select Case m_frmMnuOperations.SecAnalysisevel
    
            Case 0
108             g_RSAppSettings.Find "SettingName = 'AdmProvSec'"
            
110             m_sSecAnalysisFieldName = g_RSAppSettings.Fields.Item("SettingValue3").value
112             m_sSecAnalysisLyrName = g_RSAppSettings.Fields.Item("SettingValue1").value

114         Case 1
116             g_RSAppSettings.Find "SettingName = 'AdmDistSec'"
            
118             m_sSecAnalysisFieldName = g_RSAppSettings.Fields.Item("SettingValue3").value
120             m_sSecAnalysisLyrName = g_RSAppSettings.Fields.Item("SettingValue1").value
        
122         Case Else
124             g_RSAppSettings.Find "SettingName = 'SecurityGridZoomLevels'"
            
126             Select Case CLng(Mid(GIS10.ScaleAsText, 3))
            
                    Case Is > CLng(g_RSAppSettings.Fields.Item("SettingValue1").value)
128                     SafeMoveFirst g_RSAppSettings
130                     g_RSAppSettings.Find "SettingName = 'SecGrid1'"
132                 Case Is > CLng(g_RSAppSettings.Fields.Item("SettingValue2").value)
134                     SafeMoveFirst g_RSAppSettings
136                     g_RSAppSettings.Find "SettingName = 'SecGrid2'"
138                 Case Else
140                     SafeMoveFirst g_RSAppSettings
142                     g_RSAppSettings.Find "SettingName = 'SecGrid3'"
                End Select
            
144             m_sSecAnalysisFieldName = g_RSAppSettings.Fields.Item("SettingValue3").value
146             m_sSecAnalysisLyrName = g_RSAppSettings.Fields.Item("SettingValue1").value
    
        End Select

148     If m_sSecAnalysisLyrName = "" Then Exit Sub

150     GIS10.get(m_sSecAnalysisLyrName).Params.Visible = True

152     Command1_Click
154     'GIS10.get(m_sSecAnalysisLyrName).draw
156     'GIS10.viewer.PrintClipboard
        GIS10.UpDate
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_DoSecurityAnalysis_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_DoSecurityAnalysis " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_InsertIncident()
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_InsertIncident_Err
        '</EhHeader>
        
        'Dim sCnn As String
     'sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\DynamicData\IncidentsV2.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
        If g_bIncidentsV2 Then
        
            If Not m_frmIncidentsV2ControlPanel.Visible Then
                'm_frmIncidentsV2ControlPanel.Show vbModal, Me
                m_frmIncidentsV2DataEntry.Show vbModeless, Me
            End If
        
        Else
        
100         If Not m_fmrAddIncident.Visible Then

102         m_fmrAddIncident.Init m_Cnn, "Enter Your Name"
104         m_fmrAddIncident.Show vbModeless, Me

            End If
        End If
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_InsertIncident_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_InsertIncident " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_LayerActivatedStatus(sLayerName As String, bActivated As Boolean)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_LayerActivatedStatus_Err
        '</EhHeader>
        On Error Resume Next
100     GIS10.get(sLayerName).Active = bActivated
102     GIS10.UpDate
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_LayerActivatedStatus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_LayerActivatedStatus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_LoadOASISIncidents()
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_LoadOASISIncidents_Err
        '</EhHeader>
        On Error Resume Next
100     LoadOASISIncidentsEX
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_LoadOASISIncidents_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_LoadOASISIncidents " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_OpenLISAnalyzis()
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_OpenLISAnalyzis_Err
        '</EhHeader>
        On Error Resume Next
100     MineActionScoring1.Init m_Cnn, Me.hWnd
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_OpenLISAnalyzis_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_OpenLISAnalyzis " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_ZoomToLoc(sName As String, _
                                         sID As String)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_ZoomToLoc_Err
        '</EhHeader>
        Dim lL As TatukGIS_XDK10.XGIS_LayerVector
        Dim shp As TatukGIS_XDK10.XGIS_Shape
'GIS109 Done
        On Error Resume Next
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'MAPcodeLayer'"
        
104     Set lL = GIS10.get(g_RSAppSettings.Fields.Item("SettingValue1").value)
            
        For Each shp In lL.Loop(GIS10.Extent, g_RSAppSettings.Fields.Item("SettingValue5").value & " = '" & sID & "'", Nothing, "", True)

108         If Not shp Is Nothing Then
110             GIS10.VisibleExtent = shp.Extent
112             shp.MakeEditable
114             shp.Flash
            End If

        Next
    
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_ZoomToLoc_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.m_frmMnuOperations_ZoomToLoc " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmOvMap_MapOnMouseMove(translated As Boolean, ByVal Shift As TatukGIS_XDK10.XShiftState, ByVal x As Long, ByVal y As Long)
        '<EhHeader>
        On Error GoTo m_frmOvMap_MapOnMouseMove_Err
        '</EhHeader>
100     MouseUp x, y
        '<EhFooter>
        Exit Sub

m_frmOvMap_MapOnMouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmOvMap_MapOnMouseMove " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmOvMap_MapOnMouseUp(translated As Boolean, ByVal Button As TatukGIS_XDK10.XMouseButton, ByVal Shift As TatukGIS_XDK10.XShiftState, ByVal x As Long, ByVal y As Long)
        '<EhHeader>
        On Error GoTo m_frmOvMap_MapOnMouseUp_Err
        '</EhHeader>
100     MouseUp x, y
        '<EhFooter>
        Exit Sub

m_frmOvMap_MapOnMouseUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmOvMap_MapOnMouseUp " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

' Set bDebugMode to true. This happens only
' if the Debug.Assert call happens. It only
' happens in the IDE.
Private Function InDebugMode() As Boolean
        '<EhHeader>
        On Error GoTo InDebugMode_Err
        '</EhHeader>
100     bDebugMode = True
102     InDebugMode = True
        '<EhFooter>
        Exit Function

InDebugMode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.InDebugMode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub Init(Optional sInitialMapName As String)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        Dim oINIReader As New clIniReader
        Dim fs As New FileSystemObject
        
100     Debug.Assert InDebugMode
                    
        If fs.FileExists(g_sAppPath & "\Incidents.ini") Then m_sIncidentIni = g_sAppPath & "\Incidents"
                    
        Set fs = Nothing
                    
102     m_bLOADING = True
104     PrepareModuleMeny
106     SetMenuSystem
108     AB.RecalcLayout

112     Set m_fmrAddIncident = New frmAddIncident
114     Set m_frmOvMap = New frmOVMap
120     Set m_frmLocator = New frmLocator
122     Set m_frmOASISCharts = New frmOASISCharts
        Set m_frmTextAnnoSettings = New frmTextAnnoSettings
        Set m_frmAttributes = New frmAttributes
        Set m_frmSearch = New frmSearch
        Set m_frmUpdateSettings = New frmUpdateSettings
        Set m_frmMainSettings = New frmMainSettings
        Set m_frmSelectorReports = New frmSelectorReports
        Set m_frmSelectorSettings = New frmSelectorSettings
        Set m_frmSpatialiseDD = New frmSpatialiseDD
        Set m_frmIncidentsV2SitrepGenerator = New frmIncidentsV2SitrepGenerator
        Set m_frmIncidentsV2DataEntry = New frmIncidentsV2DataEntry
        Set m_frmIncidentsV2ControlPanel = New frmIncidentsV2ControlPanel
        Set m_frmMapLibraryDLG = New frmMapLibraryDLG
        'Set m_frmDynamicContent = New frmDynamicContent
        
124     SSC.AddObject "frmOIncident", m_fmrAddIncident, True
        
126     m_sInitialMapName = sInitialMapName

128     AddDockableInspectors
        
130     SetApplicationSettings 0
        
        'If pInitThread.IsThreadRunning Then
        '    pInitThread.ThreadPriority = THREAD_PRIORITY_NORMAL
        'End If
        
132     g_bUseSynch = CheckSynchDS

        '1^connectionString^RemoteTableprefix^ServerPath^HasEncrypt^sKey^HWND
        
        On Error Resume Next
        If g_bOnlineCheckedAtLogin Then
            If g_udtSynchUpdateOptions.AutoUpdate Then
142             ShellExecute Me.hWnd, vbNullString, g_sAppPath & "\OASIS_SynchNG_Client.exe", "CheckBackground", "C:\", 1
144             ShellExecute Me.hWnd, vbNullString, g_sAppPath & "\AUClient.exe", "CheckBackground", "C:\", 1
            End If
        End If
        
        'Call ShellExecute(Me.Hwnd, vbNullString, "C:\Program Files\Globalsat\TR Management Center\OASIS_Tracker.exe", "", "c:\", 2)
    
146     Set oINIReader = Nothing

        CheckOVBFiles
        OpenLocalAppSettings
        PrepareConvertorUTM_LL
        ApplyLocalSettings
        '<EhFooter>
        
        
                
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.Init " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub UpdateProjectFileToDB()
        '<EhHeader>
        On Error GoTo UpdateProjectFileToDB_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
        Dim oRSSQLLayer As New ADODB.Recordset
        Dim sMapProj As String
        Dim sText As String
        Dim oStream As ADODB.Stream
        Dim bInUse As String
        
        Dim oTTKLyr As New TatukGIS_XDK10.XGIS_LayerAbstract
        
        Dim sINI As String
        Dim oIni As New TatukGIS_XDK10.XGIS_Ini
        Dim oSList As New TatukGIS_XDK10.XStringList
        Dim lDude As Long
                    
100     GIS10.UpDate

102     DoEvents
             
104     With oRSSQLLayer
            
106         .Open "SELECT * from [ttkGISLayerSQLInProject] ORDER BY [Sequence] DESC", m_Cnn, adOpenDynamic, adLockBatchOptimistic
                
108         Do Until .EOF

110             Set oTTKLyr = GIS10.get(.Fields("LayerCaption").value)
                    
112             If Not oTTKLyr Is Nothing Then

116                 Set oSList = New TatukGIS_XDK10.XStringList
118                 oTTKLyr.ParamsList.SaveToStrings oSList
120                 'DebugPrint "xxxxxxxxxxxxxxxxxxxxxxxxxxxxx" & vbCrLf & "Layer : " & oTTKLyr.Name & vbCrLf & oSList.Text & vbCrLf & "Active  =" & oTTKLyr.Active
122                 .Fields("INISettings").value = oSList.Text
                    .Fields("INISettings").value = Replace$(.Fields("INISettings").value, g_sAppPath, "CLIENTDBPATH")
124                 .Fields("Transparency").value = oTTKLyr.Transparency
126                 .Fields("IsVisible").value = oTTKLyr.Active
128                 .Fields("Sequence").value = oTTKLyr.ZOrderEx
130                 .Fields("IsExpanded").value = Not oTTKLyr.Collapsed

132                 If oTTKLyr.FileInfo = "ArcView Shape Files (SHP)" Then
134                     .Fields("Dialect").value = "SHP"
136                     .Fields("ADO").value = Replace(LCase(oTTKLyr.Path), LCase(g_sAppPath), "CLIENTDBPATH")
                    End If
                    
138                 If InStr(oTTKLyr.FileInfo, "OpenGIS® Web Map Service") > 0 Then
140                     .Fields("Dialect").value = "WMS"
142                     .Fields("ADO").value = oTTKLyr.Path
                    End If
                    
144                 If oTTKLyr.FileInfo = "TatukGIS SQL Vector Coverage (TTKLS)" Then
146                     .Fields("ADO").value = Replace(LCase(oTTKLyr.SQLParameter("ADO")), LCase(g_sAppPath), "CLIENTDBPATH")
148                     .Fields("Dialect").value = oTTKLyr.SQLParameter("DIALECT")
                        If Not bSQLServerInUse Then oTTKLyr.Path = "[TatukGIS Layer]\\nStorage = Native\\nADO=" & .Fields("ADO").value
                        If Not bSQLServerInUse Then oTTKLyr.Path = oTTKLyr.Path & "Persist Security Info=False\\nDialect=MSJET\\nLayer=" & .Fields("LayerName").value & "\\n.ttkls"
                    End If

                    'If .Fields("Sequence").Value < iLowestZOrder Then iLowestZOrder = .Fields("Sequence").Value
150                 .UpdateBatch adAffectCurrent
152                 Set oTTKLyr = Nothing

                Else
154                 .Delete adAffectCurrent
156                 .UpdateBatch
                End If

158             .MoveNext
            Loop
                
160         .Close
162         lDude = 0
164         .Open "SELECT * from [ttkGISLayerSQLInProject] ORDER BY [Sequence]", m_Cnn, adOpenDynamic, adLockBatchOptimistic
                
166         Do Until .EOF
168             Set oTTKLyr = GIS10.get(.Fields("LayerCaption").value)
                
170             If oTTKLyr Is Nothing Then
172                 m_Cnn.Execute "DELETE FROM [ttkGISLayerSQLInProject] WHERE [LayerCaption] = '" & .Fields("LayerCaption").value & "'"
                Else
174                 .Fields("Sequence").value = lDude
                    GIS10.Delete .Fields("LayerCaption").value
178                 .UpdateBatch adAffectCurrent
180                 lDude = lDude + 1
                End If

182             .MoveNext
            Loop
                
184         .Close
            
        End With
                
186     Set oRSSQLLayer = Nothing
        
188     If UCase(GIS10.ProjectName) = UCase(g_sAppPath & "\data\user\Exports\tempproj.ttkgp") Then
                    
            On Error Resume Next
            
190         Kill g_sAppPath & "\data\user\Exports\tempproj.ttkgp"
            
            On Error GoTo UpdateProjectFileToDB_Err
                    
192         GIS10.viewer.SaveProjectAs g_sAppPath & "\data\user\Exports\tempproj.ttkgp" ', True
            
194         Set oStream = New ADODB.Stream
196         oStream.Open
198         oStream.Type = 2    ' type = text
200         oStream.Charset = "ascii"
202         oStream.LoadFromFile (g_sAppPath & "\data\user\Exports\tempproj.ttkgp")
204         sText = oStream.ReadText
206         oStream.Close
208         Set oStream = Nothing

        End If
                          
212     With oRS
                
214         .Open "SELECT * from [ttkGISProjectDef]", m_Cnn, adOpenDynamic, adLockBatchOptimistic

216         If Not .State = adStateClosed Then

220             If .EOF Then .AddNew

                'need to verify if this saves shapefiles to project file overlapping with SQL layers table which can support file SHP files
228             If Len(sText) > 10 Then
                    .Fields("MapData").value = sText
230                 .Fields("sGUID").value = GUIDGen
                End If

236             .Fields("XMin").value = GIS10.viewer.VisibleExtent.xmin
238             .Fields("XMax").value = GIS10.viewer.VisibleExtent.xmax
240             .Fields("YMin").value = GIS10.viewer.VisibleExtent.ymin
242             .Fields("YMax").value = GIS10.viewer.VisibleExtent.ymax

244             .UpdateBatch adAffectCurrent

            End If
                
246         .Close

        End With

248     Set oRS = Nothing
250     Set oRSSQLLayer = Nothing

        '<EhFooter>
        Exit Sub

UpdateProjectFileToDB_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.UpdateProjectFileToDB " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function AddSQLLayersToProject9()
        '<EhHeader>
        On Error GoTo AddSQLLayersToProject9_Err
        '</EhHeader>

        
100     GIS10.Mode = TatukGIS_XDK10.XgisZoomEx
        
        Dim oRS As New ADODB.Recordset
        Dim oRsDD As ADODB.Recordset
        Dim oConnectionDD As ADODB.Connection
  
        Dim oSQLLyr As TatukGIS_XDK10.XGIS_LayerSqlAdo
        Dim oWMSLyr As TatukGIS_XDK10.XGIS_LayerWMS
        Dim oSHPLyr As TatukGIS_XDK10.XGIS_LayerSHP
        
        Dim oTTKLyr As TatukGIS_XDK10.XGIS_LayerAbstract
        Dim oExtent As TatukGIS_XDK10.XGIS_Extent
        Dim oNC As New ADODB.Connection
        
        Dim sFilename As String
        Dim bLoadFailed As Boolean
        Dim sINI As String
        Dim oStream As ADODB.Stream
        Dim oIni As New TatukGIS_XDK10.XGIS_Ini
        Dim bAllLayersLoaded As Boolean
        Dim sName As String
        Dim oSList As TatukGIS_XDK10.XStringList
        
102     bAllLayersLoaded = False
        
        'oNC.Open m_Cnn.ConnectionString '- PETRI THIS SCREWS UP SQL SERVER
104     Set oNC = m_Cnn
            
106     With oRS
 
108         .Open "SELECT * from [ttkGISLayerSQLInProject] ORDER BY [Sequence] DESC", oNC
110         bAllLayersLoaded = True
            
112         Do Until .EOF

114             sName = .Fields("LayerCaption").value

                'DebugPrint "SNAME = " & sName
116             If .Fields("DIALECT").value = "SHP" Then

118                 Set oSHPLyr = New TatukGIS_XDK10.XGIS_LayerSHP
120                 oSHPLyr.Path = Replace(.Fields("ADO").value, "CLIENTDBPATH", g_sAppPath)

122                 If FileExists(oSHPLyr.Path) Then
124                     oSHPLyr.Name = .Fields("LayerName").value
126                     oSHPLyr.caption = .Fields("LayerCaption").value
128                     oSHPLyr.Open
130                     Set oTTKLyr = oSHPLyr
                    Else
132                     bLoadFailed = True
134                     bAllLayersLoaded = False
                    End If

136             ElseIf .Fields("DIALECT").value = "WMS" Then

138                 Set oWMSLyr = New TatukGIS_XDK10.XGIS_LayerWMS
140                 oWMSLyr.Path = .Fields("ADO").value
142                 oWMSLyr.Name = .Fields("LayerName").value
144                 oWMSLyr.caption = .Fields("LayerCaption").value
146                 oWMSLyr.Open
148                 Set oTTKLyr = oWMSLyr

                Else
                
150                 Set oRsDD = New ADODB.Recordset
152                 Set oConnectionDD = New ADODB.Connection

154                 If bSQLServerInUse Then
156                     Set oConnectionDD = m_Cnn
                    Else
158                     oConnectionDD.Open IIf(bSQLServerInUse, g_sGlobalConnectionString, Replace(.Fields("ADO").value, "CLIENTDBPATH", g_sAppPath))
                    End If

160                 oRsDD.Open "SELECT * FROM [ttkGISLayerSQL] WHERE [name] = '" & .Fields("LayerName").value & "'", oConnectionDD, adOpenDynamic, adLockBatchOptimistic

162                 If oRsDD.EOF Then oRsDD.AddNew
164                 oRsDD.Fields("name").value = .Fields("LayerName").value
166                 oRsDD.Fields("xmin").value = .Fields("xmin").value
168                 oRsDD.Fields("xmax").value = .Fields("xmax").value
170                 oRsDD.Fields("ymin").value = .Fields("ymin").value
172                 oRsDD.Fields("ymax").value = .Fields("ymax").value
174                 oRsDD.Fields("shapetype").value = .Fields("Shapetype").value
176                 oRsDD.UpdateBatch
178                 oRsDD.Close
                    'oConnectionDD.Close
180                 Set oRsDD = Nothing
182                 Set oConnectionDD = Nothing

184                 Set oExtent = New TatukGIS_XDK10.XGIS_Extent
186                 Set oSQLLyr = New TatukGIS_XDK10.XGIS_LayerSqlAdo
188                 oSQLLyr.Name = .Fields("LayerCaption").value
190                 oSQLLyr.SQLParameter("LAYER") = .Fields("LayerName").value
192                 oSQLLyr.SQLParameter("DIALECT") = g_sGlobalDialect
194                 oSQLLyr.SQLParameter("ADO") = IIf(bSQLServerInUse, g_sGlobalConnectionString, Replace(.Fields("ADO").value, "CLIENTDBPATH", g_sAppPath))
                    
                    'ADO=Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\\\Users\\\\OASIS\\\\Documents\\\\iMMAP - OASIS\\\\OASIS Client\\\\data\\\\db\\\\dynamicdata\\\\NOMADOASIS.MDB;Persist Security Info=False\\nDialect=MSJET\\nLayer=dd_Demographics_ddAirports\\n.ttkls
                    
196                 oSQLLyr.HideFromLegend = False
198                 oSQLLyr.Params.Visible = True
200                 oExtent.Prepare .Fields("XMIN").value, .Fields("YMIN").value, .Fields("XMAX").value, .Fields("YMAX").value
202                 oSQLLyr.Extent = oExtent
204                 Set oTTKLyr = oSQLLyr

                End If

206             If Not bLoadFailed Then
                
208                 If GIS10.get(oTTKLyr.Name) Is Nothing Then
    
210                     GIS10.Add oTTKLyr
    
212                     If Not bLoadFailed Then
214                         oTTKLyr.Transparency = .Fields("Transparency").value
216                         oTTKLyr.Collapsed = Not .Fields("IsExpanded").value
218                         oTTKLyr.Active = .Fields("IsVisible").value 'oTTKLyr.Active
              
220                         If Not IsNull(.Fields("INISettings").value) Then
222                             Set oSList = New TatukGIS_XDK10.XStringList
224                             sINI = .Fields("INISettings").value
226                             sINI = Replace$(sINI, "CLIENTDBPATH", g_sAppPath)
228                             oSList.Text = sINI
230                             oTTKLyr.ParamsList.LoadFromStrings oSList
                                On Error GoTo AddSQLLayersToProject9_Err
                            End If
                            
232                         oTTKLyr.Path = "[TatukGIS Layer]\\n"
234                         oTTKLyr.Path = oTTKLyr.Path & "Storage=Native\\n"
236                         oTTKLyr.Path = oTTKLyr.Path & "ADO=" & IIf(bSQLServerInUse, g_sGlobalConnectionString, Replace(.Fields("ADO").value, "CLIENTDBPATH", g_sAppPath)) & "\\n"
238                         oTTKLyr.Path = oTTKLyr.Path & "Dialect=MSJET\\n"
240                         oTTKLyr.Path = oTTKLyr.Path & "Layer=" & oTTKLyr.Table & "\\n"
242                         oTTKLyr.Path = oTTKLyr.Path & ".ttkls"

                        End If
                    End If
                
                End If

244             .MoveNext
246             bLoadFailed = False
            Loop
          
248         If Not .EOF Or Not .Bof Then .MoveFirst

250         Do Until .EOF
252             Set oTTKLyr = GIS10.get(.Fields("LayerCaption").value)
254             If Not oTTKLyr Is Nothing Then oTTKLyr.ZOrder = .Fields("Sequence").value
256             .MoveNext
            Loop

258         .Close
 
        End With
   
260     Set oSQLLyr = Nothing
262     Set oWMSLyr = Nothing
264     Set oSHPLyr = Nothing
266     Set oTTKLyr = Nothing

        '<EhFooter>
        Exit Function

AddSQLLayersToProject9_Err:
        DebugPrint Err.Description & vbCrLf & "in OASISClient.frmMain.AddSQLLayersToProject9 " & "at line " & Erl
        g_bMapLoadedCorrectly = False
        bLoadFailed = True
        Resume Next
        '</EhFooter>
End Function

Private Function PrepareLocalAppSettingValue(sSetting As String) As Boolean
    SafeMoveFirst g_RSLocalAppSettings
    
    With g_RSLocalAppSettings
        .Find "SettingName = '" & sSetting & "'"
    
        If .EOF Then
            .AddNew
            .Fields("SettingName").value = sSetting
            .UpDate
        Else
            PrepareLocalAppSettingValue = True
        End If
    
    End With
 End Function

Private Sub OpenLocalAppSettings()

On Error Resume Next

    With oMapSettings
        .AlwaysSaveMapStateOnExit = CBool(IIf(PrepareLocalAppSettingValue("oMapSettings.AlwaysSaveMapStateOnExit"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .AutoScroll = CBool(IIf(PrepareLocalAppSettingValue("oMapSettings.AutoScroll"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .iMAPUnits = CInt(IIf(PrepareLocalAppSettingValue("oMapSettings.iMAPUnits"), g_RSLocalAppSettings.Fields("SettingValue1").value, 0))
        .MapRotation = CInt(IIf(PrepareLocalAppSettingValue("oMapSettings.MapRotation"), g_RSLocalAppSettings.Fields("SettingValue1").value, 0))
        .ScrollBars = CInt(IIf(PrepareLocalAppSettingValue("oMapSettings.ScrollBars"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .StoreLayerParamsInProject = CBool(IIf(PrepareLocalAppSettingValue("oMapSettings.StoreLayerParamsInProject"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
    End With

    With oMapObjects
        .NorthArrowColor = CLng(IIf(PrepareLocalAppSettingValue("oMapObjects.NorthArrowColor"), g_RSLocalAppSettings.Fields("SettingValue1").value, 0))
        .NorthArrowPicture = IIf(PrepareLocalAppSettingValue("oMapObjects.NorthArrowPicture"), g_RSLocalAppSettings.Fields("SettingValue1").value, "")
        .NorthArrowPicture = Replace$(.NorthArrowPicture, "CLIENTDBPATH", g_sAppPath)
        
        .NorthArrowTransparency = CBool(IIf(PrepareLocalAppSettingValue("oMapObjects.NorthArrowTransparency"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .NorthArrowType = CInt(IIf(PrepareLocalAppSettingValue("oMapObjects.NorthArrowType"), g_RSLocalAppSettings.Fields("SettingValue1").value, 0))
        .UseNorthArrow = CBool(IIf(PrepareLocalAppSettingValue("oMapObjects.UseNorthArrow"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .UseScaleBar = CBool(IIf(PrepareLocalAppSettingValue("oMapObjects.UseScaleBar"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .UseWaterMark = CBool(IIf(PrepareLocalAppSettingValue("oMapObjects.UseWaterMark"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .WaterMarkPath = IIf(PrepareLocalAppSettingValue("oMapObjects.WaterMarkPath"), g_RSLocalAppSettings.Fields("SettingValue1").value, "")
    End With

    With oIncidentLayerSettings
        .CachedPaint = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.CachedPaint"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .ConfigFilePAth = IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.ConfigFilePAth"), g_RSLocalAppSettings.Fields("SettingValue1").value, "")
        .HideFromLegend = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.HideFromLegend"), g_RSLocalAppSettings.Fields("SettingValue1").value, True))
        .IgnoreShapeParams = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.IgnoreShapeParams"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .IncrementalPaint = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.IncrementalPaint"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .UseConfig = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.UseConfig"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .UseFileParams = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.UseFileParams"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .VisibleFromStart = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.VisibleFromStart"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
    End With

    With oCoordTransSettings
        .False_Easting = CLng(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.False_Easting"), g_RSLocalAppSettings.Fields("SettingValue1").value, 0))
        .False_Northing = CLng(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.False_Northing"), g_RSLocalAppSettings.Fields("SettingValue1").value, 0))
        .Lat_Of_Origin = CLng(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.Lat_Of_Origin"), g_RSLocalAppSettings.Fields("SettingValue1").value, 0))
        .Long_Of_Origin = CLng(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.Long_Of_Origin"), g_RSLocalAppSettings.Fields("SettingValue1").value, 0))
        .Sphere = CBool(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.Sphere"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .Inverse_Flattening = IIf(PrepareLocalAppSettingValue("oCoordTransSettings.Inverse_Flattening"), g_RSLocalAppSettings.Fields("SettingValue1").value, 6378137)
        .Semi_Major_Axis = CDbl(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.Semi_Major_Axis"), g_RSLocalAppSettings.Fields("SettingValue1").value, "298.2572236"))
    End With
        
    With oUrlLayerSettings
        .AutoShutTime = CInt(IIf(PrepareLocalAppSettingValue("oUrlLayerSettings.AutoShutTime"), g_RSLocalAppSettings.Fields("SettingValue1").value, 0))
        .AutoShutWin = CBool(IIf(PrepareLocalAppSettingValue("oUrlLayerSettings.AutoShutWin"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .UseExtendedInfoWin = CBool(IIf(PrepareLocalAppSettingValue("oUrlLayerSettings.UseExtendedInfoWin"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .WinHeight = CInt(IIf(PrepareLocalAppSettingValue("oUrlLayerSettings.WinHeight"), g_RSLocalAppSettings.Fields("SettingValue1").value, 5235))
        .WinWidth = CInt(IIf(PrepareLocalAppSettingValue("oUrlLayerSettings.WinWidth"), g_RSLocalAppSettings.Fields("SettingValue1").value, 4800))
    End With
                
    With oMapTipSetting
        .Enabled = CBool(IIf(PrepareLocalAppSettingValue("oMapTipSetting.Enabled"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .MapTipLayer = IIf(PrepareLocalAppSettingValue("oMapTipSetting.MapTipLayer"), g_RSLocalAppSettings.Fields("SettingValue1").value, "--All--")
        .MapTipField = IIf(PrepareLocalAppSettingValue("oMapTipSetting.MapTipField"), g_RSLocalAppSettings.Fields("SettingValue1").value, "UID")
        .TextColor = CLng(IIf(PrepareLocalAppSettingValue("oMapTipSetting.TextColor"), g_RSLocalAppSettings.Fields("SettingValue1").value, &H80000012))
        .TipColor = CLng(IIf(PrepareLocalAppSettingValue("oMapTipSetting.TipColor"), g_RSLocalAppSettings.Fields("SettingValue1").value, &HC0FFFF))
        .TipDelay = CInt(IIf(PrepareLocalAppSettingValue("oMapTipSetting.TipDelay"), g_RSLocalAppSettings.Fields("SettingValue1").value, 4))
        .TipBorder = CBool(IIf(PrepareLocalAppSettingValue("oMapTipSetting.TipBorder"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        tmrToolTip.Interval = .TipDelay * 1000
        tmrToolTip.Enabled = .Enabled
    End With
        
    With oSelectionStyle
        .color = CLng(IIf(PrepareLocalAppSettingValue("oSelectionStyle.Color"), g_RSLocalAppSettings.Fields("SettingValue1").value, GIS10.SelectionColor))
        .OutLineOnly = CBool(IIf(PrepareLocalAppSettingValue("oSelectionStyle.OutLineOnly"), g_RSLocalAppSettings.Fields("SettingValue1").value, GIS10.SelectionOutlineOnly))
        .Transparency = CInt(IIf(PrepareLocalAppSettingValue("oSelectionStyle.Transparency"), g_RSLocalAppSettings.Fields("SettingValue1").value, GIS10.SelectionTransparency))
        .Width = CInt(IIf(PrepareLocalAppSettingValue("oSelectionStyle.Width"), g_RSLocalAppSettings.Fields("SettingValue1").value, GIS10.SelectionWidth))
    End With
        
    oLocatorSettings.Level1 = IIf(PrepareLocalAppSettingValue("oLocatorSettings.Level1"), g_RSLocalAppSettings.Fields("SettingValue1").value, "Intersect Interior Interior") 'GisUtils.GIS_RELATE_INTERSECT_INTERIOR_INTERIOR
    oLocatorSettings.Level2 = IIf(PrepareLocalAppSettingValue("oLocatorSettings.Level2"), g_RSLocalAppSettings.Fields("SettingValue1").value, "Contains")
        
    With g_ZoomToSettings
        .SaveOnExit = CBool(IIf(PrepareLocalAppSettingValue("g_ZoomToSettings.SaveOnExit"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
        .UseMultiple = CBool(IIf(PrepareLocalAppSettingValue("g_ZoomToSettings.UseMultiple"), g_RSLocalAppSettings.Fields("SettingValue1").value, False))
    End With
        
End Sub

Private Sub SavePrivateAppSettings(Optional bPrompt As Boolean)
        '<EhHeader>
        On Error GoTo SavePrivateAppSettings_Err
        '</EhHeader>
        
        If Not oMapSettings.AlwaysSaveMapStateOnExit Then Exit Sub
        
100     With oMapSettings
            PrepareLocalAppSettingValue "oMapSettings.AlwaysSaveMapStateOnExit"
            g_RSLocalAppSettings.Fields("SettingValue1").value = CStr(.AlwaysSaveMapStateOnExit)
            g_RSLocalAppSettings.UpDate
            '            .AlwaysSaveMapStateOnExit
            '            .AutoScroll
            '            .iMAPUnits
102         PrepareLocalAppSettingValue "oMapSettings.MapRotation"
104         g_RSLocalAppSettings.Fields("SettingValue1").value = CStr(.MapRotation)
106         g_RSLocalAppSettings.UpDate
            
108         PrepareLocalAppSettingValue "oMapSettings.ScrollBars"
110         g_RSLocalAppSettings.Fields("SettingValue1").value = CStr(.ScrollBars)
112         g_RSLocalAppSettings.UpDate
      
            '            .StoreLayerParamsInProject
        End With

114     With oMapObjects

116         PrepareLocalAppSettingValue "oMapObjects.NorthArrowColor"
118         g_RSLocalAppSettings.Fields("SettingValue1").value = CStr(.NorthArrowColor)
120         g_RSLocalAppSettings.UpDate
            
            If Not .NorthArrowPicture = "" Then
122             PrepareLocalAppSettingValue "oMapObjects.NorthArrowPicture"
124             g_RSLocalAppSettings.Fields("SettingValue1").value = Replace(.NorthArrowPicture, g_sAppPath, "CLIENTDBPATH")
126             g_RSLocalAppSettings.UpDate
            End If
            
128         PrepareLocalAppSettingValue "oMapObjects.NorthArrowTransparency"
130         g_RSLocalAppSettings.Fields("SettingValue1").value = CStr(.NorthArrowTransparency)
132         g_RSLocalAppSettings.UpDate
            
134         PrepareLocalAppSettingValue "oMapObjects.NorthArrowType"
136         g_RSLocalAppSettings.Fields("SettingValue1").value = .NorthArrowType
138         g_RSLocalAppSettings.UpDate
            
140         PrepareLocalAppSettingValue "oMapObjects.UseNorthArrow"
142         g_RSLocalAppSettings.Fields("SettingValue1").value = .UseNorthArrow
144         g_RSLocalAppSettings.UpDate
            
146         PrepareLocalAppSettingValue "oMapObjects.UseScaleBar"
148         g_RSLocalAppSettings.Fields("SettingValue1").value = .UseScaleBar
150         g_RSLocalAppSettings.UpDate
            
152         PrepareLocalAppSettingValue "oMapObjects.UseWaterMark"
154         g_RSLocalAppSettings.Fields("SettingValue1").value = .UseWaterMark
156         g_RSLocalAppSettings.UpDate
            
            If Not .WaterMarkPath = "" Then
158             PrepareLocalAppSettingValue "oMapObjects.WaterMarkPath"
160             g_RSLocalAppSettings.Fields("SettingValue1").value = .WaterMarkPath
162             g_RSLocalAppSettings.UpDate
            End If
        
        End With

164     With oIncidentLayerSettings
166         PrepareLocalAppSettingValue "oIncidentLayerSettings.CachedPaint"
168         g_RSLocalAppSettings.Fields("SettingValue1").value = .CachedPaint
170         g_RSLocalAppSettings.UpDate
        
172         PrepareLocalAppSettingValue "oIncidentLayerSettings.ConfigFilePAth"
174         g_RSLocalAppSettings.Fields("SettingValue1").value = IIf(Len(.ConfigFilePAth) > 0, .ConfigFilePAth, " ")
176         g_RSLocalAppSettings.UpDate
        
178         PrepareLocalAppSettingValue "oIncidentLayerSettings.HideFromLegend"
180         g_RSLocalAppSettings.Fields("SettingValue1").value = .HideFromLegend
182         g_RSLocalAppSettings.UpDate
        
184         PrepareLocalAppSettingValue "oIncidentLayerSettings.IgnoreShapeParams"
186         g_RSLocalAppSettings.Fields("SettingValue1").value = .IgnoreShapeParams
188         g_RSLocalAppSettings.UpDate
        
190         PrepareLocalAppSettingValue "oIncidentLayerSettings.IncrementalPaint"
192         g_RSLocalAppSettings.Fields("SettingValue1").value = .IncrementalPaint
194         g_RSLocalAppSettings.UpDate
        
196         PrepareLocalAppSettingValue "oIncidentLayerSettings.UseConfig"
198         g_RSLocalAppSettings.Fields("SettingValue1").value = .UseConfig
200         g_RSLocalAppSettings.UpDate
        
202         PrepareLocalAppSettingValue "oIncidentLayerSettings.UseFileParams"
204         g_RSLocalAppSettings.Fields("SettingValue1").value = .UseFileParams
206         g_RSLocalAppSettings.UpDate
        
208         PrepareLocalAppSettingValue "oIncidentLayerSettings.VisibleFromStart"
210         g_RSLocalAppSettings.Fields("SettingValue1").value = .VisibleFromStart
212         g_RSLocalAppSettings.UpDate

        End With

214     With oCoordTransSettings

216         PrepareLocalAppSettingValue "oCoordTransSettings.False_Northing"
218         g_RSLocalAppSettings.Fields("SettingValue1").value = .Inverse_Flattening
220         g_RSLocalAppSettings.UpDate
        
222         PrepareLocalAppSettingValue "oCoordTransSettings.False_Easting"
224         g_RSLocalAppSettings.Fields("SettingValue1").value = .Inverse_Flattening
226         g_RSLocalAppSettings.UpDate
        
228         PrepareLocalAppSettingValue "oCoordTransSettings.Sphere"
230         g_RSLocalAppSettings.Fields("SettingValue1").value = .Inverse_Flattening
232         g_RSLocalAppSettings.UpDate
        
234         PrepareLocalAppSettingValue "oCoordTransSettings.Long_Of_Origin"
236         g_RSLocalAppSettings.Fields("SettingValue1").value = .Inverse_Flattening
238         g_RSLocalAppSettings.UpDate
        
240         PrepareLocalAppSettingValue "oCoordTransSettings.Lat_Of_Origin"
242         g_RSLocalAppSettings.Fields("SettingValue1").value = .Inverse_Flattening
244         g_RSLocalAppSettings.UpDate
        
246         PrepareLocalAppSettingValue "oCoordTransSettings.Inverse_Flattening"
248         g_RSLocalAppSettings.Fields("SettingValue1").value = .Inverse_Flattening
250         g_RSLocalAppSettings.UpDate
        
252         PrepareLocalAppSettingValue "oCoordTransSettings.Semi_Major_Axis"
254         g_RSLocalAppSettings.Fields("SettingValue1").value = .Semi_Major_Axis
256         g_RSLocalAppSettings.UpDate
        End With
        
258     With oUrlLayerSettings
260         PrepareLocalAppSettingValue "oUrlLayerSettings.AutoShutTime"
262         g_RSLocalAppSettings.Fields("SettingValue1").value = .AutoShutTime
264         g_RSLocalAppSettings.UpDate
        
266         PrepareLocalAppSettingValue "oUrlLayerSettings.AutoShutWin"
268         g_RSLocalAppSettings.Fields("SettingValue1").value = .AutoShutWin
270         g_RSLocalAppSettings.UpDate
        
272         PrepareLocalAppSettingValue "oUrlLayerSettings.UseExtendedInfoWin"
274         g_RSLocalAppSettings.Fields("SettingValue1").value = .UseExtendedInfoWin
276         g_RSLocalAppSettings.UpDate
        
278         PrepareLocalAppSettingValue "oUrlLayerSettings.WinHeight"
280         g_RSLocalAppSettings.Fields("SettingValue1").value = .WinHeight
282         g_RSLocalAppSettings.UpDate
        
284         PrepareLocalAppSettingValue "oUrlLayerSettings.WinWidth"
286         g_RSLocalAppSettings.Fields("SettingValue1").value = .WinWidth
288         g_RSLocalAppSettings.UpDate
        
        End With
    
290     With oMapTipSetting
292         PrepareLocalAppSettingValue "oMapTipSetting.Enabled"
294         g_RSLocalAppSettings.Fields("SettingValue1").value = .Enabled
296         g_RSLocalAppSettings.UpDate
        
298         PrepareLocalAppSettingValue "oMapTipSetting.MapTipLayer"
300         g_RSLocalAppSettings.Fields("SettingValue1").value = IIf(Len(.MapTipLayer) > 0, .MapTipLayer, " ")
302         g_RSLocalAppSettings.UpDate
        
304         PrepareLocalAppSettingValue "oMapTipSetting.MapTipField"
306         g_RSLocalAppSettings.Fields("SettingValue1").value = IIf(Len(.MapTipField) > 0, .MapTipField, " ")
308         g_RSLocalAppSettings.UpDate
        
310         PrepareLocalAppSettingValue "oMapTipSetting.TextColor"
312         g_RSLocalAppSettings.Fields("SettingValue1").value = .TextColor
314         g_RSLocalAppSettings.UpDate
        
316         PrepareLocalAppSettingValue "oMapTipSetting.TipColor"
318         g_RSLocalAppSettings.Fields("SettingValue1").value = .TipColor
320         g_RSLocalAppSettings.UpDate
        
322         PrepareLocalAppSettingValue "oMapTipSetting.TipDelay"
324         g_RSLocalAppSettings.Fields("SettingValue1").value = .TipDelay
326         g_RSLocalAppSettings.UpDate
        
328         PrepareLocalAppSettingValue "oMapTipSetting.TipBorder"
330         g_RSLocalAppSettings.Fields("SettingValue1").value = .TipBorder
332         g_RSLocalAppSettings.UpDate
           
        End With
        
334     With oSelectionStyle
336         PrepareLocalAppSettingValue "oSelectionStyle.Color"
338         g_RSLocalAppSettings.Fields("SettingValue1").value = .color
340         g_RSLocalAppSettings.UpDate
        
342         PrepareLocalAppSettingValue "oSelectionStyle.OutLineOnly"
344         g_RSLocalAppSettings.Fields("SettingValue1").value = .OutLineOnly
346         g_RSLocalAppSettings.UpDate
        
348         PrepareLocalAppSettingValue "oSelectionStyle.Transparency"
350         g_RSLocalAppSettings.Fields("SettingValue1").value = .Transparency
352         g_RSLocalAppSettings.UpDate
        
354         PrepareLocalAppSettingValue "oSelectionStyle.Width"
356         g_RSLocalAppSettings.Fields("SettingValue1").value = .Width
358         g_RSLocalAppSettings.UpDate
        End With

360     With g_ZoomToSettings
362         PrepareLocalAppSettingValue "g_ZoomToSettings.SaveOnExit"
364         g_RSLocalAppSettings.Fields("SettingValue1").value = .SaveOnExit
366         g_RSLocalAppSettings.UpDate
368         PrepareLocalAppSettingValue "g_ZoomToSettings.UseMultiple"
370         g_RSLocalAppSettings.Fields("SettingValue1").value = .UseMultiple
372         g_RSLocalAppSettings.UpDate
        End With

374     PrepareLocalAppSettingValue "oLocatorSettings.Level1"
376     g_RSLocalAppSettings.Fields("SettingValue1").value = IIf(Len(oLocatorSettings.Level1) > 0, oLocatorSettings.Level1, " ")
378     g_RSLocalAppSettings.UpDate
        'GisUtils.GIS_RELATE_INTERSECT_INTERIOR_INTERIOR
    
380     PrepareLocalAppSettingValue "oLocatorSettings.Level2"
382     g_RSLocalAppSettings.Fields("SettingValue1").value = IIf(Len(oLocatorSettings.Level2) > 0, oLocatorSettings.Level2, " ")
384     g_RSLocalAppSettings.UpDate
     
        '<EhFooter>
        Exit Sub

SavePrivateAppSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SavePrivateAppSettings " & _
               "at line " & Erl
        
        
        Resume Next
        '</EhFooter>
End Sub

Private Function CheckSynchDS() As Boolean
        '<EhHeader>
        On Error GoTo CheckSynchDS_Err
        '</EhHeader>
        Dim rsSynchDS As New ADODB.Recordset
        On Error GoTo Hell
    
100     rsSynchDS.CursorLocation = g_sGlobalCursorLocation
102     rsSynchDS.Open "SELECT * FROM SQLSynchLayers", m_Cnn, adOpenDynamic, adLockReadOnly
    
104     If rsSynchDS.RecordCount > 0 Then
106         CheckSynchDS = True
        End If
    
        Exit Function
Hell:
108     Err.Clear
110     CheckSynchDS = False
        '<EhFooter>
        Exit Function

CheckSynchDS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CheckSynchDS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub AddAllAdditionalStandardLayers()
        '<EhHeader>
        On Error GoTo AddAllAdditionalStandardLayers_Err
        '</EhHeader>
        
100     CreateOASISLayers
102     LoadSQLLyr '2

104     If Not m_oDrawLyr Is Nothing Then
106         GIS10.Add m_oDrawLyr
        End If
        
        If Not m_oSpatialiseLayer Is Nothing Then
             GIS10.Add m_oSpatialiseLayer
        End If
        
108     If Not m_oBufferLyr Is Nothing Then
110         GIS10.Add m_oBufferLyr
        End If
        
120     SafeMoveFirst g_RSAppSettings
122     g_RSAppSettings.Find "SettingName = 'EventLayerName'"

124     Set EventLayer = GIS10.get(g_RSAppSettings.Fields.Item("SettingValue1").value) '("Provinces_region")

        '<EhFooter>
        Exit Sub

AddAllAdditionalStandardLayers_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.AddAllAdditionalStandardLayers " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function LoadProjectFileFromDB(Optional sPathPassed As String = "") As Boolean
        '<EhHeader>
        On Error GoTo LoadProjectFileFromDB_Err
        '</EhHeader>

        Dim oStream As ADODB.Stream
        Dim oRS As New ADODB.Recordset
        Dim sText As String
        Dim sPath As String
        
        If Not sPathPassed = "" Then
            sPath = sPathPassed
        Else
            sPath = g_sAppPath & "\data\user\Exports\tempproj.ttkgp"
        End If
        
100     LoadProjectFileFromDB = False

102     With oRS

104         .Open "SELECT * from [ttkGISProjectDef]", m_Cnn, adOpenDynamic, adLockBatchOptimistic

106         If Not .State = adStateClosed Then
                
                Set g_DatabaseSpecExtent = Nothing
                
108             If Not .EOF Then
110
112                 LoadProjectFileFromDB = .Fields("InUse").value
                    Set g_DatabaseSpecExtent = New TatukGIS_XDK10.XGIS_Extent
                    g_DatabaseSpecExtent.Prepare .Fields("XMIN").value, .Fields("YMIN").value, .Fields("XMAX").value, .Fields("YMAX").value
                    
                    
                    ' GIS10.viewer.Extent = g_DatabaseSpecExtent
                    GIS10.viewer.Extent.Assign_ g_DatabaseSpecExtent
                        
                    If LoadProjectFileFromDB Then sText = .Fields("MapData").value

                End If
                
114             .Close
            End If

        End With
    
116     Set oRS = Nothing
    
118     If Len(sText) > 10 Then
    
120         Set oStream = New ADODB.Stream
            On Error Resume Next
122         Kill sPath
            On Error GoTo LoadProjectFileFromDB_Err
124         oStream.Open
126         oStream.Type = 2    ' Set type to text
128         oStream.Charset = "ascii"
130         oStream.WriteText sText
132         oStream.SaveToFile (sPath)
134         oStream.Close
136         Set oStream = Nothing
    
        End If
        
        '<EhFooter>
        Exit Function

LoadProjectFileFromDB_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.LoadProjectFileFromDB " & "at line " & Erl
        LoadProjectFileFromDB = False
        '</EhFooter>
End Function

Public Sub InitMapFromMapLibraryDefault()
    
    Dim sPath As String
    Dim oExtent As New XGIS_Extent

    If m_frmMnuOperations.ctlMapFileContainer1.IsMapSelected Then
        sPath = m_frmMnuOperations.ctlMapFileContainer1.GetActiveMapPath
        
        If Not Len(Replace(sPath, "Database Driven:", "")) = Len(sPath) Then
            m_frmMnuOperations.ctlMapFileContainer1.PrepareMapForLoad sPath
            InitMap sPath, m_frmMnuOperations.ctlMapFileContainer1.GetActiveMapExtent
        Else
            InitMap sPath
        End If
        
    Else

        'MsgBox "The OASIS default map has been selected.  Please add a new map or contact your OASIS adminstrator for assistance", vbInformation
        If FileExists(g_sAppPath & "\data\gis\oasisdefaultdatapack\Cultural.ttkgp") Then
            InitMap g_sAppPath & "\data\gis\oasisdefaultdatapack\Cultural.ttkgp"
            oExtent.Prepare 0, -30, 120, 60
            GIS10.VisibleExtent = oExtent
            'GIS10
        Else
            InitMap g_sAppPath & "\data\user\maps\DefaultMap.ttkgp"
        End If
        
    End If
    
End Sub

Public Sub InitMap(Optional sMapPath As String, _
                   Optional oExtent As XGIS_Extent)
        '<EhHeader>
        On Error GoTo InitMap_Err
        '</EhHeader>
        Dim bUseVisibleExtent As Boolean
        Dim sLayers() As String
        Dim i As Integer
        Dim oLyr As TatukGIS_XDK10.XGIS_LayerAbstract

        '  If Len(sMapPath) < 5 Then
        '  If LoadProjectFileFromDB Then
        '  m_sInitialMapName = g_sAppPath & "\data\user\Exports\tempproj.ttkgp"
        '    GIS10.VisibleExtent.Assign_ g_DatabaseSpecExtent
        'End If
        ' End If
        
        ' If FileExists(g_sAppPath & "\data\user\maps\immaparrow.gif") Then
        '  NArrow.Path = g_sAppPath & "\data\user\maps\immaparrow.gif"
        '  End If
        m_sCurrentMapPath = sMapPath
        
100     GIS10.Lock
    
102     m_bLOADING = True

104     m_bMapInitialized = True
105     CreateIncidentsTTKGPReference
        
106     PrepareOverViewMap IIf(Len(sMapPath) > 0, sMapPath, m_sInitialMapName)
    
108     LoadMapProducts IIf(Len(sMapPath) > 0, sMapPath, m_sInitialMapName)

 bUseVisibleExtent = Not GisUtils.GisIsContainExtent(GIS10.VisibleExtent, GIS10.Extent)
    If m_eZoomberMaxExtent Is Nothing Then
    
        If bUseVisibleExtent Then
            ctlZoomSlider1.Init GIS10.VisibleExtent, GIS10.VisibleExtent, vbWhite, &HC0&, vbWhite
        Else
            ctlZoomSlider1.Init GIS10.Extent, GIS10.VisibleExtent, vbWhite, &HC0&, vbWhite
        End If
        
        Else
        ctlZoomSlider1.Init m_eZoomberMaxExtent, GIS10.VisibleExtent, vbWhite, &HC0&, vbWhite
        End If
        
        ctlZoomSlider1.SetColours m_lZoomSelectColour, m_lZoomBackColour, m_lZoomCrosshairColour
                  
110     LoadAvailableThematics

        g_bMapLoadedCorrectly = True
         
        AddSQLLayersToProject9
112     AddAllAdditionalStandardLayers
        
114     ReDim m_oSQLGenericLyrs(0)
        
116     SafeMoveFirst g_RSAppSettings
118     g_RSAppSettings.Find "SettingName = 'W3Settings'"
        
124     FillCOPValues

126     Set g_PrevExt = GIS10.Extent
    
        Dim RS As New ADODB.Recordset

128     Set RS = m_Cnn.Execute("SELECT DefaultViewName,DefaultViewX,DefaultViewY,DefaultViewZ,LatestViewX,LatestViewY,LatestViewZ,LatestMapName FROM Personnell WHERE Personnell_ID = " & g_CurrentUserID)

130     If Not RS.Bof Then SafeMoveFirst RS

        'If Not IsNull(RS.Fields.Item("DefaultViewName").Value) Then
132     '   m_DefaultViewName = "Start View" ' RS.Fields.Item("DefaultViewName").Value
134     '  m_DefaultViewx = RS.Fields.Item("DefaultViewX").Value
136     '  m_DefaultViewy = RS.Fields.Item("DefaultViewY").Value
138     ' m_DefaultViewz = RS.Fields.Item("DefaultViewZ").Value
        'End If
   
140     ComThemes.ListIndex = 0
142     C1TabFastFunction.Width = 340

144     If GIS10.items.Count > 0 Then
146         SafeMoveFirst g_RSAppSettings
148         g_RSAppSettings.Find "SettingName = 'MADataSetFiles'"
150         sLayers = Split(g_RSAppSettings.Fields.Item("SettingValue1").value, ",")
    
152         For i = 0 To UBound(sLayers)
154             Set oLyr = GIS10.get(sLayers(i))

156             If Not oLyr Is Nothing Then oLyr.HideFromLegend = True
            Next
    
158         SafeMoveFirst g_RSAppSettings
160         g_RSAppSettings.Find "SettingName = 'HiddenLayers'"
    
            If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").value) Then
    
162             sLayers = Split(g_RSAppSettings.Fields.Item("SettingValue1").value, ",")
        
164             For i = 0 To UBound(sLayers)
166                 Set oLyr = GIS10.get(sLayers(i))

168                 If Not oLyr Is Nothing Then oLyr.HideFromLegend = True
                Next
            
            End If

        End If
        
170     m_frmMnuOperations.Init GIS10.viewer, m_Cnn
        NArrow.GIS_Viewer = GIS10.viewer
        'watermark.GIS_Viewer = GIS10.viewer
        'watermark.Visible = False
172     AB.RecalcLayout

174     m_bLOADING = False
        udRotationAngle.value = 0
        GIS10.RotationPoint = GisUtils.GisCenterPoint(GIS10.Extent)
        
        '  If Not g_DatabaseSpecExtent Is Nothing Then
        '     GIS10.VisibleExtent = g_DatabaseSpecExtent
            
        '  End If
        
       
  
        If Not oExtent Is Nothing Then
            GIS10.VisibleExtent = oExtent
            Set g_DatabaseSpecExtent = oExtent
             
        End If
        
        If GIS10.VisibleDockClientCount = 0 Then
            '   GIS10.Visible = False
                
        End If
        
        
        
176     GIS10.Unlock
        
        ReDim m_PrevExt(0)
        Set m_PrevExt(0) = GIS10.VisibleExtent

        InitSelector
        
    
        'ApplyLocalSettings
        
        
        'm_dStartZoom = GIS10.zoom
        
        'ucSlider1.Tag = 1000

        '<EhFooter>
        Exit Sub

InitMap_Err:
        MsgBox Err.Description '& vbCrLf & '       "in OASISClient.frmMain.InitMap " & '       "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub FillCOPValues()
        '<EhHeader>
        On Error GoTo FillCOPValues_Err
        '</EhHeader>
        Dim i As Integer
        
100     m_bThematicsDone = False
        
102     abGridTools.Tools.Item("comLyr").CBClear
104     comActiveLyr.Clear
106     comActiveLyr.AddItem "No Active Layer"
        
108     cmbSnap.Clear
110     cmbSnap.AddItem "No snapping"
 
        If Not g_RSGISGridTableSettings.Bof Then SafeMoveFirst g_RSGISGridTableSettings
        
        abGridTools.Tools.Item("comLyr").CBAddItem "---Nothing---"
        
        With m_frmMnuOperations
            
            mnuLayer(0).Visible = True
            mnuViewLayer.Visible = False
        
            If Not mnuLayer.LBound = mnuLayer.UBound Then

                For i = mnuLayer.UBound To 1 Step -1
                    Unload mnuLayer(i)
                    Unload .mnuLayer(i)
                Next

            End If
        
            If Not g_RSGISGridTableSettings.Bof And Not g_RSGISGridTableSettings.EOF Then
                mnuViewLayer.Visible = True
            
114             Do While Not g_RSGISGridTableSettings.EOF
            
116                 If Not GIS10.get(g_RSGISGridTableSettings.Fields.Item("name").value) Is Nothing Then
118                     If Not g_RSGISGridTableSettings.Fields.Item("alias").value = vbNull Then
120                         abGridTools.Tools.Item("comLyr").CBAddItem g_RSGISGridTableSettings.Fields.Item("alias").value
                            Load mnuLayer(mnuLayer.UBound + 1)
                            mnuLayer(mnuLayer.UBound).caption = g_RSGISGridTableSettings.Fields.Item("alias").value
                            Load .mnuLayer(.mnuLayer.UBound + 1)
                            .mnuLayer(.mnuLayer.UBound).caption = g_RSGISGridTableSettings.Fields.Item("alias").value & " Load to data view"
                            .mnuLayer(.mnuLayer.UBound).Visible = False
                        End If
                    End If
            
122                 g_RSGISGridTableSettings.MoveNext
                Loop

            End If
        
124         If m_oColUserLayers.Count > 0 Then

126             For i = 1 To m_oColUserLayers.Count
128                 abGridTools.Tools.Item("comLyr").CBAddItem m_oColUserLayers.Item(i)
                    Load mnuLayer(mnuLayer.UBound + 1)
                    mnuLayer(mnuLayer.UBound).caption = m_oColUserLayers.Item(i)
                    Load .mnuLayer(.mnuLayer.UBound + 1)
                    .mnuLayer(.mnuLayer.UBound).caption = m_oColUserLayers.Item(i) & " Load to data view"
                    .mnuLayer(.mnuLayer.UBound).Visible = False
                Next

            End If

            If mnuLayer.UBound > 0 Then mnuLayer(0).Visible = False
            .mnuLayer(0).Visible = False
        End With

130     If abGridTools.Tools.Item("comLyr").CBListCount > 0 Then abGridTools.Tools.Item("comLyr").CBListIndex = 0
            
132     If cmbSnap.ListCount > 0 Then
134         cmbSnap.ListIndex = 0
136         comActiveLyr.ListIndex = 0
        End If

        '<EhFooter>
        Exit Sub

FillCOPValues_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.FillCOPValues " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub PrepareModuleMeny()
        'Dim cBar As cExplorerBar
        'Dim cItem As cExplorerBarItem
        '<EhHeader>
        On Error GoTo PrepareModuleMeny_Err
        '</EhHeader>
        Dim sDBSUffix As String

100     sDBSUffix = g_CurrentUserID

102     g_RSStyles.Open "SELECT * FROM Style", m_Cnn, adOpenDynamic, adLockReadOnly

104     g_RSRender.Open "SELECT * FROM Render", m_Cnn, adOpenDynamic, adLockReadOnly
  
        '<EhFooter>
        Exit Sub

PrepareModuleMeny_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.PrepareModuleMeny " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub PrepareOverViewMap(Optional sMapProj As String)
        '<EhHeader>
        On Error GoTo PrepareOverViewMap_Err
        '</EhHeader>
        Dim llm As TatukGIS_XDK10.XGIS_LayerSHP

            Exit Sub

'        If sMapProj = "" Then
'            ' add layers
'100         Set llm = New TatukGIS_XDK10.XGIS_LayerSHP  'minimap states
'102         llm.Path = GisUtils.GisSamplesDataDir + "states.shp"
'104         llm.Name = "states"
'106         llm.UseConfig = False
'108         llm.Params.area.Color = RGB(255, 255, 255)
'110         llm.Params.area.OutlineColor = RGB(&HC0&, &HC0&, &HC0&)
'
'112         m_frmOvMap.GISm.Add llm  'add to minimap
'        Else
'            m_frmOvMap.GISm.Open sMapProj, False
'        End If
'
'        If Not m_frmOvMap.Visible Then
'            m_frmOvMap.Show vbModeless, Me
'        End If
        
        '<EhFooter>
        Exit Sub

PrepareOverViewMap_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.PrepareOverViewMap " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub AddDockableInspectors()
    '<EhHeader>
    On Error GoTo AddDockableInspectors_Err
    '</EhHeader>
    
    '<EhFooter>
    Exit Sub

AddDockableInspectors_Err:
    MsgBox Err.Description & vbCrLf & _
           "in OASISClient.frmMain.AddDockableInspectors " & _
           "at line " & Erl
    Resume Next
    '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        
        '    SkinFramework1.LoadSkin "C:\Program Files\Codejock Software\ActiveX\Xtreme SuitePro ActiveX v11.2.0\Samples\SkinFramework\Styles\Office2007.cjstyles", ""
        '    SkinFramework1.ApplyWindow Me.hwnd
        '    SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
100     ReDim m_shpProps(0)
102     ReDim gdpts(0)
104     ReDim ptsSel(0)
106     ReDim m_PrevExt(0)
108     Set m_PrevExt(0) = GIS10.VisibleExtent

110     If g_bOnlineCheckedAtLogin Then
112         SetIMess "Online"
            If InStr(g_sAppServerPath, "localhost") > 0 Or InStr(g_sAppServerPath, "127.0.0.1") > 0 Then SetIMess "Localhost"        ' "No Internet Status: " & m_sInternetCon
114         Set oClientInterCom = New IClient
116         oClientInterCom.RegisterDataChannel 1, Me.hWnd
118         Set oSync = New SynchWorker
120         Set m_InternetCheck = New SynchWorker
122         Set m_oSQLLyrSynch = New SynchWorker

        Else
124         SetIMess "Offline"
        End If

126     Set g_clsHotKey = New cRegHotKey
        
128     Set g_ColAlerts = New Collection
        
130     Set ClipBoard_xt = New cCustomClipboard
132     Set m_oClipboardViewer = New cClipboardViewer
134     m_oClipboardViewer.InitClipboardChangeNotification Me.hWnd

136     g_clsHotKey.Attach Me.hWnd
138     g_clsHotKey.RegisterKey "Language", vbKeyL, MOD_ALT + MOD_CONTROL
140     g_clsHotKey.RegisterKey "OPSV", vbKeyV, MOD_ALT + MOD_CONTROL
142     'g_clsHotKey.RegisterKey "doPrint", vbKeyP, MOD_ALT + MOD_CONTROL '
144     g_clsHotKey.RegisterKey "Admin", vbKeyA, MOD_ALT + MOD_CONTROL
146     g_clsHotKey.RegisterKey "Dude", vbKeyD, MOD_ALT + MOD_CONTROL
148     g_clsHotKey.RegisterKey "Script1", vbKey1, MOD_ALT + MOD_CONTROL
150     g_clsHotKey.RegisterKey "Script2", vbKey2, MOD_ALT + MOD_CONTROL
152     g_clsHotKey.RegisterKey "Script3", vbKey3, MOD_ALT + MOD_CONTROL
154     g_clsHotKey.RegisterKey "Script4", vbKey4, MOD_ALT + MOD_CONTROL
156     g_clsHotKey.RegisterKey "AppState", vbKeyS, MOD_ALT + MOD_CONTROL
158     g_clsHotKey.RegisterKey "Folders", vbKeyF, MOD_ALT + MOD_CONTROL
160     g_clsHotKey.RegisterKey "GPS", vbKeyG, MOD_ALT + MOD_CONTROL
162     g_clsHotKey.RegisterKey "OVB", vbKeyO, MOD_ALT + MOD_CONTROL
164     g_clsHotKey.RegisterKey "CompInfo", vbKeyC, MOD_ALT + MOD_CONTROL
166     g_clsHotKey.RegisterKey "Ticker", vbKeyQ, MOD_ALT + MOD_CONTROL
168     g_clsHotKey.RegisterKey "ForceIni", vbKeyI, MOD_ALT + MOD_CONTROL
170     g_clsHotKey.RegisterKey "Berserk", vbKeyB, MOD_ALT + MOD_CONTROL
172     g_clsHotKey.RegisterKey "Tracker", vbKeyX, MOD_ALT + MOD_CONTROL
174     g_clsHotKey.RegisterKey "LyrSearch", vbKeyZ, MOD_ALT + MOD_CONTROL
179     g_clsHotKey.RegisterKey "WMS", vbKeyM, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "Debug", vbKeyY, MOD_ALT + MOD_CONTROL

        'Instantiate a reference to the multithreader library

180     Set pUpdateCheckerThread = New Thread
182     Set pLoadLyrAttrToGrdThread = New Thread
184     Set pSubmitGeoMarksThread = New Thread
186     Set pInitThread = New Thread
188     Set pThread = New Thread
190     Set FormThread = New Thread
192     Set pSynchThread = New Thread
194     Set pCheckSynchThread = New Thread
196     Set incThread = New Thread
198     Set pGetIncThread = New Thread
200     Set pLoadMap = New Thread
202     Set pCheckInet = New Thread
        
204     Me.caption = App.Title & " [" & App.major & "." & App.minor & "." & App.Revision & "] powered by iMMAP"
    
        'On Error Resume Next
206     SSC.Language = "VBScript"
208     SSC.AllowUI = True
 
210     With SSC
212         .AddObject "OASISGis", GIS10, True
214         .AddObject "OASISGisUtils", GisUtils, True
216         .AddObject "RSAppSettings", g_RSAppSettings, True
218         .AddObject "RSGISGridTableSettings", g_RSGISGridTableSettings, True
220         .AddObject "CN", m_Cnn, True
222         .AddObject "OASISGISDataGrid", dxGISDataGrid, True
224         .AddObject "OASISToolBar", AB, True
226         .AddObject "OASISCharting", frmSecChart, True
228         .AddObject "OASISMum", Me, True
        End With
                
230     DE9IM.Add "Contains", "T*****FF"
232     DE9IM.Add "Cross", "T*T "
234     DE9IM.Add "Cross Line", "0 "
236     DE9IM.Add "Disjoint", "FF*FF"
238     DE9IM.Add "Equality", "T*F**FF*"
240     DE9IM.Add "Intersect", "T "
242     DE9IM.Add "Intersect Boundary Boundary", "****T "
244     DE9IM.Add "Intersect Boundary Interior", "***T "
246     DE9IM.Add "Intersect Interior Boundary", "*T "
248     DE9IM.Add "Intersect Interior Interior", "T"
250     DE9IM.Add "Intersect1", "*T"
252     DE9IM.Add "Intersect2", "***T"
254     DE9IM.Add "Intersect3", "****T"
256     DE9IM.Add "Line Cross Line", "0"
258     DE9IM.Add "Line Cross Polygon", "T*T"
260     DE9IM.Add "Line Travers Polygon", "T**F"
262     DE9IM.Add "Overlap", "T*T***T"
264     DE9IM.Add "Overlap Line", "1*T***T"
266     DE9IM.Add "Polygon Crossed By Line", "T*****T"
268     DE9IM.Add "Polygon CrossTraversed By Line", "TF**F*T"
270     DE9IM.Add "Polygon Traversed By Line", "TF"
272     DE9IM.Add "Touch", "F***T "
274     DE9IM.Add "Touch Boundary Boundary", "F***T"
276     DE9IM.Add "Touch Boundary Interior", "F**T"
278     DE9IM.Add "Touch Interior", "F**T "
280     DE9IM.Add "Touch Interior Boundary", "FT"
282     DE9IM.Add "Within", "T*F**F"

        Scale1.GIS_Viewer = GIS10.viewer
    
        m_ColZoomScales.Add "1:2000000000"
        m_ColZoomScales.Add "1:1500000000"
        m_ColZoomScales.Add "1:73000000"
        m_ColZoomScales.Add "1:36000000"
        m_ColZoomScales.Add "1:18000000"
        m_ColZoomScales.Add "1:9000000"
        m_ColZoomScales.Add "1:4000000"
        m_ColZoomScales.Add "1:2000000"
        m_ColZoomScales.Add "1:1000000"
        m_ColZoomScales.Add "1:500000"
        m_ColZoomScales.Add "1:280000"
        m_ColZoomScales.Add "1:144000"
        m_ColZoomScales.Add "1:72223"
        m_ColZoomScales.Add "1:36111"
        m_ColZoomScales.Add "1:18055"
        m_ColZoomScales.Add "1:9028"
        m_ColZoomScales.Add "1:4514"
        m_ColZoomScales.Add "1:2257"
        m_ColZoomScales.Add "1:1128"
        m_ColZoomScales.Add "1:564"
        'Unload frmSplash
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckOVBFiles()
        '<EhHeader>
        On Error GoTo CheckOVBFiles_Err
        '</EhHeader>
        Dim fs As New FileSystemObject
        Dim fd As Folder
        Dim i As Integer
        Dim sSuffix As String
        Dim oTXT As TextStream
        Dim oFile As File
    
100     Set fd = fs.GetFolder(g_sAppPath & "\data\user\")

102     For Each oFile In fd.Files

104         If Len(oFile.Name) > 4 Then
        
106             sSuffix = UCase$(right$(oFile.Name, 3))
        
108             If sSuffix = "OVB" Or sSuffix = "OJS" Then
110                 If sSuffix = "OVB" Then
112                     SSC.Language = "VBScript"
                    
                    Else
114                     SSC.Language = "JScript"
                    
                    End If
                                
116                 Set oTXT = oFile.OpenAsTextStream(ForReading)
                    
118                 If m_fmrAddIncident Is Nothing Then MsgBox ""

120                 SSC.AddCode oTXT.ReadAll
                                
122                 oTXT.Close
124                 Set oTXT = Nothing
                
                End If
            End If

        Next

126     Set fd = Nothing
128     Set fs = Nothing

        '<EhFooter>
        Exit Sub

CheckOVBFiles_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CheckOVBFiles " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub MineActionScoring1_ScoringDone()
        '<EhHeader>
        On Error GoTo MineActionScoring1_ScoringDone_Err
        '</EhHeader>
100     m_frmMnuOperations.refreshGrid
        '<EhFooter>
        Exit Sub

MineActionScoring1_ScoringDone_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.MineActionScoring1_ScoringDone " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub MineActionScoring1_ZoomTofeature(sID As String)
        '<EhHeader>
        On Error GoTo MineActionScoring1_ZoomTofeature_Err
        '</EhHeader>
100     DebugPrint sID
    
        '<EhFooter>
        Exit Sub

MineActionScoring1_ZoomTofeature_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.MineActionScoring1_ZoomTofeature " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub mnuFlash_Click()
    FindFeatureFromGrid True, False, False
End Sub

Private Sub FindFeatureFromGrid(bFlash As Boolean, _
                                bSelect As Boolean, _
                                bZoomTo As Boolean, _
                                Optional bCopy As Boolean)
                                
On Error GoTo FindFeatureFromGrid_Err
        '</EhHeader>

        Dim i As Integer
        Dim j As Long
        Dim xx As Double
        Dim yy As Double
        Dim sSQL As String
        Dim lLayer As TatukGIS_XDK10.XGIS_LayerVector
        Dim lShape As TatukGIS_XDK10.XGIS_Shape
        Dim lExtent As New XGIS_Extent
         
100     Set lLayer = GIS10.get(GISGridLayerName)

102     With dxGISDataGrid.ex
        
104         lLayer.Lock

106         For i = 0 To .SelectedCount - 1

108             Set lExtent = New TatukGIS_XDK10.XGIS_Extent
    
110             If GISGridRS.Fields.Count > 1 Then
                        
112                 j = 0

114                 Do Until j = GISGridRS.Fields.Count
                     
116                     If GISGridRS.Fields(j).Name = "GIS_UID" Then
118                         sSQL = "GIS_UID = " & .SelectedNodes(i).values(j)
                        End If
                        
120                     j = j + 1
                    Loop
                                           
                Else
122                 sSQL = "GIS_UID = 4357943657894637895643"
                End If
                
124             Set lShape = lLayer.FindFirst(lLayer.Extent, sSQL, lShape, "", True)

126             If Not IsNull(lShape) And Not lShape Is Nothing Then
                          
                    If bCopy Then
170                        If GISGridRS.Fields.Count > 1 Then
'
172                            j = 0
174                            sSQL = ""
176                             Do Until j = GISGridRS.Fields.Count
                                    If dxGISDataGrid.ex.SelectedNodes(i).IsVisible Then
178                                     sSQL = sSQL & GISGridRS.Fields.Item(j).Name & " = " & dxGISDataGrid.ex.SelectedNodes(i).values(j) & vbCrLf
                                    End If
180                                j = j + 1
                                Loop
182                             Clipboard.Clear
                                Clipboard.SetText sSQL
                            End If
                    End If
                          
128                 If bZoomTo Then
                        
130                     If 1 = 1 Or lShape.Extent.xmax = lShape.Extent.xmin Then
132                         lExtent.xmax = lShape.Extent.xmax + ((GIS10.VisibleExtent.xmax - GIS10.VisibleExtent.xmin) / 4)
134                         lExtent.xmin = lShape.Extent.xmin - ((GIS10.VisibleExtent.xmax - GIS10.VisibleExtent.xmin) / 4)
136                         lExtent.ymax = lShape.Extent.ymax + ((GIS10.VisibleExtent.ymax - GIS10.VisibleExtent.ymin) / 4)
138                         lExtent.ymin = lShape.Extent.ymin - ((GIS10.VisibleExtent.ymax - GIS10.VisibleExtent.ymin) / 4)
                        
                        Else
                        
140                         lExtent.xmax = lShape.Extent.xmax + ((lShape.Extent.xmax - lShape.Extent.xmin) / 1)
142                         lExtent.xmin = lShape.Extent.xmin - ((lShape.Extent.xmax - lShape.Extent.xmin) / 1)
144                         lExtent.ymax = lShape.Extent.ymax + ((lShape.Extent.ymax - lShape.Extent.ymin) / 1)
146                         lExtent.ymin = lShape.Extent.ymin - ((lShape.Extent.ymax - lShape.Extent.ymin) / 1)
                        End If
                        
148                     GIS10.VisibleExtent = lExtent

                    End If
                    
150                 If bFlash Then
152                     If bZoomTo Then Set lShape = lLayer.FindFirst(lLayer.Extent, sSQL, lShape, "", True)
154                     lShape.Flash
156                     lShape.Invalidate
                    End If
                    
158                 If bSelect Then
                        If bZoomTo Then Set lShape = lLayer.FindFirst(lLayer.Extent, sSQL, lShape, "", True)
160                     Set lShape = lShape.MakeEditable
162                     lShape.IsSelected = True
164                     lShape.Invalidate
                    End If
                    
                End If
                
166             Set lShape = Nothing
                
            Next

168         lLayer.Unlock
        
        End With

        '<EhFooter>
        Exit Sub

FindFeatureFromGrid_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.FindFeatureFromGrid " & "at line " & Erl
        Resume Next
        '</EhFooter>

                                
'        '<EhHeader>
'        On Error GoTo FindFeatureFromGrid_Err
'        '</EhHeader>
'        Dim k As Integer
'        Dim i As Integer
'        Dim j As Long
'        Dim xx As Double
'        Dim yy As Double
'        Dim sSQL As String
'        Dim lLayer As TatukGIS_XDK10.XGIS_LayerVector
'        Dim lShape As TatukGIS_XDK10.XGIS_Shape
'        Dim lExtent As New XGIS_Extent
'
'100     Set lLayer = GIS10.get(GISGridLayerName)
'
'102     With dxGISDataGrid.ex
'
'104         lLayer.Lock
'
'106         For i = 0 To .SelectedCount - 1
'
'108             Set lExtent = New TatukGIS_XDK10.XGIS_Extent
'
'110             If GISGridRS.Fields.Count > 1 Then
'
'112                 j = 0
'
'114                 Do Until j = GISGridRS.Fields.Count
'
'                        'Optimize This Exit when UID is found
'116                     If GISGridRS.Fields(j).Name = "GIS_UID" Then
'118                         sSQL = "GIS_UID = " & .SelectedNodes(i).values(j)
'                            Exit Do
'                        End If
'
'120                     j = j + 1
'                    Loop
'
'                Else
'122                 sSQL = "GIS_UID = 4357943657894637895643"
'                End If
'
'124             Set lShape = lLayer.FindFirst(lLayer.Extent, sSQL, lShape, "", True)
'
'126             If Not IsNull(lShape) And Not lShape Is Nothing Then
'
'128                 If bZoomTo Then
'
'130                     If 1 = 1 Or lShape.Extent.XMax = lShape.Extent.XMin Then
'132                         lExtent.XMax = lShape.Extent.XMax + ((GIS10.VisibleExtent.XMax - GIS10.VisibleExtent.XMin) / 4)
'134                         lExtent.XMin = lShape.Extent.XMin - ((GIS10.VisibleExtent.XMax - GIS10.VisibleExtent.XMin) / 4)
'136                         lExtent.YMax = lShape.Extent.YMax + ((GIS10.VisibleExtent.YMax - GIS10.VisibleExtent.YMin) / 4)
'138                         lExtent.YMin = lShape.Extent.YMin - ((GIS10.VisibleExtent.YMax - GIS10.VisibleExtent.YMin) / 4)
'
'                        Else
'
'140                         lExtent.XMax = lShape.Extent.XMax + ((lShape.Extent.XMax - lShape.Extent.XMin) / 1)
'142                         lExtent.XMin = lShape.Extent.XMin - ((lShape.Extent.XMax - lShape.Extent.XMin) / 1)
'144                         lExtent.YMax = lShape.Extent.YMax + ((lShape.Extent.YMax - lShape.Extent.YMin) / 1)
'146                         lExtent.YMin = lShape.Extent.YMin - ((lShape.Extent.YMax - lShape.Extent.YMin) / 1)
'                        End If
'
'148                     GIS10.VisibleExtent = lExtent
'
'                    End If
'
'150                 If bFlash Then
'152                     If bZoomTo Then Set lShape = lLayer.FindFirst(lLayer.Extent, sSQL, lShape, "", True)
'154                     lShape.Flash
'156                     lShape.Invalidate
'                    End If
'
'158                 If bSelect Then
'160                     If bZoomTo Then Set lShape = lLayer.FindFirst(lLayer.Extent, sSQL, lShape, "", True)
'162                     Set lShape = lShape.MakeEditable
'164                     lShape.IsSelected = True
'166                     lShape.Invalidate
'                    End If
'
'168                 If bCopy Then
'
'170                     '   If GISGridRS.Fields.Count > 1 Then
'
'172                     '       j = 0
'174                     '       sSQL = ""
'176                     '        Do Until j = GISGridRS.Fields.Count
'178                     '            sSQL = "GIS_UID = " & dxGISDataGrid.ex.SelectedNodes(i).values(j) & sSQL & vbCrLf
'180                     '            j = j + 1
'                        '        Loop
'182                     '        DebugPrint sSQL
'                        '    End If
'
'184                     ' For k = 0 To lLayer.Fields.Count - 1
'186                     '     DebugPrint lShape.GetField(lLayer.Fields.Item(k).Name)
'                        ' Next
'
'                        'End If
'                    End If
'
'188                 Set lShape = Nothing
'
'                Next
'
'190             lLayer.Unlock
'
'            End With
'
'            '<EhFooter>
'            Exit Sub
'
'FindFeatureFromGrid_Err:
'            MsgBox Err.Description & vbCrLf & "in OASISClient_4.frmMain.FindFeatureFromGrid " & "at line " & Erl
'            Resume Next
'            '</EhFooter>
    End Sub

Private Sub mnuSelectInMap_Click()
    FindFeatureFromGrid False, True, False
End Sub

Private Sub mnuClearSelections_Click()
    
    Dim lLayer As TatukGIS_XDK10.XGIS_LayerVector
    Set lLayer = GIS10.get(GISGridLayerName)
    lLayer.RevertAll
    GIS10.UpDate
    Set lLayer = Nothing

End Sub

Private Sub mnuZoomTo_Click()
    FindFeatureFromGrid True, False, True
End Sub

Private Sub SubMeny1_OASISMnuPressed(oBtn As OASISMenuButton)
        '<EhHeader>
        On Error GoTo SubMeny1_OASISMnuPressed_Err
        '</EhHeader>
100     Select Case oBtn
    
            Case OASISMenuButton.Activitieswizard
            
102         Case OASISMenuButton.IncidentAnalysis
        
104         Case OASISMenuButton.Incidentwizard
106             If m_fmrAddIncident.Visible Then Exit Sub
108             m_fmrAddIncident.Init m_Cnn, "Petri"
110             m_fmrAddIncident.Show vbModeless, Me
112         Case OASISMenuButton.LocationAnalysis
        
114         Case OASISMenuButton.Locationwizard
        
116         Case OASISMenuButton.Personellwizard
        
118         Case OASISMenuButton.Radioroom
        
        
        End Select
        '<EhFooter>
        Exit Sub

SubMeny1_OASISMnuPressed_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SubMeny1_OASISMnuPressed " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub C1TabFastFunction_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
        '<EhHeader>
        On Error GoTo C1TabFastFunction_Switch_Err
        '</EhHeader>
    
100     tmrUpdate.Enabled = False
    
   
102     If NewTab = 6 Then
        
104         If C1TabFastFunction.Width > 410 Then
106             C1TabFastFunction.Width = 340
            Else
108             C1TabFastFunction.Width = 2835
            End If
        
110         Cancel = 1
        Else
        
112         If NewTab = 5 Then
114             tmrUpdate.Enabled = True
            End If
        
116         C1TabFastFunction.Width = 2835
        End If
        '<EhFooter>
        Exit Sub

C1TabFastFunction_Switch_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.C1TabFastFunction_Switch " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ctTreeBookmrks_NodeClick(ByVal nIndex As Long, _
                                     ByVal nColumn As Integer)
        '<EhHeader>
        On Error GoTo ctTreeBookmrks_NodeClick_Err
        '</EhHeader>
100     DebugPrint ctTreeBookmrks.NodeText(nIndex) & " Level:" & ctTreeBookmrks.NodeLevel(nIndex)
   
        Exit Sub
    
        Dim RS As New ADODB.Recordset
   
102     If ctTreeBookmrks.NodeText(nIndex) = "Start View" Then 'm_DefaultViewName Then
104         'GIS10.viewer.SetViewport m_DefaultViewx, m_DefaultViewy
106         'GIS10.viewer.zoom = m_DefaultViewz
            If Not g_DatabaseSpecExtent Is Nothing Then
                GIS10.VisibleExtent = g_DatabaseSpecExtent
            Else
                GIS10.VisibleExtent = GIS10.Extent
            End If
            
108         GIS10.UpDate
            'Map1.ZoomTo m_DefaultViewz, m_DefaultViewx, m_DefaultViewy
        Else

110         If ctTreeBookmrks.NodeLevel(nIndex) = 2 Then
112             RS.Open "SELECT X,Y,Z FROM GeoBookMarks WHERE Name ='" & ctTreeBookmrks.NodeText(nIndex) & "'", m_Cnn, adOpenDynamic, adLockReadOnly

114             If Not RS.Bof Then SafeMoveFirst RS
            
                Dim ptg As New TatukGIS_XDK10.XGIS_Point
            
116             ptg.Prepare CDbl(RS.Fields.Item("X")), CDbl(RS.Fields.Item("Y"))
118             GIS10.zoom = RS.Fields.Item("Z")
120             GIS10.CenterViewport ptg
122             GIS10.UpDate
            End If
        End If

        'DebugPrint  ctTreeBookmrks.NodeLevel(nIndex)
        '<EhFooter>
        Exit Sub

ctTreeBookmrks_NodeClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ctTreeBookmrks_NodeClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdAddBookmarks_Click()
        '<EhHeader>
        On Error GoTo cmdAddBookmarks_Click_Err
        '</EhHeader>

100     FrmAddBookMark.Init m_Cnn, GIS10.viewer.CenterPtg.x, GIS10.viewer.CenterPtg.y, GIS10.viewer.zoom, GIS10.Name
102     FrmAddBookMark.Show vbModal, Me
108     LoadGeoBookMarks
        
        '<EhFooter>
        Exit Sub

cmdAddBookmarks_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdAddBookmarks_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComThemes_Click()
        '<EhHeader>
        On Error GoTo ComThemes_Click_Err
        '</EhHeader>

100     Select Case ComThemes.List(ComThemes.ListIndex)
           Case "NoActivetheme"
102         txtThemeDescription.Text = "Will Clear All Active Map Themes"
            '
104         Case "Target Scoring"
106             txtThemeDescription.Text = "This theme is based on the individual scoring/weight values you have set for different Targets if incidents"

108         Case "Type Scoring"
110             txtThemeDescription.Text = "This theme is based on the individual scoring/weight values you have set for different types of incidents"

112         Case "Scoring"
114             txtThemeDescription.Text = "This theme is based on the scoring/weight values set for incident intensity."

116         Case "Distribution of Incidents"
118             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area."

120         Case "Distribution of Victims"
122             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area."

124         Case "Distribution of Small Arms Fire"
126             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area where the result is based on Incidents of type Small Arms Fire."

128         Case "Distribution of IEDs"
130             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area where the result is based on Incidents of type IED."

132         Case "Distribution of Ambushes"
134             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area where the result is based on Incidents of type Ambush."

136         Case "Distribution of ArmedAttacks"
138             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area where the result is based on Incidents of type Armed Attack."

140         Case "Distribution of Bombings"
142             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area where the result is based on Incidents of type Bombings."

144         Case "Distribution of MortarAttacks"
146             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area where the result is based on Incidents of type Mortar Attacks."

148         Case "Distribution of RPGAttacks"
150             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area where the result is based on Incidents of type RPG Attacks."

152         Case "Distribution of Snippings"
154             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area where the result is based on Incidents of type Snippings."

156         Case "Distribution of Unknown"
158             txtThemeDescription.Text = "This theme is based on the relative distribution of Incidents within a defined geographical area where the result is based on Incidents of Unknown type."

        End Select

        '<EhFooter>
        Exit Sub

ComThemes_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ComThemes_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSetScoring_Click()
        '<EhHeader>
        On Error GoTo cmdSetScoring_Click_Err
        '</EhHeader>
100      frmScoring.Init
102      frmScoring.Show vbModal
     
104      If frmScoring.m_bApply Then
106         m_frmMnuOperations_DoSecurityAnalysis
108         frmScoring.m_bApply = False
         End If
        '<EhFooter>
        Exit Sub

cmdSetScoring_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdSetScoring_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Public Function WriteHTMLReport(sMapPath As String, Optional bExecuteRpt As Boolean = True, Optional arAttachments As Variant) As String
        '<EhHeader>
        On Error GoTo WriteHTMLReport_Err
        '</EhHeader>
        Dim sHTML As String
        Dim Lf As New LineFile
        Dim bBlue As Boolean

    

100     Lf.MakeNew g_sAppPath & "\data\user\Exports\OasisReport.html"
    
102     sHTML = "<html><head><meta http-equiv=""""Content-Type"""" content=""""text/html; charset=iso-8859-1"""">"
104     sHTML = sHTML & vbCrLf & "<title>OASIS Incident Report Receipt</title>"
106     sHTML = sHTML & vbCrLf & "</head><body><Center><h2>OASIS Incident Report Receipt</h2>"
108     sHTML = sHTML & vbCrLf & "This is your Receipt on the reported incident.<br>"
110     sHTML = sHTML & vbCrLf & "Keep this recipt in a safe place for future reference.<br>"
112     sHTML = sHTML & vbCrLf & "If you find any information that is faulty please contact your OASIS administrator <br>"
114     sHTML = sHTML & vbCrLf & "and attach this Receipt. <br>"

116     sHTML = sHTML & vbCrLf & "<img src=""" & sMapPath & """ border=""2"">"

118     sHTML = sHTML & vbCrLf & "</Center>"

120     sHTML = sHTML & vbCrLf & "<table width=""""100%"""" bgcolor=""""#999999"""">"

122     With m_fmrAddIncident

     '   sHTML = sHTML & vbCrLf & "<tr><td width=""""25%""""><strong>Entered By</strong>:</td>"
     '   sHTML = sHTML & vbCrLf & "<td width=""""25%"""">Mr J</td><td width=""""25%""""><strong>Entry Date:</strong></td><td width=""""25%"""">12/12/2006</td></tr>"

124          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#FFFFC0""><td width=""""25%""""><strong>Entered By:</strong></td>" ' Entered By
126          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtEnteredBy.Text & "</strong></td>" ' Entered By

128          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>Record date:</strong></td>" ' Record date
130          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtEntryDate.Text & "</strong></td></tr>" ' Record date

132          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#C0C0FF""><td width=""""25%""""><strong>Entry Date:</strong></td>" 'Entry Date
134          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtEntryDate.Text & "</strong></td>" 'Entry Date

136          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>Violent Incident:</strong></td>" 'Violent"
138          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & IIf(.OptIncViolence(0).value, "Yes", IIf(.OptIncViolence(2).value, "Unknown", "No")) & "</strong></td></tr>"  'Violent"

140          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#FFFFC0""><td width=""""25%""""><strong>Incident Type:</strong></td>" 'Incident type
142          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .ComIncType.List(.ComIncType.ListIndex) & "</strong></td>" 'Incident type

144          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>Incident Target:</strong></td>" 'Incident Target
146          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .ComIncTarget.List(.ComIncTarget.ListIndex) & "</strong></td></tr>" 'Incident Target

148          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#C0C0FF""><td width=""""25%""""><strong>Casualties injured:</strong></td>"  'Casualties injured
150          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtCasualtiesInjured.Text & "</strong></td>"  'Casualties injured

152          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>Casualties Affected:</strong></td>"  'Casualties Affected
                  
154          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtCasualtiesAffected.Text & "</strong></td></tr>"  'Casualties Affected

156          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#FFFFC0""><td width=""""25%""""><strong>Casualties Dead:</strong></td>"  'Casualties Dead
158          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtCasualtiesDead.Text & "</strong></td>"  'Casualties Dead

160          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>Known Casualties:</Strong><td>"
162          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & IIf(.chkUnknown.value = vbChecked, "Unknown", "Known") & "</strong></td></tr>" 'Cualties Known"

164          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#C0C0FF""><td width=""""25%""""><strong>Incident Description:</strong></td>" 'Incident Description
166          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtIncidentDescription.Text & "</strong></td>" 'Incident Description

168          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong></strong></td>" 'Space"
170          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong></strong></td></tr>" 'Space"

172          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#FFFFC0""><td width=""""25%""""><strong>Latitude</strong></td>" 'Latitude"
174          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtY.Text & "</strong></td>" 'Latitude"

176          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>Longitude</strong></td>" 'Longitude"
178          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtX.Text & "</strong></td></tr>" 'Longitude"

180          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#C0C0FF""><td width=""""25%""""><strong>Province:</strong></td>" 'Province
182          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & Mid(.lblProvince_.caption, 10) & "</strong></td>" 'Province

184          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>District:</strong></td>" 'District
186          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & Mid(.lblDistrict_.caption, 10) & "</strong></td></tr>"  'District

188          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>Nearest Town:</strong></td>" 'Nearest Town
190          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .lblNearestTown.caption & "</strong></td>" 'Nearest Town

192          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong></strong></td>" 'Space"
194          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong></strong></td></tr>" 'Space"

196          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#FFFFC0"" ><td width=""""25%""""><strong>Incident Date:</strong></td>" 'Incident Date
198          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & CStr(.MVIncident) & "</strong></td>"  'Incident Date

200          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong></strong></td>" 'Space"
202          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong></strong></td></tr>" 'Space"

204          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#C0C0FF""><td width=""""25%""""><strong>Incident Exact Time:</strong></td>" 'Incident Exact Time
206          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtHour.Text & ":" & .txtMinutes.Text & "</strong></td>" 'Incident Exact Time

208          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>Incident Time:</strong></td>" 'Incident Time
210          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & .txtHour.Text & ":" & .txtMinutes.Text & "</strong></td></tr>" 'Incident Time

    '         sHTML = sHTML & vbCrLf & "<tr><td width=""""25%""""><strong>" & ",True" '& CBool(txtWizValue(16).Text) 'Attachments

212          sHTML = sHTML & vbCrLf & "<tr bgcolor=""#FFFFC0""><td width=""""25%""""><strong>Verify When published:</strong></td>" 'Verify When published"
214          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & IIf(.chkSendMe.value = vbChecked, "Yes", "No") & "</strong></td>" 'Verify When published"

216          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>Allow Other Users To Contact me:</strong></td>" 'Allow Contact"
218          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & IIf(.chkAllowOther.value = vbChecked, "Yes", "No") & "</strong></td></tr>" 'Allow Contact"
    
        End With


220     sHTML = sHTML & vbCrLf & "</table></body></html>"

222     Lf.WriteElt sHTML
224     Lf.CloseFile

226     WriteHTMLReport = sHTML 'g_sAppPath & "\data\user\Exports\OasisReport.html"
    
228     If bExecuteRpt Then
230         ShellExecute Me.hWnd, vbNullString, g_sAppPath & "\data\user\Exports\OasisReport.html", vbNullString, "C:\", 1
        End If
    
        'TxtRpt sHTML
    
        '<EhFooter>
        Exit Function

WriteHTMLReport_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.WriteHTMLReport " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub udRotationAngle_Change()
  ' calculate the angle for set value
  Exit Sub
  edRotationAngle.Text = udRotationAngle.value
  GIS10.RotationPoint.x = GIS10.CenterPtg.x
  GIS10.RotationPoint.y = GIS10.CenterPtg.y
  GIS10.RotationAngle = DegToRad(udRotationAngle.value)
  GIS10.UpDate
End Sub

Private Sub ListAllAnnotationShps()
        '<EhHeader>
        On Error GoTo ListAllShps_Err
        '</EhHeader>

        m_frmTextAnnoSettings.lstTexts.Clear

100     If m_oDrawLyr Is Nothing Then Exit Sub
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
        
102     With m_oDrawLyr
104         For Each oShp9 In .Loop(.Extent, "", Nothing, "", True)
        
                If Not g_lPinUID = oShp9.uID Then
110                 m_frmTextAnnoSettings.lstTexts.AddItem oShp9.Params.labels.value
112                 m_frmTextAnnoSettings.lstTexts.ItemData(m_frmTextAnnoSettings.lstTexts.ListCount - 1) = oShp9.uID
                    
                    On Error Resume Next
                    m_frmTextAnnoSettings.panColorBack.BackColor = .color
                    m_frmTextAnnoSettings.panColorFore.BackColor = .FontColor
                    On Error GoTo ListAllShps_Err
                End If

            Next
    
        End With
    
        If m_frmTextAnnoSettings.lstTexts.ListCount > 0 Then m_frmTextAnnoSettings.lstTexts.ListIndex = 0
    
        '<EhFooter>
        Exit Sub

ListAllShps_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.ListAllShps " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub RemoveAllShps(oLyr As TatukGIS_XDK10.XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo RemoveAllShps_Err
        '</EhHeader>
        Dim oShp9 As TatukGIS_XDK10.XGIS_Shape
        
100     For Each oShp9 In oLyr.Loop(oLyr.Extent, "", Nothing, "", False)
            If Not g_lPinUID = oShp9.uID Then
104             oLyr.Delete oShp9.GetField("GIS_UID")
            End If
        Next
    
        '<EhFooter>
        Exit Sub

RemoveAllShps_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.RemoveAllShps " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub AddAnnoText(ptg As TatukGIS_XDK10.XGIS_Point)
        '<EhHeader>
        On Error GoTo AddAnnoText_Err
        '</EhHeader>
        Dim oshp As TatukGIS_XDK10.XGIS_Shape

        If Not m_frmTextAnnoSettings.chkMultipleText.value = vbChecked Then
            RemoveAllShps m_oDrawLyr
           m_frmTextAnnoSettings.lstTexts.Clear
        End If
        
100     Set oshp = m_oDrawLyr.CreateShape(XgisShapeTypePoint)
                        
102     If oshp Is Nothing Then Exit Sub
                        
104     With oshp
106         .Lock TatukGIS_XDK10.XgisLockExtent
108         .AddPart
110         .AddPoint ptg

            With .Params.labels
                .rotate = m_frmTextAnnoSettings.txtRotation.Text
                .Alignment = TatukGIS_XDK10.XgisLabelAlignmentCenter
                .FontColor = m_frmTextAnnoSettings.panColorFore.BackColor
                .color = m_frmTextAnnoSettings.panColorBack.BackColor
112             .Allocator = False
114             .Duplicates = True
                
                If IsNumeric(m_frmTextAnnoSettings.comFontSize.Text) Then
116                 .Font.Size = m_frmTextAnnoSettings.comFontSize.Text
                Else
                    .Font.Size = 12
                End If
                
120             .OutlineWidth = 0
122             .Pattern = XbsClear
124             .Position = TatukGIS_XDK10.XgisLabelPositionMiddleCenter
126             .value = m_frmTextAnnoSettings.txtAnnoText.Text
            End With
                
              m_frmTextAnnoSettings.lstTexts.AddItem m_frmTextAnnoSettings.txtAnnoText.Text
127           m_frmTextAnnoSettings.lstTexts.ItemData(m_frmTextAnnoSettings.lstTexts.ListCount - 1) = .uID
              m_frmTextAnnoSettings.lstTexts.ListIndex = m_frmTextAnnoSettings.lstTexts.ListCount - 1
                
128         .Params.Marker.Size = 1
            .Unlock
        End With

130     GIS10.UpDate

        '<EhFooter>
        Exit Sub

AddAnnoText_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.AddAnnoText " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Public Sub InitSelector()
        '<EhHeader>
        On Error GoTo InitSelector_Err
        '</EhHeader>
        Dim i As Integer
        Dim sCurrItem As String
        Dim sCurFeatItm As String
        Dim keyArray() As Variant
        Dim element As Variant

100     Set m_SelLyrCol = New Collection

        m_udtSelectorSettings.bAutoClear = True
104     m_udtSelectorSettings.dBuffeLevel = CDbl(1 / 2)
        
106     If ComSelLayer.ListCount > 0 Then
        
108         sCurrItem = ComSelLayer.List(ComSelLayer.ListIndex)
        
        End If
        
110     If ComFeatureLayer.ListCount > 0 Then
112         sCurFeatItm = ComFeatureLayer.List(ComFeatureLayer.ListIndex)
        End If
        
        m_udtSelectorSettings.sSpatialOperator = DE9IM.Item("Intersect")
        
124     ComSelLayer.Clear
126     ComFeatureLayer.Clear
        
128     For i = 0 To GIS10.items.Count - 1
            On Error Resume Next
130         If GisUtils.IsInherited(GIS10.items.Item(i), "XGIS_LayerVector") Then
132             m_SelLyrCol.Add GIS10.items.Item(i).Name, GIS10.items.Item(i).caption
134             ComSelLayer.AddItem GIS10.items.Item(i).caption 'Name
136             ComFeatureLayer.AddItem GIS10.items.Item(i).caption
            End If
        Next
       
138     If ComSelLayer.ListCount > 0 Then
            
140         If Len(sCurrItem) > 0 Then
142             FindIndexStrEx ComSelLayer, sCurrItem
            Else
144             ComSelLayer.ListIndex = 0 'FindIndexStrEx ComSelLayer, "--All--"
            End If
            
        End If

146     If ComFeatureLayer.ListCount > 0 Then
148         If Len(sCurFeatItm) > 0 Then
150             FindIndexStrEx ComFeatureLayer, sCurFeatItm
            Else
152             ComFeatureLayer.ListIndex = 0
            End If
        End If

        ComFeatureLayer.Enabled = False
        '<EhFooter>
        Exit Sub

InitSelector_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.InitSelector " & _
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

Private Sub RemoveAlltabs()
        '<EhHeader>
        On Error GoTo RemoveAlltabs_Err
        '</EhHeader>
        Dim i As Integer

102     ReDim m_shpProps(0)
        
112     If AttribTabs.NumTabs < 1 Then Exit Sub
    
114     For i = AttribTabs.NumTabs To 1 Step -1

116         If Not i - 1 = 0 Then
118             Unload SelAttributes1(i - 1)
120             Unload elDynHolder(i - 1)
            Else
122             SelAttributes1(0).Clear
                elDynHolder.Item(0).Visible = False
            End If
        
124         AttribTabs.RemoveTab i - 1
        Next

   
        '<EhFooter>
        Exit Sub

RemoveAlltabs_Err:
        DebugPrint Err.Description & vbCrLf & _
               "in OASISClient.frmMain.RemoveAlltabs " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetGISSelectionStyle()

End Sub

Private Sub AddTabs(sCaption As String, _
                    Optional oshp As TatukGIS_XDK10.XGIS_Shape, _
                    Optional bShowGeo As Boolean = True)
        '<EhHeader>
        On Error GoTo AddTabs_Err
        '</EhHeader>
        
100     Set SelAttributes1(SelAttributes1.UBound).Container = elDynHolder(elDynHolder.UBound)
102     SelAttributes1(SelAttributes1.UBound).Visible = True
        SelAttributes1(SelAttributes1.UBound).ReadOnly = Not m_udtSelectorSettings.bEdit
    
104     If Not oshp Is Nothing Then
        
106         SelAttributes1(SelAttributes1.UBound).AllowRestructure = True
108         SelAttributes1(SelAttributes1.UBound).ShowShape oshp
    
        End If

112     If SelAttributes1.UBound > 0 Then
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
136     Load SelAttributes1(SelAttributes1.UBound + 1)
        If AttribTabs.NumTabs > 0 Then AttribTabs.CurrTab = 0
        '<EhFooter>
        Exit Sub

AddTabs_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.AddTabs " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub PrepareConvertorUTM_LL()
Dim strResult As String
Set objProj = New GeoMercator

    Exit Sub

    With oCoordTransSettings

        'Initialise the ellipsoid variables in the GeoMercator class module
        strResult = objProj.SetEllipsoid(.Semi_Major_Axis, .Inverse_Flattening, .Sphere)

        'If strResult is not a zero-length string, it means there was an error with the
        'initialisation. The error message will be contained in strResult, so display the error message.
        If Len(strResult) > 0 Then
            MsgBox strResult, vbExclamation, "Mercator"
            Exit Sub
        End If

        'Initialise the projection variables in the GeoMercator class module
        strResult = objProj.SetProjection(.False_Northing, .False_Easting, .Lat_Of_Origin, .Long_Of_Origin)

        'If strResult is not a zero-length string, it means there was an error with the
        'initialisation. The error message will be contained in strResult, so display the error message.
        If Len(strResult) > 0 Then
            MsgBox strResult, vbExclamation, "Mercator"
        End If
        
    End With

End Sub

Private Sub LL_2_UTM(x As Double, _
                     y As Double, _
                     dblNorthing As Double, _
                     dblEasting As Double, _
                     strMessage As String)
       
    dblNorthing = DegToRad(y)
      
    dblEasting = DegToRad(x)
        
    strMessage = objProj.Forward(dblNorthing, dblEasting)

End Sub

Private Sub UTM_2_LL(x As Double, y As Double, dblNorthing As Double, _
                     dblEasting As Double, _
                     strMessage As String)

            strMessage = objProj.Inverse(dblNorthing, dblEasting)
      
            y = RadToDeg(dblNorthing)
            x = RadToDeg(dblEasting)
End Sub

Private Sub ProjTest(x As Double, _
                     y As Double, _
                     Optional Zone As Integer, _
                     Optional bIsUTM As Boolean)

 
End Sub

Private Function MyPi1() As Double
MyPi1 = 4 * Atn(1)
End Function

Private Sub SetGridOption(gOption, _
                  chStatus)
    dxGISDataGrid.Option = gOption
    dxGISDataGrid.OptionEnabled = chStatus
End Sub

Private Sub tmrToolTip_Timer()
    Dim Pt As POINTAPI
    Dim mWnd As Long, WR As RECT

    picToolTip.Visible = False

    With oMapTipSetting

        If .Enabled Then

            'First Make Sure the PTTip is Initiated
            If ptTip.x + ptTip.y > 0 Then
                'Get Current Cursor Position to see if has changed since last time
                GetCursorPos Pt
         
                If Pt.x - ptTip.x = 0 And Pt.y - ptTip.y = 0 Then
         
                    If Not m_oToolTipSHP Is Nothing Then
                        mWnd = WindowFromPoint(ptTip.x, ptTip.y)
                        'Get the window's position in Pixels
                        GetWindowRect mWnd, WR
                        
                        If .MapTipField = "UID" Then
                            lblToolTip.caption = m_oToolTipSHP.uID 'GetField("Name")
                        Else
                            lblToolTip.caption = m_oToolTipSHP.GetField(.MapTipField)
                        End If
                        
                        picToolTip.Move elMap.left + ScaleX(ptTip.x - WR.left + 7, vbPixels, vbTwips), ScaleY(ptTip.y - WR.top - 17, vbPixels, vbTwips), lblToolTip.Width + 140, lblToolTip.Height + 80

                        'DebugPrint "Pic:" & ScaleX(WR.Left + ptTip.X - 2, vbPixels, vbTwips) & " Win:" & Me.Left & " Mouse:" & ptTip.X & " Rect:" & WR.Left
                        picToolTip.BackColor = .TipColor
                        lblToolTip.BackColor = .TipColor
                        lblToolTip.ForeColor = .TextColor
                        picToolTip.BorderStyle = IIf(.TipBorder, 1, 0)
                        picToolTip.Visible = True
                        picToolTip.ZOrder 0
                    End If
                End If
        
            End If
        
        Else
            picToolTip.Visible = False
        End If
        
    End With
    
End Sub

'GEOTRANS Projection acronyms (used in the DK 8.x and IS 8.x versions):
'
'  GED -Geodetic(Unprojected)
'  GEC -Geocentic(Unprojected)
'  ALB - Albers Equal Area Conic
'  MER -Mercator
'  ROB -Robinson
'  bOn -Bonne
'  CAS -Cassini
'  CEA - Cylindrical Equal Area
'  TCE - Transverse Cylindrical Equal Area
'  EQC - Equidistant Cylindrical
'  MIL - Miller Cylindrical
'  MOL - Mollweide Pseudocylindrical
'  EC4 - Eckert IV
'  EC6 - Eckert VI
'  WA4 - Wagner IV
'  WA7 - Wagner VII
'  GRI - Van der Grinten
'  LAM - Lambert Conformal Conic
'  ORT -Orthographic
'  POL - Polar Stereographic
'  PLC -Polyconic
'  Sin -Sinusoid
'  TMR - Transverse Mercator
'  UTM - Universal Transverse Mercator (UTM) HOT - Hotine Oblique Mercator Azimuth Center
'  HT2 - Hotine Oblique Mercator Two Point NZMG - New Zealand Map
'
'Datum acronyms
'
'  WGE - World Geodetic System 1984
'  WGD - World Geodetic System 1972
'  ADI -m - ADINDAN, Mean
'  ADI -a - ADINDAN, Ethiopia
'  ADI -b - ADINDAN, Sudan
'  ADI -c - ADINDAN, Mali
'  ADI -d - ADINDAN, Senegal
'  ADI-E - ADINDAN, Burkina Faso
'  ADI -f - ADINDAN, Cameroon
'  AFG -AFGOOYE, Somalia
'  AIA - ANTIGUA ISLAND ASTRO 1943
'  AIN-A - AIN EL ABD 1970, Bahrain
'  AIN-B - AIN EL ABD 1970, Saudi Arabia
'  AMA - AMERICAN SAMOA 1962
'  ANO - ANNA 1 ASTRO 1965, Cocos Is.
'  ARF-M - ARC 1950, Mean
'  ARF-A - ARC 1950, Botswana
'  ARF-B - ARC 1950, Lesotho
'  ARF-C - ARC 1950, Malawi
'  ARF-D - ARC 1950, Swaziland
'  ARF-E - ARC 1950, Zaire
'  ARF-F - ARC 1950, Zambia
'  ARF-G - ARC 1950, Zimbabwe
'  ARF-H - ARC 1950, Burundi
'  ARS-M - ARC 1960, Kenya & Tanzania
'  ARS-A - ARC 1960, Kenya
'  ARS-B - ARC 1960, Tanzania
'  ASC - ASCENSION ISLAND 1958
'  ASM - MONTSERRAT ISLAND ASTRO 1958
'  ASQ - ASTRO STATION 1952, Marcus Is.
'  ATF - ASTRO BEACON E 1945, Iwo Jima
'  AUA - AUSTRALIAN GEODETIC 1966
'  AUG - AUSTRALIAN GEODETIC 1984
'  BAT -DJAKARTA, INDONESIA
'  BID -Bissau, Guinea - Bissau
'  BER - BERMUDA 1957, Bermuda Islands
'  BOO - BOGOTA OBSERVATORY, Columbia
'  BUR - BUKIT RIMPAH, Banka & Belitung
'  CAC - CAPE CANAVERAL, Fla & Bahamas
'  CAI - CAMPO INCHAUSPE 1969, Arg.
'  CAO - CANTON ASTRO 1966, Phoenix Is.
'  CAP - CAPE, South Africa
'  CAZ - CAMP AREA ASTRO, Camp McMurdo
'  CCD - S-JTSK, Czech Republic
'  CGE -CARTHAGE, Tunisia
'  CHI - CHATHAM ISLAND ASTRO 1971, NZ
'  CHU - CHUA ASTRO, Paraguay
'  COA - CORREGO ALEGRE, Brazil
'  DAL -DABOLA, Guinea
'  DID - DECEPTION ISLAND
'  DOB - GUX 1 ASTRO, Guadalcanal Is.
'  EAS - EASTER ISLAND 1967
'  ENW - WAKE-ENIWETOK 1960
'  EST -ESTONIA, 1937
'  EUR-M - EUROPEAN 1950, Mean (3 Param)
'  EUR-A - EUROPEAN 1950, Western Europe
'  EUR-B - EUROPEAN 1950, Greece
'  EUR-C - EUROPEAN 1950, Norway & Finland
'  EUR-D - EUROPEAN 1950, Portugal & Spain
'  EUR-E - EUROPEAN 1950, Cyprus
'  EUR-F - EUROPEAN 1950, Egypt
'  EUR-G - EUROPEAN 1950, England, Channel
'  EUR-H - EUROPEAN 1950, Iran
'  EUR-I - EUROPEAN 1950, Sardinia(Italy)
'  EUR-J - EUROPEAN 1950, Sicily(Italy)
'  EUR-K - EUROPEAN 1950, England, Ireland
'  EUR-L - EUROPEAN 1950, Malta
'  EUR-S - EUROPEAN 1950, Iraq, Israel
'  EUR-T - EUROPEAN 1950, Tunisia
'  EUS - EUROPEAN 1979
'  FAH -OMAN
'  FLO - OBSERVATORIO MET. 1939, Flores
'  FOT - FORT THOMAS 1955, Leeward Is.
'  GAA - GAN 1970, Rep. of Maldives
'  GEO - GEODETIC DATUM 1949, NZ
'  GIZ - DOS 1968, Gizo Island
'  GRA - GRACIOSA BASE SW 1948, Azores
'  GUA - GUAM 1963
'  GSE - GUNUNG SEGARA, Indonesia
'  HEN - HERAT NORTH, Afghanistan
'  HER - HERMANNSKOGEL, old Yugoslavia
'  HIT - PROVISIONAL SOUTH CHILEAN 1963
'  HJO - HJORSEY 1955, Iceland
'  HKD - HONG KONG 1963
'  HTN -HU - TZU - SHAN, Taiwan
'  IBE - BELLEVUE (IGN ), Efate Is.
'  IDN - INDONESIAN 1974
'  ind -b - INDIAN, Bangladesh
'  ind -i - INDIAN, India & Nepal
'  ind -P - INDIAN, Pakistan
'  INF-A - INDIAN 1954, Thailand
'  ING-A - INDIAN 1960, Vietnam 16N
'  ING-B - INDIAN 1960, Con Son Island
'  INH-A - INDIAN 1975, Thailand
'  IRL - IRELAND 1965
'  ISG - ISTS 061 ASTRO 1968, S Georgia
'  IST - ISTS 073 ASTRO 1969, Diego Garc
'  JOH - JOHNSTON ISLAND 1961
'  KAN - KANDAWALA, Sri Lanka
'  KEG - KERGUELEN ISLAND 1949
'  KEA - KERTAU 1948, W Malaysia & Sing.
'  KUS - KUSAIE ASTRO 1951, Caroline Is.
'  LCF - L.C. 5 ASTRO 1961, Cayman Brac
'  LEH -LEIGON, Ghana
'  LIB - LIBERIA 1964
'  LUZ -a - LUZON, Philippines
'  LUZ-B - LUZON, Mindanao Island
'  MAS -MASSAWA, Ethiopia
'  MER -MERCHICH, Morocco
'  MID - MIDWAY ASTRO 1961, Midway Is.
'  MIK - MAHE 1971, Mahe Is.
'  Min -a - MINNA, Cameroon
'  Min -b - MINNA, Nigeria
'  MOD - ROME 1940, Sardinia
'  MPO -m 'PORALOKO, Gabon
'  MVS - VITI LEVU 1916, Viti Levu Is.
'  NAH-A - NAHRWAN, Masirah Island (Oman)
'  NAH-B - NAHRWAN, United Arab Emirates
'  NAH-C - NAHRWAN, Saudi Arabia
'  NAP -NAPARIMA, Trinidad & Tobago
'  NAS-A - NORTH AMERICAN 1927, Eastern US
'  NAS-B - NORTH AMERICAN 1927, Western US
'  NAS-C - NORTH AMERICAN 1927, CONUS
'  NAS-D - NORTH AMERICAN 1927, Alaska
'  NAS-E - NORTH AMERICAN 1927, Canada
'  NAS-F - NORTH AMERICAN 1927, Alberta/BC
'  NAS-G - NORTH AMERICAN 1927, E. Canada
'  NAS-H - NORTH AMERICAN 1927, Man/Ont
'  NAS-I - NORTH AMERICAN 1927, NW Terr.
'  NAS-J - NORTH AMERICAN 1927, Yukon
'  NAS-L - NORTH AMERICAN 1927, Mexico
'  NAS-N - NORTH AMERICAN 1927, C. America
'  NAS-O - NORTH AMERICAN 1927, Canal Zone
'  NAS-P - NORTH AMERICAN 1927, Caribbean
'  NAS-Q - NORTH AMERICAN 1927, Bahamas
'  NAS-R - NORTH AMERICAN 1927, San Salvador
'  NAS-T - NORTH AMERICAN 1927, Cuba
'  NAS-U - NORTH AMERICAN 1927, Greenland
'  NAS-V - NORTH AMERICAN 1927, Aleutian E
'  NAS-W - NORTH AMERICAN 1927, Aleutian W
'  NAR-A - NORTH AMERICAN 1983, Alaska
'  NAR-B - NORTH AMERICAN 1983, Canada
'  NAR-C - NORTH AMERICAN 1983, CONUS
'  NAR-D - NORTH AMERICAN 1983, Mexico
'  NAR-E - NORTH AMERICAN 1983, Aleutian
'  NAR-H - NORTH AMERICAN 1983, Hawaiwi
' NSD - NORTH SAHARA 1959, Algeria
'  OEG - OLD EGYPTIAN 1907
'  OGB-M - ORDNANCE GB 1936, Mean (3 Para)
'  OGB-A - ORDNANCE GB 1936, England
'  OGB-B - ORDNANCE GB 1936, Eng., Wales
'  OGB-C - ORDNANCE GB 1936, Scotland
'  OGB-D - ORDNANCE GB 1936, Wales
'  OHA-M - OLD HAWAI'IAN (CC ), Mean
'  OHA-A - OLD HAWAI'IAN (CC ), Hawai'i
'  OHA-B - OLD HAWAI'IAN (CC ), Kauai
'  OHA-C - OLD HAWAI'IAN (CC ), Maui
'  OHA-D - OLD HAWAI'IAN (CC ), Oahu
'  OHI-M - OLD HAWAI'IAN (IN ), Mean
'  OHI-A - OLD HAWAI'IAN (IN ), Hawai'i
'  OHI-B - OLD HAWAI'IAN (IN ), Kauai
'  OHI-C - OLD HAWAI'IAN (IN ), Maui
'  OHI-D - OLD HAWAI'IAN (IN ), Oahu
'  PHA - AYABELLA LIGHTHOUSE, Bjibouti
'  PIT - PITCAIRN ASTRO 1967
'  PLN - PICO DE LAS NIEVES, Canary Is.
'  POS - PORTO SANTO 1936, Madeira Is.
'  PRP-A - PROV. S AMERICAN 1956, Bolivia
'  PRP-B - PROV. S AMERICAN 1956, N Chile
'  PRP-C - PROV. S AMERICAN 1956, S Chile
'  PRP-D - PROV. S AMERICAN 1956, Colombia
'  PRP-E - PROV. S AMERICAN 1956, Ecuador
'  PRP-F - PROV. S AMERICAN 1956, Guyana
'  PRP-G - PROV. S AMERICAN 1956, Peru
'  PRP-H - PROV. S AMERICAN 1956, Venez
'  PRP-M - PROV. S AMERICAN 1956, Mean
'  PTB - POINT 58, Burkina Faso & Niger
'  PTN - POINT NOIRE 1948
'  PUK - PULKOVO 1942, Russia
'  PUR - PUERTO RICO & Virgin Is.
'  QAT - QATAR NATIONAL
'  QUO - QORNOQ, South Greenland
'  REU - REUNION, Mascarene Is.
'  SAE - SANTO (DOS) 1965
'  SAO - SAO BRAZ, Santa Maria Is.
'  SAP - SAPPER HILL 1943, E Falkland Is
'  SAN-M - SOUTH AMERICAN 1969, Mean
'  SAN-A - SOUTH AMERICAN 1969, Argentina
'  SAN-B - SOUTH AMERICAN 1969, Bolivia
'  SAN-C - SOUTH AMERICAN 1969, Brazil
'  SAN-D - SOUTH AMERICAN 1969, Chile
'  SAN-E - SOUTH AMERICAN 1969, Colombia
'  SAN-F - SOUTH AMERICAN 1969, Ecuador
'  SAN-G - SOUTH AMERICAN 1969, Guyana
'  SAN-H - SOUTH AMERICAN 1969, Paraguay
'  SAN-I - SOUTH AMERICAN 1969, Peru
'  SAN-J - SOUTH AMERICAN 1969, Baltra
'  SAN-K - SOUTH AMERICAN 1969, Trinidad
'  SAN-L - SOUTH AMERICAN 1969, Venezuela
'  SCK -SCHWARZECK, Namibia
'  SGM - SELVAGEM GRANDE 1938, Salvage Is
'  SHB - ASTRO DOS 71/4, St. Helena Is.
'  SOA - SOUTH ASIA, Singapore
'  SPK-A - S-42 (PULKOVO 1942 ), Hungary
'  SPK-B - S-42 (PULKOVO 1942 ), Poland
'  SPK-C - S-41 (PK42) Former Czechoslov.
'  SPK-D - S-41 (PULKOVO 1942 ), Latvia
'  SPK-E - S-41 (PK 1942 ), Kazakhstan
'  SPK-F - S-41 (PULKOVO 1942 ), Albania
'  SPK-G - S-41 (PULKOVO 1942 ), Romania
'  SRL - SIERRA LEONE 1960
'  TAN - TANANARIVE OBSERVATORY 1925
'  TDC - TRISTAN ASTRO 1968
'  TIL - TIMBALAI 1948, Brunei & E Malay
'  TOY -a - TOKYO, Japan
'  TOY-B - TOKYO, South Korea
'  TOY -c - TOKYO, Okinawa
'  TOY -m - TOKYO, Mean
'  TRN - ASTRO TERN ISLAND (FRIG) 1961
'  VOI - VOIROL 1874, Algeria
'  VOR - VOIROL 1960, Algeria
'  WAK - WAKE ISLAND ASTRO 1952
'  YAC -YACARE, Uruguay
'  ZAN -ZANDERIJ, Suriname
'  EUR-7 - EUROPEAN 1950, Mean (7 Param)
'  OGB-7 - ORDNANCE GB 1936, Mean (7 Param)
'  CH1903 - Switzerland CH-1903



Private Sub TestToXport(sLyrName As String, sFile As String, bGPSExport As Boolean, bMapInfoTAB As Boolean, bGoogleKML As Boolean, bAutocadDWG As Boolean, bShapeFile As Boolean)

        Dim oLyr As XGIS_LayerVector
        Dim xportLyr As XGIS_LayerSHP
468     Set oLyr = GIS10.get(sLyrName)

470     If Not oLyr Is Nothing Then
                        
478         If bShapeFile Then
490             Set xportLyr = New XGIS_LayerSHP
496             xportLyr.Path = sFile & ".shp"
498             oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
500             xportLyr.SaveAll
            End If
                
502         If bAutocadDWG = vbChecked Then
518             Set xportLyr = New XGIS_LayerDXF
520             xportLyr.Path = sFile & ".dxf"
522             oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
524             xportLyr.SaveAll
            End If
                        
526         If bGoogleKML Then
542             Set xportLyr = New XGIS_LayerKML
544             xportLyr.Path = sFile & ".kml"
546             oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
548             xportLyr.SaveAll
            End If
                        
550         If bMapInfoTAB Then
566             Set xportLyr = New XGIS_LayerGML
568             xportLyr.Path = sFile & ".gml"
570             oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
572             xportLyr.SaveAll
            End If
                
574         If bGPSExport Then
588             Set xportLyr = New XGIS_LayerGPX
590             xportLyr.Path = sFile & ".gpx"
592             oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
594             xportLyr.SaveAll
            End If
                
        End If

End Sub







