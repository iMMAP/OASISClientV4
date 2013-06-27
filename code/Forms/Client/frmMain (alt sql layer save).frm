VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{13C02136-A0FB-4F12-9894-5D9DC001170C}#6.0#0"; "ctTree.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E760686B-BC9E-4802-9ECF-175FDF4062CE}#5.0#0"; "MAPX50.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{98C28BA7-605C-46CF-927B-9AA96B3230E0}#1.1#0"; "Scoring_Module.ocx"
Object = "{978D0654-1960-4687-8496-61E47D034439}#1.0#0"; "OASISWebBrowser.ocx"
Begin VB.Form frmMain 
   Caption         =   "s"
   ClientHeight    =   8415
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCommand6 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5160
      TabIndex        =   174
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAttachments 
      Caption         =   "Attachments"
      Height          =   495
      Left            =   4080
      TabIndex        =   173
      Top             =   4980
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSendMess 
      Caption         =   "SendMess"
      Height          =   255
      Left            =   4020
      TabIndex        =   172
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAttmtLister 
      Caption         =   "AttmtLister"
      Height          =   255
      Left            =   4020
      TabIndex        =   171
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   195
      Left            =   3960
      TabIndex        =   170
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
      TabIndex        =   167
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      Begin VB.Label lblToolTip 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   168
         Top             =   30
         Width           =   45
      End
   End
   Begin VB.Frame FraFrmConvert 
      Caption         =   "frmConvert"
      Height          =   765
      Left            =   13260
      TabIndex        =   96
      Top             =   960
      Visible         =   0   'False
      Width           =   3105
      Begin VB.TextBox txtUTMY 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   2460
         Width           =   2235
      End
      Begin VB.TextBox txtUTMX 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   166
         Top             =   2040
         Width           =   2235
      End
      Begin VB.TextBox txtDDY 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   165
         Top             =   1680
         Width           =   2235
      End
      Begin VB.TextBox txtDDX 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   164
         Top             =   1320
         Width           =   2235
      End
      Begin VB.Frame FraCoords 
         Caption         =   "Coords"
         Height          =   945
         Left            =   300
         TabIndex        =   98
         Top             =   300
         Width           =   2355
         Begin VB.TextBox txtYOrg 
            Height          =   315
            Left            =   390
            TabIndex        =   100
            Top             =   510
            Width           =   1815
         End
         Begin VB.TextBox txtXOrg 
            Height          =   285
            Left            =   390
            TabIndex        =   99
            Top             =   180
            Width           =   1785
         End
         Begin VB.Label lblY 
            AutoSize        =   -1  'True
            Caption         =   "Y:"
            Height          =   195
            Left            =   180
            TabIndex        =   102
            Top             =   510
            Width           =   150
         End
         Begin VB.Label lblX 
            AutoSize        =   -1  'True
            Caption         =   "X:"
            Height          =   195
            Left            =   210
            TabIndex        =   101
            Top             =   180
            Width           =   150
         End
      End
      Begin VB.CommandButton cmdDoConversion 
         Caption         =   "DoConversion"
         Height          =   285
         Left            =   1410
         TabIndex        =   97
         Top             =   150
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdShowSelectedInfo 
      Caption         =   "ShowSelectedInfo"
      Height          =   285
      Left            =   3960
      TabIndex        =   93
      Top             =   3900
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdTracker 
      Caption         =   "Tracker"
      Height          =   285
      Left            =   3960
      TabIndex        =   91
      Top             =   3600
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdLineSelect 
      Caption         =   "Selectors"
      Height          =   285
      Left            =   4020
      TabIndex        =   88
      Top             =   3360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdSingleSelect 
      Caption         =   "Single Select"
      Height          =   285
      Left            =   3960
      TabIndex        =   89
      Top             =   3180
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame FraRotation 
      Caption         =   "Rotation"
      Height          =   555
      Left            =   13260
      TabIndex        =   80
      Top             =   2340
      Visible         =   0   'False
      Width           =   915
      Begin VB.TextBox edRotationAngle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   60
         TabIndex        =   81
         Top             =   210
         Width           =   240
      End
      Begin MSComCtl2.UpDown udRotationAngle 
         Height          =   285
         Left            =   301
         TabIndex        =   82
         Top             =   210
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "edRotationAngle"
         BuddyDispid     =   196633
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
      TabIndex        =   78
      Top             =   810
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdJoinAttribute 
      Caption         =   "Join Attribute"
      Height          =   285
      Left            =   3960
      TabIndex        =   77
      Top             =   2700
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdCommand4 
      Caption         =   "Gen Edit lyr"
      Height          =   285
      Left            =   3960
      TabIndex        =   76
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
   Begin VB.Timer tmrDataPackCheck 
      Interval        =   10000
      Left            =   5985
      Top             =   3960
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
      TabIndex        =   71
      Top             =   1020
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton cmdCreateThematics 
      Caption         =   "Themat"
      Height          =   285
      Left            =   4020
      TabIndex        =   70
      Top             =   1260
      Visible         =   0   'False
      Width           =   1170
   End
   Begin C1SizerLibCtl.C1Elastic elW3 
      Height          =   5430
      Left            =   13140
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   7875
      _cx             =   13891
      _cy             =   9578
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
      BorderWidth     =   4
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
      GridRows        =   4
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmMain.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic elNav 
         Height          =   420
         Left            =   60
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2355
         Width           =   7755
         _cx             =   13679
         _cy             =   741
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
         BorderWidth     =   2
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
         _GridInfo       =   $"frmMain.frx":68AD
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic elCommands 
            Height          =   360
            Left            =   30
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   30
            Width           =   3795
            _cx             =   6694
            _cy             =   635
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
            Begin VB.CommandButton cmdExportData 
               Height          =   330
               Left            =   4005
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmMain.frx":68EA
               Style           =   1  'Graphical
               TabIndex        =   69
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   600
            End
            Begin VB.CommandButton cmdMoveLastWhat 
               Height          =   330
               Left            =   3330
               Picture         =   "frmMain.frx":6CDE
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdMoveNextWhat 
               Height          =   330
               Left            =   2655
               Picture         =   "frmMain.frx":D530
               Style           =   1  'Graphical
               TabIndex        =   52
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdMoveFirstWhat 
               Height          =   330
               Left            =   0
               Picture         =   "frmMain.frx":13D82
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdBackWhat 
               Height          =   330
               Left            =   667
               Picture         =   "frmMain.frx":1A5D4
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdEditWhat 
               Caption         =   "Edit"
               Height          =   330
               Left            =   2668
               TabIndex        =   49
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdDeleteWhat 
               Height          =   330
               Left            =   2001
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmMain.frx":20E26
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   600
            End
            Begin VB.CommandButton cmdAddWhat 
               Height          =   330
               Left            =   1334
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmMain.frx":21223
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   600
            End
         End
         Begin C1SizerLibCtl.C1Elastic elNavigator 
            Height          =   360
            Left            =   3825
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   3900
            _cx             =   6879
            _cy             =   635
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
            Begin VB.CommandButton cmdMoveNext 
               Height          =   330
               Left            =   3420
               Picture         =   "frmMain.frx":21605
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdMoveLastWher 
               Height          =   330
               Left            =   4230
               Picture         =   "frmMain.frx":27E57
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdMovePrevious 
               Height          =   330
               Left            =   694
               Picture         =   "frmMain.frx":2E6A9
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdMoveFirstWhere 
               Height          =   330
               Left            =   450
               Picture         =   "frmMain.frx":34EFB
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdAddWhere 
               Caption         =   "Add"
               Height          =   330
               Left            =   1388
               TabIndex        =   44
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdEditWhere 
               Caption         =   "Edit"
               Height          =   330
               Left            =   2082
               TabIndex        =   43
               Top             =   0
               Width           =   600
            End
            Begin VB.CommandButton cmdDeleteWhere 
               Caption         =   "Delete"
               Height          =   330
               Left            =   2776
               TabIndex        =   42
               Top             =   0
               Width           =   600
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic elW3Wiz 
         Height          =   2340
         Left            =   60
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   435
         Width           =   7755
         _cx             =   13679
         _cy             =   4128
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
         BorderWidth     =   3
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
         _GridInfo       =   $"frmMain.frx":3B74D
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.PictureBox ctrWhere1 
            BackColor       =   &H000000FF&
            Height          =   2250
            Left            =   45
            ScaleHeight     =   2190
            ScaleWidth      =   7605
            TabIndex        =   60
            Top             =   45
            Width           =   7665
         End
         Begin VB.PictureBox strWho1 
            BackColor       =   &H000000FF&
            Height          =   2250
            Left            =   45
            ScaleHeight     =   2190
            ScaleWidth      =   7605
            TabIndex        =   61
            Top             =   45
            Width           =   7665
         End
         Begin VB.PictureBox ctrWhat1 
            BackColor       =   &H000000FF&
            Height          =   2250
            Left            =   45
            ScaleHeight     =   2190
            ScaleWidth      =   7605
            TabIndex        =   62
            Top             =   45
            Width           =   7665
         End
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxW3DBGrid 
         Height          =   2595
         Left            =   60
         OleObjectBlob   =   "frmMain.frx":3B77F
         TabIndex        =   45
         Top             =   2775
         Width           =   7755
      End
      Begin VB.Label lblLblW3 
         Caption         =   "lblW3"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   60
         TabIndex        =   59
         Top             =   2775
         Width           =   7755
      End
   End
   Begin VB.CommandButton cmdThree 
      Caption         =   "Where?"
      Height          =   240
      Left            =   11745
      TabIndex        =   36
      Top             =   -90
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdTwo 
      Caption         =   "What?"
      Height          =   240
      Left            =   10665
      TabIndex        =   35
      Top             =   45
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdONE 
      Caption         =   "Who?"
      Height          =   240
      Left            =   9630
      TabIndex        =   34
      Top             =   -45
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCommand3 
      Caption         =   "Create gen ed lyr"
      Height          =   285
      Left            =   4020
      TabIndex        =   7
      Top             =   1500
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
      Left            =   9840
      TabIndex        =   1
      Top             =   2580
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   166723586
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
      Left            =   1290
      Top             =   30
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
            Picture         =   "frmMain.frx":3EC49
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3ED5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EE6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EF7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F091
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F1A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F2B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F3C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F4D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F5EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F6FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F80F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F921
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FA33
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FB45
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FC57
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FD69
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FE7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FF8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4009F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ActiveBar3LibraryCtl.ActiveBar3 AB 
      Align           =   1  'Align Top
      Height          =   8415
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11400
      _LayoutVersion  =   2
      _ExtentX        =   20108
      _ExtentY        =   14843
      _DataPath       =   ""
      Bands           =   "frmMain.frx":401B1
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
               Picture         =   "frmMain.frx":581FB
               Key             =   ""
               Object.Tag             =   "Adress"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5830D
               Key             =   ""
               Object.Tag             =   "Check"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5841F
               Key             =   ""
               Object.Tag             =   "Company"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58531
               Key             =   ""
               Object.Tag             =   "Count"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58643
               Key             =   ""
               Object.Tag             =   "Time"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58755
               Key             =   ""
               Object.Tag             =   "Man"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58867
               Key             =   ""
               Object.Tag             =   "Phone"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58979
               Key             =   ""
               Object.Tag             =   "Apple"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":58A8B
               Key             =   ""
               Object.Tag             =   "Salary"
            EndProperty
         EndProperty
      End
      Begin MapXLib.Map Map1 
         Height          =   435
         Left            =   13620
         TabIndex        =   163
         Top             =   840
         Visible         =   0   'False
         Width           =   615
         _Version        =   500012
         _ExtentX        =   1085
         _ExtentY        =   767
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
         Title.X         =   204
         Title.Y         =   28
         Map.NumericCoordSys.ProjectionInfo=   "frmMain.frx":58B9D
         Map.DisplayCoordSys.ProjectionInfo=   "frmMain.frx":58CCD
      End
      Begin C1SizerLibCtl.C1Elastic elDynamicReports 
         Height          =   855
         Left            =   4680
         TabIndex        =   160
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
         _GridInfo       =   $"frmMain.frx":58DFD
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin OASISClient.DynamicDataReports DynamicDataReports1 
            Height          =   855
            Left            =   0
            TabIndex        =   162
            Top             =   0
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1508
         End
      End
      Begin C1SizerLibCtl.C1Elastic elDynamicData 
         Height          =   705
         Left            =   4200
         TabIndex        =   159
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
         _GridInfo       =   $"frmMain.frx":58E31
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin OASISClient.DynamicDataModule DynamicDataModule1 
            Height          =   705
            Left            =   0
            TabIndex        =   161
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1244
         End
      End
      Begin C1SizerLibCtl.C1Elastic elAddons 
         Height          =   1725
         Left            =   2295
         TabIndex        =   58
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
         _GridInfo       =   $"frmMain.frx":58E65
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin OASISWebBrowser.Browser WebBrowser2 
            Height          =   1725
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   3043
         End
      End
      Begin C1SizerLibCtl.C1Elastic elOasisProfile 
         Height          =   5235
         Left            =   7200
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   -240
         Visible         =   0   'False
         Width           =   6600
         _cx             =   11642
         _cy             =   9234
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
         _GridInfo       =   $"frmMain.frx":58E97
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin OASISWebBrowser.Browser WebBrowser1 
            Height          =   4935
            Left            =   150
            TabIndex        =   75
            Top             =   150
            Width           =   6300
            _ExtentX        =   11113
            _ExtentY        =   8705
         End
      End
      Begin C1SizerLibCtl.C1Elastic elRSSTool 
         Height          =   7350
         Left            =   12120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   660
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
         _GridInfo       =   $"frmMain.frx":58ECE
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin OASISClient.OASISRSSBrowser RSSBrowser1 
            Height          =   7050
            Left            =   150
            TabIndex        =   169
            Top             =   150
            Visible         =   0   'False
            Width           =   5145
            _ExtentX        =   2778
            _ExtentY        =   3413
         End
      End
      Begin C1SizerLibCtl.C1Elastic elMap 
         Height          =   6780
         Left            =   5280
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   900
         Visible         =   0   'False
         Width           =   7935
         _cx             =   13996
         _cy             =   11959
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
         _GridInfo       =   $"frmMain.frx":58F05
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Tab C1TTab1Tab2 
            Height          =   2160
            Left            =   150
            TabIndex        =   120
            Top             =   4470
            Width           =   7635
            _cx             =   13467
            _cy             =   3810
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
            Caption         =   "Grid|Selector"
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
            TabHeight       =   400
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Begin C1SizerLibCtl.C1Elastic elSelector 
               Height          =   2070
               Left            =   8685
               TabIndex        =   127
               TabStop         =   0   'False
               Top             =   45
               Width           =   7140
               _cx             =   12594
               _cy             =   3651
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
               _GridInfo       =   $"frmMain.frx":58F4A
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic elSelTools 
                  Height          =   375
                  Left            =   15
                  TabIndex        =   142
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   7110
                  _cx             =   12541
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
                  Begin VB.ComboBox ComFeatureLayer 
                     Height          =   315
                     Left            =   6420
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   145
                     Top             =   0
                     Width           =   2535
                  End
                  Begin VB.ComboBox ComSelLayer 
                     Height          =   315
                     Left            =   1620
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   143
                     Top             =   0
                     Width           =   2115
                  End
                  Begin OASISClient.OASISButton OASSelect 
                     Height          =   315
                     Left            =   3720
                     TabIndex        =   146
                     Top             =   0
                     Width           =   1515
                     _ExtentX        =   2672
                     _ExtentY        =   556
                     ButtonStyle     =   1
                     Text            =   "Select Tool"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   204
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
                  Begin VB.Label lblFeatureLayer 
                     AutoSize        =   -1  'True
                     Caption         =   "Feature Layer:"
                     Height          =   195
                     Left            =   5340
                     TabIndex        =   147
                     Top             =   60
                     Width           =   1020
                  End
                  Begin VB.Label lblActiveInfo 
                     AutoSize        =   -1  'True
                     Caption         =   "Selection/Reporting:"
                     Height          =   195
                     Left            =   60
                     TabIndex        =   144
                     Top             =   60
                     Width           =   1470
                  End
               End
               Begin C1SizerLibCtl.C1Tab C1TTabSelector 
                  Height          =   1665
                  Left            =   15
                  TabIndex        =   128
                  Top             =   390
                  Width           =   7110
                  _cx             =   12541
                  _cy             =   2937
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
                  Begin TatukGIS_DK.XGIS_ControlAttributes GEO1 
                     Height          =   1635
                     Left            =   7725
                     TabIndex        =   129
                     Top             =   15
                     Width           =   7080
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
                  Begin C1SizerLibCtl.C1Elastic elAttribCloseHolder 
                     Height          =   1635
                     Left            =   15
                     TabIndex        =   151
                     TabStop         =   0   'False
                     Top             =   15
                     Width           =   7080
                     _cx             =   12488
                     _cy             =   2884
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
                     _GridInfo       =   $"frmMain.frx":58F8C
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin C1SizerLibCtl.C1Elastic elSelCloseTools 
                        Height          =   1635
                        Left            =   6660
                        TabIndex        =   155
                        TabStop         =   0   'False
                        Top             =   0
                        Width           =   420
                        _cx             =   741
                        _cy             =   2884
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
                           Picture         =   "frmMain.frx":58FCB
                           Style           =   1  'Graphical
                           TabIndex        =   73
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
                           Picture         =   "frmMain.frx":5F81D
                           Style           =   1  'Graphical
                           TabIndex        =   158
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
                           Picture         =   "frmMain.frx":6606F
                           Style           =   1  'Graphical
                           TabIndex        =   157
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
                           Picture         =   "frmMain.frx":6C8C1
                           Style           =   1  'Graphical
                           TabIndex        =   156
                           ToolTipText     =   "Close Current Info"
                           Top             =   0
                           Width           =   420
                        End
                     End
                     Begin C1SizerLibCtl.C1Tab AttribTabs 
                        Height          =   1635
                        Left            =   0
                        TabIndex        =   152
                        Top             =   0
                        Width           =   6660
                        _cx             =   11747
                        _cy             =   2884
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
                           Height          =   1305
                           Index           =   0
                           Left            =   45
                           TabIndex        =   153
                           TabStop         =   0   'False
                           Top             =   45
                           Width           =   6570
                           _cx             =   11589
                           _cy             =   2302
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
                           _GridInfo       =   $"frmMain.frx":73113
                           AccessibleName  =   ""
                           AccessibleDescription=   ""
                           AccessibleValue =   ""
                           AccessibleRole  =   9
                           Begin TatukGIS_DK.XGIS_ControlAttributes SelAttributes1 
                              Height          =   1125
                              Index           =   0
                              Left            =   90
                              TabIndex        =   154
                              Top             =   90
                              Width           =   6390
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
                  Begin C1SizerLibCtl.C1Elastic elSelReports 
                     Height          =   1635
                     Left            =   8025
                     TabIndex        =   130
                     TabStop         =   0   'False
                     Top             =   15
                     Width           =   7080
                     _cx             =   12488
                     _cy             =   2884
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
                        TabIndex        =   131
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
                           TabIndex        =   148
                           Top             =   480
                           Width           =   2445
                           Begin VB.ComboBox txtSpatialOperation 
                              Height          =   315
                              ItemData        =   "frmMain.frx":73153
                              Left            =   150
                              List            =   "frmMain.frx":7315A
                              Style           =   2  'Dropdown List
                              TabIndex        =   150
                              Top             =   540
                              Width           =   2175
                           End
                           Begin VB.ComboBox ComBuffLevel 
                              Height          =   315
                              ItemData        =   "frmMain.frx":73169
                              Left            =   150
                              List            =   "frmMain.frx":73185
                              Style           =   2  'Dropdown List
                              TabIndex        =   149
                              Top             =   210
                              Width           =   1335
                           End
                        End
                        Begin VB.CheckBox chkIncludeGeo 
                           Caption         =   "Include Geo ID"
                           Height          =   285
                           Left            =   120
                           TabIndex        =   139
                           Top             =   990
                           Width           =   1575
                        End
                        Begin VB.CommandButton cmdPrint 
                           Caption         =   "Print"
                           Height          =   285
                           Left            =   3780
                           TabIndex        =   138
                           Top             =   120
                           Width           =   1005
                        End
                        Begin VB.TextBox txtTitle 
                           Height          =   255
                           Left            =   1080
                           TabIndex        =   137
                           Top             =   120
                           Width           =   2655
                        End
                        Begin VB.CheckBox chkIncludeArea 
                           Caption         =   "Include Area"
                           Height          =   255
                           Left            =   120
                           TabIndex        =   136
                           Top             =   1260
                           Width           =   1305
                        End
                        Begin VB.CheckBox chkIncludeLength 
                           Caption         =   "Include Length"
                           Height          =   225
                           Left            =   1740
                           TabIndex        =   135
                           Top             =   1020
                           Width           =   1605
                        End
                        Begin VB.CheckBox chkIncludeCentroid 
                           Caption         =   "Include Centroid"
                           Height          =   285
                           Left            =   1740
                           TabIndex        =   134
                           Top             =   1260
                           Width           =   1485
                        End
                        Begin VB.TextBox txtMapTitle 
                           Height          =   315
                           Left            =   1020
                           TabIndex        =   133
                           Top             =   720
                           Width           =   2625
                        End
                        Begin VB.CheckBox chkIncludeMap 
                           Caption         =   "Include Map"
                           Height          =   285
                           Left            =   120
                           TabIndex        =   132
                           Top             =   420
                           Width           =   1245
                        End
                        Begin VB.Label lblTitle 
                           AutoSize        =   -1  'True
                           Caption         =   "Report Title:"
                           Height          =   195
                           Left            =   60
                           TabIndex        =   141
                           Top             =   180
                           Width           =   870
                        End
                        Begin VB.Label lblMapTitle 
                           AutoSize        =   -1  'True
                           Caption         =   "Map Title:"
                           Height          =   195
                           Left            =   90
                           TabIndex        =   140
                           Top             =   750
                           Width           =   705
                        End
                     End
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic elGisAttr 
               Height          =   2070
               Left            =   450
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   45
               Width           =   7140
               _cx             =   12594
               _cy             =   3651
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
               _GridInfo       =   $"frmMain.frx":731B1
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin ActiveBar3LibraryCtl.ActiveBar3 abGridTools 
                  Height          =   390
                  Left            =   0
                  TabIndex        =   122
                  Top             =   0
                  Width           =   7140
                  _LayoutVersion  =   2
                  _ExtentX        =   12594
                  _ExtentY        =   688
                  _DataPath       =   ""
                  Bands           =   "frmMain.frx":731EF
                  Begin VB.CheckBox chkFilterIn 
                     Caption         =   "Filter In Map"
                     Height          =   285
                     Left            =   5650
                     TabIndex        =   125
                     Top             =   75
                     Width           =   1185
                  End
                  Begin VB.CheckBox chkSelectIn 
                     Caption         =   "Select In Map"
                     Height          =   315
                     Left            =   6910
                     TabIndex        =   124
                     Top             =   60
                     Width           =   1305
                  End
                  Begin VB.CheckBox chkOnlyVisible 
                     Caption         =   "Load Visible"
                     Height          =   255
                     Left            =   4450
                     TabIndex        =   123
                     Top             =   90
                     Value           =   1  'Checked
                     Width           =   1215
                  End
               End
               Begin DXDBGRIDLibCtl.dxDBGrid dxGISDataGrid 
                  Height          =   1680
                  Left            =   0
                  OleObjectBlob   =   "frmMain.frx":74F11
                  TabIndex        =   126
                  Top             =   390
                  Width           =   7140
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elTatukGIS 
            Height          =   4290
            Left            =   150
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   150
            Width           =   7635
            _cx             =   13467
            _cy             =   7567
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
            Begin TatukGIS_DK.XGIS_ViewerWnd GIS 
               Height          =   4545
               Left            =   30
               TabIndex        =   13
               Top             =   -330
               Width           =   9240
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
               SelectionPattern=   "frmMain.frx":75B99
               SelectionTransparency=   100
               SelectionWidth  =   100
               SelectionOutlineOnly=   0   'False
               OldCachedPaint  =   0   'False
               PrinterModeDraft=   0   'False
               PrinterModeForceBitmap=   0   'False
               Mode            =   0
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
               ParentColor     =   0   'False
               ParentCtl3D     =   0   'False
               Object.Visible         =   -1  'True
               Cursor          =   16
               DoubleBuffered  =   0   'False
               ModeMouseButton =   0
               CursorForUserDefined=   0
            End
            Begin C1SizerLibCtl.C1Tab C1TabFastFunction 
               Height          =   3915
               Left            =   0
               TabIndex        =   14
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
                  TabIndex        =   63
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
                     TabIndex        =   115
                     Top             =   690
                     Width           =   2310
                     Begin MSMask.MaskEdBox me2Long 
                        Height          =   300
                        Left            =   990
                        TabIndex        =   116
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
                        TabIndex        =   117
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
                        TabIndex        =   119
                        Top             =   630
                        Width           =   750
                     End
                     Begin VB.Label Label1 
                        AutoSize        =   -1  'True
                        Caption         =   "Latitude:"
                        Height          =   195
                        Index           =   1
                        Left            =   135
                        TabIndex        =   118
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
                     TabIndex        =   113
                     Top             =   3030
                     Visible         =   0   'False
                     Width           =   2265
                     Begin VB.TextBox txtMGRS 
                        Height          =   285
                        Left            =   90
                        TabIndex        =   114
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
                     TabIndex        =   108
                     Top             =   1770
                     Visible         =   0   'False
                     Width           =   2265
                     Begin MSMask.MaskEdBox me1Long 
                        Height          =   330
                        Left            =   945
                        TabIndex        =   109
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
                        TabIndex        =   110
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
                        TabIndex        =   112
                        Top             =   315
                        Width           =   615
                     End
                     Begin VB.Label Label2 
                        AutoSize        =   -1  'True
                        Caption         =   "Longitude:"
                        Height          =   195
                        Index           =   0
                        Left            =   135
                        TabIndex        =   111
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
                     TabIndex        =   103
                     Top             =   2070
                     Visible         =   0   'False
                     Width           =   2265
                     Begin MSMask.MaskEdBox me3Long 
                        Height          =   330
                        Left            =   945
                        TabIndex        =   104
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
                        TabIndex        =   105
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
                        TabIndex        =   107
                        Top             =   315
                        Width           =   615
                     End
                     Begin VB.Label Label2 
                        AutoSize        =   -1  'True
                        Caption         =   "Longitude:"
                        Height          =   195
                        Index           =   2
                        Left            =   90
                        TabIndex        =   106
                        Top             =   630
                        Width           =   750
                     End
                  End
                  Begin VB.CommandButton cmdPolyEdit 
                     Height          =   330
                     Left            =   930
                     MaskColor       =   &H0000C000&
                     Picture         =   "frmMain.frx":75BFB
                     Style           =   1  'Graphical
                     TabIndex        =   95
                     ToolTipText     =   "Create scribble Area"
                     Top             =   3480
                     UseMaskColor    =   -1  'True
                     Width           =   450
                  End
                  Begin VB.CommandButton cmdZoomToSettings 
                     Caption         =   "..."
                     Height          =   255
                     Left            =   1800
                     TabIndex        =   94
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
                     TabIndex        =   85
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
                     TabIndex        =   84
                     Top             =   420
                     Width           =   255
                  End
                  Begin VB.CommandButton cmdRemove 
                     Height          =   330
                     Left            =   1410
                     MaskColor       =   &H000000FF&
                     Picture         =   "frmMain.frx":75F3D
                     Style           =   1  'Graphical
                     TabIndex        =   83
                     ToolTipText     =   "Remove Markers"
                     Top             =   3480
                     Width           =   450
                  End
                  Begin VB.CheckBox chkUseMarker 
                     Caption         =   "Marker"
                     Height          =   195
                     Left            =   90
                     TabIndex        =   79
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
                     TabIndex        =   67
                     Top             =   1650
                     Width           =   2265
                  End
                  Begin VB.ComboBox comConversionType 
                     Height          =   315
                     ItemData        =   "frmMain.frx":7C78F
                     Left            =   660
                     List            =   "frmMain.frx":7C79F
                     Style           =   2  'Dropdown List
                     TabIndex        =   66
                     Top             =   60
                     Width           =   1710
                  End
                  Begin VB.CommandButton cmdZoomTo 
                     Height          =   330
                     Left            =   1890
                     Picture         =   "frmMain.frx":7C7CB
                     Style           =   1  'Graphical
                     TabIndex        =   65
                     ToolTipText     =   "Zoom to/Create Marker"
                     Top             =   3480
                     Width           =   450
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Format:"
                     Height          =   195
                     Left            =   90
                     TabIndex        =   68
                     Top             =   150
                     Width           =   585
                  End
               End
               Begin C1SizerLibCtl.C1Elastic elMagnifier 
                  Height          =   3825
                  Left            =   4080
                  TabIndex        =   64
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
                  TabIndex        =   24
                  Top             =   45
                  Width           =   2475
                  Begin VB.CommandButton cmdLocationAnalysis 
                     Caption         =   "Location Analysis"
                     Height          =   240
                     Left            =   90
                     TabIndex        =   32
                     Top             =   3510
                     Visible         =   0   'False
                     Width           =   2235
                  End
                  Begin VB.OptionButton OptThemeBoundary 
                     Caption         =   "OASIS Security Cells"
                     Height          =   375
                     Index           =   2
                     Left            =   90
                     TabIndex        =   31
                     Top             =   1350
                     Visible         =   0   'False
                     Width           =   2220
                  End
                  Begin VB.CheckBox chkDynamicUpdate 
                     Caption         =   "Dynamic Update"
                     Height          =   420
                     Left            =   90
                     TabIndex        =   30
                     Top             =   1755
                     Visible         =   0   'False
                     Width           =   1605
                  End
                  Begin VB.OptionButton OptThemeBoundary 
                     Caption         =   "Urban Administrative Areas"
                     Height          =   555
                     Index           =   1
                     Left            =   90
                     TabIndex        =   29
                     Top             =   675
                     Visible         =   0   'False
                     Width           =   2190
                  End
                  Begin VB.OptionButton OptThemeBoundary 
                     Caption         =   "Administrave borders"
                     Height          =   375
                     Index           =   0
                     Left            =   90
                     TabIndex        =   28
                     Top             =   225
                     Value           =   -1  'True
                     Visible         =   0   'False
                     Width           =   2190
                  End
                  Begin VB.CommandButton cmdCommand1 
                     Caption         =   "Activate"
                     Height          =   465
                     Left            =   90
                     TabIndex        =   27
                     Top             =   2655
                     Visible         =   0   'False
                     Width           =   2235
                  End
                  Begin VB.CommandButton cmdSetScoring 
                     Caption         =   "Set Scoring"
                     Height          =   465
                     Left            =   90
                     TabIndex        =   26
                     Top             =   3240
                     Width           =   2235
                  End
                  Begin VB.CommandButton cmdTimeAnalysis 
                     Caption         =   "Time Analysis"
                     Height          =   240
                     Left            =   90
                     TabIndex        =   25
                     Top             =   3285
                     Visible         =   0   'False
                     Width           =   2235
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic1 
                     Height          =   3810
                     Left            =   0
                     TabIndex        =   72
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
                  TabIndex        =   15
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
                     TabIndex        =   16
                     Top             =   60
                     Width           =   480
                  End
                  Begin TreeViewLibCtl.ctTree ctTreeBookmrks 
                     Height          =   3225
                     Left            =   90
                     TabIndex        =   33
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
                     PictureClose    =   "frmMain.frx":8301D
                     PictureMinus    =   "frmMain.frx":83039
                     PictureOpen     =   "frmMain.frx":83055
                     PicturePlus     =   "frmMain.frx":83071
                     PictureLeaf     =   "frmMain.frx":8308D
                     TitleBackImage  =   "frmMain.frx":830A9
                     CheckPicDown    =   "frmMain.frx":830C5
                     CheckPicUp      =   "frmMain.frx":830E1
                     CheckPicDisabled=   "frmMain.frx":830FD
                     RadioPicDown    =   "frmMain.frx":83119
                     RadioPicUp      =   "frmMain.frx":83135
                     BackImage       =   "frmMain.frx":83151
                     MouseIcon       =   "frmMain.frx":8316D
                     NoFocusBackColor=   -2147483633
                     HeaderData      =   "frmMain.frx":83189
                     PicArray0       =   "frmMain.frx":831B1
                     PicArray1       =   "frmMain.frx":831CD
                     PicArray2       =   "frmMain.frx":831E9
                     PicArray3       =   "frmMain.frx":83205
                     PicArray4       =   "frmMain.frx":83221
                     PicArray5       =   "frmMain.frx":8323D
                     PicArray6       =   "frmMain.frx":83259
                     PicArray7       =   "frmMain.frx":83275
                     PicArray8       =   "frmMain.frx":83291
                     PicArray9       =   "frmMain.frx":832AD
                     PicArray10      =   "frmMain.frx":832C9
                     PicArray11      =   "frmMain.frx":832E5
                     PicArray12      =   "frmMain.frx":83301
                     PicArray13      =   "frmMain.frx":8331D
                     PicArray14      =   "frmMain.frx":83339
                     PicArray15      =   "frmMain.frx":83355
                     PicArray16      =   "frmMain.frx":83371
                     PicArray17      =   "frmMain.frx":8338D
                     PicArray18      =   "frmMain.frx":833A9
                     PicArray19      =   "frmMain.frx":833C5
                     PicArray20      =   "frmMain.frx":833E1
                     PicArray21      =   "frmMain.frx":833FD
                     PicArray22      =   "frmMain.frx":83419
                     PicArray23      =   "frmMain.frx":83435
                     PicArray24      =   "frmMain.frx":83451
                     PicArray25      =   "frmMain.frx":8346D
                     PicArray26      =   "frmMain.frx":83489
                     PicArray27      =   "frmMain.frx":834A5
                     PicArray28      =   "frmMain.frx":834C1
                     PicArray29      =   "frmMain.frx":834DD
                     PicArray30      =   "frmMain.frx":834F9
                     PicArray31      =   "frmMain.frx":83515
                     PicArray32      =   "frmMain.frx":83531
                     PicArray33      =   "frmMain.frx":8354D
                     PicArray34      =   "frmMain.frx":83569
                     PicArray35      =   "frmMain.frx":83585
                     PicArray36      =   "frmMain.frx":835A1
                     PicArray37      =   "frmMain.frx":835BD
                     PicArray38      =   "frmMain.frx":835D9
                     PicArray39      =   "frmMain.frx":835F5
                     PicArray40      =   "frmMain.frx":83611
                     PicArray41      =   "frmMain.frx":8362D
                     PicArray42      =   "frmMain.frx":83649
                     PicArray43      =   "frmMain.frx":83665
                     PicArray44      =   "frmMain.frx":83681
                     PicArray45      =   "frmMain.frx":8369D
                     PicArray46      =   "frmMain.frx":836B9
                     PicArray47      =   "frmMain.frx":836D5
                     PicArray48      =   "frmMain.frx":836F1
                     PicArray49      =   "frmMain.frx":8370D
                     PicArray50      =   "frmMain.frx":83729
                     PicArray51      =   "frmMain.frx":83745
                     PicArray52      =   "frmMain.frx":83761
                     PicArray53      =   "frmMain.frx":8377D
                     PicArray54      =   "frmMain.frx":83799
                     PicArray55      =   "frmMain.frx":837B5
                     PicArray56      =   "frmMain.frx":837D1
                     PicArray57      =   "frmMain.frx":837ED
                     PicArray58      =   "frmMain.frx":83809
                     PicArray59      =   "frmMain.frx":83825
                     PicArray60      =   "frmMain.frx":83841
                     PicArray61      =   "frmMain.frx":8385D
                     PicArray62      =   "frmMain.frx":83879
                     PicArray63      =   "frmMain.frx":83895
                     PicArray64      =   "frmMain.frx":838B1
                     PicArray65      =   "frmMain.frx":838CD
                     PicArray66      =   "frmMain.frx":838E9
                     PicArray67      =   "frmMain.frx":83905
                     PicArray68      =   "frmMain.frx":83921
                     PicArray69      =   "frmMain.frx":8393D
                     PicArray70      =   "frmMain.frx":83959
                     PicArray71      =   "frmMain.frx":83975
                     PicArray72      =   "frmMain.frx":83991
                     PicArray73      =   "frmMain.frx":839AD
                     PicArray74      =   "frmMain.frx":839C9
                     PicArray75      =   "frmMain.frx":839E5
                     PicArray76      =   "frmMain.frx":83A01
                     PicArray77      =   "frmMain.frx":83A1D
                     PicArray78      =   "frmMain.frx":83A39
                     PicArray79      =   "frmMain.frx":83A55
                     PicArray80      =   "frmMain.frx":83A71
                     PicArray81      =   "frmMain.frx":83A8D
                     PicArray82      =   "frmMain.frx":83AA9
                     PicArray83      =   "frmMain.frx":83AC5
                     PicArray84      =   "frmMain.frx":83AE1
                     PicArray85      =   "frmMain.frx":83AFD
                     PicArray86      =   "frmMain.frx":83B19
                     PicArray87      =   "frmMain.frx":83B35
                     PicArray88      =   "frmMain.frx":83B51
                     PicArray89      =   "frmMain.frx":83B6D
                     PicArray90      =   "frmMain.frx":83B89
                     PicArray91      =   "frmMain.frx":83BA5
                     PicArray92      =   "frmMain.frx":83BC1
                     PicArray93      =   "frmMain.frx":83BDD
                     PicArray94      =   "frmMain.frx":83BF9
                     PicArray95      =   "frmMain.frx":83C15
                     PicArray96      =   "frmMain.frx":83C31
                     PicArray97      =   "frmMain.frx":83C4D
                     PicArray98      =   "frmMain.frx":83C69
                     PicArray99      =   "frmMain.frx":83C85
                  End
                  Begin VB.Label lblGeoMarks 
                     Caption         =   "Geo Marks"
                     Height          =   180
                     Left            =   780
                     TabIndex        =   17
                     Top             =   60
                     Width           =   990
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Bookmarks 
                  Height          =   3825
                  Left            =   -3690
                  TabIndex        =   18
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
                     ItemData        =   "frmMain.frx":83CA1
                     Left            =   45
                     List            =   "frmMain.frx":83CD2
                     Style           =   2  'Dropdown List
                     TabIndex        =   22
                     Top             =   345
                     Width           =   2325
                  End
                  Begin VB.Frame FraDetails 
                     Caption         =   "Details"
                     Height          =   2220
                     Left            =   75
                     TabIndex        =   20
                     Top             =   720
                     Width           =   2295
                     Begin VB.TextBox txtThemeDescription 
                        Height          =   1875
                        Left            =   105
                        MultiLine       =   -1  'True
                        TabIndex        =   21
                        Text            =   "frmMain.frx":83E38
                        Top             =   270
                        Width           =   2100
                     End
                  End
                  Begin VB.CommandButton cmdActivateTheme 
                     Caption         =   "Activate Theme"
                     Height          =   360
                     Left            =   90
                     TabIndex        =   19
                     Top             =   3150
                     Width           =   2205
                  End
                  Begin VB.Label lblAvailableThemes 
                     Caption         =   "Available Themes:"
                     Height          =   270
                     Left            =   120
                     TabIndex        =   23
                     Top             =   90
                     Width           =   1500
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1EGOTO 
                  Height          =   3825
                  Left            =   3480
                  TabIndex        =   39
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
               TabIndex        =   86
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
               _GridInfo       =   $"frmMain.frx":83E4A
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
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   204
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
                  TabIndex        =   87
                  Top             =   0
                  Width           =   270
               End
            End
         End
      End
      Begin CONTROLSLibCtl.dxProgressBar dxProgressBar1 
         Height          =   495
         Left            =   5580
         TabIndex        =   92
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
         Caption         =   "OPS View"
      End
      Begin VB.Menu mnuMapSpes 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainSettings 
         Caption         =   "Settings"
      End
   End
   Begin VB.Menu mnuGridAction 
      Caption         =   "dxGridAction"
      Visible         =   0   'False
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
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowSelected 
         Caption         =   "Show Selected"
      End
      Begin VB.Menu mnuHideSelected 
         Caption         =   "Hide Selected"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadAll 
         Caption         =   "Load All"
      End
      Begin VB.Menu mnuVisibleExtent 
         Caption         =   "Load Only Visible"
      End
      Begin VB.Menu mnuGridSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowRecordCount 
         Caption         =   "Show Record Counts"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAutoSize 
         Caption         =   "Auto Size Field Header"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoZoom 
         Caption         =   "Auto Zoom"
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

'TODO Remove Below
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
Private WithEvents EventLayer As XGIS_LayerVector
Attribute EventLayer.VB_VarHelpID = -1
'Attribute EventLayer.VB_VarHelpID = -1
Private WithEvents m_frmOPSView As frmOPSView
Attribute m_frmOPSView.VB_VarHelpID = -1
'Attribute m_frmOPSView.VB_VarHelpID = -1
Private WithEvents m_frmOvMap As frmOVMap
Attribute m_frmOvMap.VB_VarHelpID = -1
'Attribute m_frmOvMap.VB_VarHelpID = -1
Private WithEvents m_frmAddSHPWiz As frmAddPointWZ
Attribute m_frmAddSHPWiz.VB_VarHelpID = -1
'Attribute m_frmAddSHPWiz.VB_VarHelpID = -1
Private WithEvents m_fmrAddIncident As frmAddIncident
Attribute m_fmrAddIncident.VB_VarHelpID = -1
'Attribute m_fmrAddIncident.VB_VarHelpID = -1
Private WithEvents m_frmMAModule As frmMAModule
Attribute m_frmMAModule.VB_VarHelpID = -1
'Attribute m_frmMAModule.VB_VarHelpID = -1
Private WithEvents m_frmAddWhere As frmAddWhere
Attribute m_frmAddWhere.VB_VarHelpID = -1
'Attribute m_frmAddWhere.VB_VarHelpID = -1
Private m_frmCannedReports As frmCannedReports
'Attribute m_frmCannedReports.VB_VarHelpID = -1
Private WithEvents m_frmW3Wizard As frmW3Wizard
Attribute m_frmW3Wizard.VB_VarHelpID = -1
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
Private WithEvents m_frmIOMJOC As frmIOMJOC
Attribute m_frmIOMJOC.VB_VarHelpID = -1
'Attribute m_frmIOMJOC.VB_VarHelpID = -1
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
Private WithEvents m_frmSpatialize As frmSpatialize
Attribute m_frmSpatialize.VB_VarHelpID = -1
'Attribute m_frmSpatialize.VB_VarHelpID = -1
Private WithEvents m_frmTextAnnoSettings As frmTextAnnoSettings
Attribute m_frmTextAnnoSettings.VB_VarHelpID = -1
'Attribute m_frmTextAnnoSettings.VB_VarHelpID = -1
Private WithEvents m_oClipboardViewer As cClipboardViewer
Attribute m_oClipboardViewer.VB_VarHelpID = -1
'Attribute m_oClipboardViewer.VB_VarHelpID = -1
Private WithEvents m_frmMainSettings As frmMainSettings
Attribute m_frmMainSettings.VB_VarHelpID = -1
'Attribute m_frmMainSettings.VB_VarHelpID = -1
Private WithEvents m_frmSelections As frmSelections
Attribute m_frmSelections.VB_VarHelpID = -1
'Attribute m_frmSelections.VB_VarHelpID = -1
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
Private WithEvents m_frmResourcesFinder As frmResourcesFinder
Attribute m_frmResourcesFinder.VB_VarHelpID = -1
Private m_frmWVISitrepGenerator As frmWVISitrepGenerator
Private WithEvents m_frmWVIIncidents As frmWVIIncidents
Attribute m_frmWVIIncidents.VB_VarHelpID = -1
Private WithEvents m_frmWVIControlPanel As frmWVIControlPanel
Attribute m_frmWVIControlPanel.VB_VarHelpID = -1
'Attribute m_frmSpatialiseDD.VB_VarHelpID = -1
Private ClipBoard_xt As cCustomClipboard

Private m_frmAttributes As frmAttributes
Private m_frmSearch As frmSearch
Private m_bSelectRadiusTool As Boolean
Private m_oCosmeticLayer As XGIS_LayerVector
'Private m_oDrawLyr As XGIS_LayerVector
Private m_oDrawLyr As XGIS_LayerSqlAdo
Private m_oBufferLyr As XGIS_LayerVector

'Both for spatialise DD tool
Private m_oSpatialiseLayer As XGIS_LayerVector
Private mDDLayer As XGIS_LayerVector


Private m_oIncidentLyr As XGIS_LayerVector
Private m_oW3Lyr As XGIS_LayerVector
Private m_oW3WHOLyr As XGIS_LayerVector
Private m_oSQLIncLyr As XGIS_LayerSqlAdo
Private m_oSQLOpsLyr As XGIS_LayerSqlAdo
Private m_oIncRS As adodb.Recordset
Private FirstRun As Boolean
Private g_oEditLayer As XGIS_LayerAbstract
Private ProjectionList As New XGIS_ProjectionList
Private m_oSQLGenericLyrs() As XGIS_LayerSqlAdo

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private m_prevTool As OASIS_TOOLS
Private m_bUseDistrictOnly As Boolean
Private m_oPtRubberStart As New XGIS_Point
Private m_oLineRubber As New XGIS_ShapeArc
Private m_bShowRubberBand As Boolean

Private m_bLOADING As Boolean

'OASIS FEATURES
Private m_oShpArc As XGIS_ShapeArc
Private m_oShpPoint As XGIS_ShapePoint
Private m_oShpPolygon As XGIS_ShapePolygon
Private m_oShpUnknown As XGIS_Shape
Private m_oIncShpPt As XGIS_ShapePoint

Private m_sAdmVal1 As String
Private m_sAdmVal2 As String
Private m_sAdmLoc As String

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
Private m_oXgisThemeWiz As New XGIS_ControlLegendVectorWiz
Private m_bThematicsDone As Boolean
Private m_sInternetCon As String
Private m_bInternetCon As Boolean
Private m_cUploadThreads() As clsThreads

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

Private m_bDLDataPacks As Boolean
Private RSDataPacks As New adodb.Recordset

'For the Geometry Edit
Dim vkControl As Boolean
Private menuPos As XGIS_Point

'For the GIS Grid
Dim GISGridRS As adodb.Recordset
Dim GlobalGISGridLayer As XGIS_LayerVector

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

Private m_PrevExt() As XGIS_Extent
Private m_bPrevActionUsed As Boolean

Private m_bScrib As Boolean
Private m_bPolyL As Boolean
Private m_bPolyG As Boolean
Private m_bRECT As Boolean

Private m_lyrFtrSelector As XGIS_LayerVector
Private m_lyrFtrTarget As XGIS_LayerVector

Private gdpts() As POINTAPI
Private ptsSel() As XGIS_Point
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
Private m_oToolTipSHP As XGIS_Shape
Private m_LyrCol As Collection
Private m_sIncidentIni As String
Private m_bAlreadyinitedProfile As Boolean

Private SecurityLayerDateFrom As Date
Private SecurityLayerDateTill As Date

Private Sub AB_BandOpen(ByVal Band As ActiveBar3LibraryCtl.Band, ByVal Cancel As ActiveBar3LibraryCtl.ReturnBool)



    If Band.Name = "SysCustomize" Then

        Cancel = True

    End If


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
Dim oLyr As XGIS_LayerVector
Dim shp As XGIS_Shape

    Set oLyr = GIS.get(sLayerName)
    
    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    Set shp = oLyr.FindFirst(oLyr.Extent, "GIS_UID = " & uID, Nothing, "", True)

    If Not shp Is Nothing Then
        GIS.VisibleExtent = shp.Extent
    End If
End Sub

Private Sub SelAutoSelect(uID As Long, sLayerName As String)
    Dim oLyr As XGIS_LayerVector
    Dim shp As XGIS_Shape

    Set oLyr = GIS.get(sLayerName)
    
    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    Set shp = oLyr.FindFirst(oLyr.Extent, "GIS_UID = " & uID, Nothing, "", True)

    If Not shp Is Nothing Then
        shp.Lock XgisLockExtent
        Set shp = shp.MakeEditable
        shp.IsSelected = True
        shp.Unlock
        shp.draw
        GIS.UpDate
    End If
End Sub

Private Sub chkFilterIn_Click()
        '<EhHeader>
        On Error GoTo chkFilterIn_Click_Err
        '</EhHeader>

        Dim FilterText As String
        
        Dim lStart As Long
        Dim lEnd As Long
        Dim lEndOld As Long
        Dim sField As String
        Dim sDate As String
        
        
100     FilterText = dxGISDataGrid.Filter.FilterText
        lEndOld = 1

        If Not FilterText = "" Then

            If InStr(FilterText, "Date]") > 0 Then
            
                
                lStart = InStr(FilterText, "Date]") + 5
                lStart = InStr(lStart, FilterText, "'", vbTextCompare) + 1
                lEnd = InStr(lStart + 1, FilterText, "'", vbTextCompare)
                sDate = Mid(FilterText, lStart, lEnd - lStart)
                FilterText = Replace(FilterText, sDate, Format(sDate, "yyyymmdd"), , 1, vbTextCompare)
                
                lStart = InStr(FilterText, "Date]") + 5
                lStart = InStrRev(FilterText, "[", lStart)
                lEndOld = lEnd
                lEnd = InStr(lStart + 1, FilterText, "]", vbTextCompare)
                sField = Mid(FilterText, lStart, lEnd - lStart + 1)
                FilterText = Replace(FilterText, sField, LCase("Format(" & sField & ", 'yyyymmdd')"), , 1, vbTextCompare)
                
            
            End If
            
             If InStr(lEndOld, FilterText, "Date]", vbTextCompare) > 0 Then
            
                lStart = InStr(lEnd, FilterText, "Date]") + 5
                lStart = InStr(lStart, FilterText, "'", vbTextCompare) + 1
                lEndOld = lEnd
                lEnd = InStr(lStart + 1, FilterText, "'", vbTextCompare)
                sDate = Mid(FilterText, lStart, lEnd - lStart)
                FilterText = Replace(FilterText, sDate, Format(sDate, "yyyymmdd"), , 1, vbTextCompare)
                
                lStart = InStr(lEndOld, FilterText, "Date]") + 5
                lStart = InStrRev(FilterText, "[", lStart)
                lEnd = InStr(lStart + 1, FilterText, "]", vbTextCompare)
                sField = Mid(FilterText, lStart, lEnd - lStart + 1)
                FilterText = Replace(FilterText, sField & " ", "Format(" & sField & ", 'yyyymmdd') ")
                
            
            End If
        
        End If

102     FilterText = Replace(FilterText, "[", "")
104     FilterText = Replace(FilterText, "]", "")

106     If Not chkFilterIn.Value = vbChecked Then FilterText = ""

108     If Not abGridTools.Tools.Item("comLyr").Text = "---Nothing---" Then

110         If abGridTools.Tools.Item("comLyr").Text = m_oSQLIncLyr.Name Then
            
112             LoadOASISIncidentsEX False, False
                
114             If Not Len(FilterText) = 0 Then

116                 If Not Len(m_oSQLIncLyr.scope) = 0 Then
118                     FilterText = "(" & FilterText & ") AND "
                    End If
                    
                End If
                
120             If Not Len(m_oSQLIncLyr.scope) = 0 Then
                
122                 FilterText = FilterText & "(" & m_oSQLIncLyr.scope & ")"
                
                End If
                
            End If
    
124         GIS.Lock
            Dim lL As XGIS_LayerVector
126         SafeMoveFirst g_RSGISGridTableSettings
128         g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"
                
130         If Not g_RSGISGridTableSettings.EOF Then
132             Set lL = GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
            Else
134             Set lL = GIS.get(abGridTools.Tools.Item("comLyr").Text)
            End If

136         DoEvents
138         lL.MoveFirst lL.Extent, "", Nothing, "", True

140         DoEvents
142         lL.scope = FilterText
144         lL.Params.Visible = True
146         GIS.UpDate
148         GIS.Unlock
    
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

Private Sub chkSelectIn_Click()
        '<EhHeader>
        On Error GoTo chkSelectIn_Click_Err
        '</EhHeader>

        Dim FilterText As String
100     FilterText = dxGISDataGrid.Filter.FilterText
102     FilterText = Replace(FilterText, "[", "")
104     FilterText = Replace(FilterText, "]", "")

106     If Not chkSelectIn.Value = vbChecked Then FilterText = ""

108     If Not abGridTools.Tools.Item("comLyr").Text = "---Nothing---" Then

110         GIS.Lock
112         Call mnuClearSelections_Click
            Dim lL As XGIS_LayerVector
114         SafeMoveFirst g_RSGISGridTableSettings
116         g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"
            
118         Set lL = GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
120         lL.DeselectAll
122         lL.MoveFirst lL.Extent, FilterText, Nothing, "", True
            
            Dim shp As XGIS_Shape
            
124         Set shp = lL.FindFirst(lL.Extent, FilterText, Nothing, "", True)

126         Do While Not shp Is Nothing
            
128             Set shp = shp.MakeEditable
130             If Not chkSelectIn.Value = vbChecked Then
132                 shp.IsSelected = False
                Else
134                 shp.IsSelected = True
                End If
136             Set shp = lL.FindNext
            
            Loop

138         GIS.Unlock
    
        End If

        '<EhFooter>
        Exit Sub

chkSelectIn_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.chkSelectIn_Click " & _
               "at line " & Erl
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
        
        Debug.Print "INSERT INTO Attachments (incidentID,AttachmentTable,Description,sGUID,FileType,FileSize,other) VALUES ('" & incidentID & "','" & AttachmentTable & "','" & Description & "','" & sGUID & "','" & FileType & "','" & FileSize & "')"
        m_Cnn.Execute "INSERT INTO Attachments (incidentID,AttachmentTable,Description,sGUID,FileType,FileSize,other) VALUES ('" & incidentID & "','" & AttachmentTable & "','" & Description & "','" & sGUID & "','" & FileType & "','" & FileSize & "','" & subdomain & "')"

    End If

End Sub

Private Sub cmdAttmtLister_Click()
    frmAttachmentUploader.Show vbModal, Me
    frmAttachmentViewer.Show
End Sub

Private Sub cmdCommand6_Click()
        frmMapPrint.Init GIS.viewer, m_frmMnuOperations.Legend1.Legend, g_bMapTplUpdate, g_sRemoteTablePrefix
        frmMapPrint.Show vbModeless, Me
End Sub

Private Sub cmdSearch_Click()

    If m_frmSearch Is Nothing Then
        Set m_frmSearch = New frmSearch
    End If
                
    m_frmSearch.Init GIS.viewer
                
    m_frmSearch.Show vbModeless, Me
End Sub

Private Sub cmdSendMess_Click()
    frmSendMess.Show
End Sub



Private Sub dxGISDataGrid_OnShowCellTip(ByVal Node As DXDBGRIDLibCtl.IdxGridNode, _
                                        ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, _
                                        TipText As String, _
                                        l As Single, _
                                        T As Single, _
                                        r As Single, _
                                        b As Single, _
                                        NeedShowTip As Boolean)
    
    Dim i As Integer
    
    On Error Resume Next

    If Node.HasChildren Then
        TipText = "Grouping:" & vbNewLine & Node.Strings(0)
    Else
        TipText = ""

        For i = 1 To Node.ValuesCount - 1
            TipText = TipText & Node.values(i) & vbNewLine
        Next

    End If

End Sub

Private Sub DynamicDataModule1_GetSpatialLoc(oLayer As XGIS_LayerVector)



    m_prevTool = oPan ' g_CurrentTool
    elMap.Visible = True
    AB.ClientAreaControl = elMap
    elDynamicData.Visible = False
    
    AB.Enabled = False
'    GIS.Mode = XgisSelect

    If Not m_frmSpatialiseDD Is Nothing Then
        If Not m_frmSpatialiseDD.Visible Then
            Set m_frmSpatialiseDD = New frmSpatialiseDD
        End If
    Else
        Set m_frmSpatialiseDD = New frmSpatialiseDD
    End If
    
    Set mDDLayer = oLayer
    m_frmSpatialiseDD.Show vbModeless, Me
    m_frmSpatialiseDD.SetLayer mDDLayer
    
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

Private Sub m_frmMainSettings_GetFields(sName As String)
    Dim i As Integer
    Dim oVec As XGIS_LayerVector

    With m_frmMainSettings
        .ComTipField.Clear
        
        Set oVec = GIS.get(m_LyrCol.Item(sName))
        
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

Private Sub m_frmResourcesFinder_GetResources()

    If Not g_CurrentTool = oCreateLocationPoint Then
        m_prevTool = g_CurrentTool
        g_CurrentTool = oCreateLocationPoint
    
    End If

    If Not g_CurrentFeatureType = 666 Then
        g_CurrentFeatureType = 666
    End If

    GIS.Mode = XgisSelect

End Sub

Private Sub m_frmResourcesFinder_GetGISImage()

    GIS.viewer.PrintClipboard
     
End Sub

Public Sub m_frmSpatialize_RestoreDDWindow()

    elMap.Visible = False
    elDynamicData.Visible = True
    AB.ClientAreaControl = elDynamicData
    AB.Enabled = True
    
End Sub

Private Sub m_frmSpatialiseDD_FlashIt()

    If m_oSpatialiseLayer.Items.Count > 0 Then
        m_oSpatialiseLayer.GetShape(m_oSpatialiseLayer.GetLastUid).Flash
    Else
        MsgBox "there is no shape to flash!"
    End If

End Sub

Private Sub m_frmSpatialiseDD_PanLeftRight(bLeft As Boolean)
    Dim dDiffer As Double
    Dim eExtent As New XGIS_Extent
    
    With GIS.VisibleExtent
    
        dDiffer = GIS.VisibleExtent.YMax - GIS.VisibleExtent.YMin
        
        If bLeft Then
        
            eExtent.Prepare .XMin - dDiffer / 4, .YMin, .XMax - dDiffer / 4, .YMax
        
        Else
        
            eExtent.Prepare .XMin + dDiffer / 4, .YMin, .XMax + dDiffer / 4, .YMax
        
        End If
    
    End With
    
    GIS.VisibleExtent = eExtent
End Sub

Private Sub m_frmSpatialiseDD_PanUpDown(bUp As Boolean)

    Dim dDiffer As Double
    Dim eExtent As New XGIS_Extent
    
    With GIS.VisibleExtent
    
        dDiffer = GIS.VisibleExtent.YMax - GIS.VisibleExtent.YMin
        
        If Not bUp Then
        
            eExtent.Prepare .XMin, .YMin - dDiffer / 4, .XMax, .YMax - dDiffer / 4
        
        Else
        
            eExtent.Prepare .XMin, .YMin + dDiffer / 4, .XMax, .YMax + dDiffer / 4
        
        End If
    
    End With
    
    GIS.VisibleExtent = eExtent

End Sub

Private Sub m_frmSpatialiseDD_UpdateShape(oShape As XGIS_Shape)
    
    If Not oShape.IsInsideExtent(GIS.VisibleExtent, XgisInsideTypeFull) Then
        GIS.VisibleExtent = oShape.Extent
        GIS.zoom = GIS.zoom / 3
    End If
    
    If m_oSpatialiseLayer Is Nothing Then
        Set m_oSpatialiseLayer = New XGIS_LayerVector
        m_oSpatialiseLayer.Name = GUIDGen
    End If
    
    If GIS.get(m_oSpatialiseLayer.Name) Is Nothing Then
        GIS.Add m_oSpatialiseLayer
    End If
    
    Do Until Not m_oSpatialiseLayer.Items.Count > 0
        m_oSpatialiseLayer.Delete m_oSpatialiseLayer.GetLastUid
        m_oSpatialiseLayer.SaveData
    Loop
    
    mDDLayer.SaveAll
    m_oSpatialiseLayer.SaveAll
    Set oShape = mDDLayer.AddShape(oShape, True)
    mDDLayer.SaveAll
    UpdateSpatialiseDDList oShape

End Sub

Private Sub m_frmSpatialiseDD_ZoomIn(bZoomIn As Boolean)

    If bZoomIn Then
        GIS.zoom = GIS.zoom * 2
    Else
        GIS.zoom = GIS.zoom / 2
    End If

End Sub

Private Sub m_frmSpatialiseDD_CreateObject(sObjectType As String)

    If m_oSpatialiseLayer Is Nothing Then
        Set m_oSpatialiseLayer = New XGIS_LayerVector
        m_oSpatialiseLayer.Name = GUIDGen
    Else
        m_oSpatialiseLayer.SaveAll
    End If
    
    If GIS.get(m_oSpatialiseLayer.Name) Is Nothing Then
        GIS.Add m_oSpatialiseLayer
    End If
    
    Do Until Not m_oSpatialiseLayer.Items.Count > 0
        m_oSpatialiseLayer.Delete m_oSpatialiseLayer.GetLastUid
        m_oSpatialiseLayer.SaveData
    Loop

    ReDim ptsSel(0)
    ReDim gdpts(0)

    'GIS.Mode = XgisSelect
    'GIS.Editor.endEdit
    'endEdit
    
    Select Case sObjectType

        Case "Point"
        
            g_CurrentTool = oCreateLocationPoint 'oCreateLocationPoint
            g_CurrentFeatureType = Custom
            m_bPolyL = False
            m_bPolyG = False
            SetOASISTool

        Case "Polyline"
        
            GIS.Mode = XgisEdit
            g_CurrentTool = oLineSelect
            m_bPolyL = True
            m_bPolyG = False
            SetOASISTool
        
        Case "Polygon"

            GIS.Mode = XgisEdit ' XgisUserDefined
            g_CurrentTool = oAreaSelect
            m_bPolyL = False
            m_bPolyG = True
            SetOASISTool

        Case Else
            
            GIS.Mode = XgisDrag
            'g_CurrentTool = oZoom ' m_prevTool
            'g_CurrentFeatureType = Custom
            'GIS.Mode = XgisZoom
            'm_bPolyL = False
            'm_bPolyG = False
            endEdit
            
    End Select
    
End Sub

Public Sub m_frmSpatialiseDD_RestoreDDWindow(bCommitShape As Boolean)
    
    If Not m_oSpatialiseLayer Is Nothing Then
        GIS.Delete m_oSpatialiseLayer.Name
        Set m_oSpatialiseLayer = Nothing
    End If
    
    endEdit
    m_bPolyG = False
    m_bPolyL = False
    g_CurrentTool = m_prevTool
    
    elMap.Visible = False
    elDynamicData.Visible = True
    AB.ClientAreaControl = elDynamicData
    AB.Enabled = True

    If Not bCommitShape Then
        Set mDDLayer = Nothing
        DynamicDataModule1.SaveChanges False, Nothing, m_frmSpatialiseDD.GetKMLForShape
    Else
        DynamicDataModule1.SaveChanges True, mDDLayer, m_frmSpatialiseDD.GetKMLForShape
        
    End If
    
End Sub

Private Sub m_frmMnuDynamicDataModule_DDListClicked()
    DynamicDataModule1.ListDataElements_Click
End Sub

Private Sub m_frmMnuDynamicReportsModule_DRComboClicked()
    DynamicDataReports1.cmbDatabases_Click
End Sub

Private Sub m_frmMnuDynamicReportsModule_DRListClicked()
    DynamicDataReports1.listQueries_Click
End Sub


Public Sub m_frmMnuDynamicDataModule_DDComboClicked()
    DynamicDataModule1.ListDatabases_Click ' ComboDatabases_Click
End Sub

Private Sub cmdDoConversion_Click()
    Dim oPt As XGIS_Point
    Dim oLyr As XGIS_LayerAbstract
    

    
    Debug.Print ""
End Sub

Private Sub cmdLineSelect_Click()
        '<EhHeader>
        On Error GoTo cmdLineSelect_Click_Err
        '</EhHeader>
                
        If Not m_frmSelections.Visible Then
            m_frmSelections.Show vbModeless, Me
            m_frmSelections.Init GIS
        End If
        
        'oPolyLineSelect, oPolySelect, oRectSelect, oMultiSelect
        '<EhFooter>
        Exit Sub

cmdLineSelect_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.cmdLineSelect_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
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
    GIS.Mode = XgisUserDefined
    g_CurrentTool = oAreaSelect
    m_bPolyG = True
End Sub

Private Sub cmdPolyEdit_Click()
' enmCurTool = GeoFencing

    DoPolyLineSelect

End Sub

Private Sub cmdSelectCircle_Click()
    g_CurrentTool = oRadiusSelect
    GIS.Mode = XgisSelect
End Sub

Private Sub cmdMSGScroller_Click()

    If cmdMSGScroller.caption = "6" Then 'Expand
        cmdMSGScroller.caption = "5"
        elScroller.Height = 255
        elScroller.Width = elTatukGIS.Width
        If Not mnuOPSView.Checked Then GIS.Height = elTatukGIS.Height + 150
        MsgScroll.AutoScroll = True
    Else
        cmdMSGScroller.caption = "6"
        elScroller.Height = 255
        elScroller.Width = cmdMSGScroller.Width
        MsgScroll.AutoScroll = False
        If Not mnuOPSView.Checked Then GIS.Height = elTatukGIS.Height + 400
    End If

End Sub

Private Sub cmdRemove_Click()
    If g_lPinUID > 0 Then
        m_oDrawLyr.Delete g_lPinUID
        GIS.UpDate
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
    Dim oLyr As XGIS_LayerVector
    Dim shp As XGIS_Shape

    Dim aTemp() As ShpProps
    Dim i As Long
    
    Select Case Index
    
        Case 0
            'Close
            If AttribTabs.NumTabs = 0 Then
                 elDynHolder(0).Visible = False
                Exit Sub
            End If
            
            For i = 0 To UBound(m_shpProps)

                If AttribTabs.CurrTab = 0 And i = 0 Then ReDim aTemp(0)

                If i <> AttribTabs.CurrTab Then
                    If i = 0 Then
                        ReDim aTemp(0)
                    Else
                        ReDim Preserve aTemp(UBound(aTemp) + 1)
                    End If
                    
                    aTemp(UBound(aTemp)) = m_shpProps(i)
                End If

            Next i
            
            m_shpProps = aTemp
            
            AttribTabs.RemoveTab AttribTabs.CurrTab
            
            If AttribTabs.NumTabs = 0 Then
                elDynHolder(0).Visible = False
            Else
                AttribTabs.CurrTab = AttribTabs.NumTabs - 1
            End If

        Case 1

            If Not frmGeoSummary.Visible Then
                frmGeoSummary.Show vbModeless, Me
            End If
 
            Set oLyr = GIS.get(m_shpProps(AttribTabs.CurrTab).sLayerName)
            
            If Not oLyr Is Nothing Then
                frmGeoSummary.SetGeoSum oLyr
                frmGeoSummary.SetFocus
            End If
 
        Case 2
            m_frmSelectorReports.Show vbModal, Me

        Case 3

            With m_udtSelectorSettings
                m_frmSelectorSettings.Init .dBuffeLevel, .sSpatialOperator, .bAutoZoom, .bAutoSelect, .bAutoFlash, .bAutoClear, .bEdit
                m_frmSelectorSettings.Move cmdSelTool(0).Container.Left$, cmdSelTool(0).Container.Top
                m_frmSelectorSettings.Show vbModal, Me
            End With

    End Select

End Sub

Private Sub cmdShowSelectedInfo_Click()

    SafeMoveFirst g_RSGISGridTableSettings

    If Not g_RSGISGridTableSettings.EOF Then

        g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"

        m_frmAttributes.GEO1.ShowSelected GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
    
    End If
    
End Sub

Private Sub cmdSingleSelect_Click()
    GIS.Mode = XgisSelect
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
'134             sString = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT * FROM oincidents WHERE " & sSQL)
'                'Set oRSServer = OpenSilentHttpCommsRS(sString, True)
'136             Set oRSServer = New ADODB.Recordset
'138             oRSServer.Open sString
'
'            Else
'140             sString = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT * FROM oincidents")
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

Private Sub elDynHolder_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print ""
End Sub

Private Sub Form_Resize()
'Stop

'Debug.Print GIS.Height
'
'
'
'Dim TotPos As Long
'
'Dim obj As Object
'
'
'    Debug.Print cmdSelX.Top
'
'    Set obj = AttribTabs.Container
'    TotPos = TotPos + obj.Top
'    On Error Resume Next
'
'    Do While Not obj Is Nothing
'        Debug.Print obj.Left & obj.Name
'        TotPos = TotPos + obj.Top
'
'        If obj.Name = Me.Name Then
'           ' Debug.Print TotPos '+ AttribTabs.Width
'           ' cmdSelX.Top = TotPos
'           ' cmdSelX.Refresh
'           cmdSelX.Move Me.Width - 800, TotPos 'GIS.Height
'            Exit Sub
'        End If
'        Set obj = obj.Container
'    Loop
'
'
'
    
    
    
    
End Sub

Private Sub FraPinColor_Click(Index As Integer)
Dim c As cCommonDialog
    Set c = New cCommonDialog
    c.ShowColor

    FraPinColor(Index).BackColor = c.Color

End Sub


Private Sub GIS_OnDblClick(translated As Boolean)
        
        Debug.Print "----- **DBL Click Start** ----"
        Debug.Print "gdpts:" & UBound(gdpts)
        Debug.Print "ptsSel:" & UBound(ptsSel)
        
    m_bScrib = False
    m_bRECT = False
    
    If m_bPolyG Or m_bPolyL Then
    
        RemoveAlltabs
        
        If UBound(gdpts) = 0 Then Exit Sub
        ReDim Preserve gdpts(UBound(gdpts) - 1)
        ReDim Preserve ptsSel(UBound(ptsSel) - 1)
        GIS.PaintMinimum
        
        If m_bPolyG Then
            m_SelUID = CreatePolygonEX
        '    DoPolyLineSelect
        Else
            m_SelUID = CreatePolyLine
        '    DoLineSelect
            
        End If
        
        m_bPolyG = False
        m_bPolyL = False
        endEdit
        
        ReDim gdpts(0)
        ReDim ptsSel(0)
        
        m_bPolyL = False
        m_bPolyG = False
        
        'GDI_Polygon GIS.hdc, gdpts
        m_bDrawFinished = True
        translated = True
    End If
    
        Debug.Print "----- **DblClick End** ----"
        Debug.Print "gdpts:" & UBound(gdpts)
        Debug.Print "ptsSel:" & UBound(ptsSel)
        
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
102         ResetScoring GIS.get(EventLayer.Name)
'104         DoEvents
'106         EventLayer.draw
'108         DoEvents
'110         Set EventLayer = Nothing
'112         DoEvents
114         GIS.UpDate
            DoEvents
            'Set EventLayer = Nothing
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
    
    SetParent GIS.hWnd, elTatukGIS.hWnd
    elTatukGIS.Refresh
    C1TabFastFunction.Width = 340
    GIS.Width = elTatukGIS.Width
    GIS.Height = elTatukGIS.Height
    GIS.ZOrder 0
    C1TabFastFunction.Width = 340
    C1TabFastFunction.Height = 3900
    C1TabFastFunction.ZOrder 0
    mnuOPSView.Checked = False
    
    On Error Resume Next
    Unload m_frmOPSView
    Set m_frmOPSView = Nothing
    Call elTatukGIS_RealignFrame
    
End Sub

Private Sub m_frmSelections_FlashInMap(uID As Long, sLayerName As String)
Dim oLyr As XGIS_LayerVector
Dim shp As XGIS_Shape

    Set oLyr = GIS.get(sLayerName)
    
    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    Set shp = oLyr.FindFirst(oLyr.Extent, "GIS_UID = " & uID, Nothing, "", True)

    If Not shp Is Nothing Then
        shp.Flash
    End If

End Sub

Private Sub m_frmSelections_GetShp(uID As Long, sLayerName As String, oShp As TatukGIS_DK.XGIS_Shape)
    Dim oLyr As XGIS_LayerVector

    Set oLyr = GIS.get(sLayerName)
    
    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    Set oShp = oLyr.FindFirst(oLyr.Extent, "GIS_UID = " & uID, Nothing, "", True)

End Sub

Private Sub GetSHP(uID As Long, sLayerName As String, oShp As TatukGIS_DK.XGIS_Shape)
    Dim oLyr As XGIS_LayerVector

    Set oLyr = GIS.get(sLayerName)
    
    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    Set oShp = oLyr.FindFirst(oLyr.Extent, "GIS_UID = " & uID, Nothing, "", True)
End Sub

Private Function GetNearestShape(PassedPoint As XGIS_Point, _
                                 sLayerName As String, Optional bFlash As Boolean = False) As Double
    
    Dim oLyr As XGIS_LayerVector
    Dim oPoint As XGIS_Point
    Dim oShp As XGIS_Shape
    Dim oDistance As Double
    Dim oPart As Long
    Dim oUtil As New XGIS_Utils
    Dim oShapeForPoint As New XGIS_Shape

    Set oLyr = GIS.get(sLayerName)
    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    Set oShp = oLyr.Locate(PassedPoint, 1000000000, True)
    oDistance = oShp.distance(PassedPoint, 1000000000)
    oDistance = oDistance * 111.3199
    If Not oLyr.Locate(PassedPoint, 5 / GIS.zoom, True) Is Nothing Then
    
        GetNearestShape = 0
    Else
        
        'oShp.Flash
        If bFlash Then oShp.Flash 6
        GetNearestShape = oDistance
    
    End If
    
End Function

Private Function GetDistanceToAllShapes(PassedPoint As XGIS_Point, _
                                        sLayerName As String, sDescField As String) As String
    
    Dim oLyr As XGIS_LayerVector
    Dim oPoint As XGIS_Point
    Dim oShp As XGIS_Shape
    Dim oDistance As Double
    Dim oPart As Long
    Dim oUtil As New XGIS_Utils
    Dim oShapeForPoint As New XGIS_Shape
    Dim sDistance As String
    Dim i As Long

    Set oLyr = GIS.get(sLayerName)
    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    Set oShp = oLyr.Shape
    
    Do Until oLyr.EOF
    
        If Not GetDistanceToAllShapes = "" Then GetDistanceToAllShapes = GetDistanceToAllShapes & "|||"
        oDistance = Round(oShp.distance(PassedPoint, 1000000000) / 1000, 0)
        sDistance = CStr(oDistance)
        If Len(sDistance) < 3 Then sDistance = "0" & sDistance
        If Len(sDistance) < 3 Then sDistance = "0" & sDistance
        If Len(sDistance) < 3 Then sDistance = "0" & sDistance
        If Len(sDistance) < 3 Then sDistance = "0" & sDistance
        GetDistanceToAllShapes = GetDistanceToAllShapes & sDistance & "km to " & oShp.GetField(sDescField) & ".   "
        
        i = 3
        Do Until i >= oLyr.Fields.Count
        
            GetDistanceToAllShapes = GetDistanceToAllShapes & "[" & Replace(oLyr.FieldInfo(i).Name, "_", " ") & ": "
            GetDistanceToAllShapes = GetDistanceToAllShapes & oShp.GetField(oLyr.FieldInfo(i).Name) & "]    "
            i = i + 1
            
        Loop
        
        oLyr.MoveNext
        
    Loop
    
    GetDistanceToAllShapes = GetDistanceToAllShapes & "|||"
    
End Function

Private Sub m_frmSelections_SelectInMap(uID As Long, _
                                        sLayerName As String, _
                                        bSelected As Boolean)
    Dim oLyr As XGIS_LayerVector
    Dim shp As XGIS_Shape

    Set oLyr = GIS.get(sLayerName)
    
    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    Set shp = oLyr.FindFirst(oLyr.Extent, "GIS_UID = " & uID, Nothing, "", True)

    If Not shp Is Nothing Then
        shp.Lock XgisLockExtent
        Set shp = shp.MakeEditable
        shp.IsSelected = bSelected
        shp.Unlock
        shp.draw
        GIS.UpDate
    End If

End Sub

Private Sub ChangeSelectionType(enmtool As OASIS_TOOLS)
        '<EhHeader>
        On Error GoTo ChangeSelectionType_Err
        '</EhHeader>
        
100     If m_SelUID > 0 Then
102         m_oDrawLyr.Delete m_SelUID
104         m_SelUID = 0
        End If
        
106     If m_lBufUID > 0 Then
108         GIS.get("Buffers").Delete m_lBufUID
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

Private Sub m_frmSelections_SelectionTypeChanged(enmtool As OASIS_TOOLS)
        '<EhHeader>
        On Error GoTo m_frmSelections_SelectionTypeChanged_Err
        '</EhHeader>
        
100     If m_SelUID > 0 Then
102         m_oDrawLyr.Delete m_SelUID
104         m_SelUID = 0
        End If
        
106     If m_lBufUID > 0 Then
108         GIS.get("Buffers").Delete m_lBufUID
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

        '<EhFooter>
        Exit Sub

m_frmSelections_SelectionTypeChanged_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmSelections_SelectionTypeChanged " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub doFeatureSelect2()
    g_CurrentTool = oFeatureSelect
    GIS.Mode = XgisUserDefined
End Sub

Private Sub doCircleSelect()
    g_CurrentTool = oCircleSelect
    GIS.Mode = XgisSelect
End Sub

Private Sub m_frmSelections_UpdateGeoSum(sLayerName As String)
    Dim oLyr As XGIS_LayerVector
    Dim shp As XGIS_Shape

    Set oLyr = GIS.get(sLayerName)
    
    m_frmSelections.SetGeoSum oLyr


End Sub

Private Sub m_frmSelections_ZoomTo(uID As Long, sLayerName As String)
Dim oLyr As XGIS_LayerVector
Dim shp As XGIS_Shape

    Set oLyr = GIS.get(sLayerName)
    
    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    Set shp = oLyr.FindFirst(oLyr.Extent, "GIS_UID = " & uID, Nothing, "", True)

    If Not shp Is Nothing Then
        GIS.VisibleExtent = shp.Extent
    End If
End Sub

Private Sub m_frmSelectorReports_DoPrint()
    
        Dim oLyr As XGIS_LayerVector
        Dim i As Integer
        Dim oRS As New adodb.Recordset
        Dim iType As Integer
        Dim iFldCount As Integer
        Dim k As Integer
        Dim oShp As XGIS_Shape
        Dim bTableTemplateCreated As Boolean
    
108     If AttribTabs.NumTabs < 1 Then Exit Sub
    
        With m_frmSelectorReports
    
110         For k = 0 To AttribTabs.NumTabs - 1

                If m_shpProps(k).UseInReport And (m_shpProps(k).sLayerCaption = ComSelLayer.List(ComSelLayer.ListIndex)) Then
                
                    GetSHP m_shpProps(k).uID, m_shpProps(k).sLayerName, oShp

                    If Not oShp Is Nothing Then
                        Set oLyr = oShp.layer
         
                        If Not oLyr Is Nothing Then
                            If Not bTableTemplateCreated Then

                                For i = 0 To oLyr.Fields.Count - 1
                                    m_frmDebug.DebugPrint oLyr.FieldInfo(i).Name
                            
                                    Select Case oLyr.FieldInfo(i).FieldType
           
                                        Case XgisFieldTypeBoolean
                                            iType = adBoolean

                                        Case XgisFieldTypeDate
                                            iType = adDate

                                        Case XgisFieldTypeFloat
                                            iType = adDouble

                                        Case XgisFieldTypeNumber
                                            iType = adDouble

                                        Case XgisFieldTypeString
                                            iType = adVarChar
                                    End Select

                                    oRS.Fields.Append oLyr.Fields.Item(i).Name, iType, oLyr.FieldInfo(i).Width ', , oSHP.GetField(oLyr.Fields.Item(i).Name)
                                Next
                    
                                iFldCount = oRS.Fields.Count - 1

                                If .chkIncludeArea.Value = vbChecked Then
                                    oRS.Fields.Append "Geo_AREA", adDouble
                                End If
            
                                If .chkIncludeGeo.Value = vbChecked Then
                                    oRS.Fields.Append "GEO_ID", adDouble
                                End If
            
                                If .chkIncludeLength.Value = vbChecked Then
                                    oRS.Fields.Append "GEO_Length", adDouble
                                End If
            
                                If .chkIncludeCentroid.Value = vbChecked Then
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
                                oRS.Fields.Item(i).Value = oShp.GetField(oRS.Fields.Item(i).Name)
                            Next
                
                            If .chkIncludeArea.Value = vbChecked Then
                                oRS.Fields.Item("Geo_AREA").Value = oShp.area
                            End If
                
                            If .chkIncludeGeo.Value = vbChecked Then
                                oRS.Fields.Item("GEO_ID").Value = oShp.uID
                            End If
            
                            If .chkIncludeLength.Value = vbChecked Then
                                oRS.Fields.Item("GEO_Length").Value = oShp.Length
                            End If
            
                            If .chkIncludeCentroid.Value = vbChecked Then
                                oRS.Fields.Item("GEO_CenterX").Value = oShp.Centroid.x
                                oRS.Fields.Item("GEO_CenterY").Value = oShp.Centroid.y
                            End If
            
                        End If
                    End If
                End If

            Next

            If oRS.RecordCount = 0 Then Exit Sub

            If .chkIncludeMap.Value = vbChecked Then
                Clipboard.Clear
                oLyr.viewer.PrintClipboard
                frmReportsFromRS.SetReportRS .txtTitle.Text, oRS, "", Clipboard.GetData(vbCFEMetafile), .txtMapTitle.Text, ""
            Else
                Clipboard.Clear
                frmReportsFromRS.SetReportRS .txtTitle.Text, oRS, ""
            End If
            
        End With
            
        frmReportsFromRS.ShowReport
        frmReportsFromRS.Show vbModal, Me

End Sub

Private Sub m_frmSelectorSettings_SettingsDone()

    With m_udtSelectorSettings
        .sSpatialOperator = DE9IM.Item(m_frmSelectorSettings.txtSpatialOperation.List(m_frmSelectorSettings.txtSpatialOperation.ListIndex))
        .dBuffeLevel = CDbl(m_frmSelectorSettings.ComBuffLevel.List(m_frmSelectorSettings.ComBuffLevel.ListIndex))
        .bAutoClear = IIf(m_frmSelectorSettings.chkAutomaticClear.Value = vbChecked, True, False)
        .bAutoFlash = IIf(m_frmSelectorSettings.chkAutoFlash.Value = vbChecked, True, False)
        .bAutoSelect = IIf(m_frmSelectorSettings.chkAutoSelect.Value = vbChecked, True, False)
        .bAutoZoom = IIf(m_frmSelectorSettings.chkAutoZoom.Value = vbChecked, True, False)
        .bEdit = IIf(m_frmSelectorSettings.chkAllowEdit.Value = vbChecked, True, False)
    End With
    
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
        Dim lL As XGIS_LayerAbstract
        Dim sLocalAttribs() As String
        Dim iNumberOfAttribs As Integer
    
100     iNumberOfAttribs = 0
         
102     Set lL = GIS.get(sLayerName)
    
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
        Dim oProvLayer As XGIS_LayerVector
        Dim shp As XGIS_Shape
    
100         Set oProvLayer = GIS.get(sLayer)
102         Set shp = oProvLayer.FindFirst(GIS.Extent, sField & " = '" & sValue & "'", Nothing, "", True)

104         If Not shp Is Nothing Then
106             GIS.VisibleExtent = shp.Extent
108             Set shp = shp.MakeEditable
110             shp.Flash
            End If

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

        '"tbUtils"

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
        
       ' sProc = SSC.Procedures.Item("OASISToolBar_ComboSelChange").Name
       '
       ' If Len(sProc) > 0 Then
       '     SSC.Run "OASISToolBar_ComboSelChange", Tool
       ' End If

        '    Dim oLyr As XGIS_LayerSHP
        '
        '    oLyr.Path = ""
        '    oLyr.Open
        '
        '    Set oLyr = CreateObject("TatukGIS_DK.XGIS_LayerVector")
        '
        '
        '    GIS.Add oLyr

        '<EhFooter>
        Exit Sub

AB_ComboSelChange_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.AB_ComboSelChange " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CreateGenericEditLyr(oGenericLyr As XGIS_LayerSqlAdo, _
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
110     GIS.Mode = XgisEdit
    
112     With GIS.Editor
114         .ShowTracking = bShowTracking
116         .ShowFast = bShowFast
118         .ShowPointsNumber = bShowPointsNumber
        End With
    
120     g_CurrentFeatureType = Location
122     m_prevTool = g_CurrentTool
        'g_CurrentTool = oCreateLocationPolyline
        
124     m_frmAddSHPWiz.Show vbModal, Me

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
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdCommand4_Click " & _
               "at line " & Erl
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
        GIS.Mode = XgisSelect

'        frmTransparencer.Init GIS.viewer
'100     frmTransparencer.Show vbModeless, Me

    'frmChartProviderSettings.Show vbModeless, Me
    
    
    '    frmDynamicReports.Init "DUDE", "C:\OASIS\Client\data\db\NWIND-IMMAP.xml"
    '    frmDynamicReports.Show
        '<EhFooter>
        Exit Sub

cmdCommand5_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdCommand5_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCreateThematics_Click()
'    Dim i As Integer
'    Dim OASISLyr As String
'    Dim oLyrOASIS As XGIS_LayerVector
'    Dim oLyrProv As XGIS_LayerVector
'
'    'cosLabel
'    'm_oCosmeticLayer
'    ' m_oSQLIncLyr
'
'
'    g_RSAppSettings.MoveFirst
'    g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"
'    OASISLyr = g_RSAppSettings.Fields.Item("SettingValue1").value
'
'    g_RSAppSettings.MoveFirst
'    g_RSAppSettings.Find "SettingName = 'AdmProvSec'"
'
'    'm_sSecAnalysisFieldName = g_RSAppSettings.Fields.Item("SettingValue3").value
'    Set oLyrProv = GIS.Get(g_RSAppSettings.Fields.Item("SettingValue1").value)
'
'    oLyrProv.MoveFirst oLyrProv.Extent, "", Nothing, "", True
'
'    Dim shp As New XGIS_ShapePoint
'    Dim ptg As New XGIS_Point
'
'    Do While Not oLyrProv.EOF
'        m_frmDebug.DebugPrint  oLyrProv.Shape.GetField(g_RSAppSettings.Fields.Item("SettingValue3").value)
'        'Set shp = New XGIS_ShapePoint
'        Set ptg = oLyrProv.Shape.Centroid
'
'        Set shp = m_oCosmeticLayer.CreateShape(XgisShapeTypePoint)
'
'        'm_frmDebug.DebugPrint  shp.Centroid.X
'        'm_frmDebug.DebugPrint  shp.Centroid.Y
'        shp.Lock XgisLockExtent
'        shp.AddPart
'        shp.AddPoint oLyrProv.Shape.Centroid
'        shp.Unlock
'
'        oLyrProv.MoveNext
'    Loop
'
End Sub

Private Sub cmdExportData_Click()
        '<EhHeader>
        On Error GoTo cmdExportData_Click_Err
        '</EhHeader>
100             frmExportFormats.FraExportFormats.Visible = False
102             frmExportFormats.Show vbModal, Me

104             If frmExportFormats.bExport Then
106                 If frmExportFormats.chkFormats(0).Value = vbChecked Then
108                     dxGISDataGrid.M.ExportToXLS g_sAppPath & "\data\user\Exports\W3_OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xls"
                    End If
110                 If frmExportFormats.chkFormats(1).Value = vbChecked Then
112                     dxGISDataGrid.M.ExportToXML g_sAppPath & "\data\user\Exports\W3_OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xml"
                    End If
114                 If frmExportFormats.chkFormats(2).Value = vbChecked Then
116                     dxGISDataGrid.M.ExportToHTML g_sAppPath & "\data\user\Exports\W3_OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".htm"
                    End If
118                 If frmExportFormats.chkFormats(3).Value = vbChecked Then
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

Private Sub cmdJoinAttribute_Click()
        '<EhHeader>
        On Error GoTo cmdJoinAttribute_Click_Err
        '</EhHeader>
        Dim sSpatialKey As String
        Dim sAttrKey As String
        Dim oRS As New adodb.Recordset

100     frmLayerSelection.Init GIS
102     frmLayerSelection.Show vbModal, Me
    
104     oRS.Open "Select * FROM 111Table", m_Cnn, adOpenDynamic, adLockBatchOptimistic
    
106     m_frmDebug.DebugPrint frmLayerSelection.GetItem
    
108     Set m_oSQLOpsLyr = New XGIS_LayerSqlAdo

110     With m_oSQLOpsLyr
        
112         .Name = "Operations11"
114         .SQLParameter("LAYER") = "oOperations"
116         .SQLParameter("DIALECT") = "MSJET"
118         '.SQLParameter("ADO") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\Oasisclient.mdb;" '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\db\Oasisclient.mdb;pwd=<#mypassword#>"
            .SQLParameter("ADO") = GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")
120         .HideFromLegend = False
            
        End With

        '        m_oSQLOpsLyr.Path = "C:\Program Files\TatukGIS\DK8\Samples\VB6\Set1\SimpleEdit\editSQL.ttkls"
        '        m_oSQLOpsLyr.SQLParameter("ADO") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\Oasisclient.mdb;"
        '        m_oSQLOpsLyr.HideFromLegend = False

122     GIS.Add m_oSQLOpsLyr
124     GIS.Mode = XgisSelect
126     JoinAttributeDataToGeometries "Operations11", oRS, "UID", "id"
    
128     Unload frmLayerSelection
    
        '<EhFooter>
        Exit Sub

cmdJoinAttribute_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdJoinAttribute_Click " & _
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
        Dim ptg As New XGIS_Point
        Dim oShp As XGIS_Shape
        Dim sFont As String
        Dim SymbolList As XGIS_SymbolList
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
                    'm_oDrawLyr.MoveFirst m_oDrawLyr,
108                 Set oShp = m_oDrawLyr.GetShape(g_lPinUID)
                 
110                 If oShp Is Nothing Then Exit Sub
                
112                 With oShp
114                     .Lock XgisLockExtent
116                     .SetPosition ptg, m_oDrawLyr, 0
118                     .Unlock
                    End With
                
                Else
120                 Set oShp = m_oDrawLyr.CreateShape(XgisShapeTypePoint)
                    
122                 If oShp Is Nothing Then Exit Sub
                        
124                 With oShp
126                     .Lock XgisLockExtent
128                     .AddPart
130                     .AddPoint ptg

132                     With .Params.labels
134                         .Alignment = XgisLabelAlignmentCenter
136                         .Color = vbRed
138                         .FontColor = vbGreen
140                         .Allocator = False
142                         .Duplicates = True
144                         .Font.Size = 45
146                         .OutlineWidth = 0
148                         .Pattern = XbsClear
150                         .Position = XgisLabelPositionMiddleCenter
152                         .Value = ""
                        End With
                                        
154                     sFont = txtFont(0).FontName
156                     sFont = sFont & ":" & Asc(txtFont(0).Text) & ":NORMAL"

158                     Set SymbolList = New XGIS_SymbolList

160                     .Params.Marker.Color = FraPinColor(0).BackColor
162                     .Params.Marker.OutlineColor = 16711680
164                     .Params.Marker.Symbol = SymbolList.Prepare(sFont)
166                     .Params.Marker.Size = 500
'168                     .Params.Marker.ShowLegend = 1
'170                     .Params.Legend = "Uncategorized"
                
'172                     .Params.Marker.OutlineStyle = XpsSolid
'174                     .Params.Marker.OutlineWidth = 20
'176                     .Params.Marker.OutlineColor = vbRed
'178                     .Params.Marker.Style = XgisMarkerStyleTriangleDown
'180                     .Params.Marker.Color = vbBlue
'182                     .Params.Marker.Size = 200
                        Dim ostr As New XStringList
                        
                        .Params.Marker.SaveToStrings ostr
                        
                        On Error Resume Next
                        
                        .SetField "StringStyle", ostr.Text
                        m_frmDebug.DebugPrint ostr.Text
184                     .Unlock
186                     g_lPinUID = .uID
                    End With

                    '                GIS.UpDate
                End If

            End If
        End If

188     GIS.CenterViewport ptg
        If Not oShp Is Nothing Then oShp.Flash 12
        
       ' m_oDrawLyr.SaveAll
        
        
190     GIS.UpDate
    
    
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
100     'm_frmDebug.DebugPrint "onEndDragGroupColumn"
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

104     If chkFilterIn.Value = vbChecked Then Call chkFilterIn_Click
106    '     SafeMoveFirst g_RSGISGridTableSettings
        
108    '     g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"
    
        '    Dim lL As XGIS_LayerVector
                
110     '    If Not g_RSGISGridTableSettings.EOF Then
112     '        Set lL = GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
        '    Else
114     '        Set lL = GIS.get(abGridTools.Tools.Item("comLyr").Text)
        '    End If

116     '    DoEvents
118     '    lL.MoveFirst lL.Extent, "", Nothing, "", True

120      '   DoEvents
            
122      '   lL.scope = FilterText
                
124     '  lL.Params.Visible = True
126    '     GIS.UpDate
      '  End If
        
128     If chkSelectIn.Value = vbChecked Then Call chkSelectIn_Click
'130         GIS.Lock
'132         SafeMoveFirst g_RSGISGridTableSettings
'134         g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"
            
'136         Set lL = GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
'138         lL.DeselectAll
'140         lL.MoveFirst lL.Extent, FilterText, Nothing, "", True
            
'            Dim shp As XGIS_Shape
            
'142         Set shp = lL.FindFirst(lL.Extent, FilterText, Nothing, "", True)

'144         Do While Not shp Is Nothing
            
'146             Set shp = shp.MakeEditable
'148             shp.IsSelected = True
'150             Set shp = lL.FindNext
'            Loop

'152         GIS.Unlock

            'Also intead or using FindFirst, use MoveFirst:
            '
            '---------------
            'll.MoveFirst (ll.Extent, criteria);
            'shp := ll.Shape;
            '--------------
            '--------------
            'll.MoveNext;
            'shp := ll.Shape;
            '--------------
 '       End If
        
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

        Case "ResourcesFinder"
        
            m_frmResourcesFinder.Show vbModeless, Me
        
        Case "WVISitRep"
        
            If FileExists(g_sAppPath & "\data\db\dynamicdata\WorldVision.mdb") Then
                m_frmWVIControlPanel.Show vbModal, Me
            End If
            
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
                
            m_frmSearch.Init GIS.viewer
                
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
            sAppsets = sAppsets & "GIS Project: " & GIS.ProjectName & vbCrLf
            sAppsets = sAppsets & "SERVER: " & g_sAppServerPath & vbCrLf
            MsgBox sAppsets

        Case "Folders"
            frmFolders.Show vbModal, Me

        Case "GPS"
            frmGPS.Show vbModeless, Me
            'OASISLocator1.Init GIS.viewer
            'LocatorHandle.Visible = Not LocatorHandle.Visible

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
                            .Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
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
                        .AddScriptingObject "OASISGis", GIS
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
             frmMapPrint.Init GIS.viewer, m_frmMnuOperations.Legend1.Legend, g_bMapTplUpdate, g_sRemoteTablePrefix
             frmMapPrint.Show vbModeless, Me
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
            elW3.Visible = False
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
        
        'SSC.AddCode  "Call OASISGis.AboutBox()"
        
106     If Not g_RSAppSettings.EOF Then
108         If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
110             If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = "" Then
                    
112                 sScripts = Split(g_RSAppSettings.Fields.Item("SettingValue1").Value, ",")
                                
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
'  tlbMain.Buttons("undo").Enabled = GIS.Editor.CanUndo
'  tlbMain.Buttons("redo").Enabled = GIS.Editor.CanRedo
'  GIS.Editor.CurrentShape.
  GIS.UpDate
End Sub

Private Sub GIS_OnKeyUp(translated As Boolean, Key As Integer, ByVal Shift As TatukGIS_DK.XShiftState)
        '<EhHeader>
        On Error GoTo GIS_OnKeyUp_Err
        '</EhHeader>
100   If Key = VK_CONTROL Then
    '    If tlbMain.Buttons("edit").Value = tbrPressed Then
    '      GIS.Mode = XgisEdit
    '    Else
    '      If tlbMain.Buttons("drag").Value = tbrPressed Then
    '        GIS.Mode = XgisDrag
    '      End If
    '    End If
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
        Dim RSIncidentForOASISReports As adodb.Recordset
        Dim iFieldCount As Integer
        Dim m_frmReportsFromRS As frmReportsFromRS ' Dim i As Integer
       
        Dim sPDFPath As String
        Dim sEMHTML As String
        Dim RSUpdater As adodb.Recordset
        Dim ptg As XGIS_Point
        
'100     sGUID = GetGuid
        
102     g_CurrentTool = oZoom
104     GIS.Mode = XgisZoomEx
        'idiot
            
106     If m_oIncShpPt Is Nothing Then
108         Set m_oIncShpPt = m_oSQLIncLyr.CreateShape(XgisShapeTypePoint)
            
            Set ptg = New XGIS_Point
            
            With m_fmrAddIncident
            
               ptg.Prepare CDbl(.txtX.Text), CDbl(.txtY.Text)
            
            End With
            
            
110         m_oIncShpPt.Lock XgisLockExtent
112         m_oIncShpPt.AddPart

114         m_oIncShpPt.AddPoint ptg

        End If
            
136     If Not m_oIncShpPt Is Nothing Then
138         m_oIncShpPt.Lock XgisLockExtent
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
160             m_oIncShpPt.SetField "Incident_DATE", m_fmrAddIncident.MVIncident.Value
            End If
            
162         If Not IsEmpty(m_oIncShpPt.GetField("Incident_DATESERIAL")) Then
164             m_oIncShpPt.SetField "Incident_DATESERIAL", ConvertDateToSerial(m_fmrAddIncident.MVIncident.Value)
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
            
194         If m_fmrAddIncident.OptCasualties(0).Value Then
                        
196             If m_fmrAddIncident.chkUnknown.Value = vbUnchecked Then
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

210         ElseIf m_fmrAddIncident.OptCasualties(0).Value Then

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
226             If m_fmrAddIncident.OptIncViolence(0).Value Then
228                 m_oIncShpPt.SetField "Violent", 0 '0,1,2
230             ElseIf m_fmrAddIncident.OptIncViolence(1).Value = True Then
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
        
282         GIS.Center = m_oIncShpPt.Centroid
284         m_oSQLIncLyr.SaveData

286         SetNewSynchDBElement GetGuid, sGUID, "OASIS Incidents", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, "oincidents", True
            
288         If m_fmrAddIncident.chkOASISReports = vbChecked Then

290             Set RSIncidentForOASISReports = New adodb.Recordset

292             iFieldCount = m_oIncShpPt.layer.Fields.Count
294             i = 0

296             Do Until i = iFieldCount
            
298                 If Not m_oIncShpPt.layer.Fields.Item(i).Name = "GUID" And Not m_oIncShpPt.layer.Fields.Item(i).Name = "ID" Then
        
300                     Select Case m_oIncShpPt.layer.Fields.Item(i).FieldType
                
                            Case XgisFieldTypeBoolean
302                             RSIncidentForOASISReports.Fields.Append m_oIncShpPt.layer.Fields.Item(i).Name, adBoolean
                
304                         Case XgisFieldTypeDate
306                             RSIncidentForOASISReports.Fields.Append m_oIncShpPt.layer.Fields.Item(i).Name, adDate

308                         Case XgisFieldTypeFloat
310                             RSIncidentForOASISReports.Fields.Append m_oIncShpPt.layer.Fields.Item(i).Name, adDouble

312                         Case XgisFieldTypeNumber
314                             RSIncidentForOASISReports.Fields.Append m_oIncShpPt.layer.Fields.Item(i).Name, adBigInt

316                         Case XgisFieldTypeString
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
344                     RSIncidentForOASISReports.Fields("UID").Value = m_oIncShpPt.uID
                    Else
346                     RSIncidentForOASISReports.Fields(i).Value = m_oIncShpPt.GetField(RSIncidentForOASISReports.Fields(i).Name)
                    End If

348                 i = i + 1
                Loop
            
350             Set m_frmReportsFromRS = New frmReportsFromRS
                        
352             If frmGridRptProprties.chkIncludeMap = vbChecked Then
354                 Clipboard.Clear
356                 GIS.viewer.PrintClipboard

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
                                
388             Set RSUpdater = New adodb.Recordset

390             With RSUpdater

392                 .Open "SELECT * FROM Attachments", m_Cnn, adOpenDynamic, adLockBatchOptimistic
394                 .AddNew
396                 .Fields("incidentID").Value = dID
398                 .Fields("FilePath").Value = varAttachments(i)
400                 .Fields("DateInserted").Value = Now()
402                 .Fields("AttachmentTable").Value = "oincidents"
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
                
410             UploadAttachments g_RSAppSettings.Fields.Item("SettingValue1").Value, CStr(varAttachments(i))
            Next
        
        Else

412         If Not varAttachments(0) = "" Then
                
414             Set RSUpdater = New adodb.Recordset

416             With RSUpdater
            
418                 .Open "SELECT * FROM Attachments", m_Cnn, adOpenDynamic, adLockBatchOptimistic
420                 .AddNew
422                 .Fields("incidentID").Value = dID
424                 .Fields("FilePath").Value = varAttachments(0)
426                 .Fields("DateInserted").Value = Now()
428                 .Fields("AttachmentTable").Value = "oincidents"
430                 .UpdateBatch adAffectCurrent
432                 .Close

                End With

434             Set RSUpdater = Nothing

436             UploadAttachments g_RSAppSettings.Fields.Item("SettingValue1").Value, CStr(varAttachments(0))
            End If
        End If
        
438     If m_fmrAddIncident.chkHTMLReport = vbChecked Then
        
            On Error Resume Next
        
440         Kill g_sAppPath & "\data\user\Exports\" & "exp.jpg"
442         Kill g_sAppPath & "\data\user\Exports\" & "exp.JGW"
444         Kill g_sAppPath & "\data\user\Exports\" & "exp.JFW"
446         Kill g_sAppPath & "\data\user\Exports\" & "exp.tab"
448         Kill g_sAppPath & "\data\user\Exports\" & "OasisReport.html"
        
450         GIS.UpDate
            
452         GIS.viewer.ExportToImage g_sAppPath & "\data\user\Exports\" & "exp.jpg", GIS.VisibleExtent, ScaleX(GIS.Width, vbTwips, vbPixels), ScaleY(GIS.Height, vbTwips, vbPixels), 100, 0, 96

454         sEMHTML = WriteHTMLReport("exp.jpg")
            'GIS.Unlock
        
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
    
        Dim RSUpdater As adodb.Recordset
100     Set RSUpdater = New adodb.Recordset
102     With RSUpdater
            
104         .Open "SELECT * FROM SynchHistory", m_Cnn, adOpenDynamic, adLockBatchOptimistic

108         If Not .EOF Then
                .AddNew
110             .Fields("sID").Value = sID
112             .Fields("sGUID").Value = sNewGUID
114             .Fields("sTableName").Value = sTableName
116             .Fields("swhen").Value = sRFC3339DateTime
118             .Fields("sStatus").Value = "pending"
120             .Fields("sequence").Value = 1
122             .Fields("sBy").Value = sBy
124             .Fields("sdelete").Value = "false"
126             .Fields("updates").Value = supdates
128             .Fields("noconflict").Value = "local"
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
                                     Optional CN As adodb.Connection)
    
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
    Dim oVLayer As XGIS_LayerVector
100         m_frmDebug.DebugPrint sScope
102     Set oVLayer = GIS.get(sLayer)
    
104     If Not oVLayer Is Nothing Then
106         oVLayer.scope = sScope
        End If
    
108     GIS.UpDate
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
    Dim oVLayer As XGIS_LayerVector
    Dim i As Integer

100     Set oVLayer = GIS.get(sLayer)

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
    Dim oVLayer As XGIS_LayerVector
    Dim i As Integer

100     Set oVLayer = GIS.get(sLayer)
        'oVLayer.Scope = "DISTINCT " & sField
    
102     oVLayer.MoveFirst oVLayer.Extent, "", Nothing, "", True
    
104     m_frmChangeTracer.lvUniqueValues.ColumnHeaders.Clear
106     m_frmChangeTracer.lvUniqueValues.ListItems.Clear
    
108     m_frmChangeTracer.lvUniqueValues.ColumnHeaders.Add Text:="Unique Value"
    
110     Do While Not oVLayer.EOF
112         m_frmDebug.DebugPrint oVLayer.Shape.GetField(sField)
114         m_frmChangeTracer.lvUniqueValues.ListItems.Add Text:=oVLayer.Shape.GetField(sField)
116         oVLayer.MoveNext
        Loop
    
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
        Dim oVLayer As XGIS_LayerVector
        Dim i As Integer

100     Set oVLayer = GIS.get(sLayer)

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
        Dim oVLayer As XGIS_LayerVector
        Dim i As Integer

100     Set oVLayer = GIS.get(sLayer)
        'oVLayer.Scope = "DISTINCT " & sField
    
102     oVLayer.MoveFirst oVLayer.Extent, "", Nothing, "", True
    
104     With m_frmFreqSettings
106         With .lvUniqueValues
        
108             .ColumnHeaders.Clear
110             .ListItems.Clear
    
112             .ColumnHeaders.Add Text:="Unique Value"
    
114             Do While Not oVLayer.EOF
116                 m_frmDebug.DebugPrint oVLayer.Shape.GetField(sField)
118                 .ListItems.Add Text:=oVLayer.Shape.GetField(sField)
120                 oVLayer.MoveNext
                Loop
    
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

Private Sub m_frmFreqSettings_ShowSpatialAnalysisLegend()
        '<EhHeader>
        On Error GoTo m_frmFreqSettings_ShowSpatialAnalysisLegend_Err
        '</EhHeader>
100     frmSpatialAnalysisLegend.GIS.Add GIS.get(m_frmFreqSettings.OverLayLayer)
102     frmSpatialAnalysisLegend.Legend1.GIS_Viewer = frmSpatialAnalysisLegend.GIS.viewer
104     frmSpatialAnalysisLegend.Show vbModal, Me
        '<EhFooter>
        Exit Sub

m_frmFreqSettings_ShowSpatialAnalysisLegend_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmFreqSettings_ShowSpatialAnalysisLegend " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadAvailableThematics()
        '<EhHeader>
        On Error GoTo LoadAvailableThematics_Err
        '</EhHeader>
        Dim rsThGroup As New adodb.Recordset

        'Set g_RSThemeSettings = New ADODB.Recordset
                
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'ThemeTool'"
                
104     If g_RSAppSettings.Fields.Item("SettingValue1").Value = "0" Then
106         m_frmMnuOperations.elLegend.Grid(gsRowHeight, 1) = (0)
            Exit Sub
        End If
    
108     rsThGroup.Open "SELECT * FROM ThemeGroups", m_Cnn, adOpenDynamic, adLockReadOnly
        
110     m_frmMnuOperations.ComThematicsGroup.Clear
112     m_frmMnuOperations.ComThematicsGroup.AddItem "--- No Thematic Group ---"
        
        If Not rsThGroup.BOF And Not rsThGroup.EOF Then
114         SafeMoveFirst rsThGroup
    
116         Do While Not rsThGroup.EOF
                If Not IsNull(rsThGroup.Fields.Item("Name").Value) Then
118                 m_frmMnuOperations.ComThematicsGroup.AddItem rsThGroup.Fields.Item("Name").Value
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

Private Sub m_frmIOMJOC_CreateTheme(sName As String)
        '<EhHeader>
        On Error GoTo m_frmIOMJOC_CreateTheme_Err
        '</EhHeader>
        Dim oLyr As XGIS_LayerVector
        
100     Set oLyr = GIS.get("IOM_Progress_Average_Per_Item_point")
            
102     If Not oLyr Is Nothing Then
            
104         With oLyr
106             .ConfigName = g_sAppPath & "\data\user\Maps\Layers\" & sName  '"C:\OASIS\NFI1"
108             .StoreParamsInProject = False
110             .UseConfig = True
112             .RereadConfig
114             .draw
116             GIS.InvalidateExtent .Extent ' UpDate
            End With
        
            'ShowThemelegend oLyr
        
        Else
        
            On Error Resume Next
        
        End If
    
    
        '<EhFooter>
        Exit Sub

m_frmIOMJOC_CreateTheme_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmIOMJOC_CreateTheme " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmIOMJOC_ShowInMap(sID As String)
        '<EhHeader>
        On Error GoTo m_frmIOMJOC_ShowInMap_Err
        '</EhHeader>
    Dim oLyr As XGIS_LayerVector
    Dim oShp As XGIS_ShapePoint

100     If sID = "" Then Exit Sub

102     Set oLyr = GIS.get("IOM_Progress_Average_Per_Item_point")
    
104     If oLyr Is Nothing Then Exit Sub
    
106     Set oShp = oLyr.FindFirst(oLyr.Extent, "ID = " & CInt(sID), Nothing, "", True)
    
108     If Not oShp Is Nothing Then
110         GIS.ZoomBy GIS.zoom / 2, oShp.Centroid.x, oShp.Centroid.y
        End If
    
        '<EhFooter>
        Exit Sub

m_frmIOMJOC_ShowInMap_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmIOMJOC_ShowInMap " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmLocator_hglAdminLevel0()
        '<EhHeader>
        On Error GoTo m_frmLocator_hglAdminLevel0_Err
        '</EhHeader>
        Dim oProvLayer As XGIS_LayerVector
        Dim shp As XGIS_Shape
    
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel0'"
        
104     Set oProvLayer = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
        
        If oProvLayer Is Nothing Then Exit Sub
        
        oProvLayer.MoveFirst oProvLayer.Extent, "", Nothing, "", True
        
106     Set shp = oProvLayer.FindFirst(oProvLayer.Extent, g_RSAppSettings.Fields.Item("SettingValue2").Value & " = '" & m_frmLocator.ComProvince.List(m_frmLocator.ComProvince.ListIndex) & "'", Nothing, "", True)
        
        If shp Is Nothing Then Exit Sub
        
       ' m_frmDebug.DebugPrint  shp.GetField("PRV_NAME")
        
108     GIS.VisibleExtent = shp.Extent
    oProvLayer.MoveFirst oProvLayer.Extent, "", Nothing, "", True
    Set shp = oProvLayer.FindFirst(oProvLayer.Extent, g_RSAppSettings.Fields.Item("SettingValue2").Value & " = '" & m_frmLocator.ComProvince.List(m_frmLocator.ComProvince.ListIndex) & "'", Nothing, "", True)

    shp.Flash 8, 10
    
        'DoEvents
110     'shp.Flash

        '<EhFooter>
        Exit Sub

m_frmLocator_hglAdminLevel0_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmLocator_hglAdminLevel0 " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmLocator_hglAdminLevel1()
        '<EhHeader>
        On Error GoTo m_frmLocator_hglAdminLevel1_Err
        '</EhHeader>
        Dim oDistLayer As XGIS_LayerVector
        Dim shp As XGIS_Shape
    
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'PCodeAdminLevel1'"
    
104     Set oDistLayer = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
        
        If oDistLayer Is Nothing Then Exit Sub
        
106     Set shp = oDistLayer.FindFirst(oDistLayer.Extent, g_RSAppSettings.Fields.Item("SettingValue2").Value & " = '" & m_frmLocator.ComDistrict.List(m_frmLocator.ComDistrict.ListIndex) & "'", Nothing, "", True)
    
108     If shp Is Nothing Then
            Exit Sub
        End If
    
110     GIS.VisibleExtent = shp.Extent
    oDistLayer.MoveFirst oDistLayer.Extent, "", Nothing, "", True
    Set shp = oDistLayer.FindFirst(oDistLayer.Extent, g_RSAppSettings.Fields.Item("SettingValue2").Value & " = '" & m_frmLocator.ComDistrict.List(m_frmLocator.ComDistrict.ListIndex) & "'", Nothing, "", True)

    shp.Flash 8, 10


        '<EhFooter>
        Exit Sub

m_frmLocator_hglAdminLevel1_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmLocator_hglAdminLevel1 " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmLocator_hglAdminLocation()
        '<EhHeader>
        On Error GoTo m_frmLocator_hglAdminLocation_Err
        '</EhHeader>
        Dim oVillageLayer As XGIS_LayerVector
        Dim shp As XGIS_Shape
    
    
100     If Not m_frmLocator.ComPlace.List(m_frmLocator.ComPlace.ListIndex) = "--ALL--" Then
102         SafeMoveFirst g_RSAppSettings
104         g_RSAppSettings.Find "SettingName = 'HICPcodeLayer'"

106         Set oVillageLayer = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
        
108         If oVillageLayer Is Nothing Then Exit Sub
        
110         Set shp = oVillageLayer.FindFirst(oVillageLayer.Extent, g_RSAppSettings.Fields.Item("SettingValue4").Value & " = '" & Replace(m_frmLocator.ComPlace.List(m_frmLocator.ComPlace.ListIndex), "'", "''") & "'", Nothing, "", True)
                
112         If Not shp Is Nothing Then
114             GIS.VisibleExtent = shp.Extent
Set shp = oVillageLayer.FindFirst(oVillageLayer.Extent, g_RSAppSettings.Fields.Item("SettingValue4").Value & " = '" & Replace(m_frmLocator.ComPlace.List(m_frmLocator.ComPlace.ListIndex), "'", "''") & "'", Nothing, "", True)
            
        shp.Flash 8, 10
                shp.Invalidate
            End If
        End If


        '<EhFooter>
        Exit Sub

m_frmLocator_hglAdminLocation_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmLocator_hglAdminLocation " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_ActivateTheme(sTheme As String, _
                                             sThemeGR As String)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_ActivateTheme_Err
        '</EhHeader>
        Dim oLyr As XGIS_LayerVector
        Dim rsTHGR As New adodb.Recordset
        Dim rsTh As New adodb.Recordset
        Dim sFile As String
        Dim oIni As New XGIS_Ini
        
100     rsTHGR.Open "SELECT * FROM ThemeGroups WHERE Name = '" & sThemeGR & "'", m_Cnn, adOpenDynamic, adLockReadOnly
    
102     If (Not rsTHGR.EOF) And (Not rsTHGR.BOF) Then
104         SafeMoveFirst rsTHGR
        Else
            On Error Resume Next
    
106         rsTHGR.Close
108         Set rsTHGR = Nothing
110         Set rsTh = Nothing
        
112         MsgBox "The Theme Group: " & sThemeGR & " is not correct. Please contact your OASIS administrator and try again.", vbInformation, "OASIS Theme Manager"
            Exit Sub
        End If
    
114     rsTh.Open "SELECT * FROM Themes WHERE ThemeGroup = " & rsTHGR.Fields.Item("ID").Value, m_Cnn, adOpenDynamic, adLockReadOnly

        If Not IsNull(rsTh.Fields.Item("AnalysisLayer").Value) Then
116         Set oLyr = GIS.get(rsTh.Fields.Item("AnalysisLayer").Value)
            
118         If Not oLyr Is Nothing Then
                Dim ofs As New FileSystemObject
                    
                If ofs.FileExists(g_sAppPath & "\data\user\Maps\Layers\" & rsTh.Fields.Item("ThemeConfigName").Value) Then

120                 With oLyr
                        '.ConfigFile
                        .UseConfig = True
                        sFile = g_sAppPath & "\data\user\Maps\Layers\" & Replace$(rsTh.Fields.Item("ThemeConfigName").Value, ".ini", "")
122                     .ConfigName = sFile '"C:\OASIS\NFI1"
                        'oINI.Create_
124                     '.Params.LoadFromIni sFIle
                        '.StoreParamsInProject = False
126                     '.UseConfig = True
128                     .RereadConfig
130                     .Params.Visible = True
                        .draw
132                     'GIS.viewer.InvalidateExtent oLyr.Extent  ' UpDate
                    
                    End With
        
134                 ShowThemelegend oLyr
                    'm_frmMnuOperations.Legend1.Invalidate
                    'm_frmMnuOperations.Legend1.GIS_Viewer = Nothing
                    'm_frmMnuOperations.Legend1.GIS_Viewer = GIS.viewer  ' UpDate
                Else
                    MsgBox "Seems Like File " & g_sAppPath & "\data\user\Maps\Layers\" & rsTh.Fields.Item("ThemeConfigName").Value & " is missing or configuration is incorrect.", vbInformation, "OASIS Client"
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

Private Sub m_frmMnuOperations_LoadMap(sMapName As String)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_LoadMap_Err
        '</EhHeader>

100     LoadAvailableThematics

        'm_frmDebug.DebugPrint  m_frmMnuOperations.elLegend.Align
    
102     GIS.Lock
    
104     m_bLOADING = True

106     m_bMapInitialized = True
    
108     PrepareOverViewMap g_sAppPath & "\data\user\Maps\" & sMapName
    
        'GIS.Open g_sAppPath & "\data\user\Maps\" & sMapName

110     LoadMapProducts g_sAppPath & "\data\user\Maps\" & sMapName
        
112     AddAllAdditionalStandardLayers
        
114     CreateW3Layers
116     CreateW3WhoLayers
120     FillCOPValues

122     Set g_PrevExt = GIS.VisibleExtent
    
        Dim RS As New adodb.Recordset
        RS.Open "SELECT DefaultViewName,DefaultViewX,DefaultViewY,DefaultViewZ,LatestViewX,LatestViewY,LatestViewZ,LatestMapName FROM Personnell WHERE Personnell_ID = " & g_CurrentUserID, m_Cnn, adOpenDynamic, adLockBatchOptimistic
124     'Set rs = m_Cnn.Execute("SELECT DefaultViewName,DefaultViewX,DefaultViewY,DefaultViewZ,LatestViewX,LatestViewY,LatestViewZ,LatestMapName FROM Personnell WHERE Personnell_ID = " & g_CurrentUserID)
        
126     If Not RS.BOF Then SafeMoveFirst RS
        
128     m_DefaultViewName = RS.Fields.Item("DefaultViewName").Value
130     m_DefaultViewx = RS.Fields.Item("DefaultViewX").Value
132     m_DefaultViewy = RS.Fields.Item("DefaultViewY").Value
134     m_DefaultViewz = RS.Fields.Item("DefaultViewZ").Value
    
136     LoadGeoBookMarks
138     ComThemes.ListIndex = 0
140     C1TabFastFunction.Width = 340

142     If GIS.Items.Count > 0 Then
144         SafeMoveFirst g_RSAppSettings
146         g_RSAppSettings.Find "SettingName = 'MADataSetFiles'"
    
            Dim sLayers() As String
            Dim i As Integer
            Dim oLyr As Object 'XGIS_LayerAbstract
            
148         sLayers = Split(g_RSAppSettings.Fields.Item("SettingValue1").Value, ",")
    
150         For i = 0 To UBound(sLayers)
152             Set oLyr = GIS.get(sLayers(i))

154             If Not oLyr Is Nothing Then oLyr.HideFromLegend = True
            Next
    
156         SafeMoveFirst g_RSAppSettings
158         g_RSAppSettings.Find "SettingName = 'HiddenLayers'"
    
            If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").Value) Then
    
160             sLayers = Split(g_RSAppSettings.Fields.Item("SettingValue1").Value, ",")
    
162         For i = 0 To UBound(sLayers)
164             Set oLyr = GIS.get(sLayers(i))

166             If Not oLyr Is Nothing Then oLyr.HideFromLegend = True
            Next

            End If

        End If
        
168     m_frmMnuOperations.Init GIS.viewer, m_Cnn

170     AB.RecalcLayout
172     m_bLOADING = False

174     GIS.Unlock

        '<EhFooter>
        Exit Sub

m_frmMnuOperations_LoadMap_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.m_frmMnuOperations_LoadMap " & "at line " & Erl
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

Private Sub m_frmMnuOperations_ShowIOM()
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_ShowIOM_Err
        '</EhHeader>
    
100     If m_frmIOMJOC.Visible Then
            Exit Sub
        Else
102         m_frmIOMJOC.Init
        End If
    
104     m_frmIOMJOC.Show vbModeless, Me
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_ShowIOM_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_ShowIOM " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMnuOperations_UpdateNFI(sLocation As String, iNeed As Integer, iDelivery As Integer)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_UpdateNFI_Err
        '</EhHeader>
    Dim oLyr As XGIS_LayerVector
    Dim oShp As XGIS_Shape
100 Set oLyr = GIS.get("hic_gov_poly1")

102     m_frmDebug.DebugPrint sLocation
    
            
            
104             With oLyr
                    '.Lock
                    '.MoveFirst .Extent, "", Nothing, "", True
106                 .MoveFirst .Extent, "Adm2Name = '" & sLocation & "'", Nothing, "", True
108                 Do While Not .EOF
110                     Set oShp = .Shape
112                     oShp.MakeEditable
114                     oShp.SetField "Need", iNeed
116                     oShp.SetField "Delivery", iDelivery
                    
118                     .MoveNext
                    Loop
                
                    '.ConfigName = "C:\OASIS\gov"
                   ' .ConfigName = "C:\OASIS\demographics"
120                 .ConfigName = "C:\OASIS\NFI1"
122                 .StoreParamsInProject = False
                    '.UseFileParams = True
                    '.IgnoreShapeParams = True
                    '.Params.LoadFromIni
124                 .UseConfig = True
126                 .RereadConfig
                    '.Unlock
128                 .draw
130                 GIS.UpDate
                End With
    
    
    '            g_RSAppSettings.MoveFirst
    '        g_RSAppSettings.Find "SettingName = 'MAPcodeLayer'"
    '
    '        Set ll = GIS.Get(g_RSAppSettings.Fields.Item("SettingValue1").value)
    '        Set shp = ll.FindFirst(GIS.Extent, g_RSAppSettings.Fields.Item("SettingValue5").value & " = '" & sID & "'", Nothing, "", True)
    '
    '
    '        'm_frmDebug.DebugPrint  shp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value)
    '        If Not shp Is Nothing Then
    '            GIS.VisibleExtent = shp.Extent
    '            'GIS.Update
    '            'shp.IsSelected = True
    '            shp.MakeEditable
    '            shp.Flash
    '        End If

        '<EhFooter>
        Exit Sub

m_frmMnuOperations_UpdateNFI_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_UpdateNFI " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmTextAnnoSettings_DeleteAnnoText()
        '<EhHeader>
        On Error GoTo m_frmTextAnnoSettings_DeleteAnnoText_Err
        '</EhHeader>

100     If m_frmTextAnnoSettings.lstTexts.ListIndex = -1 Then Exit Sub

102     m_oDrawLyr.MoveFirst m_oDrawLyr.Extent, "GIS_UID = " & m_frmTextAnnoSettings.lstTexts.ItemData(m_frmTextAnnoSettings.lstTexts.ListIndex), Nothing, "", True
    
104     If Not m_oDrawLyr.Shape Is Nothing Then
106         m_oDrawLyr.Delete m_oDrawLyr.Shape.uID
108         GIS.UpDate
        End If

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
    m_oDrawLyr.MoveFirst m_oDrawLyr.Extent, "GIS_UID = " & m_frmTextAnnoSettings.lstTexts.ItemData(m_frmTextAnnoSettings.lstTexts.ListIndex), Nothing, "", True
      
    If Not m_oDrawLyr.Shape Is Nothing Then

        With m_oDrawLyr.Shape.Params.labels
            m_frmTextAnnoSettings.panColorBack.BackColor = .Color
            m_frmTextAnnoSettings.panColorFore.BackColor = .FontColor
            m_frmTextAnnoSettings.txtAnnoText.Text = .Value
            m_frmTextAnnoSettings.txtRotation.Text = .rotate
        End With
    
        GIS.UpDate
    End If

End Sub

Private Sub m_frmTextAnnoSettings_ResetAll()
    RemoveAllShps m_oDrawLyr
    GIS.UpDate
    ListAllAnnotationShps
End Sub

Private Sub m_frmTextAnnoSettings_UpdateAnnoText()
        '<EhHeader>
        On Error GoTo m_frmTextAnnoSettings_UpdateAnnoText_Err
        '</EhHeader>
        Dim oShp As XGIS_Shape

100     If m_frmTextAnnoSettings.lstTexts.ListIndex = -1 Then Exit Sub

102     m_oDrawLyr.MoveFirst m_oDrawLyr.Extent, "GIS_UID = " & m_frmTextAnnoSettings.lstTexts.ItemData(m_frmTextAnnoSettings.lstTexts.ListIndex), Nothing, "", True
    
104     If Not m_oDrawLyr.Shape Is Nothing Then

106         With m_oDrawLyr.Shape.Params.labels
108             .Value = m_frmTextAnnoSettings.txtAnnoText.Text
110             .rotate = m_frmTextAnnoSettings.txtRotation.Text
                .Color = m_frmTextAnnoSettings.panColorBack.BackColor
                .FontColor = m_frmTextAnnoSettings.panColorFore.BackColor
            End With
    
112         GIS.UpDate
        End If

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

Private Sub m_frmWVIControlPanel_LoadWVIIncidents()
    m_frmWVIIncidents.Show vbModeless, Me
End Sub

Private Sub m_frmWVIControlPanel_LoadWVISitrep()
    Dim ppic As StdPicture
    m_frmWVISitrepGenerator.Show vbModeless, Me
    m_frmWVISitrepGenerator.ShowWait True

    DoEvents
                
    Clipboard.Clear
    GIS.PrintClipboard
    Set ppic = Clipboard.GetData(vbCFEMetafile)
    m_frmWVISitrepGenerator.Init g_sUserName, dxGISDataGrid.Filter.FilterText, ppic
    m_frmWVISitrepGenerator.ShowWait False
End Sub

Private Sub m_frmWVIIncidents_GetLocationOnMap()

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
    m_frmDebug.DebugPrint errDesc
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
   m_frmDebug.DebugPrint "Clipboard changed"
   With ClipBoard_Ext
      lCount = .GetCurrentFormats(Me.hWnd)
      For lFormat = 1 To lCount
         lID = .GetCurrentFormatID(lFormat)
         m_frmDebug.DebugPrint "Clipboard Format:" & .GetCurrentFormatName(lFormat)
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

Private Sub mnuAddToClipboard_Click()
    GIS.viewer.PrintClipboard
    MsgBox "The map has been added to your clipboard....", vbInformation
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

    Debug.Print "Original Center: " & Map1.centerX & ", " & Map1.centerY

    iProjectionType = miRobinson
    iDatumNumber = 62             'North American 1927 (NAD 27)
    iUnits = miUnitMeter
    dOriginLongitude = 0
    Map1.NumericCoordSys.Set iProjectionType, iDatumNumber, iUnits, dOriginLongitude

    Debug.Print "New Center: " & Map1.centerX & ", " & Map1.centerY
End Sub

Private Sub SetDefCoordSys()
        '<EhHeader>
        On Error GoTo SetDefCoordSys_Err
        '</EhHeader>
        Dim ProjectionList As New XGIS_ProjectionList
        Dim DatumLST As New XGIS_DatumList
        
        Exit Sub
        
100     GIS.Projection = ProjectionList.FindEx("GED")
        GIS.Projection.Datum = DatumLST.FindEx("WGE")
102     GIS.Units.Units = XgisUnitsTypeMeter
        GIS.RecalcExtent
        
        '<EhFooter>
        Exit Sub

SetDefCoordSys_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SetDefCoordSys " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CoordTrans(oPt As XGIS_Point) As XGIS_Coordinate
Dim coord As New XGIS_Coordinate
Dim utmProj As New XGIS_ProjectionAbstract
Dim PRList As New XGIS_ProjectionList

coord.Latitude = oPt.x
coord.Longitude = oPt.y
coord.Height = 0

'set PRList.ProjectionList.FindEx("UTM")
'proj.SetUp(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
'
'TGIS_Coordinate c = proj.Projected(coord);
'Lon = c.Longitude;
'Lat = c.Latitude;


End Function

Private Sub mnuGridSettings_Click()
    'TODO Some Optimization stuff
End Sub

Private Sub mnuHideSelected_Click()
        '<EhHeader>
        On Error GoTo mnuHideSelected_Click_Err
        '</EhHeader>
        Dim lL As XGIS_LayerVector

100     SafeMoveFirst g_RSGISGridTableSettings
102     g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"

104     If Not g_RSGISGridTableSettings.EOF Then
106         Set lL = GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
        Else
108         Set lL = GIS.get(abGridTools.Tools.Item("comLyr").Text)
        End If
        
110     If Not lL Is Nothing Then HideShowSelected lL, False
        '<EhFooter>
        Exit Sub

mnuHideSelected_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.mnuHideSelected_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuLoadAll_Click()

        LoadLayerAttrDataToGridInit True

End Sub

Private Sub mnuMainSettings_Click()
        Dim i As Integer

        With m_frmMainSettings
            .ComTipLyr.Clear
102         .ComTipLyr.AddItem "--All--"

            Set m_LyrCol = New Collection
            
            m_LyrCol.Add "--All--", "--All--"
            
104         For i = 0 To GIS.Items.Count - 1
                
106             If GisUtils.IsInherited(GIS.Items.Item(i), "XGIS_LayerVector") Then
                    m_LyrCol.Add GIS.Items.Item(i).Name, GIS.Items.Item(i).caption
108                 .ComTipLyr.AddItem GIS.Items.Item(i).caption
                End If
        
            Next
      
            FindIndexStrEx .ComTipLyr, oMapTipSetting.MapTipLayer
            
            If .ComTipLyr.ListIndex = -1 Then FindIndexStrEx .ComTipLyr, "--All--"
            
            .Show vbModal, Me
            
            picToolTip.Visible = False
            tmrToolTip.Enabled = False
            tmrToolTip.Interval = 1000 * oMapTipSetting.TipDelay
            tmrToolTip.Enabled = oMapTipSetting.Enabled
        End With
    
     With oSelectionStyle
        GIS.SelectionColor = .Color
        GIS.SelectionOutlineOnly = .OutLineOnly
        GIS.SelectionTransparency = .Transparency
        GIS.SelectionWidth = .Width
     End With
    
    If Not m_oSQLIncLyr Is Nothing Then
        With oIncidentLayerSettings
            m_oSQLIncLyr.CachedPaint = .CachedPaint
            '.ConfigFilePAth
            m_oSQLIncLyr.HideFromLegend = .HideFromLegend
            m_oSQLIncLyr.IgnoreShapeParams = False  '.IgnoreShapeParams
            m_oSQLIncLyr.IncrementalPaint = .IncrementalPaint
            '.UseConfig
            '.UseFileParams
            m_oSQLIncLyr.draw
        End With
    End If
    
    With oMapSettings
'        .AutoScroll
'        .iMAPUnits
'        .MapRotation
'        .ScrollBars
'        .StoreLayerParamsInProject
    End With
End Sub

Private Sub mnuMapTip_Click()
    mnuMapTip.Checked = Not mnuMapTip.Checked
    oMapTipSetting.Enabled = mnuMapTip.Checked
End Sub

Private Sub mnuOPSView_Click()

    If mnuOPSView.Checked Then
    
        If MsgScroll.ListCount > 0 Then
            elScroller.Visible = True
            elScroller.Move 5, elTatukGIS.Height - MsgScroll.Height, elTatukGIS.Width, 255
        End If
        
        If Not m_frmOPSView Is Nothing Then Unload m_frmOPSView
        mnuOPSView.Checked = False
        'Call elTatukGIS_RealignFrame
    Else
        
        Set m_frmOPSView = New frmOPSView
        m_frmOPSView.Show
        elScroller.Visible = False
        mnuOPSView.Checked = True
        SetParent GIS.hWnd, m_frmOPSView.elHolder.hWnd
        
    End If

End Sub

Private Sub mnuPanTool_Click()
    GIS.Mode = XgisDrag
End Sub

Private Sub mnuPrevious_Click()
        '<EhHeader>
        On Error GoTo mnuPrevious_Click_Err
        '</EhHeader>
        Dim iDeduct As Integer

100     m_bPrevActionUsed = True
102     GIS.Lock
    
104     If UBound(m_PrevExt) <> LBound(m_PrevExt) Then
106         If GisUtils.GisIsSameExtent(GIS.VisibleExtent, m_PrevExt(UBound(m_PrevExt))) Then
108             iDeduct = 1
            End If
        
110         GIS.VisibleExtent = m_PrevExt(UBound(m_PrevExt) - iDeduct)
112         iDeduct = iDeduct + 1

114         If UBound(m_PrevExt) <> LBound(m_PrevExt) Then ReDim Preserve m_PrevExt(UBound(m_PrevExt) - iDeduct)
        Else
116         GIS.VisibleExtent = m_PrevExt(UBound(m_PrevExt))
        End If
        
118     GIS.Unlock

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
Dim oLyr As XGIS_LayerVector
Dim shp As XGIS_Shape

    Set oLyr = GIS.get(sLayerName)
    
    oLyr.MoveFirst oLyr.Extent, "", Nothing, "", True
    Set shp = oLyr.FindFirst(oLyr.Extent, "GIS_UID = " & uID, Nothing, "", True)

    If Not shp Is Nothing Then
        shp.Flash
    End If
    
End Sub

Private Sub mnuSendToSMS_Click()
    frmSMSMain.Show vbModal
End Sub

Private Sub mnuShowRecordCount_Click()
    Dim sumgroup As DXDBGRIDLibCtl.dxGridSummaryGroup
    Dim sumitem As DXDBGRIDLibCtl.dxGridSummaryItem
    Dim i As Integer
    Dim col As DXDBGRIDLibCtl.dxGridColumn

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
                Set col = .Columns(i)

                If col.Visible Then
                    col.SummaryFooterType = cstCount
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
        Dim lL As XGIS_LayerVector

100     SafeMoveFirst g_RSGISGridTableSettings
102     g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"

104     If Not g_RSGISGridTableSettings.EOF Then
106         Set lL = GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
        Else
108         Set lL = GIS.get(abGridTools.Tools.Item("comLyr").Text)
        End If
        
110     If Not lL Is Nothing Then HideShowSelected lL, True

        '<EhFooter>
        Exit Sub

mnuShowSelected_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.mnuShowSelected_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuVisibleExtent_Click()
    LoadLayerAttrDataToGridInit , True
End Sub

Private Sub mnuZoomIN_Click()
    GIS.zoom = GIS.zoom * 2
End Sub

Private Sub mnuZoomOUT_Click()
    GIS.zoom = GIS.zoom / 2
End Sub

Private Sub mnuZoomRectanletool_Click()
    GIS.Mode = XgisZoom
End Sub

Private Sub m_frmSpatialize_CheckCoordInMap(x As Double, y As Double, sMGRS As String)
   ' GIS.SetFocus
    GIS.ZoomBy 4, x, y, 8
End Sub

Private Sub m_frmSpatialize_FocusRecieved(Name As String)
    'GIS.SetFocus
End Sub

Private Sub m_frmSpatialize_GetCoordFromMapTool()
    
    'GIS.SetFocus
    m_prevTool = g_CurrentTool
    g_CurrentTool = oCreateLocationPoint
    g_CurrentFeatureType = Custom
   ' GIS.SetFocus
End Sub

Private Sub m_frmSpatialize_LocRadiusSelectTool()
  '  GIS.SetFocus
    GIS.Mode = XgisSelect
    g_CurrentTool = oRadiusSelect
   ' GIS.SetFocus
End Sub

Private Sub ShowDlWins()
        '<EhHeader>
        On Error GoTo ShowDlWins_Err
        '</EhHeader>
        Dim sWebsite As String
        'Dim ofrmOASISClientSynch As New frmOASISClientSynch

100     sWebsite = g_sAppServerPath

102     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
104         sWebsite = sWebsite & "/"
        End If
        

106         With frmOASISClientSynch
        
108             .txtServerURL.Text = sWebsite
110             .GetSynchFileFolders
112             Set .dxDataPacks.DataSource = RSDataPacks
114             .dxDataPacks.Columns.RetrieveFields
116             .dxDataPacks.Columns(0).Visible = False
118             .dxDataPacks.Columns(1).Visible = False
120             .dxDataPacks.Columns(5).Visible = False
122             .dxDataPacks.Columns(6).Visible = False
124             .dxDataPacks.Columns(7).Visible = False

126             .ResetDownloadPossibilities
128             frmLog.Show vbModeless, Me
130             .Show vbModal, Me
            End With

            'Unload ofrmOASISClientSynch

132         GIS.UpDate
134         m_bDLDataPacks = False
        
        '<EhFooter>
        Exit Sub

ShowDlWins_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.ShowDlWins " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
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
 
        Dim lShape As New XGIS_Shape
        Dim lLayer As New XGIS_LayerVector
    
100     m_oSQLIncLyr.MoveFirst m_oSQLIncLyr.Extent, "", Nothing, "", True
102     Set lShape = m_oSQLIncLyr.FindFirst(m_oSQLIncLyr.Extent, "ID = '" & MsgScroll.GetKey(Index) & "'", Nothing, "", True)
        
104     If Not m_oSQLIncLyr.Shape Is Nothing Then
    
            'If lShape.Extent Is Nothing Then
    
106             GIS.VisibleExtent = m_oSQLIncLyr.Shape.Extent

                'After changing the extent the shape needs to be found once again
                'this is the workaround for tatuk.
                '
                'It looks like once you use an XGIS_Shape it automatically performs a "movenext" operation

108             m_oSQLIncLyr.MoveFirst m_oSQLIncLyr.Extent, "", Nothing, "", True
110             Set lShape = m_oSQLIncLyr.FindFirst(m_oSQLIncLyr.Extent, "ID = '" & MsgScroll.GetKey(Index) & "'", Nothing, "", True)
112             lShape.Flash 8, 10
            'End If

        Else
114         GIS.Unlock
        End If

116     MsgScroll.ItemColor(Index) = vbBlue
    
        '    'm_frmDebug.DebugPrint  MsgScroll.Item
        '    MsgScroll.AutoScroll = False
        '    MsgScroll.RemoveItem Index
        '    MsgScroll.AutoScroll = True
        '
        '    If MsgScroll.ListCount = 0 Then
        '        MsgScroll.AutoScroll = False
        '        elScroller.Visible = False
        '        GIS.Height = elTatukGIS.Height + 400
        '    End If

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
    'Prepare the Cursor.....
    GIS.CursorPrepare 2048, LoadResPicture(101, vbResCursor) 'LoadCursorFromFile("C:\Documents and Settings\iMMAP\Desktop\Cursores Zelda\immap.ani")
    
    GIS.Mode = XgisUserDefined
    GIS.Cursor = 2048
    

End Sub

Private Sub OASSelect_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print ""
End Sub

Private Sub oClientInterCom_OnDataArrival(bData As String)
    
    'm_frmDebug.DebugPrint bData
    m_frmDebug.DebugPrint bData
    Select Case Left(CStr(bData), 2)
    
        Case "??"
            'm_frmDebug.DebugPrint "INFO ONLY"
        Case "!!"
            'm_frmDebug.DebugPrint bData
        Case "**"
            'm_frmDebug.DebugPrint "Completed Action"
            InterComHandling bData
        Case Else
            
    End Select
    
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
    Dim col As Collection
    Dim Item As Variant
    Dim oRS As adodb.Recordset
    Dim i As Integer
    Dim j As Integer
    Dim sITM() As String

    Select Case Mid$(sComString, 3, 2)
    
        Case "00"
    
        Case "01"
            Set col = oSync.NewIncidents
            
            SetMess "Synchronised items recieved." & " Time: " & GetTimeFormatted

            sITM = Split(Mid$(sComString, 5), "//")
            
            For j = LBound(sITM) To UBound(sITM)

                If Len(sITM(j)) > 1 Then
                    Set oRS = New adodb.Recordset
                    oRS.Open "SELECT ID, Province, District, TYPE, TARGET, Incident_DATE FROM oincidents_FEA WHERE (((oincidents_FEA.Incident_DATE)>Now()-7)) AND [ID] = '" & sITM(j) & "'", m_Cnn, adOpenForwardOnly, adLockReadOnly
                    'oRS.Open "SELECT ID, Province, District, TYPE, TARGET, Incident_DATE FROM oincidents_FEA WHERE [ID] = '" & sITM(j) & "'", m_Cnn, adOpenForwardOnly, adLockReadOnly
                    g_ColAlerts.Add sITM(j) ', CStr(Item)
                    i = i + 1

                    If Not oRS.EOF Then

                        With oRS.Fields
                            MsgScroll.AddItem "New Incident #" & g_ColAlerts.Count & " In: " & .Item("Province").Value & " - " & .Item("District").Value & " - " & .Item("Incident_DATE").Value, sITM(j)
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

                If Not mnuOPSView.Checked Then GIS.Height = elTatukGIS.Height + 150

                If Not elScroller.Visible Then

                    SafeMoveFirst g_RSAppSettings
                    g_RSAppSettings.Find "SettingName = 'Notifier'"

                    If Not g_RSAppSettings.EOF And Not g_RSAppSettings.BOF Then
                        If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").Value) Then
                            MsgScroll.BackColor = g_RSAppSettings.Fields.Item("SettingValue1").Value
                            MsgScroll.ForeColor = g_RSAppSettings.Fields.Item("SettingValue2").Value
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
                    GIS.Height = elTatukGIS.Height + 400
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
            Else
                SetIMess "Offline" ' "No Internet Status: " & m_sInternetCon
            End If

        Case "05"

        Case "06"

        Case "07"

        Case "08"

        Case "09"
        
        Case "??"

            If Len(sComString) > 7 Then SetMess Right(sComString, Len(sComString) - 4)
            
    End Select

End Sub

Private Sub EmulateTicker()
        '<EhHeader>
        On Error GoTo EmulateTicker_Err
        '</EhHeader>
        Dim i As Integer
        Dim oRS As adodb.Recordset

100     SetMess "Synchronised items recieved." & " Time: " & GetTimeFormatted
   
102     Set oRS = New adodb.Recordset
104     oRS.Open "SELECT UID, ID, Province, District, TYPE, TARGET, Incident_DATE FROM oincidents_FEA", m_Cnn, adOpenForwardOnly, adLockReadOnly

        Do While Not oRS.EOF

106         g_ColAlerts.Add oRS.Fields.Item("ID").Value
        
108         i = i + 1

110         With oRS.Fields
112             MsgScroll.AddItem "New Incident #" & i & " In: " & .Item("Province").Value & " - " & .Item("District").Value & " - " & .Item("Incident_DATE").Value, .Item("ID").Value
            End With

            oRS.MoveNext

        Loop

114     SetMess "Synchronised items recieved. #" & i & " Time: " & GetTimeFormatted
   
116     If Not mnuOPSView.Checked Then
118         GIS.Height = elTatukGIS.Height + 150
        End If

120     If Not elScroller.Visible Then

122         SafeMoveFirst g_RSAppSettings
124         g_RSAppSettings.Find "SettingName = 'Notifier'"

126         If Not g_RSAppSettings.EOF And Not g_RSAppSettings.BOF Then
128             If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").Value) Then
130                 MsgScroll.BackColor = g_RSAppSettings.Fields.Item("SettingValue1").Value
132                 MsgScroll.ForeColor = g_RSAppSettings.Fields.Item("SettingValue2").Value
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
        Dim col As Collection
        Dim i As Integer
        Dim oRS As adodb.Recordset

100     Select Case enmRunnerType
    
            Case OASIS_SynchNG.enmPaddy
    
102         Case OASIS_SynchNG.GeoMarks
    
104         Case OASIS_SynchNG.IncidentSynch
106             Set col = oSync.NewIncidents

                SetMess "Synchronised items recieved." & " Time: " & GetTimeFormatted

108             For Each Item In col
110                 Set oRS = New adodb.Recordset
112                 oRS.Open "SELECT ID, Province, District, TYPE, TARGET, Incident_DATE FROM oincidents_FEA WHERE [ID] = '" & Item & "'", m_Cnn, adOpenForwardOnly, adLockReadOnly
114                 g_ColAlerts.Add Item ', CStr(Item)
                    i = i + 1

116                 With oRS.Fields
118                     MsgScroll.AddItem "New Incident #" & i & " In: " & .Item("Province").Value & " - " & .Item("District").Value & " - " & .Item("Incident_DATE").Value, CStr(Item)
                    End With

                    SetMess "Synchronised items recieved. #" & i & " Time: " & GetTimeFormatted
                Next
           
120             If MsgScroll.ListCount > 0 Then
                    If Not mnuOPSView.Checked Then
122                     GIS.Height = elTatukGIS.Height + 150
                    End If
                    
124                 If Not elScroller.Visible Then
                        
                        SafeMoveFirst g_RSAppSettings
                        g_RSAppSettings.Find "SettingName = 'Notifier'"
            
                        If Not g_RSAppSettings.EOF And Not g_RSAppSettings.BOF Then
                            If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").Value) Then
                                MsgScroll.BackColor = g_RSAppSettings.Fields.Item("SettingValue1").Value
                                MsgScroll.ForeColor = g_RSAppSettings.Fields.Item("SettingValue2").Value
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
134                     GIS.Height = elTatukGIS.Height + 400
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
    m_frmDebug.DebugPrint i & "  " & lTotal & "  " & sMess
    SetMess "Working # " & i & " of " & lTotal & " " & sMess
End Sub

Private Sub oSync_WorkerError(errnum As Long, _
                              errDesc As String, _
                              errSource As String)
    SetMess "Synch Error..." & errDesc & " " & errSource & " " & errnum & " Time: " & GetTimeFormatted
End Sub

Private Sub SelAttributes1_OnClick(Index As Integer, translated As Boolean)
    Debug.Print ""
End Sub

Private Sub SelAttributes1_OnEnter_(Index As Integer, translated As Boolean)
    Debug.Print ""
End Sub

Private Sub SelAttributes1_OnMouseDown(Index As Integer, _
                                       translated As Boolean, _
                                       ByVal Button As TatukGIS_DK.XMouseButton, _
                                       ByVal Shift As TatukGIS_DK.XShiftState, _
                                       ByVal x As Long, _
                                       ByVal y As Long)

100     If Button = TatukGIS_DK.XmbRight Then
    
            ' Get the position of the cursor
102         GetCursorPos ptPopUpPos
            'abGridPop.Bands("popGrid").PopupMenu , ScaleX(Point.X, vbPixels, vbTwips) - Me.left, ScaleY(Point.Y, vbPixels, vbTwips) - Me.top

106         PopupMenu mnuSelector, x:=ScaleX(ptPopUpPos.x, vbPixels, vbTwips) - Me.Left, y:=ScaleY(ptPopUpPos.y, vbPixels, vbTwips) - (Me.Top + 250)
  
        End If

End Sub

Private Sub SelAttributes1_OnMouseUp(Index As Integer, translated As Boolean, ByVal Button As TatukGIS_DK.XMouseButton, ByVal Shift As TatukGIS_DK.XShiftState, ByVal x As Long, ByVal y As Long)
    Debug.Print ""
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

'Private Sub strWho1_Addlocation(sID As String)
'        '<EhHeader>
'        On Error GoTo strWho1_Addlocation_Err
'        '</EhHeader>
'100     frmAddLocToW3.ctrWhere1.Init m_Cnn, False, sID, 1
'
'102     frmAddLocToW3.Show vbModal
'104     strWho1.AttachLocationID frmAddLocToW3.ctrWhere1.GetLastAddedLocationID
'106     strWho1.MovePrevious
'108     strWho1.MoveNext
'
'110     Unload frmAddLocToW3
'        '<EhFooter>
'        Exit Sub
'
'strWho1_Addlocation_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.strWho1_Addlocation " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub

Public Sub RunSynch(DummyArgument As Variant)
        '<EhHeader>
        On Error GoTo RunSynch_Err
        '</EhHeader>
        Dim sWebsite As String
        Dim sDataPackGUIDs As String
        Dim oRS As New adodb.Recordset
        Dim oRSGeo As New adodb.Recordset
        Dim RSCurSynchDS As New adodb.Recordset
        Dim RSRemCurSynchDS As New adodb.Recordset
        Dim sSQL As String
        Dim MsXmlHttp As New MSXML2.XMLHTTP
        Dim MsXmlDoc As New MSXML2.DOMDocument
        Dim sString As String
        
100     If Not g_bUseSynch Then Exit Sub
        Exit Sub
        
102     sWebsite = g_sAppServerPath '
        On Error Resume Next
    
104     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
106         sWebsite = sWebsite & "/"
        End If
        
108     If g_sRemoteTablePrefix = "" Then

            sString = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT SettingTablePrefix FROM UserGroups WHERE ID IN (SELECT UserGroupID FROM Users WHERE [user] = '" & g_sUserName & "' AND [pwd] = '" & g_sUserPass & "')")
110         'Set mRSUGSettings = OpenSilentHttpCommsRS(sString, True)
            Set mRSUGSettings = New adodb.Recordset
            mRSUGSettings.Open sString
    
114         If mRSUGSettings.State = 0 Then
                Exit Sub
            End If
            
116         g_sRemoteTablePrefix = mRSUGSettings.Fields.Item("SettingTablePrefix").Value
            
118         If g_sRemoteTablePrefix = "" Then
                Exit Sub
            End If
        
        End If

120     With RSCurSynchDS
122         .CursorLocation = adUseClient
124         .Open "SELECT * FROM SQLSynchLayers", m_Cnn, adOpenDynamic, adLockBatchOptimistic

126         If Not .EOF And Not .BOF Then
128             SafeMoveFirst RSCurSynchDS
130             sSQL = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & .Fields("SQLLyrName").Value & "_FEA")
            Else
132             sSQL = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & .Fields("SQLLyrName").Value)
            End If
    
        End With

        'Set RSRemCurSynchDS = OpenSilentHttpCommsRS(sSQL, True)
        Set RSRemCurSynchDS = New adodb.Recordset
        RSRemCurSynchDS.Open sSQL
    
136     If RSRemCurSynchDS.State = 0 Then
            Exit Sub
        End If

138     With RSCurSynchDS
        
            Dim sFields As String
            Dim h As Integer
        
140         If Not RSRemCurSynchDS.EOF And Not RSRemCurSynchDS.BOF Then
142             SafeMoveFirst RSRemCurSynchDS

                'If Not RSRemCurSynchDS.EOF Then
144             Do While Not RSRemCurSynchDS.EOF

146                 If Not .EOF And Not .BOF Then
148                     SafeMoveFirst RSCurSynchDS
150                     .Find "GUID = '" & RSRemCurSynchDS.Fields.Item("GUID").Value & "'"
                    End If
                
152                 If Not .EOF Then  'Updated GeoMark

154                     For h = 1 To RSRemCurSynchDS.Fields.Count - 1
156                         .Fields.Item(h).Value = RSRemCurSynchDS.Fields.Item(h).Value
158                         .UpDate
                        Next

                    Else 'New Geo Mark
160                     .AddNew 'Split(sFields, ","), Split(RSRemCurSynchDS.GetString(, , ","), ",")

162                     For h = 1 To RSRemCurSynchDS.Fields.Count - 1
164                         .Fields.Item(h).Value = RSRemCurSynchDS.Fields.Item(h).Value
                        Next

166                     .Fields.Item(.Fields.Count - 1).Value = True
168                     .UpDate
                    End If
                
170                 RSRemCurSynchDS.MoveNext
                Loop

            Else 'ONLY NEW ITEMS ON CLIEN>T SIDE!

172             If Not .BOF Then
174                 SafeMoveFirst RSCurSynchDS
                End If
                
176             oRS.Open "SELECT * FROM " & .Fields("SQLLyrName").Value & "_FEA", m_Cnn
178             oRSGeo.Open "SELECT * FROM " & .Fields("SQLLyrName").Value & "_GEO", m_Cnn
                
180             Do While Not oRS.EOF
                    
182                 RSRemCurSynchDS.AddNew
                    
184                 For h = 0 To oRS.Fields.Count - 2
186                     RSRemCurSynchDS.Fields.Item(h).Value = oRS.Fields.Item(h).Value
                    Next

188                 RSRemCurSynchDS.Fields.Item(oRS.Fields.Count - 1).Value = 1
190                 oRS.MoveNext
                Loop
            
192             RSRemCurSynchDS.Filter = adFilterPendingRecords
            
194             If Not RSRemCurSynchDS.EOF And Not RSRemCurSynchDS.BOF Then
196                 MsXmlHttp.Open "POST", sWebsite & "oasis.asp", 0
198                 RSRemCurSynchDS.Save MsXmlDoc, 1
200                 MsXmlHttp.Send MsXmlDoc

                    'SaveSilentHttpCommsRS RSRemCurSynchDS, sWebsite & "oasis.asp", True

                End If
   
204             sSQL = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & .Fields("SQLLyrName").Value & "_GEO")
206             'Set RSRemCurSynchDS = OpenSilentHttpCommsRS(sSQL, True)
                Set RSRemCurSynchDS = New adodb.Recordset
                RSRemCurSynchDS.Open sSQL
                
208             Do While Not oRSGeo.EOF
                    
210                 RSRemCurSynchDS.AddNew
                    
212                 For h = 0 To oRSGeo.Fields.Count - 1
214                     RSRemCurSynchDS.Fields.Item(h).Value = oRSGeo.Fields.Item(h).Value
                    Next

                    'RSRemCurSynchDS.Fields.Item(oRS.Fields.Count - 1).Value = 1
216                 oRSGeo.MoveNext
                Loop
            
218             RSRemCurSynchDS.Filter = adFilterPendingRecords
            
220             If Not RSRemCurSynchDS.EOF And Not RSRemCurSynchDS.BOF Then
222                 MsXmlHttp.Open "POST", sWebsite & "oasis.asp", 0
224                 RSRemCurSynchDS.Save MsXmlDoc, 1
226                 MsXmlHttp.Send MsXmlDoc

                    'SaveSilentHttpCommsRS RSRemCurSynchDS, sWebsite & "oasis.asp", True

                End If
            
                'End If
            End If

228         .UpdateBatch
        End With

        On Error Resume Next
230     mRSUGSettings.Close
232     RSCurSynchDS.Close
234     RSRemCurSynchDS.Close

236     Set mRSUGSettings = Nothing
238     Set RSCurSynchDS = Nothing
240     Set RSRemCurSynchDS = Nothing

        '<EhFooter>
        Exit Sub

RunSynch_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.RunSynch " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub tmrDataPackCheck_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    If m_bDLDataPacks Then
        m_bDLDataPacks = False
        tmrDataPackCheck.Enabled = False
        ShowDlWins
    End If

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
    
            m_sInternetCon = "Internet Available for Synchronisation Through: " & m_sInternetCon
        
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'StartSynchWorker
            '
            If g_lHw < 1 Then
        
                m_frmDebug.DebugPrint "Started Running Synch... " & " Time: " & GetTimeFormatted
                SetMess "Started Running Synch... " & " Time: " & GetTimeFormatted
    
                g_lHw = GetDesktopWindow()
                ShellExecute g_lHw, vbNullString, App.Path & "\bin\OASISCommsMon.exe", "3^" & m_Cnn.ConnectionString & "^" & g_sRemoteTablePrefix & "^" & g_sAppServerPath & "^" & g_bHasEncrypt & "^" & g_sKey & "^" & Me.hWnd & "^" & tmrInternetCheck.Interval & "^" & g_udtSynchUpdateOptions.GeoMarks & "^" & g_udtSynchUpdateOptions.SynchLayersSettings, "C:\", 1
            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Else
            m_sInternetCon = "Internet Not Available for Synchronisation: " & m_sInternetCon
        End If

        Me.caption = "OASIS Client Version: " & App.major & "." & App.minor & "." & App.Revision & " " & App.Comments '& " " & m_sInternetCon
    
    End If

End Sub

Public Sub CheckSynchFeedStatus(DummyVariant As Variant)
        '<EhHeader>
        On Error GoTo CheckSynchFeedStatus_Err
        '</EhHeader>
        Dim sWebsite As String
        Dim sDataPackGUIDs As String
        Dim RSUGSettings As New adodb.Recordset
        Dim RSCurSynchTables As New adodb.Recordset
        Dim RSSerSynchTables As New adodb.Recordset
        Dim sSQL As String
        Dim h As Integer
        Dim sString As String
           Dim MsXmlHttp As New MSXML2.XMLHTTP
       Dim MsXmlDoc As New MSXML2.DOMDocument
        On Error Resume Next
    
100     With RSCurSynchTables
            
102         .Open "SELECT * FROM SynchTables", m_Cnn, adOpenDynamic, adLockOptimistic

104         If Not .EOF And Not .BOF Then
106             SafeMoveFirst RSCurSynchTables
108             sWebsite = g_sAppServerPath '
    
110             If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
112                 sWebsite = sWebsite & "/"
                End If
            
                sString = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT SettingTablePrefix FROM UserGroups WHERE ID IN (SELECT UserGroupID FROM Users WHERE [user] = '" & g_sUserName & "' AND [pwd] = '" & g_sUserPass & "')")
                'Set RSUGSettings = OpenSilentHttpCommsRS(sString, True)
                Set RSUGSettings = New adodb.Recordset
                RSUGSettings.Open sString
    
116             If RSUGSettings.State = 0 Then
                    Exit Sub
                End If
             
118             sSQL = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & RSUGSettings.Fields.Item("SettingTablePrefix").Value & "SynchTables") ' AND WHERE ID IN (SELECT UserGroupID FROM Users WHERE user = '" & g_sUserName & "' AND pwd = '" & g_sUserPass & "')"
                'Set RSSerSynchTables = OpenSilentHttpCommsRS(sSQL, True)
                Set RSSerSynchTables = New adodb.Recordset
                RSSerSynchTables.Open sSQL
                
122             If RSSerSynchTables.State = 0 Then
                    Exit Sub
                End If
            
124             Do While Not .EOF
126                 RSSerSynchTables.AddNew

128                 For h = 1 To RSSerSynchTables.Fields.Count - 1
130                     RSSerSynchTables.Fields.Item(h).Value = .Fields.Item(h).Value
                    Next

132                 .Fields.Item(.Fields.Count - 1).Value = True
134                 .UpDate
136                 .MoveNext
                Loop

                'SaveSilentHttpCommsRS RSSerSynchTables, sWebsite & "oasis.asp", True
                MsXmlHttp.Open "POST", sWebsite & "oasis.asp", 0
                RSSerSynchTables.Save MsXmlDoc, 1
                MsXmlHttp.Send MsXmlDoc
144             RSSerSynchTables.UpdateBatch

            End If

        End With
    
        '<EhFooter>
        Exit Sub

CheckSynchFeedStatus_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.CheckSynchFeedStatus " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub RunFeedSynch(feedName As Variant)
        '<EhHeader>
        On Error GoTo RunFeedSynch_Err
        '</EhHeader>
        Dim rsSynch As adodb.Recordset
    
100     Set rsSynch = New adodb.Recordset
102     rsSynch.Open "SELECT * FROM FeedsHistory", m_Cnn, adOpenDynamic, adLockOptimistic
104     Set rsSynch = Nothing
    
        '<EhFooter>
        Exit Sub

RunFeedSynch_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.RunFeedSynch " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub CheckDataPacks(DummyArgument As Variant)
        '<EhHeader>
        On Error GoTo CheckDataPacks_Err
        '</EhHeader>
        Dim sSQL As String

        '100     g_RSAppSettings.MoveFirst
        '102     g_RSAppSettings.Find "SettingName = 'DataPacksDownloads'"
        '
        '104     m_frmDebug.DebugPrint  ""
        '
        '106     If g_RSAppSettings.EOF Then
        '108         sSQL = "INSERT INTO AppSettings (SettingName, SettingValue1, SettingValue2, SettingValue3, SettingValue4, SettingValue5"
        '110         sSQL = sSQL & ", SettingValue6, SettingValue7, SettingValue8, SettingValue9, SettingValue10)"
        '112         sSQL = sSQL & " VALUES ('DataPacksDownloads',' ',' ',' ',' ',' ',' ',' ',' ',' ',' ')"
        '
        '114         m_Cnn.Execute sSQL
        '116         sSQL = ""
        '            g_RSAppSettings.Requery
        '        End If
        '
        '117     g_RSAppSettings.MoveFirst
        '        g_RSAppSettings.Find "SettingName = 'DataPacksDownloads'"
        '
        '118     With g_RSAppSettings.Fields
        '120         sSQL = "WHERE " & IIf(Len(.Item("SettingValue1").Value) > 4, .Item("SettingValue1").Value, "")
        '122         If Len(sSQL) > 8 Then sSQL = sSQL & " AND "
        '124         sSQL = sSQL & IIf(Len(.Item("SettingValue2").Value) > 4, .Item("SettingValue2").Value, "")
        '
        '126         If Len(sSQL) > 8 Then sSQL = sSQL & " AND "
        '128         sSQL = sSQL & IIf(Len(.Item("SettingValue3").Value) > 4, .Item("SettingValue3").Value, "")
        '
        '130         If Len(sSQL) > 8 Then sSQL = sSQL & " AND "
        '132         sSQL = sSQL & IIf(Len(.Item("SettingValue4").Value) > 4, .Item("SettingValue4").Value, "")
        '
        '134         If Len(sSQL) > 8 Then sSQL = sSQL & " AND "
        '136         sSQL = sSQL & IIf(Len(.Item("SettingValue5").Value) > 4, .Item("SettingValue5").Value, "")
        '
        '138         If Len(sSQL) > 8 Then sSQL = sSQL & " AND "
        '140         sSQL = sSQL & IIf(Len(.Item("SettingValue6").Value) > 4, .Item("SettingValue6").Value, "")
        '
        '142         If Len(sSQL) > 8 Then sSQL = sSQL & " AND "
        '144         sSQL = sSQL & IIf(Len(.Item("SettingValue7").Value) > 4, .Item("SettingValue7").Value, "")
        '
        '146         If Len(sSQL) > 8 Then sSQL = sSQL & " AND "
        '148         sSQL = sSQL & IIf(Len(.Item("SettingValue8").Value) > 4, .Item("SettingValue8").Value, "")
        '
        '150         If Len(sSQL) > 8 Then sSQL = sSQL & " AND "
        '152         sSQL = sSQL & IIf(Len(.Item("SettingValue9").Value) > 4, .Item("SettingValue9").Value, "")
        '
        '154         If Len(sSQL) > 8 Then sSQL = sSQL & " AND "
        '156         sSQL = sSQL & IIf(Len(.Item("SettingValue10").Value) > 4, .Item("SettingValue10").Value, "")
        '
        '
        '        End With
        '        'Clean Up Where Clause
        '
        '        If Len(sSQL) < 12 Then sSQL = ""
    
        Dim sWebsite As String
        Dim sDataPackGUIDs As String
        Dim RSUGSettings As New adodb.Recordset
        Dim RSCurDataPacks As New adodb.Recordset
        Dim sString As String
        
100     m_bDLDataPacks = False
    
102     Set RSDataPacks = New adodb.Recordset
    
104     sWebsite = g_sAppServerPath ' "http://www.immap.org/"

106     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
108         sWebsite = sWebsite & "/"
        End If
        
        On Error Resume Next
        
        sString = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT * FROM Users WHERE [user] = '" & g_sUserName & "' AND [pwd] = '" & g_sUserPass & "'")
        'Set RSUGSettings = OpenSilentHttpCommsRS(sString, True)
        Set RSUGSettings = New adodb.Recordset
        RSUGSettings.Open sString
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If this thread trys to communicate to the server when
        'there is communication already underway then the recordset
        'returned = nothing and hence the following "if" clause is
        'necessary in order to prevent the thread from breaking the
        'application
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If RSUGSettings Is Nothing Then
            Exit Sub
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
112     If RSUGSettings.State = 0 Then
            Exit Sub
        End If
        
        On Error GoTo CheckDataPacks_Err

        'MsgBox "TODO add Where Clause for Existing Datasets... Also make sure Dynamic Sata Pack server is availble..."
        
114     With RSCurDataPacks
            
116         .Open "SELECT GUID FROM DataPacks", m_Cnn, adOpenDynamic, adLockReadOnly

118         If Not .EOF Then
120             SafeMoveFirst RSCurDataPacks

122             Do While Not .EOF
            
124                 If Len(sDataPackGUIDs) > 5 Then
126                     sDataPackGUIDs = sDataPackGUIDs & " AND NOT GUID = '" & .Fields("GUID").Value & "'"
                    Else
128                     sDataPackGUIDs = " AND NOT GUID = '" & .Fields("GUID").Value & "'"
                    End If
                
130                 .MoveNext
                Loop

            End If
        
        End With
        
132     sDataPackGUIDs = sDataPackGUIDs & " AND NOT Update = false"

        sString = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT * FROM DataPacks WHERE UserGroupID = " & RSUGSettings.Fields.Item("UserGroupID").Value & sDataPackGUIDs)
        'Set RSDataPacks = OpenSilentHttpCommsRS(sString, True)
        Set RSDataPacks = New adodb.Recordset
        RSDataPacks.Open sString
  
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'If this thread trys to communicate to the server when
        'there is communication already underway then the recordset
        'returned = nothing and hence the following "if" clause is
        'necessary in order to prevent the thread from breaking the
        'application
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If RSDataPacks Is Nothing Then
            Exit Sub
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
138     If Not RSDataPacks.EOF And Not RSDataPacks.BOF Then
140         If MsgBox("There are OASIS Data Packs Available for Download." & vbCrLf & "Do you wish to review & download these?", vbYesNo, "OASIS Synch modules") = vbYes Then
                
142             m_bDLDataPacks = True
                
                '134             frmOASISClientSynch.txtServerURL.Text = sWebSite
                '136             frmOASISClientSynch.GetSynchFileFolders
                '                Set frmOASISClientSynch.dxDataPacks.DataSource = RSDataPacks
                '138             frmOASISClientSynch.dxDataPacks.Columns.RetrieveFields
                '140             frmOASISClientSynch.dxDataPacks.Columns(0).Visible = False
                '142             frmOASISClientSynch.dxDataPacks.Columns(1).Visible = False
                '144             frmOASISClientSynch.dxDataPacks.Columns(5).Visible = False
                '146             frmOASISClientSynch.dxDataPacks.Columns(6).Visible = False
                '148             frmOASISClientSynch.dxDataPacks.Columns(7).Visible = False
                '
                '150             frmOASISClientSynch.ResetDownloadPossibilities
                '                frmLog.Show vbModeless, Me
                '152             frmOASISClientSynch.Show vbModal, Me
                '154             GIS.Update
            End If
        End If
    
        '<EhFooter>
        Exit Sub

CheckDataPacks_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.CheckDataPacks " & "at line " & Erl
        Resume Next
        '</EhFooter>
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
    
102     me2Long.Text = "0" & CStr(Left(x, Len(me2Long.Text) - 1))
104     me2Lat.Text = CStr(Left(y, Len(me2Long.Text) - 1))
    
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
        Dim de As Single, mi As Single, se As Single, Lt As String
        Dim Temp1 As Single, temp2 As Single, temp3 As Single
    
100     Select Case FrameNumber

            Case 1 ' dd:mm:ss.ss
102             Lt = Replace(me1Lat.Text, "_", "0")
104             de = Val(Left$(Lt, 2))
106             mi = Val(Mid$(Lt, 4, 2))
108             se = Val(Right$(Lt, 5))

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
130             Lt = Replace(me2Lat.Text, "_", "0")
132             de = Val(Lt)
134             mi = 0
136             se = 0
            
138             If de > 59.99999! Then
140                 MsgBox "Degrees must be from 0 to 59.99999", , "Entry Error"
142                 me2Lat.SetFocus
                    Exit Sub
                End If
        
144         Case 3 ' dd:mm.mmmm
146             Lt = Replace(me3Lat.Text, "_", "0")

148             de = Val(Left$(Lt, 2))
150             mi = Val(Right$(Lt, 7))
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
        Dim de As Single, mi As Single, se As Single, Lt As String
        Dim Temp1 As Single, temp2 As Single, temp3 As Single
    
100     Select Case FrameNumber
            Case 1 ' ddd:mm:ss.ss
102             Lt = Replace(me1Long.Text, "_", "0")
104             de = Val(Left$(Lt, 3))
106             mi = Val(Mid$(Lt, 5, 2))
108             se = Val(Right$(Lt, 5))

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
130             Lt = Replace(me2Long.Text, "_", "0")
132             de = Val(Lt)
134             mi = 0
136             se = 0
            
138             If de > 180! Then
140                 MsgBox "Degrees must be from 0 to 179.99999", , "Entry Error"
142                 me2Long.SetFocus
                    Exit Sub
                End If
        
144         Case 3 ' ddd:mm.mmmm
146             Lt = Replace(me3Long.Text, "_", "0")

148             de = Val(Left$(Lt, 3))
150             mi = Val(Right$(Lt, 7))
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

                If IsNull(g_RSAppSettings.Fields.Item("SettingValue1").Value) Then Exit Sub

106             sTools = Split(g_RSAppSettings.Fields.Item("SettingValue1").Value, ",")


                Dim tl As ActiveBar3LibraryCtl.Tool
                
                'Activate Again In version 2.5
                AB.Bands.Item("tbLayer").Tools.Remove "btnPrint"
                         
108             For Each tl In AB.Bands.Item("tbLayer").Tools
110                 If InStr(g_RSAppSettings.Fields.Item("SettingValue1").Value, tl.Name) Then
112                     m_frmDebug.DebugPrint "add: " & tl.Name
                    Else
114                     m_frmDebug.DebugPrint "remove: " & tl.Name
116                     AB.Bands.Item("tbLayer").Tools.Remove tl.Name
                    End If
                Next

118             For Each tl In AB.Bands.Item("tbExtents").Tools
120                 If InStr(g_RSAppSettings.Fields.Item("SettingValue2").Value, tl.Name) Then
122                     m_frmDebug.DebugPrint "add: " & tl.Name
                    Else
124                     m_frmDebug.DebugPrint "remove: " & tl.Name
126                     AB.Bands.Item("tbExtents").Tools.Remove tl.Name
                    End If
                Next

128             For Each tl In AB.Bands.Item("sdToolBar").Tools
130                 If InStr(g_RSAppSettings.Fields.Item("SettingValue3").Value, tl.Name) Then
132                     m_frmDebug.DebugPrint "add: " & tl.Name
                    Else
134                     m_frmDebug.DebugPrint "remove: " & tl.Name
136                     AB.Bands.Item("sdToolBar").Tools.Remove tl.Name
                    End If
                Next
                
138             For Each tl In AB.Bands.Item("tbUtils").Tools
140                 If InStr(g_RSAppSettings.Fields.Item("SettingValue4").Value, tl.Name) Then
142                     m_frmDebug.DebugPrint "add: " & tl.Name
                    Else
144                     m_frmDebug.DebugPrint "remove: " & tl.Name
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



Private Sub ActivateOperations()
        '<EhHeader>
        On Error GoTo ActivateOperations_Err
        '</EhHeader>
        Dim i As Integer
        
100     AB.ClientAreaControl = elMap
102     elMap.Visible = True
            
        'Tool Panels
104     SafeMoveFirst g_RSAppSettings
106     g_RSAppSettings.Find "SettingName = 'MainMapTools'"
        
108     AB.Bands.Item("sdToolBar").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue1").Value) = 1, True, False)
110     AB.Bands.Item("tbExtents").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue2").Value) = 1, True, False)
112     AB.Bands.Item("tbLayer").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue3").Value) = 1, True, False)
114     AB.Bands.Item("bToolbarStyle").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue4").Value) = 1, True, False)
116     AB.Bands.Item("tbUtils").Visible = IIf(Trim$(g_RSAppSettings.Fields("SettingValue5").Value) = "1", True, False)
            
        '                'AvailableMapTools
        '126             g_RSAppSettings.MoveFirst
        '128             g_RSAppSettings.Find "SettingName = 'AvailableMapTools'"
        '
        '                Dim sTools() As String
        '
        '130             sTools = Split(g_RSAppSettings.Fields.Item("SettingValue1").Value, ",")
        '
        '
        '                Dim tl As ActiveBar3LibraryCtl.Tool
        '
        '                'm_frmDebug.DebugPrint  "ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ  Layer " & g_RSAppSettings.Fields.Item("SettingValue1").Value
        '
        '                For Each tl In AB.Bands.Item("tbLayer").Tools
        '                    If InStr(g_RSAppSettings.Fields.Item("SettingValue1").Value, tl.Name) Then
        '                        m_frmDebug.DebugPrint  tl.Name
        '                    Else
        '                        m_frmDebug.DebugPrint  tl.Name
        '                        AB.Bands.Item("tbLayer").Tools.Remove tl.Name
        '                    End If
        '                Next
        '
        '               'm_frmDebug.DebugPrint  "ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ  EXTENTS " & g_RSAppSettings.Fields.Item("SettingValue2").Value
        '                For Each tl In AB.Bands.Item("tbExtents").Tools
        '                    If InStr(g_RSAppSettings.Fields.Item("SettingValue2").Value, tl.Name) Then
        '                        m_frmDebug.DebugPrint  tl.Name
        '                    Else
        '                        m_frmDebug.DebugPrint  tl.Name
        '                        AB.Bands.Item("tbExtents").Tools.Remove tl.Name
        '                    End If
        '                Next
        '
        '                'm_frmDebug.DebugPrint  " ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ STANDARD TOOLS " & g_RSAppSettings.Fields.Item("SettingValue3").Value
        '
        '                For Each tl In AB.Bands.Item("sdToolBar").Tools
        '                    If InStr(g_RSAppSettings.Fields.Item("SettingValue3").Value, tl.Name) Then
        '                        m_frmDebug.DebugPrint  tl.Name
        '                    Else
        '                        m_frmDebug.DebugPrint  tl.Name
        '                        AB.Bands.Item("sdToolBar").Tools.Remove tl.Name
        '                    End If
        '                Next
        '
        '                For Each tl In AB.Bands.Item("tbUtils").Tools
        '                    If InStr(g_RSAppSettings.Fields.Item("SettingValue4").Value, tl.Name) Then
        '                        m_frmDebug.DebugPrint  tl.Name
        '                    Else
        '                        m_frmDebug.DebugPrint  tl.Name
        '                        AB.Bands.Item("tbUtils").Tools.Remove tl.Name
        '                    End If
        '                Next
        '
        '
        ''                For i = 0 To AB.Bands.Item("tbLayer").Tools.Count - 1
        ''                    If InStr(g_RSAppSettings.Fields.Item("SettingValue1").Value, AB.Bands.Item("tbLayer").Tools.Item(i).Name) Then
        ''
        ''                    Else
        ''                        AB.Bands.Item("tbLayer").Tools.Remove AB.Bands.Item("tbLayer").Tools.Item(i).Name
        ''                        'AB.Bands.Item("tbLayer").Tools.Item(i).Visible = False
        ''                    End If
        ''                Next
        '
        '
        ''                For i = 0 To AB.Bands.Item("tbExtents").Tools.Count - 1
        ''                    AB.Bands.Item("tbExtents").Tools.Item(i).Visible = False
        ''                Next
        ''
        ''                For i = 0 To AB.Bands.Item("sdToolBar").Tools.Count - 1
        ''                    AB.Bands.Item("sdToolBar").Tools.Item(i).Visible = False
        ''                Next
        '
        '                'Dim sTools() As String
        ''
        ''                sTools = Split(g_RSAppSettings.Fields.Item("SettingValue1").Value, ",")
        ''
        ''                For i = LBound(sTools) To UBound(sTools)
        ''                    AB.Bands.Item("tbLayer").Tools.Item(sTools(i)).Visible = True
        ''                Next
        ''
        ''                sTools = Split(g_RSAppSettings.Fields.Item("SettingValue2").Value, ",")
        ''
        ''                For i = LBound(sTools) To UBound(sTools)
        ''                    AB.Bands.Item("tbExtents").Tools.Item(sTools(i)).Visible = True
        ''                Next
        ''
        ''                sTools = Split(g_RSAppSettings.Fields.Item("SettingValue3").Value, ",")
        ''
        ''                For i = LBound(sTools) To UBound(sTools)
        ''                    AB.Bands.Item("sdToolBar").Tools.Item(sTools(i)).Visible = True
        ''                Next
        ''
        ''                'sdToolBar=btnZoomin,btnZoomout,btnZoom,btnPan,btnSelect,btnDeselect,btnInfo 'tbLayer=btnAddLyr,btnRemoveLyr,btnLyrSettings,btnLegend,btnPrint'tbExtents=btnFullExtent,btnLayerExtent,btnSelectionExtent,btnPrevExtent,btnRefreshMap
        '
        '
        '                'Tool Panels
        '148             g_RSAppSettings.MoveFirst
        '150             g_RSAppSettings.Find "SettingName = 'MainMapTools'"
        '
        '152             AB.Bands.Item("sdToolBar").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue1").Value) = 1, True, False)
        '154             AB.Bands.Item("tbExtents").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue2").Value) = 1, True, False)
        '156             AB.Bands.Item("tbLayer").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue3").Value) = 1, True, False)
        '158             AB.Bands.Item("bToolbarStyle").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue4").Value) = 1, True, False)
        '
        '                AB.Bands.Item("tbUtils").Visible = IIf(Trim$(g_RSAppSettings.Fields("SettingValue5").Value) = "1", True, False)
        '
        '                'AB.Bands.Item("sdToolBar").Visible = True
        '                'AB.Bands.Item("tbExtents").Visible = True
        '                'AB.Bands.Item("tbLayer").Visible = True
                
118     If Not m_bMapInitialized Then
            'If Not bDebugMode Then
            '    If Not pGetIncThread.IsThreadRunning Then
            '        pGetIncThread.CreateWin32Thread Me, "GetNewIncidents", 1
            '    End If
            'End If
                    
120         If m_bInternetCon Then

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
142         InitMap
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
    
110     elRSSTool.Visible = False
112     elMap.Visible = False
114     elOasisProfile.Visible = False
116     elW3.Visible = False
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
138                     If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").Value) Then
140                         WebBrowser1.Navigate2 g_RSAppSettings.Fields.Item("SettingValue1").Value
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
                '   Debug.Print m_frmDynamicContent.Visible
                
156             If Len(RSSBrowser1.ServerURL) = 0 Then
                    RSSBrowser1.Visible = True
158                 RSSBrowser1.ServerURL = g_sAppServerPath & "/oasis.asp"
160                 RSSBrowser1.UserGroupPrefix = g_sRemoteTablePrefix
162                 RSSBrowser1.Init2 m_Cnn.ConnectionString, g_bFeedUpdate
                End If
                               
164         Case "cbDynamicData"
166             AB.ClientAreaControl = elDynamicData
168             elDynamicData.Visible = True
             
170             If Not m_frmMnuDynamicDataModule.listDatabases.ListCount > 0 Then
172                 DynamicDataModule1.Init m_frmMnuDynamicDataModule.listDatabases, m_frmMnuDynamicDataModule.lstDataElements
        
                End If
                
174         Case "cbReports"
176             AB.ClientAreaControl = elDynamicReports
178             elDynamicReports.Visible = True
             
180             If Not m_frmMnuDynamicReportsModule.Combo1.ListCount > 0 Then
             
182                 DynamicDataReports1.Init m_frmMnuDynamicReportsModule.listQueries, m_frmMnuDynamicReportsModule.Combo1
         
                End If
             
184         Case "cbW3"
                '190             AB.ClientAreaControl = elW3
                '192             elW3.Visible = True
                '
                '194             dxW3DBGrid.Dataset.ADODataset.ConnectionString = m_Cnn.ConnectionString
                '196             dxW3DBGrid.Dataset.Active = True
                '198             dxW3DBGrid.Dataset.ADODataset.Requery
                '
                '                'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OASIS\W3Import.mdb;Persist Security Info=False
                '                'Dim cn As New ADODB.Connection
                '
                '                'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OASIS\W3Import.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                '
                '200             If Not frmW3Wizard.HasInitialized Then
                '202                 ctrWhat1.Init m_Cnn, Me.hwnd
                '204                 ctrWhere1.Init m_Cnn, False
                '206                 strWho1.Init m_Cnn, True, False
                '208                 m_frmW3Wizard.Init m_Cnn
                '                End If
                '
                '210             m_frmW3Wizard_MyOrganization

186         Case "cbSync", "cbJournal"
        
188             MsgBox "This functionality is currently not available for your profile." & vbCrLf & "Please contact your OASIS system administrator if you want to activate these modules.", vbInformation, "OASIS Client Support"

                'AB.ClientAreaControl = CalendarControl
                'CalendarControl.Visible = True
                'GIS.Visible = False
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
        
102     Set m_frmAddWhere = New frmAddWhere
104     Set m_frmW3Wizard = New frmW3Wizard
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
116     AB.Bands("bNavPane").ChildBands("cbProfile").Tools("frmMail").Custom = m_frmMnuOASISProfile 'm_frmAddWhere
118     AB.Bands("bNavPane").ChildBands("cbOperations").Tools("frmCalendar").Custom = m_frmMnuOperations
120     AB.Bands("bNavPane").ChildBands("cbContent").Tools("frmContacts").Custom = m_frmDynamicContent
        'AB.Bands("bNavPane").ChildBands("cbSync").Tools("frmTasks").Custom = frmAddIncident
    AB.Bands("bNavPane").ChildBands("cbDynamicData").Tools("frmDynamicData").Custom = m_frmMnuDynamicDataModule
    AB.Bands("bNavPane").ChildBands("cbReports").Tools("frmReports").Custom = m_frmMnuDynamicReportsModule
    
122     AB.Bands("bNavPane").ChildBands("cbW3").Tools("frmW3").Custom = m_frmW3Wizard
    
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
102     m_frmSpatialAnalysis.SetGISViewer GIS.viewer
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
        Dim oSQLLyr As New XGIS_LayerSqlAdo
        
        m_bPolyG = False
        m_bPolyL = False

        If m_SelUID > 0 Then
            m_oDrawLyr.Delete m_SelUID
            m_SelUID = 0
        End If
        
100     g_CurrentTool = oZoom

102     Select Case Tool.Name

            Case Is = "btnSearchLayers"

                If m_frmSearch Is Nothing Then
                    Set m_frmSearch = New frmSearch
                End If
                
                m_frmSearch.Init GIS.viewer
                
                m_frmSearch.Show vbModeless, Me

            Case Is = "btnAddAnnotation"

104             If m_frmTextAnnoSettings Is Nothing Then Set m_frmTextAnnoSettings = New frmTextAnnoSettings
                
106             ListAllAnnotationShps
108             m_frmTextAnnoSettings.Show vbModeless, Me
110             g_CurrentTool = oCreateLocationText
112             GIS.Mode = XgisSelect

114         Case Is = "btnCharting"
            
116             frmOASISChartOCTfiles.Show vbModeless, Me
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
118         Case Is = "btnAdminLocator"
120             m_frmLocator.Show vbModeless, Me
122             m_frmLocator.Init GIS.viewer

124         Case Is = "mnuResetMap"
126             GIS.FullExtent
                
128         Case Is = "btnZoomin"
130             GIS.zoom = GIS.zoom * 2

132         Case Is = "btnZoomout"
134             GIS.zoom = GIS.zoom / 2

136         Case Is = "btnZoom"
138             GIS.Mode = XgisZoomEx

140         Case Is = "btnZoomRect"
142             GIS.Mode = XgisZoom

144         Case Is = "btnPan"
146             GIS.Mode = XgisDrag

148         Case Is = "btnSelect"
150             GIS.Mode = XgisSelect

152         Case Is = "mnuAddLayer"
154             AddLayer
                
156         Case "mnuFile"

158         Case "mnuEdit"

160         Case "mnuView"

162         Case "mnuTools"

164         Case "mnuHelp"

166         Case "mnuMaps"

168         Case "mnuabout"
            
170         Case "btnCannedReports"
                'Case "btnDynReports"
                Set m_frmCannedReports = New frmCannedReports
                m_frmCannedReports.Init True
                m_frmCannedReports.Show vbModal, Me
                Set m_frmCannedReports = Nothing

176         Case "btnDynamicDataEntry"
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

190         Case "btnSpatialAnalysis"
            
192             Call OpenSpatialAnalysis

            Case "btnOASISv1Charts"
                'MsgBox "to be implemented"
                frmChartSettings.Show vbModeless, Me

194         Case "mnuCoorpWebSite"
            
196             ShellExecute Me.hWnd, vbNullString, "mailto:support@immap.org", vbNullString, "C:\", SW_SHOWNORMAL

198         Case "mnuOASISWebSite"

200         Case "mnuHelpGo"

202         Case "mnuOnlineSupport"

204         Case "mnuRemoveLayer"

206         Case "mnuSaveSession"

208         Case "mnuSaveMap"

210         Case "mnuOpenMap"

212         Case "mnuOpenSession"

214         Case "btnZoom"

216         Case "btnSelect"

218         Case "btnDeselect"

220         Case "btnAddLyr"
222             AddLayer
224             FillCOPValues
226             frmLegend.UpDate

228         Case "btnRemoveLyr"
230             frmLayerSelection.Init GIS, True
232             frmLayerSelection.Show vbModal, Me

234             If frmLayerSelection.GetItem <> "" Then
236                 If MsgBox("Do you really want to remove map layer: " & frmLayerSelection.GetItem, vbYesNo, "OASIS Client") = vbYes Then
238                     GIS.Delete frmLayerSelection.GetItem
240                     FillCOPValues
242                     frmLegend.UpDate
                        GIS.UpDate
                    End If
                End If

244         Case "btnLyrSettings"
                Dim oLyr As New XGIS_LayerVector
                
246         Case "btnLegend"
                
                '            abCOP.Bands("bnMapLegend").Visible = Not abCOP.Bands("bnMapLegend").Visible
                ''            abCOP.Bands("bnMapLegend").Tools("frmLegend").Visible = Not abCOP.Bands("bnMapLegend").Tools("frmLegend").Visible
                '            frmLegend.Update
                
248         Case "btnInfo"

250             If m_frmAttributes Is Nothing Then
252                 Set m_frmAttributes = New frmAttributes
                End If
                
254             m_frmAttributes.Init GIS.viewer
                
256             m_frmAttributes.Show vbModeless, Me

258             g_CurrentTool = oInfo
260             GIS.Mode = XgisSelect

262         Case "btnFullExtent"
264             GIS.FullExtent

266         Case "btnLayerExtent"
268             frmLayerSelection.Init GIS, True
270             frmLayerSelection.Show vbModal, Me
                
272             If Not frmLayerSelection.GetItem = "" Then
274                 GIS.VisibleExtent = GIS.get((frmLayerSelection.GetItem)).Extent
                End If
                
                frmLayerSelection.ClearItem

276         Case "btnSelectionExtent"

278         Case "btnAddClippBoard"
280             GIS.viewer.PrintClipboard

282         Case "btnPrevExtent"
284             SetMapExtent g_PrevExt

286         Case "btnRefreshMap"
288             GIS.UpDate

290         Case "btnPrint"
                
292             frmMapPrint.Init GIS.viewer, m_frmMnuOperations.Legend1.Legend, g_bMapTplUpdate, g_sRemoteTablePrefix
                'frmMapPrint.Show vbModal, Me
294             frmMapPrint.Show vbModeless, Me
        
296         Case "btnOpenMap"

298             With c
300                 .DialogTitle = "Open Map Definition File"
                    '.CancelError = True
302                 .hWnd = Me.hWnd
304                 .Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
306                 .InitDir = g_sAppPath & "\data\user\Maps\"
308                 .Filter = "Map Definition files (*.TTKGP)|*.TTKGP"
310                 .FilterIndex = 1
312                 .ShowOpen
        
                End With
                
314             If Len(c.Filename) > 0 Then
                    'GIS.Open c.Filename, False
316                 InitMap c.Filename
                    LoadLayerAttrDataToGridInit
                End If
                
318         Case "btnRepairDB"
                Dim sConnString As String
                'TODO m_cnn Must be Closed and opened!!!
320             sConnString = m_Cnn.ConnectionString
322             m_Cnn.Close
324             Set m_Cnn = Nothing
        
326             FixDB True
        
328             Set m_Cnn = New adodb.Connection
330             m_Cnn.ConnectionString = sConnString
332             m_Cnn.Open
        
334         Case "btnCreateDBLyr"
        
336             frmLayerSelection.Init GIS
        
338             frmLayerSelection.Show vbModal, Me

340             If frmLayerSelection.GetItem <> "" Then
                    
342                 Set oLyr = GIS.get((frmLayerSelection.GetItem))

344                 With c
                
346                     .Filter = "Microsoft Access (*.mdb)|*.mdb"

348                     .DialogTitle = "Create OASIS SQL Layers"
350                     .InitDir = g_sAppPath & "\data\db\"
352                     .ShowOpen
                    End With
                    
354                 If Len(c.Filename) > 0 Then
356                     Set xportLyr = New XGIS_LayerSqlAdo
                
358                     sLayerName = InputBox("Insert the name of the layer", "OASIS GIS SQL Layers", "My Layer")
                
360                     If sLayerName = "" Then
362                         MsgBox "It seems like you have not entered a proper name of the table to be created. please try again!", vbInformation, "OASIS Data Creator"
                            frmLayerSelection.ClearItem
                            Exit Sub
                        End If
                        
364                     If MsgBox("Is this the OASIS database?", vbYesNo, "Confirm if this is the OASIS database") = vbYes Then
                        
366                         sAdoString = GetConnectionString(c.Filename)
                        Else
                        
368                         sAdoString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & c.Filename & ";" 'Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                        End If
                
370                     sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = Native" & vbCrLf & "layer=" & sLayerName & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString
                    
372                     xportLyr.Path = sAdoString
374                     oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
376                     xportLyr.SaveAll
                
                        Dim mFileSysObj As New FileSystemObject
                        Dim sPath As String
                
378                     sPath = Mid$(c.Filename, 1, InStrRev(c.Filename, "\")) & sLayerName & ".ttkls"
                
380                     mFileSysObj.CreateTextFile sPath, True
            
382                     Open sPath For Output As #1
384                     Print #1, sAdoString
386                     Close #1

                        '                        Dim iwidth As Integer
                        '                        Dim oPixLayer As New XGIS_LayerPixel
                        '                        Set oPixLayer = GIS.Get("1m_herat_geo")
                        '                        iwidth = Round((GIS.Extent.XMax - GIS.Extent.XMin) / (oPixLayer.Extent.XMax - oPixLayer.Extent.XMin) * oPixLayer.BitWidth)
                        '                        GIS.ExportToImage "C:\OASIS\Client\data\db\PixelStore2_test.ttkps", GIS.Extent, iwidth, 0
                    
                    End If
                End If
                
                frmLayerSelection.ClearItem
                
                'm_oIncidentLyr.ExportLayer
388         Case "btnExportMapDefFile"

390             With c
                
392                 .Filter = "Map Definition files (*.TTKGP)|*.TTKGP"

394                 .DialogTitle = "Save Map Definition File"
396                 .InitDir = g_sAppPath & "\data\user\Maps\"
398                 .ShowSave

400                 If Len(.Filename) > 0 Then
402                     GIS.SaveProjectAs .Filename & ".TTKGP"
                    End If

                End With

404         Case "btnExportToShape"

406             With c
                
408                 frmLayerSelection.Init GIS
        
410                 frmLayerSelection.Show vbModal, Me

                    'Dim olyr As XGIS_LayerVector

412                 If frmLayerSelection.GetItem <> "" Then
414                     Set oLyr = GIS.get((frmLayerSelection.GetItem))

416                     If Not oLyr Is Nothing Then
418                         frmExportFormats.FraExportFormats.Visible = True
420                         frmExportFormats.Show vbModal, Me
                
422                         If frmExportFormats.bExport Then
                        
424                             If frmExportFormats.chkShapeFile.Value = vbChecked Then
                                    
426                                 .Filter = "ESRI Shape files (*.shp)|*.shp"
428                                 .CancelError = True
430                                 .DialogTitle = "Export to ESRI shape file"
432                                 .InitDir = g_sAppPath & "\data\user\Maps\"
434                                 .ShowSave
                    
436                                 Set xportLyr = New XGIS_LayerSHP
                                    
438                                 If Len(.Filename) < 1 Then
                                        frmLayerSelection.ClearItem
                                        Exit Sub
                                    End If
                                    
440                                 xportLyr.Path = .Filename & ".shp"
442                                 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
444                                 xportLyr.SaveAll
                                End If
                
446                             If frmExportFormats.chkAutocadDWG.Value = vbChecked Then
448                                 .Filter = "Autocad DXF files (*.dxf)|*.dxf"
450                                 .CancelError = True
452                                 .DialogTitle = "Export to Autocad DXF file"
454                                 .InitDir = g_sAppPath & "\data\user\Maps\"
456                                 .ShowSave
                    
458                                 If Len(.Filename) < 1 Then
                                        frmLayerSelection.ClearItem
                                        Exit Sub
                                    End If
                    
460                                 Set xportLyr = New XGIS_LayerDXF
                    
462                                 xportLyr.Path = .Filename & ".dxf"
464                                 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
466                                 xportLyr.SaveAll
                                End If
                        
468                             If frmExportFormats.chkGoogleKML.Value = vbChecked Then
470                                 .Filter = "GOOGLE KML files (*.kml)|*.kml"
472                                 .CancelError = True
474                                 .DialogTitle = "Export to Google KML file"
476                                 .InitDir = g_sAppPath & "\data\user\Maps\"
478                                 .ShowSave
                    
480                                 If Len(.Filename) < 1 Then
                                        frmLayerSelection.ClearItem
                                        Exit Sub
                                    End If
                                        
482                                 Set xportLyr = New XGIS_LayerKML
                    
484                                 xportLyr.Path = .Filename & ".kml"
486                                 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
488                                 xportLyr.SaveAll
                                End If
                        
490                             If frmExportFormats.chkMapInfoTAB.Value = vbChecked Then
492                                 .Filter = "OGIS GML files (*.GML)|*.GML"
494                                 .CancelError = True
496                                 .DialogTitle = "Export to OGIS GML  file"
498                                 .InitDir = g_sAppPath & "\data\user\Maps\"
500                                 .ShowSave
                                    
502                                 If Len(.Filename) < 1 Then
                                        frmLayerSelection.ClearItem
                                        Exit Sub
                                    End If
                    
504                                 Set xportLyr = New XGIS_LayerGML
                    
506                                 xportLyr.Path = .Filename & ".gml"
                            
508                                 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
510                                 xportLyr.SaveAll
                                End If
                
512                             If frmExportFormats.chkGPSExport.Value = vbChecked Then
514                                 .Filter = "GPS Export files (*.gpx)|*.gpx"
                                    '.CancelError = True
516                                 .DialogTitle = "Export to GPS Export file"
518                                 .InitDir = g_sAppPath & "\data\user\Maps\"
520                                 .ShowSave
                    
522                                 If Len(.Filename) < 1 Then
                                        frmLayerSelection.ClearItem
                                        Exit Sub
                                    End If
                                        
524                                 Set xportLyr = New XGIS_LayerGPX
                    
526                                 xportLyr.Path = .Filename & ".gpx"
528                                 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
530                                 xportLyr.SaveAll
                                End If
                
                            End If
                        End If
                    End If

                End With

                frmLayerSelection.ClearItem
                
532         Case "btnLoadSQLLyr"

534             If MsgBox("Would You like to use existing SQL layers from the OASIS Databases?", vbYesNo, "Add OASIS Layers") = vbNo Then

536                 With c
                
538                     .Filter = "OASIS SQL Layer files (*.ttkls)|*.ttkls"
540                     .CancelError = True
542                     .DialogTitle = "Open OASIS SQL Layer file"
544                     .InitDir = g_sAppPath & "\data\db\"
546                     .ShowOpen

548                     If Len(.Filename) < 1 Then
                            Exit Sub
                        End If
            
550                     oSQLLyr.Path = .Filename
552                     oSQLLyr.Open
            
554                     GIS.Add oSQLLyr
                    End With

                Else
                                    
556                 frmSQLLayers.Show vbModal, Me
                    
558                 If frmSQLLayers.m_bOK Then
560                     If Not frmSQLLayers.comSQLLayers.ListIndex < 1 Then

562                         frmSQLLayers.comPath.ListIndex = frmSQLLayers.comSQLLayers.ListIndex
                            
564                         If frmSQLLayers.comPath.Text = "ClientDB" Then
                                'sAdoString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\Data\db\OasisClient.mdb" & ";"
566                             sAdoString = GetConnectionString(g_sAppPath & "\Data\db\OasisClient.mdb")
                           sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = Native" & vbCrLf & "layer=" & frmSQLLayers.comSQLLayers.List(frmSQLLayers.comSQLLayers.ListIndex) & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString

                            Else
568                             sAdoString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & frmSQLLayers.comPath.Text & ";"
                           sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = Native" & vbCrLf & "layer=dd_" & frmSQLLayers.comSQLLayers.List(frmSQLLayers.comSQLLayers.ListIndex) & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString
                                oSQLLyr.caption = frmSQLLayers.comSQLLayers.List(frmSQLLayers.comSQLLayers.ListIndex)
                            End If
                            
570                         oSQLLyr.Path = sAdoString
574                         oSQLLyr.Open
576                         GIS.Add oSQLLyr
                        End If
                    End If
                End If

                '502             If MsgBox("Would You like to use existing SQL layers from the OASIS Database?", vbYesNo, "Add OASIS Layers") = vbNo Then
                '
                '504                 With c
                '
                '506                     .Filter = "OASIS SQL Layer files (*.ttkls)|*.ttkls"
                '508                     .CancelError = True
                '510                     .DialogTitle = "Open OASIS SQL Layer file"
                '512                     .InitDir = g_sAppPath & "\data\db\"
                '514                     .ShowOpen
                '
                '516                     If Len(.Filename) < 1 Then
                '                            Exit Sub
                '                        End If
                '
                '518                     oSQLLyr.Path = .Filename
                '520                     oSQLLyr.Open
                '
                '522                     GIS.Add oSQLLyr
                '                    End With
                '
                '                Else
                '
                '524                 frmSQLLayers.Show vbModal, Me
                '
                '526                 If frmSQLLayers.m_bOK Then
                '528                     If Not frmSQLLayers.comSQLLayers.ListIndex < 1 Then
                '530                         sAdoString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\Data\db\OasisClient.mdb" & ";"
                '532                         sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = Native" & vbCrLf & "layer=" & frmSQLLayers.comSQLLayers.List(frmSQLLayers.comSQLLayers.ListIndex) & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString
                '534                         oSQLLyr.Path = sAdoString
                '536                         oSQLLyr.Open
                '538                         GIS.Add oSQLLyr
                '                        End If
                '                    End If
                '                End If

578         Case "btnAdvancedDBManagement"
580             frmAccessTest.Show vbModal, Me

            Case "btnEmergency"
                Dim ActLyrs() As String
                Dim DeActLyrs() As String
                Dim oActLyr As XGIS_LayerAbstract
                Dim i As Integer
                
                g_RSAppSettings.MoveFirst
                g_RSAppSettings.Find "SettingName = 'EmergencyLayers'"
                
                m_frmMnuOperations.Legend1.AllowParams = True
                
                If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").Value) Then
                    If Len(g_RSAppSettings.Fields.Item("SettingValue1").Value) > 0 Then
                        ActLyrs = Split(g_RSAppSettings.Fields.Item("SettingValue1").Value, ",")

                        For i = LBound(ActLyrs) To UBound(ActLyrs)
                            Set oActLyr = GIS.get(ActLyrs(i))

                            If Not oActLyr Is Nothing Then
                                oActLyr.Params.Visible = True
                                oActLyr.draw
                                
                            End If

                        Next

                    End If
                End If
                
                If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue2").Value) Then
                    DeActLyrs = Split(g_RSAppSettings.Fields.Item("SettingValue2").Value, ",")
                End If
                
                '                For i = 0 To GIS.Items.Count - 1
                '                    m_frmDebug.DebugPrint GIS.Items.Item(i).Name
                '
                '                    If aScan(ActLyrs, GIS.Items.Item(i).Name) > 0 Then
                '                        GIS.Items.Item(i).Params.Visible = True
                '                        GIS.UpDate
                '                    End If
                '
                '                Next
                
582         Case "btnScript0"

584             DOScripts 0

586         Case "btnScript1"

588             DOScripts 1

590         Case "btnScript2"

592             DOScripts 2

594         Case "btnScript3"

596             DOScripts 3

598         Case "btnScript4"

600             DOScripts 4

602         Case Else
                
604             DOScripts 100
        
        End Select

        On Error Resume Next
        
        Dim sName As String
606     sName = SSC.Procedures.Item("OASISToolBar_ToolClick").Name
        
608     If Len(sName) > 0 Then
610         SSC.Run "OASISToolBar_ToolClick", Tool
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
        Dim oRS As adodb.Recordset
        Dim sSummaryGroups As String
        Dim m_frmReportsFromRS As frmReportsFromRS
        Dim i As Integer
        
100     Set m_frmReportsFromRS = New frmReportsFromRS
        
102     Select Case Tool.Name
    
            Case Is = "btnReport"
104             frmExportFormats.FraExportFormats.Visible = False
106             frmExportFormats.Show vbModal, Me

108             If frmExportFormats.bExport Then
                    
110                 If frmExportFormats.chkFormats(0).Value = vbChecked Or frmExportFormats.chkFormats(1).Value = vbChecked Or frmExportFormats.chkFormats(2).Value = vbChecked Or frmExportFormats.chkFormats(3).Value = vbChecked Then

112                     With c
114                         .CancelError = False
116                         .DialogTitle = "Export Grid Data to..."
118                         .InitDir = g_sAppPath & "\data\gis\"
                        
120                         If frmExportFormats.chkFormats(0).Value = vbChecked Then
122                             .Filter = "Microsoft Excel (.xls)|*.xls"
124                             .DefaultExt = ".xls"
126                             .ShowSave

                                'g_sAppPath & "\data\user\Exports\OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xls"
128                             If Not .Filename = "" Then dxGISDataGrid.M.ExportToXLS .Filename
                            End If

130                         If frmExportFormats.chkFormats(1).Value = vbChecked Then
132                             .Filter = "XML data format (.xml)|*.xml"
134                             .DefaultExt = ".xml"
136                             .ShowSave

138                             If Not .Filename = "" Then dxGISDataGrid.M.ExportToXML .Filename 'g_sAppPath & "\data\user\Exports\OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".xml"
                            End If

140                         If frmExportFormats.chkFormats(2).Value = vbChecked Then
142                             .Filter = "HTML Web page (.html)|*.html"
144                             .DefaultExt = ".html"
146                             .ShowSave

148                             If Not .Filename = "" Then dxGISDataGrid.M.ExportToHTML .Filename 'g_sAppPath & "\data\user\Exports\OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".htm"
                            End If

150                         If frmExportFormats.chkFormats(3).Value = vbChecked Then
152                             .Filter = "Tab Limited Text Format (.txt)|*.txt"
154                             .DefaultExt = ".txt"
156                             .ShowSave

158                             If Not .Filename = "" Then dxGISDataGrid.M.SaveAllToTextFile .Filename 'g_sAppPath & "\data\user\Exports\OASIS_Export_" & Day(Now) & Month(Now) & Year(Now) & Hour(Now) & Minute(Now) & Second(Now) & ".txt"
                            End If

                        End With

                    End If
                    
160                 If frmExportFormats.chkFormats(4).Value = vbChecked Then

162                     With dxGISDataGrid.Dataset

164                         Set oRS = New adodb.Recordset

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
218                                         oRS.Fields.Item(i).Value = .FieldValues(oRS.Fields.Item(i).Name)
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
232                             GIS.viewer.PrintClipboard
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
250             m_frmDebug.DebugPrint ""
            
252             SafeMoveFirst g_RSGISGridTableSettings
254             g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"

                Dim lL As XGIS_LayerVector
            
256             Set m_oXgisThemeWiz = New XGIS_ControlLegendVectorWiz
            
258             If Not g_RSGISGridTableSettings.EOF Then
260                 Set lL = GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
                Else
262                 Set lL = GIS.get(abGridTools.Tools.Item("comLyr").Text)
                End If
            
264             lL.MoveFirst lL.Extent, "", Nothing, "", True
266             m_oXgisThemeWiz.Execute lL, lL.Shape.ShapeType, lL.ParamsList
            
268             lL.Params.Visible = True
270             GIS.UpDate

272         Case Is = "btnSpatialAnalysis"

274             With m_frmFreqSettings

276                 If Not .Visible Then
278                     .Init
280                     .Show vbModeless, Me
                    End If

                End With

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
    'Private Sub cmdAddWhat_Click()
    '        '<EhHeader>
    '        On Error GoTo cmdAddWhat_Click_Err
    '        '</EhHeader>
    '
    '100     Select Case m_frmW3Wizard.CurrentW3Module
    '
    '            Case 0 'Who
    '102             strWho1.AddRecord
    '
    '104         Case 1 'What
    '106             ctrWhat1.AddRecord
    '
    '108         Case 2 'Where
    '110             ctrWhere1.AddRecord
    '        End Select
    '
    '        '<EhFooter>
    '        Exit Sub
    '
    'cmdAddWhat_Click_Err:
    '        MsgBox Err.Description & vbCrLf & _
    '               "in OASISClient.frmMain.cmdAddWhat_Click " & _
    '               "at line " & Erl
    '        Resume Next
    '        '</EhFooter>
    'End Sub

    'Private Sub cmdAddWhere_Click()
    '        '<EhHeader>
    '        On Error GoTo cmdAddWhere_Click_Err
    '        '</EhHeader>
    '100     ctrWhere1.AddRecord
    '        '<EhFooter>
    '        Exit Sub
    '
    'cmdAddWhere_Click_Err:
    '        MsgBox Err.Description & vbCrLf & _
    '               "in OASISClient.frmMain.cmdAddWhere_Click " & _
    '               "at line " & Erl
    '        Resume Next
    '        '</EhFooter>
    'End Sub

    'Private Sub cmdBackWhat_Click()
    '        '<EhHeader>
    '        On Error GoTo cmdBackWhat_Click_Err
    '        '</EhHeader>
    '100     Select Case m_frmW3Wizard.CurrentW3Module
    '            Case 0 'Who
    '102             strWho1.MovePrevious
    '104         Case 1 'What
    '106             ctrWhat1.MovePrevious
    '108         Case 2 'Where
    '110             ctrWhere1.MovePrevious
    '        End Select
    '        '<EhFooter>
    '        Exit Sub
    '
    'cmdBackWhat_Click_Err:
    '        MsgBox Err.Description & vbCrLf & _
    '               "in OASISClient.frmMain.cmdBackWhat_Click " & _
    '               "at line " & Erl
    '        Resume Next
    '        '</EhFooter>
    'End Sub

Private Sub cmdCommand3_Click()
        '<EhHeader>
        On Error GoTo cmdCommand3_Click_Err
        '</EhHeader>

        
        
100         CreateGenericLayers
        
            Exit Sub

    '    frmMapPrint.Init GIS.viewer, m_frmMnuOperations.Legend1.Legend,
    '    frmMapPrint.Show vbModal, Me
    '
    '    m_frmMAModule.Init g_RSAppSettings
    '    m_frmMAModule.Show vbModeless, Me
        '    Vector2DBConverter "P", m_oIncidentLyr, "D:\OASIS\Client\data\db\gistest.ttkls"
        '<EhFooter>
        Exit Sub

cmdCommand3_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdCommand3_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
'
'
'Private Sub cmdDeleteWhat_Click()
'        '<EhHeader>
'        On Error GoTo cmdDeleteWhat_Click_Err
'        '</EhHeader>
'100     Select Case m_frmW3Wizard.CurrentW3Module
'
'            Case 0 'Who
'102             strWho1.DeleteRecord
'104         Case 1 'What
'106             ctrWhat1.DeleteRecord
'108         Case 2 'Where
'110             ctrWhere1.DeleteRecord
'        End Select
'        '<EhFooter>
'        Exit Sub
'
'cmdDeleteWhat_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.cmdDeleteWhat_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdDeleteWhere_Click()
'        '<EhHeader>
'        On Error GoTo cmdDeleteWhere_Click_Err
'        '</EhHeader>
'100     ctrWhere1.DeleteRecord
'        '<EhFooter>
'        Exit Sub
'
'cmdDeleteWhere_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.cmdDeleteWhere_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
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
    Dim oSQLIncLyr As New XGIS_LayerSqlAdo
   
    SafeMoveFirst g_RSAppSettings
    g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"

    oSQLIncLyr.Name = g_RSAppSettings.Fields.Item("SettingValue1").Value & Now()
    oSQLIncLyr.SQLParameter("LAYER") = "oincidents"
    oSQLIncLyr.SQLParameter("DIALECT") = "MSJET"
    oSQLIncLyr.SQLParameter("ADO") = GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")

    With oIncidentLayerSettings
        oSQLIncLyr.CachedPaint = .CachedPaint
        oSQLIncLyr.HideFromLegend = False
        oSQLIncLyr.IgnoreShapeParams = .IgnoreShapeParams
        oSQLIncLyr.IncrementalPaint = .IncrementalPaint
        '.UseConfig
        '.UseFileParams
    End With

    oSQLIncLyr.Params.Visible = True
        
    GIS.Add oSQLIncLyr

    CategorizeIncidentsByType oSQLIncLyr.Name, True

    aTESTING
'118     oSQLIncLyr.Active = False

End Sub

Private Sub aTESTING()
        '<EhHeader>
        On Error GoTo aTESTING_Err
        '</EhHeader>
        Dim oSQLIncLyr As New XGIS_LayerSqlAdo
    
100     oSQLIncLyr.Name = "TestCities" & Now()
102     oSQLIncLyr.SQLParameter("LAYER") = "TestCities"
104     oSQLIncLyr.SQLParameter("DIALECT") = "MSJET"
106     oSQLIncLyr.SQLParameter("ADO") = GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")

108     With oIncidentLayerSettings
110         oSQLIncLyr.CachedPaint = .CachedPaint
112         oSQLIncLyr.HideFromLegend = False
114         oSQLIncLyr.IgnoreShapeParams = .IgnoreShapeParams
116         oSQLIncLyr.IncrementalPaint = .IncrementalPaint
        End With

118     oSQLIncLyr.Params.Visible = True
        
120     GIS.Add oSQLIncLyr

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
        Dim shp As XGIS_Shape
        Dim lL As XGIS_LayerVector
        Dim SymbolList As New XGIS_SymbolList
        Dim RS As New adodb.Recordset
        Dim sFont As String
        Dim sDBFieldName As String
        Dim sGISFieldName As String
   
100     Set lL = GIS.get(sLyrName)
        
102     lL.ParamsList.Clear

104     lL.ParamsList.Add

106     sFont = "ERS v2 Incidents"
108     sFont = sFont & ":" & 35 & ":NORMAL"
110     lL.Params.Marker.Color = 16777215
112     lL.Params.Marker.OutlineColor = 16711680
114     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
116     lL.Params.Marker.Size = 600
118     lL.Params.Marker.ShowLegend = 1
120     lL.Params.Legend = "Uncategorized"

122     lL.ParamsList.Add
124     lL.Params.Query = "STATUS = 'Provincial capital'"
126     sFont = "ESRI Crime Analysis"
128     sFont = sFont & ":41:NORMAL"
130     lL.Params.Marker.Color = vbRed
132     lL.Params.Marker.OutlineColor = vbGreen
134     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
136     lL.Params.Marker.Size = 800
138     lL.Params.Marker.ShowLegend = 1
140     lL.Params.Legend = "Provincial Capital"
    
    
142     lL.ParamsList.Add
144     lL.Params.Query = "STATUS = 'National and provincial capital'"
146     sFont = "ERS v2 Incidents"
148     sFont = sFont & ":37:NORMAL"
150     lL.Params.Marker.Color = vbGreen
152     lL.Params.Marker.OutlineColor = vbBlue
154     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
156     lL.Params.Marker.Size = 500
158     lL.Params.Marker.ShowLegend = 1
160     lL.Params.Legend = "National & Provincial Capital"
    
    
162     lL.ParamsList.Add
164     lL.Params.Query = "STATUS = 'National capital'"
166     sFont = "ERS v2 Incidents"
168     sFont = sFont & ":21:NORMAL"
170     lL.Params.Marker.Color = vbBlack
172     lL.Params.Marker.OutlineColor = vbWhite
174     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
176     lL.Params.Marker.Size = 600
178     lL.Params.Marker.ShowLegend = 1
180     lL.Params.Legend = "National Capital"
    
    '    lL.ParamsList.Add
    '    lL.Params.Query = "STATUS = 'Other'"
    '    sFont = "ESRI Crime Analysis"
    '    sFont = sFont & ":41:NORMAL"
    '    lL.Params.Marker.Color = vbRed
    '    lL.Params.Marker.OutlineColor = vbGreen
    '    lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
    '    lL.Params.Marker.Size = 800
    '    lL.Params.Marker.ShowLegend = 1
    '    lL.Params.Legend = "Provincial Capital"
        
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
    
        
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'stdLyrs'"
    
104     If g_RSAppSettings.EOF Or g_RSAppSettings.BOF Then Exit Sub
    
106     If g_RSAppSettings.Fields.Item("SettingValue1").Value = "1" Then

108         Set m_oSQLIncLyr = New XGIS_LayerSqlAdo
    
110         SafeMoveFirst g_RSAppSettings
112         g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"
114         m_oSQLIncLyr.Name = g_RSAppSettings.Fields.Item("SettingValue1").Value
116         m_oSQLIncLyr.SQLParameter("LAYER") = "oincidents"
118         m_oSQLIncLyr.SQLParameter("DIALECT") = "MSJET"
            m_oSQLIncLyr.SQLParameter("ADO") = GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")
            
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

124         GIS.Add m_oSQLIncLyr

126         CategorizeIncidentsByType '1

130         m_oSQLIncLyr.Active = False

        End If

132     If g_RSAppSettings.Fields.Item("SettingValue4").Value = "1" Then

            'OASIS Operations
134         Set m_oSQLOpsLyr = New XGIS_LayerSqlAdo

136         With m_oSQLOpsLyr
        
138             .SQLParameter("LAYER") = "oOperations"
140             .SQLParameter("DIALECT") = "MSJET"
144             .SQLParameter("ADO") = GetConnectionString(g_sAppPath & "\data\db\Oasisclient.mdb")
                .HideFromLegend = False
            
            End With

146         GIS.Add m_oSQLOpsLyr

148         ReadLayerStyleFromDB m_oSQLOpsLyr

150         m_oSQLOpsLyr.Active = False

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

Public Sub CategorizeIncidentsByType(Optional sLyrName As String, Optional bIgnoreINI As Boolean)
        '<EhHeader>
        On Error GoTo CategorizeIncidentsByType_Err
        '</EhHeader>
        Dim shp As XGIS_Shape
        Dim lL As XGIS_LayerVector
        Dim SymbolList As New XGIS_SymbolList
        Dim RS As New adodb.Recordset
        Dim sFont As String
        Dim sDBFieldName As String
        Dim sGISFieldName As String
  
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"

104     If sLyrName = "" Then sLyrName = g_RSAppSettings.Fields.Item("SettingValue1").Value
        
106     Set lL = GIS.get(sLyrName)

        
        If Not bIgnoreINI Then
108     If Not m_sIncidentIni = "" Then

110         With lL
112             .ConfigName = m_sIncidentIni
114             .StoreParamsInProject = False
116             .UseConfig = True
118             .RereadConfig
120             .draw
            End With

            Exit Sub
        End If
        End If
        
122     Set RS = New adodb.Recordset
                
124     Set RS.ActiveConnection = g_RSAppSettings.ActiveConnection
                
126     RS.CursorType = adOpenDynamic
128     RS.CursorLocation = adUseClient
130     RS.LockType = adLockReadOnly
                
132     Select Case m_frmMnuOperations.abGridTools.Tools.Item("comIncCategory").Text  'shp.GetField("Type")

            Case "Target"
134             sDBFieldName = "Name"
136             sGISFieldName = "TARGET"
138             RS.Open "SELECT * FROM IncTarget ORDER BY [NAME]"
140             m_frmDebug.DebugPrint "Target Choosen"
142         Case "Time"
144             sDBFieldName = "Incident_Time_Name"
146             sGISFieldName = "TIME00"
148             RS.Open "SELECT * FROM IncTimeCategory ORDER BY [Incident_Time_Name]"
150             m_frmDebug.DebugPrint "Time Choosen"
152         Case "Type"
154             sDBFieldName = "Incident_Type_Name"
156             sGISFieldName = "Type"
158             RS.Open "SELECT * FROM IncTypeCategory ORDER BY [Incident_Type_Name]"
160             m_frmDebug.DebugPrint "Type Choosen"
162         Case Else
164             sDBFieldName = "Incident_Type_Name"
166             sGISFieldName = "Type"
168             RS.Open "SELECT * FROM IncTypeCategory ORDER BY [Incident_Type_Name]"
170             m_frmDebug.DebugPrint "Default Choosen"
        End Select
                
        '    sDBFieldName = "Incident_Type_Name"
        '    sGISFieldName = "Type"
        '    RS.Open "SELECT * FROM IncTypeCategory"
                
172     SafeMoveFirst RS
174     lL.ParamsList.Clear

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

176     lL.ParamsList.Add
        'll.Params.AreaColor=RGB(102:102:102)
        'm_frmDebug.DebugPrint  .Item("Incident_Type_Name").value
178     sFont = "ERS v2 Incidents"
180     sFont = sFont & ":" & 35 & ":NORMAL"
    
182     lL.Params.Marker.Color = 16777215
184     lL.Params.Marker.OutlineColor = 16711680
186     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
188     lL.Params.Marker.Size = 600
190     lL.Params.Marker.ShowLegend = 1
192     lL.Params.Legend = "Uncategorized"

194     m_frmDebug.DebugPrint "Current Font:" & sFont
        
196     Do While Not RS.EOF

198         With RS.Fields
        
200             lL.ParamsList.Add
202             lL.Params.Query = sGISFieldName & " = '" & .Item(sDBFieldName).Value & "'"
204             '''m_frmDebug.DebugPrint "THIS IS QUERY:" & sGISFieldName & " = '" & .Item(sDBFieldName).Value & "'"
206             '''m_frmDebug.DebugPrint "FieldName" & sGISFieldName & " The DB Value:" & sDBFieldName
                'll.Params.AreaColor=RGB(102:102:102)
                'm_frmDebug.DebugPrint  .Item("Incident_Type_Name").value
208             sFont = .Item("Font_Name").Value
210             sFont = sFont & ":" & .Item("Ascii").Value & ":NORMAL"
212             '''m_frmDebug.DebugPrint "FONT=" & sFont
214             lL.Params.Marker.Color = .Item("bgColor").Value
216             lL.Params.Marker.OutlineColor = .Item("color").Value
218             lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
220             lL.Params.Marker.Size = .Item("size").Value
222             lL.Params.Marker.ShowLegend = 1
224             lL.Params.Legend = .Item(sDBFieldName).Value
226             '''m_frmDebug.DebugPrint "Legend Source" & .Item(sDBFieldName).Value
228             RS.MoveNext
            End With

        Loop
        
230     If sLyrName = abGridTools.Tools.Item("comLyr").Text Then
232         LoadLayerAttrDataToGridInit
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
        Dim shp As XGIS_Shape
        Dim lL As XGIS_LayerVector
        Dim SymbolList As New XGIS_SymbolList
        Dim RS As New adodb.Recordset
        Dim sFont As String
        Dim sDBFieldName As String
        Dim sGISFieldName As String
        Dim sLyrName As String
  
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"

        sLyrName = g_RSAppSettings.Fields.Item("SettingValue1").Value
        
104     Set lL = GIS.get(sLyrName)

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
    
106     Set RS = New adodb.Recordset
                
108     Set RS.ActiveConnection = m_Cnn
                
110     RS.CursorType = adOpenDynamic
112     RS.CursorLocation = adUseClient
114     RS.LockType = adLockReadOnly

118     sDBFieldName = "Name"
120     sGISFieldName = "TARGET"
122     RS.Open "SELECT * FROM IncTarget ORDER BY [Name]"
        m_frmDebug.DebugPrint "Target Choosen"
                
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
    
160     lL.Params.Marker.Color = 16777215
162     lL.Params.Marker.OutlineColor = 16711680
164     lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
166     lL.Params.Marker.Size = 600
168     lL.Params.Marker.ShowLegend = 1
170     lL.Params.Legend = "Uncategorized"

        m_frmDebug.DebugPrint "Current Font:" & sFont
        
172     Do While Not RS.EOF

174         With RS.Fields
        
176             lL.ParamsList.Add
178             lL.Params.Query = sGISFieldName & " = '" & .Item(sDBFieldName).Value & "'"
                m_frmDebug.DebugPrint "THIS IS QUERY:" & sGISFieldName & " = '" & .Item(sDBFieldName).Value & "'"
                m_frmDebug.DebugPrint "FieldName" & sGISFieldName & " The DB Value:" & sDBFieldName

180             sFont = .Item("Font_Name").Value
182             sFont = sFont & ":" & .Item("Ascii").Value & ":NORMAL"
                m_frmDebug.DebugPrint "FONT=" & sFont
184             lL.Params.Marker.Color = .Item("bgColor").Value
186             lL.Params.Marker.OutlineColor = .Item("color").Value
188             lL.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
190             lL.Params.Marker.Size = .Item("size").Value
192             lL.Params.Marker.ShowLegend = 1
194             lL.Params.Legend = .Item(sDBFieldName).Value
                m_frmDebug.DebugPrint "Legend Source" & .Item(sDBFieldName).Value
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
    Dim oLyrAbs As XGIS_LayerAbstract
    Dim i As Integer
    
    '174                 For i = 0 To GIS.Items.Count - 1
    '
    '176                     If GisUtils.IsInherited(GIS.Items.Item(i), "XGIS_LayerVector") Then
    '178                         Set oVecLyr = GIS.get(GIS.Items.Item(i).Name)
    '180                         Set oShp = oVecLyr.Locate(ptg, 10 / GIS.Zoom, True)
    With oLyrAbs
 
        For i = 0 To GIS.Items.Count - 1
            Set oLyrAbs = GIS.get(GIS.Items.Item(i).Name)
            oLyrAbs.StoreParamsInProject = False
            oLyrAbs.ConfigName = g_sAppPath & "\Data\gis\temp\" & oLyrAbs.Name
            oLyrAbs.WriteConfig
            oLyrAbs.SaveAll
        Next
    
    End With

End Sub

Public Sub ReadLayerStyleFromDB(oLayer As XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo ReadLayerStyleFromDB_Err
        '</EhHeader>

100     With oLayer
        
102         With .Params
104             With .area
            
106                 m_frmDebug.DebugPrint .Color
                    'm_frmDebug.DebugPrint  .Symbol.FontName
                    'm_frmDebug.DebugPrint  .SymbolSize
                    'm_frmDebug.DebugPrint  .SymbolGap
                    'm_frmDebug.DebugPrint  .SymbolRotate
                    'm_frmDebug.DebugPrint  .Bitmap
                    'm_frmDebug.DebugPrint  .Pattern
                    'm_frmDebug.DebugPrint  .OutlineColor
                    'm_frmDebug.DebugPrint  .OutlineWidth
                    'm_frmDebug.DebugPrint  .OutlineStyle
                    'm_frmDebug.DebugPrint  .OutlineSymbol
                    'm_frmDebug.DebugPrint  .OutlineSymbolGap
                    'm_frmDebug.DebugPrint  .OutlineSymbolRotate
                    'm_frmDebug.DebugPrint  .OutlineBitmap
                    'm_frmDebug.DebugPrint  .OutlinePattern
                    'm_frmDebug.DebugPrint  .SmartSize
                    'm_frmDebug.DebugPrint  .SmartSizeField
            
108                 .Color = RGB(0, 0, 255)
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
           
112                 m_frmDebug.DebugPrint .Color
                    'm_frmDebug.DebugPrint  .Marker.Symbol
                    '.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
                    'm_frmDebug.DebugPrint  .SymbolSize
                    'm_frmDebug.DebugPrint  .SymbolGap
                    'm_frmDebug.DebugPrint  .SymbolRotate
                    'm_frmDebug.DebugPrint  .Bitmap
                    'm_frmDebug.DebugPrint  .Pattern
                    'm_frmDebug.DebugPrint  .OutlineColor
                    'm_frmDebug.DebugPrint  .OutlineWidth
                    'm_frmDebug.DebugPrint  .OutlineStyle
                    'm_frmDebug.DebugPrint  .OutlineSymbol
                    'm_frmDebug.DebugPrint  .OutlineSymbolGap
                    'm_frmDebug.DebugPrint  .OutlineSymbolRotate
                    'm_frmDebug.DebugPrint  .OutlineBitmap
                    'm_frmDebug.DebugPrint  .OutlinePattern
                    'm_frmDebug.DebugPrint  .SmartSize
                    'm_frmDebug.DebugPrint  .SmartSizeField
                End With
           
114             With .Line
           
116                 m_frmDebug.DebugPrint .Color
                    'm_frmDebug.DebugPrint  .Marker.Symbol
                    '.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
                    'm_frmDebug.DebugPrint  .SymbolSize
                    'm_frmDebug.DebugPrint  .SymbolGap
                    'm_frmDebug.DebugPrint  .SymbolRotate
                    'm_frmDebug.DebugPrint  .Bitmap
                    'm_frmDebug.DebugPrint  .Pattern
                    'm_frmDebug.DebugPrint  .OutlineColor
                    'm_frmDebug.DebugPrint  .OutlineWidth
                    'm_frmDebug.DebugPrint  .OutlineStyle
                    'm_frmDebug.DebugPrint  .OutlineSymbol
                    'm_frmDebug.DebugPrint  .OutlineSymbolGap
                    'm_frmDebug.DebugPrint  .OutlineSymbolRotate
                    'm_frmDebug.DebugPrint  .OutlineBitmap
                    'm_frmDebug.DebugPrint  .OutlinePattern
                    'm_frmDebug.DebugPrint  .SmartSize
                    'm_frmDebug.DebugPrint  .SmartSizeField
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

'Private Sub cmdMoveFirstWhat_Click()
'        '<EhHeader>
'        On Error GoTo cmdMoveFirstWhat_Click_Err
'        '</EhHeader>
'100     Select Case m_frmW3Wizard.CurrentW3Module
'
'            Case 0 'Who
'102             strWho1.MoveFirst
'104         Case 1 'What
'106             ctrWhat1.MoveFirst
'108         Case 2 'Where
'110             ctrWhere1.MoveFirst
'        End Select
'        '<EhFooter>
'        Exit Sub
'
'cmdMoveFirstWhat_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.cmdMoveFirstWhat_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdMoveFirstWhere_Click()
'        '<EhHeader>
'        On Error GoTo cmdMoveFirstWhere_Click_Err
'        '</EhHeader>
'100     ctrWhere1.MoveFirst
'        '<EhFooter>
'        Exit Sub
'
'cmdMoveFirstWhere_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.cmdMoveFirstWhere_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdMoveLastWhat_Click()
'        '<EhHeader>
'        On Error GoTo cmdMoveLastWhat_Click_Err
'        '</EhHeader>
'100      Select Case m_frmW3Wizard.CurrentW3Module
'
'            Case 0 'Who
'102             strWho1.MoveLast
'104         Case 1 'What
'106             ctrWhat1.MoveLast
'108         Case 2 'Where
'110             ctrWhere1.MoveLast
'        End Select
'
'        '<EhFooter>
'        Exit Sub
'
'cmdMoveLastWhat_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.cmdMoveLastWhat_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdMoveLastWher_Click()
'        '<EhHeader>
'        On Error GoTo cmdMoveLastWher_Click_Err
'        '</EhHeader>
'100     ctrWhere1.MoveLast
'        '<EhFooter>
'        Exit Sub
'
'cmdMoveLastWher_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.cmdMoveLastWher_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdMoveNext_Click()
'        '<EhHeader>
'        On Error GoTo cmdMoveNext_Click_Err
'        '</EhHeader>
'100     ctrWhere1.MoveNext
'        '<EhFooter>
'        Exit Sub
'
'cmdMoveNext_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.cmdMoveNext_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdMoveNextWhat_Click()
'        '<EhHeader>
'        On Error GoTo cmdMoveNextWhat_Click_Err
'        '</EhHeader>
'100     Select Case m_frmW3Wizard.CurrentW3Module
'
'            Case 0 'Who
'102             strWho1.MoveNext
'104         Case 1 'What
'106             ctrWhat1.MoveNext
'108         Case 2 'Where
'110             ctrWhere1.MoveNext
'        End Select
'        '<EhFooter>
'        Exit Sub
'
'cmdMoveNextWhat_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.cmdMoveNextWhat_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdMovePrevious_Click()
'        '<EhHeader>
'        On Error GoTo cmdMovePrevious_Click_Err
'        '</EhHeader>
'100     ctrWhere1.MovePrevious
'        '<EhFooter>
'        Exit Sub
'
'cmdMovePrevious_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmMain.cmdMovePrevious_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub



Private Sub cmdONE_Click()
    'Dim cn As New ADODB.Connection
        '<EhHeader>
        On Error GoTo cmdONE_Click_Err
        '</EhHeader>

        'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OASIS\W3Import.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"

100     frmAddorganisation.Init m_Cnn
102     frmAddorganisation.Show
    
        '<EhFooter>
        Exit Sub

cmdONE_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdONE_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdThree_Click()
    'Dim cn As New ADODB.Connection
        '<EhHeader>
        On Error GoTo cmdThree_Click_Err
        '</EhHeader>

    '    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OASIS\W3Import.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
100     frmAddWhere.Init m_Cnn
102     frmAddWhere.Show
        '<EhFooter>
        Exit Sub

cmdThree_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdThree_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTwo_Click()
    '    Dim cn As New ADODB.Connection
        '<EhHeader>
        On Error GoTo cmdTwo_Click_Err
        '</EhHeader>

    '100     cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OASIS\W3Import.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
 
100     frmAddWhat.Init m_Cnn
102     frmAddWhat.Show
        '<EhFooter>
        Exit Sub

cmdTwo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.cmdTwo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub comActiveLyr_Click()
        '<EhHeader>
        On Error GoTo comActiveLyr_Click_Err
        '</EhHeader>
  
        Dim i As Integer
        Dim lL As XGIS_LayerAbstract
  
100     If Not m_bLOADING Then

102         For i = 1 To comActiveLyr.ListCount - 1
104             Set lL = GIS.get(comActiveLyr.List(i))

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
  
116         GIS.Invalidate
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
        ' m_frmDebug.DebugPrint  "ctTreeBookmrks_Change" & nIndex
        '<EhHeader>
        On Error GoTo ctTreeBookmrks_Change_Err
        '</EhHeader>
    
100     m_frmDebug.DebugPrint ctTreeBookmrks.NodeText(nIndex) & " Change Level:" & ctTreeBookmrks.NodeLevel(nIndex)
   
        ' Exit Sub
    
        Dim RS As New adodb.Recordset
   
102     If ctTreeBookmrks.NodeText(nIndex) = m_DefaultViewName Then
104         GIS.viewer.SetViewport m_DefaultViewx, m_DefaultViewy
106         GIS.viewer.zoom = m_DefaultViewz
108         GIS.UpDate
            'Map1.ZoomTo m_DefaultViewz, m_DefaultViewx, m_DefaultViewy
        Else

110         If ctTreeBookmrks.NodeLevel(nIndex) = 2 Then
112             RS.Open "SELECT X,Y,Z FROM GeoBookMarks WHERE Name ='" & ctTreeBookmrks.NodeText(nIndex) & "'", m_Cnn, adOpenDynamic, adLockReadOnly

114             If Not RS.BOF Then SafeMoveFirst RS
            
                Dim ptg As New XGIS_Point
            
                'Set ptg = gis.Center  .CreateShape(XgisShapeTypePoint)
                If IsNull(RS.Fields.Item("X")) Or IsNull(RS.Fields.Item("Y")) Or IsNull(RS.Fields.Item("Z")) Then
                
                    MsgBox "XY coordinates or zoom factor for this bookmark are invalid!"
                
                Else
                
116                 ptg.Prepare CDbl(RS.Fields.Item("X")), CDbl(RS.Fields.Item("Y"))
                
                    'GIS.Lock
                    'GIS.CenterPtg.Prepare CDbl(RS.Fields.Item("X")), CDbl(RS.Fields.Item("Y"))
    
                
118                 GIS.zoom = RS.Fields.Item("Z")
120                 GIS.CenterViewport ptg
                    'GIS.Unlock
122                 GIS.UpDate
    
    
                End If
            End If
        End If

        'm_frmDebug.DebugPrint  ctTreeBookmrks.NodeLevel(nIndex)

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
100     m_frmDebug.DebugPrint "TreeHeaderClicked" & nColumn
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
        Dim fl As XGIS_Shape
        Dim i As Integer

100     translated = True
        
        
        With Shape
        
102         factor = .GetField(m_sSecAnalysisFieldName) - 100

            With .Params.area
                        
                If factor < arRangeVals(0) Then
                   .Color = &H8C07&
                   .Pattern = XBrushStyle.XbsClear
                Else

                    For i = 1 To UBound(arRangeColors)

                        If factor < arRangeVals(i) Then
                            .Color = arRangeColors(i)
                            .Pattern = XBrushStyle.XbsSolid
                            Exit For
                        End If

                    Next

                End If
                
'                'ReDim arRangeVals(6)
'
'104             .Color = &H8C07&
'
'106             If factor < 101 Then
'
'108                 .Color = BlendColors(RGB(255, 0, 0), RGB(0, 255, 0), 100)
'110                 .Pattern = XBrushStyle.XbsClear
'                    'Shape.Params.Line.Style = XPenStyle.XpsDot
'112             ElseIf factor < 105 Then
'114                 .Color = BlendColors(RGB(255, 0, 0), RGB(0, 105, 0), 80)
'                    .Pattern = XBrushStyle.XbsSolid
'116             ElseIf factor < 110 Then
'118                 .Color = BlendColors(RGB(255, 0, 0), RGB(0, 105, 0), 60)
'                    .Pattern = XBrushStyle.XbsSolid
'120             ElseIf factor < 115 Then
'122                 .Color = BlendColors(RGB(255, 0, 0), RGB(0, 105, 0), 40)
'                    .Pattern = XBrushStyle.XbsSolid
'124             ElseIf factor < 120 Then
'126                 .Color = BlendColors(RGB(255, 0, 0), RGB(0, 55, 0), 20)
'                    .Pattern = XBrushStyle.XbsSolid
'128             ElseIf factor < 125 Then
'130                 .Color = BlendColors(RGB(255, 0, 0), RGB(0, 50, 0), 0)
'                    .Pattern = XBrushStyle.XbsSolid
'132             ElseIf factor < 150 And factor < 300 Then
'134                 .Color = BlendColors(RGB(255, 0, 0), RGB(0, 0, 0), 20)
'                    .Pattern = XBrushStyle.XbsSolid
'                Else
'136                 .Color = BlendColors(RGB(255, 0, 0), RGB(0, 0, 0), 40)
'                    .Pattern = XBrushStyle.XbsSolid
'                End If
'
            End With
  
138         .draw
        End With
        
        '<EhFooter>
        Exit Sub

EventLayer_OnPaintShape_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.EventLayer_OnPaintShape " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Public Function SaveAsShape(sLayerName As String, sNewLAyerName As String) As XGIS_LayerSHP
        '<EhHeader>
        On Error GoTo SaveAsShape_Err
        '</EhHeader>
    Dim oTargetLayer As New XGIS_LayerSHP
    Dim oSourceLyr As XGIS_LayerVector

100     Set oSourceLyr = GIS.get(sLayerName)
    
102     oTargetLayer.Path = g_sAppPath & "\data\gis\temp\" & sNewLAyerName & ".shp"
104     oTargetLayer.ExportLayer oSourceLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", False
106     GIS.Add oTargetLayer
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
        Dim oTestShp As XGIS_Shape
        Dim oOverLayLayer As XGIS_LayerVector
        Dim oOverlayLayer2 As XGIS_LayerVector
        
        Dim lTot As Long
        Dim lnow As Long
        Dim sResult() As String
    
100     ReDim sResult(0)
    
102     GIS.get(sTargetLayer).MoveFirst GisUtils.GisWholeWorld, "", Nothing, "", True
104     GIS.get(sTargetLayer).scope = sContraint
    
106     Set oTestShp = GIS.get(sTargetLayer).FindFirst(GisUtils.GisWholeWorld, "", Nothing, "", True)
        
108     GIS.get(sOverLayLayer).scope = sScope
110     Set oOverLayLayer = GIS.get(sOverLayLayer)
        
        'This workaround fixes the date issue (tatuk limitation)
112     Set oOverlayLayer2 = New XGIS_LayerVector
114     oOverlayLayer2.ImportLayer oOverLayLayer, oOverLayLayer.Extent, XgisShapeTypeUnknown, sScope, True

116     While Not oTestShp Is Nothing
118         lnow = GetNumWithinGeometry(oTestShp, oOverlayLayer2, "", sDE_9IM)
120         lTot = lTot + lnow
122         sResult(UBound(sResult)) = oTestShp.GetField(sTargetID) & ",,," & lnow
124         Set oTestShp = GIS.get(sTargetLayer).FindNext()

126         If Not oTestShp Is Nothing Then ReDim Preserve sResult(UBound(sResult) + 1)
        Wend
        
128     GIS.get(sTargetLayer).scope = ""
130     sSpatialString = Join(sResult, ":::")

        'GetSimpleSpatialAnalysis = Join(sResult, ":::")
        'MsgBox Join(sResult, ":::")

    'Avoid Unnecessary Use Of "DoEvents" [0148]
    'Don 't fill your code with unnecessary "DoEvents" statements, especially within time-critical loops. If you use DoEvents just to trap mouse and keyboard activity, you can call it only if there are pending items in the event queue; you can check this condition with the GetInputState API function:
    '
    'Declare Function GetInputState Lib "user32" () As Long
    '' ...
    'If GetInputState() <> 0 Then
    '    DoEvents
    'End If
    '
    'If you can't avoid that, at least you can reduce the overhead by invoking DoEvents only every N iterations of the loop, using a statement like this:
    '
    'If (loopNdx Mod 10) = 0 Then
    '    DoEvents
    'End If


        '<EhFooter>
        Exit Sub

m_frmSpatialAnalysis_GetSimpleSpatialAnalysis_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.m_frmSpatialAnalysis_GetSimpleSpatialAnalysis " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetNumWithinGeometry(oShpToSearch As XGIS_Shape, oLyr As XGIS_LayerVector, sScope As String, sDE_9IM As String) As Long
        '<EhHeader>
        On Error GoTo GetNumWithinGeometry_Err
        '</EhHeader>
    Dim i As Long
    Dim aShape As XGIS_Shape
    
100     If oLyr Is Nothing Then Exit Function
    
102     With oLyr
    
104     .scope = sScope
106     .MoveFirst GisUtils.GisWholeWorld, "", Nothing, "", True
     
108     Set aShape = .FindFirst(oShpToSearch.Extent, "", oShpToSearch, sDE_9IM, True)
    
110     While Not aShape Is Nothing
112         i = i + 1
114         Set aShape = oLyr.FindNext()
        Wend
116     .scope = ""
    
        End With
    
118     GetNumWithinGeometry = i
    
        '<EhFooter>
        Exit Function

GetNumWithinGeometry_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GetNumWithinGeometry " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function


Public Sub CheckIncidentFrequency(oOverLayLayer As XGIS_LayerVector, _
                                  oTargetLayer As XGIS_LayerVector, _
                                  Optional bResetPreviousScoring As Boolean = True)
        '<EhHeader>
        On Error GoTo CheckIncidentFrequency_Err
        '</EhHeader>
        Dim Shh As Shape
        Dim admShape As Shape
        Dim i As Integer
        Dim shp As XGIS_ShapePoint
        Dim aShape As Object
        Dim oPolyShape As XGIS_ShapePolygon
        Dim iVal As Integer
        Dim iTimeVal As Integer
        Dim iTargetVal As Integer
        Dim iTypeVal As Integer
        Dim oRSType As adodb.Recordset
        Dim oRSTarget As adodb.Recordset
        Dim oRSTime As adodb.Recordset
        
        oTargetLayer.IgnoreShapeParams = True
        oTargetLayer.CachedPaint = False
100     If frmScoring.m_bApply Then
        
102         If frmScoring.chkChkActivate(0).Value = vbChecked Then
104             Set oRSType = New adodb.Recordset
        
106             With oRSType
108                 .Open "SELECT Incident_Type_Name, Scoring FROM IncTypeCategory ORDER BY [Incident_Type_Name]", m_Cnn, adOpenDynamic, adLockReadOnly
                End With

            End If
        
110         If frmScoring.chkChkActivate(1).Value = vbChecked Then
        
112             Set oRSTarget = New adodb.Recordset

114             With oRSTarget
116                 .Open "SELECT Name, Scoring FROM IncTarget ORDER BY [NAME]", m_Cnn, adOpenDynamic, adLockReadOnly
                End With

            End If
        
118         If frmScoring.chkChkActivate(2).Value = vbChecked Then

120             Set oRSTime = New adodb.Recordset

122             With oRSTime
124                 .Open "SELECT Incident_Time_Name, Scoring FROM IncTimeCategory ORDER BY [Incident_Time_Name]", m_Cnn, adOpenDynamic, adLockReadOnly
                End With

            End If
        
        End If

126     If bResetPreviousScoring Then
128         ResetScoring oTargetLayer
        End If

        '    For i = 0 To llRel.Items.Count - 1
        '
        '        Set Shh = llRel.Items(i)
        '
        '        'Check if the Admin Shape relates
        '        For j = 0 To llAdm.Items.Count - 1
        '            Set admShape = llAdm.Items(j)
        '
        '            If Shh.IsInsidePolygon(admShape, XgisInsideTypeCentroid) Then
        '                m_frmDebug.DebugPrint  "Hit! NUM:" & i
        '                Exit For
        '            End If
        '
        '        Next j
        '
        '        i = i + 1
        '    Next
        '
        'To
        '

        'If oTargetLayer.Items("Scoring") Is Nothing Then
        'oTargetLayer.AddField "Scoring", XgisFieldTypeNumber, 16, 0
        'End If

130     oOverLayLayer.MoveFirst GIS.viewer.VisibleExtent, "", Nothing, "", True
    
        'STOP HEre

132     Set shp = oOverLayLayer.FindFirst(GIS.viewer.VisibleExtent, "", Nothing, "", True)

134     While Not shp Is Nothing

136         Set aShape = oTargetLayer.FindFirst(shp.Extent, "", shp, GisUtils.GIS_RELATE_WITHIN, True)
        
138         If aShape Is Nothing Then
                'm_frmDebug.DebugPrint  "Hit! NUM:" & i & " INCIDENT NAME:" & shp.GetField("Name") & " ADM Name: Out of Bounds"
          
            Else

                'aShape.Lock XgisLockExtent
                'Set oPolyShape = aShape

140             If Not aShape.GetField(m_sSecAnalysisFieldName) = vbNull Then
142                 Set oPolyShape = aShape.MakeEditable
                
144                 If frmScoring.m_bApply Then
                    
146                     iTargetVal = 0
148                     iTimeVal = 0
150                     iTypeVal = 0

152                     If frmScoring.chkChkActivate(0).Value = vbChecked Then

154                         With oRSType
156                             SafeMoveFirst oRSType
158                             .Find "Incident_Type_Name = '" & shp.GetField("TYPE") & "'"
160                             If Not .EOF Then iTypeVal = CInt(.Fields.Item("Scoring").Value)
                            End With

                        End If
                    
162                     If frmScoring.chkChkActivate(1).Value = vbChecked Then

164                         With oRSTarget
166                             SafeMoveFirst oRSTarget
168                             .Find "Name = '" & shp.GetField("TARGET") & "'"
170                             If Not .EOF Then iTargetVal = CInt(.Fields.Item("Scoring").Value)
                            End With

                        End If
                    
172                     If frmScoring.chkChkActivate(2).Value = vbChecked Then

174                         With oRSTime
176                             SafeMoveFirst oRSTime
178                             .Find "Incident_Time_Name = '" & shp.GetField("TIME00") & "'"
180                             If Not .EOF Then iTimeVal = CInt(.Fields.Item("Scoring").Value)
                            End With

                        End If
                    End If
                
182                 iVal = aShape.GetField(m_sSecAnalysisFieldName) + 1
184                 oPolyShape.SetField m_sSecAnalysisFieldName, CLng(iVal + iTargetVal + iTimeVal + iTypeVal)  'CInt(aShape.GetField("Scoring") + 1)
                Else
                    'TODO CHECK IF THIS IS NEEDED!
                    'Set oPolyShape = aShape.MakeEditable
                    'oPolyShape.SetField "Scoring", 100
                End If
            
                'm_frmDebug.DebugPrint  "Hit! NUM:" & i & " INCIDENT NAME:" & shp.GetField("Name") & " ADM Name: " & oPolyShape.GetField("adm2Name") & " Score: " & oPolyShape.GetField(m_sSecAnalysisFieldName)
                'oTargetLayer.SaveData
                'aShape.Unlock
                'aShape.Layer.SaveData
            End If

186         i = i + 1

            ' Set aShape = oTargetLayer.FindNext()
            'Wend
188         Set shp = oOverLayLayer.FindNext()
            'oTargetLayer.SaveData
            'EventLayer_OnPaintShape False, oPolyShape
            'oTargetLayer.SaveData
        Wend
        'GIS.RereadConfig
        'oTargetLayer.RereadConfig
        'oTargetLayer
    
        '  oTargetLayer.ParamsList.Add
        '  oTargetLayer.Params.Query = "INTERN_ID = 100"
        '
        '  oTargetLayer.Params.area.Color = RGB(0, 255, 0)
        '
        '
        '  oTargetLayer.ParamsList.Add
        '  oTargetLayer.Params.Query = "INTERN_ID > 110"
        '  oTargetLayer.Params.area.Color = RGB(155, 0, 0)
    
        'oTargetLayer.Params.AreaColor = vbRed
        'oTargetLayer.Params.area.Color = vbRed
    
        'oTargetLayer.Params.Marker.OutlineColor = vbBlue
        'oTargetLayer.Params.Marker.Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
        'oTargetLayer.Params.Marker.Size = 440
        'oTargetLayer.Params.Marker.ShowLegend = 1
        'oTargetLayer.Params.Legend = .Item("Incident_Type_Name").value
190     oTargetLayer.Active = True 'Params.Visible = True
        'GIS.Update
        ' GIS.SaveAll
    
        '<EhFooter>
        Exit Sub

CheckIncidentFrequency_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CheckIncidentFrequency " & _
               "at line " & Erl
          
        '</EhFooter>
End Sub

Public Sub ResetScoring(oLayer As XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo ResetScoring_Err
        '</EhHeader>
        On Error Resume Next
        Dim shp As Object

100     Set shp = oLayer.FindFirst(oLayer.Extent, "", Nothing, "", True)
    
102     While Not shp Is Nothing
104         Set shp = shp.MakeEditable
106         shp.SetField m_sSecAnalysisFieldName, 100
108         m_frmDebug.DebugPrint shp.GetField(m_sSecAnalysisFieldName)
110         Set shp = oLayer.FindNext()
        Wend
        'oLayer.SaveData
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
        Dim oAbsLayer As XGIS_LayerAbstract
        Dim c As New cCommonDialog
    
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
                
                        'IS VECTOR FILES
116                     If InStr(UCase(.Filename), ".SHP") Then
118                         Set oLayer = New XGIS_LayerSHP
120                     ElseIf InStr(UCase(.Filename), ".KML") Then
122                         Set oLayer = New XGIS_LayerKML
124                     ElseIf InStr(UCase(.Filename), ".TAB") Then
126                         Set oLayer = New XGIS_LayerTAB
128                     ElseIf InStr(UCase(.Filename), ".MIF") Then
130                         Set oLayer = New XGIS_LayerMIF
132                     ElseIf InStr(UCase(.Filename), ".GML") Then
134                         Set oLayer = New XGIS_LayerGML
136                     ElseIf InStr(UCase(.Filename), ".DXF") Then
138                         Set oLayer = New XGIS_LayerDXF
                        ElseIf InStr(UCase(.Filename), ".GPX") Then
                            Set oLayer = New XGIS_LayerGPX
                        Else
                            Set oLayer = New XGIS_LayerVector
                        End If
                
140                     oLayer.Path = .Filename
                        
141                     Set oLayer = GisUtils.GisCreateLayer(Mid$(.FileTitle, 1, Len(.FileTitle) - 4), .Filename)

142                     'oLayer.Open
144                     GIS.Add oLayer
146                     m_oColUserLayers.Add oLayer.Name
148                     FillCOPValues
150                 ElseIf InStr(UCase(GisUtils.GisSupportedFiles(XGIS_FileType.XgisFileTypePixel, False)), UCase(Mid(.Filename, InStr(.Filename, ".")))) > 0 Then
            
                        Dim sExt As String
                    
                        sExt = UCase$(Right$(.Filename, Len(.Filename) - InStr(.Filename, ".")))
                
                        If InStr("*.toc)|*.TOC|", sExt) Then
                            Set oLayer = New XGIS_LayerCADRG
                        ElseIf InStr("(*.tif;*.tiff)|*.TIF;*.TIFF|", sExt) Then
                            Set oLayer = New XGIS_LayerTIFF
                        ElseIf InStr("(*.bil)|*.BIL|", sExt) Then
                            Set oLayer = New XGIS_LayerBIL
                        ElseIf InStr("(*.ddf)|*.DDF|", sExt) Then
                            Set oLayer = New XGIS_LayerSDTS_RPE
                        ElseIf InStr("(*.png)|*.PNG|", sExt) Then
                            Set oLayer = New XGIS_LayerPNG
                        ElseIf InStr("(*.psi)|*.PSI|", sExt) Then
                            Set oLayer = New XGIS_LayerPixel
                        ElseIf InStr("(*.sid)|*.SID|", sExt) Then
                            Set oLayer = New XGIS_LayerMrSID
                        ElseIf InStr("(*.jpg;*.jpeg)|*.JPG;*.JPEG|", sExt) Then
                            Set oLayer = New XGIS_LayerJPG
                        ElseIf InStr("(*.jp2;*.j2k;*.jpf;*.jpx;*.jpc)|*.JP2;*.J2K;*.JPF;*.JPX;*.JPC|", sExt) Then
                            Set oLayer = New XGIS_LayerJPG
                        ElseIf InStr("(*.gif)|*.GIF|", sExt) Then
                            Set oLayer = New XGIS_LayerGIF
                        ElseIf InStr("(*.img)|*.IMG|", sExt) Then
                            Set oLayer = New XGIS_LayerIMG
                        ElseIf InStr("(*.ecw)|*.ECW|", sExt) Then
                            Set oLayer = New XGIS_LayerECW
                        ElseIf InStr("(*.bmp)|*.BMP|", sExt) Then
                            Set oLayer = New XGIS_LayerBMP
                        ElseIf InStr("(*.ttkps)|*.TTKPS|", sExt) Then
                            Set oLayer = New XGIS_LayerPixelStoreAdo2
                        Else
                            Set oLayer = New XGIS_LayerPixel
                        End If
                        
180                     oLayer.Path = .Filename
182                    Set oLayer = GisUtils.GisCreateLayer(Mid$(.FileTitle, 1, Len(.FileTitle) - 4), .Filename) ' oLayer.Open
                
184                     GIS.Add oLayer
186                     FillCOPValues
188                 ElseIf InStr(UCase(GisUtils.GisSupportedFiles(XGIS_FileType.XgisFileTypeProject, False)), UCase(Mid(.Filename, InStr(.Filename, ".")))) > 0 Then
            
190                     GIS.Open .Filename
            
                    End If
                End If

                '110         GIS. CommonDialog1.FileName
            Else
192             GIS.Add oLayer
194             GIS.UpDate
            End If
            
196         GIS.UpDate
        End With
        
       ' g_RSAppSettings.MoveFirst
       ' g_RSAppSettings.Find "SettingName = 'Permission'"
       '
       ' If Not g_RSAppSettings.EOF And Not g_RSAppSettings.BOF Then
       '     g_RSAppSettings
       ' End If
        
        
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
                                          oRS As adodb.Recordset, _
                                          sSpatialJoinKey As String, _
                                          sAttributeJoinKey As String)
        '<EhHeader>
        On Error GoTo JoinAttributeDataToGeometries_Err
        '</EhHeader>
        Dim oLyr As XGIS_LayerVector

100     Set oLyr = GIS.get(sLayer)

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

Public Sub CreateGenericLayers()
        '<EhHeader>
        On Error GoTo CreateGenericLayers_Err
        '</EhHeader>
        Dim cGenCombo As ActiveBar3LibraryCtl.Tool
        Dim oSynchTables As New adodb.Recordset

100     With oSynchTables
        
102         .CursorLocation = adUseClient
104         .CursorType = adOpenDynamic
106         .Open "SELECT * FROM SynchTables", m_Cnn, , adLockReadOnly
        
108         If Not .EOF And Not .BOF Then
                'Set cGenCombo = AB.Bands("tbUtils").Tools.Item("comGenLyr")
    
110             If cGenCombo Is Nothing Then
112                 Set cGenCombo = AB.Bands("tbUtils").Tools.Add(Rnd(10000), "comGenLyr")
114                 cGenCombo.caption = "hello"
116                 cGenCombo.ControlType = ddTTCombobox
118                 cGenCombo.Style = ddSText
                End If
                
120             cGenCombo.CBClear
                
122             SafeMoveFirst oSynchTables

124             Do While Not .EOF
126                 Set m_oSQLGenericLyrs(UBound(m_oSQLGenericLyrs)) = New XGIS_LayerSqlAdo
            
128                 If Not IsNull(.Fields.Item("sName").Value) Then
130                     m_oSQLGenericLyrs(UBound(m_oSQLGenericLyrs)).Name = .Fields.Item("sName").Value
                    End If
                
132                 m_oSQLGenericLyrs(UBound(m_oSQLGenericLyrs)).SQLParameter("LAYER") = .Fields.Item("sTableName").Value
134                 m_oSQLGenericLyrs(UBound(m_oSQLGenericLyrs)).SQLParameter("DIALECT") = "MSJET"
136                 m_oSQLGenericLyrs(UBound(m_oSQLGenericLyrs)).SQLParameter("ADO") = m_Cnn.ConnectionString '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=data\db\Oasisclient.mdb;pwd=<#mypassword#>"
                    '.Fields.Item("AllowWrite").Value
                
138                 If Not IsNull(.Fields.Item("InLegend").Value) Then
140                     m_oSQLGenericLyrs(UBound(m_oSQLGenericLyrs)).HideFromLegend = CBool(.Fields.Item("InLegend").Value)
                    End If
                
142                 If Not IsNull(.Fields.Item("sCaption").Value) Then
144                     m_oSQLGenericLyrs(UBound(m_oSQLGenericLyrs)).caption = .Fields.Item("sCaption").Value
                    End If
146                 AB.Refresh
148                 cGenCombo.CBAddItem .Fields.Item("sName").Value
                
150                 GIS.Add m_oSQLGenericLyrs(UBound(m_oSQLGenericLyrs))
                
152                 If Not IsNull(.Fields.Item("IsActive").Value) Then
154                     m_oSQLGenericLyrs(UBound(m_oSQLGenericLyrs)).Active = CBool(.Fields.Item("IsActive").Value) 'False
                    End If
                
156                 .MoveNext
                
158                 If Not .EOF Then ReDim Preserve m_oSQLGenericLyrs(UBound(m_oSQLGenericLyrs) + 1)

                Loop
            
160             cGenCombo.CBListIndex = 0

            End If
        
        End With

162     AB.RecalcLayout

        '<EhFooter>
        Exit Sub

CreateGenericLayers_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CreateGenericLayers " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadOASISIncidents()
        '<EhHeader>
        On Error GoTo LoadOASISIncidents_Err
        '</EhHeader>
    
        Dim shpInc As XGIS_ShapePoint
        Dim sWhereClause As String
        
100     ClearOASISIncidentLyr
        
102     Set m_oIncRS = New adodb.Recordset
    
        'TODO MAKE SURE TO CHECK THE VALUE OF DATE Incident_Date
    
104     Select Case m_frmMnuOperations.abGridTools.Tools.Item("comDateFrom").Text
       
            Case "1 week"

106             If MonthView1.Week = 1 Then
108                 MonthView1.Week = 52
                Else
110                 MonthView1.Week = MonthView1.Week - 1
                End If

112             sWhereClause = "Incident_DATE > #" & MonthView1.Value & "#"
                SecurityLayerDateFrom = MonthView1.Value
                SecurityLayerDateTill = Now()

                'm_frmDebug.DebugPrint  1 & "/" & Month( & "/" & Year
114         Case "2 weeks"

116             If MonthView1.Week = 1 Then
118                 MonthView1.Week = 51
120             ElseIf MonthView1.Week = 2 Then
122                 MonthView1.Week = 52
                Else
124                 MonthView1.Week = MonthView1.Week - 2
                End If

126             sWhereClause = "Incident_DATE > #" & MonthView1.Value & "#"
                SecurityLayerDateFrom = MonthView1.Value
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

                SecurityLayerDateFrom = MonthView1.Value
                SecurityLayerDateTill = Now()
144             sWhereClause = "Incident_DATE > #" & MonthView1.Value & "#"

146         Case "1 month"

148             If MonthView1.Month = 1 Then
150                 MonthView1.Month = 12
                Else
152                 MonthView1.Month = MonthView1.Month - 1
                End If
               
                SecurityLayerDateFrom = MonthView1.Value
                SecurityLayerDateTill = Now()
154             sWhereClause = "Incident_DATE > #" & MonthView1.Value & "#"

156         Case "Custom"
158             frmDatePicker.Show vbModal, Me
               
                SecurityLayerDateFrom = MonthView1.Value
                SecurityLayerDateTill = Now()
160             sWhereClause = "Incident_DATE BETWEEN #" & frmDatePicker.dtFrom.Value & "# AND #" & frmDatePicker.dtTo.Value & "#  "
               
162         Case Else
164             sWhereClause = ""
                SecurityLayerDateFrom = Format("1-1-1900", "Medium Date")
                SecurityLayerDateTill = Now()
166             MonthView1.Value = Format(Now(), "dd/mm/yyyy")
        End Select
    
168     With m_oIncRS
170         .Open "SELECT * FROM incReport" & sWhereClause, m_Cnn

172         SafeMoveFirst m_oIncRS

174         If Not .BOF Then

176             Do While Not .EOF

178                 If Not .Fields.Item("Longitude (E)").Value = vbNull Or Not .Fields.Item("Latitude (N)").Value = vbNull Then
180                     Set shpInc = m_oIncidentLyr.CreateShape(XgisShapeTypePoint)
            
182                     shpInc.Lock XgisLockExtent
184                     shpInc.AddPart
            
186                     With .Fields
188                         shpInc.AddPoint GisUtils.GISPoint(.Item("Longitude (E)").Value, .Item("Latitude (N)").Value)

                            '"ID"
                            '"Name"
                            '"Type"
                            '"Target"
                            '"Time"
                            '"Decription"
190                         shpInc.SetField "ID", .Item("ID").Value
192                         shpInc.SetField "Name", .Item("IncidentID").Value
194                         shpInc.SetField "Type", .Item("Incident Type").Value
196                         shpInc.SetField "Target", .Item("Incident Target").Value
198                         shpInc.SetField "Time", .Item("Incident Time").Value
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
    
        Dim shpInc As XGIS_ShapePoint
        Dim sWhereClause As String
               
100     Set m_oIncRS = New adodb.Recordset
    
        'TODO MAKE SURE TO CHECK THE VALUE OF DATE Incident_Date
    
102     MonthView1.Value = Format(Now(), "yyyy/mm/dd")
    
104     Select Case m_frmMnuOperations.abGridTools.Tools.Item("comDateFrom").Text
       
            Case "1 week"

106             If MonthView1.Week = 1 Then
108                 MonthView1.Week = 52
                    MonthView1.Year = MonthView1.Year - 1
                Else
110                 MonthView1.Week = MonthView1.Week - 1
                End If

                SecurityLayerDateFrom = MonthView1.Value
                SecurityLayerDateTill = Now()
112             sWhereClause = "Incident_DATESERIAL > " & ConvertDateToSerial(MonthView1.Value)
114             'sWhereClause = "Incident_DATE > #" & MonthView1.Value & "#"

                'm_frmDebug.DebugPrint  1 & "/" & Month( & "/" & Year
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
                
                SecurityLayerDateFrom = MonthView1.Value
                SecurityLayerDateTill = Now()
128             sWhereClause = "Incident_DATESERIAL > " & ConvertDateToSerial(MonthView1.Value)
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

                SecurityLayerDateFrom = MonthView1.Value
                SecurityLayerDateTill = Now()
146             sWhereClause = "Incident_DATESERIAL > " & ConvertDateToSerial(MonthView1.Value)
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
               
                SecurityLayerDateFrom = MonthView1.Value
                SecurityLayerDateTill = Now()
156             sWhereClause = "Incident_DATESERIAL > " & ConvertDateToSerial(MonthView1.Value)
                'sWhereClause = "Incident_DATE > #" & MonthView1.Value & "#"

158         Case "Custom"
160             If bShowCustomDialog Then frmDatePicker.Show vbModal, Me
               
                SecurityLayerDateFrom = frmDatePicker.dtFrom.Value
                SecurityLayerDateTill = frmDatePicker.dtTo.Value
162             sWhereClause = "Incident_DATESERIAL >= " & ConvertDateToSerial(frmDatePicker.dtFrom.Value) & " AND Incident_DATESERIAL <= " & ConvertDateToSerial(frmDatePicker.dtTo.Value)
                'sWhereClause = "Incident_DATE BETWEEN #" & frmDatePicker.dtFrom.Value & "# AND #" & frmDatePicker.dtTo.Value & "#  "
               
164         Case Else
166             sWhereClause = ""
                SecurityLayerDateFrom = Format("1-1-1900", "Medium Date")
                SecurityLayerDateTill = Now()
168             MonthView1.Value = Format(Now(), "dd/mm/yyyy")
        End Select
        
        'this code was not used since it would make transparent layers not transparent any more!
        '170     m_oSQLIncLyr.scope = sWhereClause  '"Incident_DATE BETWEEN #2007-07-18# AND #2007-07-19#"  '"DATE00 > #2007-07-18#"  sWhereClause
        '        If m_oSQLIncLyr.Params.Visible Then
        '172         GIS.draw
        '        End If

170     m_oSQLIncLyr.scope = sWhereClause  '"Incident_DATE BETWEEN #2007-07-18# AND #2007-07-19#"  '"DATE00 > #2007-07-18#"  sWhereClause

        If m_oSQLIncLyr.Active And bCommitFilter Then
            'had to change to update since it would not always remove the old layer filter
172         'm_oSQLIncLyr.draw
            GIS.UpDate
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
        Dim RS As New adodb.Recordset
        Dim WRS As New adodb.Recordset
        'Dim cn As New Connection
        Dim shpInc As XGIS_Shape
        Dim SymbolList As New XGIS_SymbolList

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

110         If Not .BOF Then

112             Do While Not .EOF

114                 If Not .Fields.Item("latX").Value = vbNull Or Not .Fields.Item("longY").Value = vbNull Then
116                     Set shpInc = m_oW3Lyr.CreateShape(XgisShapeTypePoint)
            
118                     shpInc.Lock XgisLockExtent
120                     shpInc.AddPart
            
122                     With .Fields
124                         shpInc.AddPoint GisUtils.GISPoint(CDbl(Replace(.Item("longY").Value, ",", ".")), CDbl(Replace(.Item("latX").Value, ",", ".")))

                            '"ID"
                            '"Name"
                            '"Type"
                            '"Target"
                            '"Time"
                            '"Decription"
                            
126                         shpInc.SetField "Name", .Item("Name").Value
                            
                            'Add the lookUp type
128                         SafeMoveFirst RS
130                         RS.Find "id = '" & .Item("placeTypeId").Value & "'"
                            
132                         shpInc.SetField "Type", RS.Fields.Item("name").Value
134                         shpInc.SetField "Description", .Item("Description").Value
136                         shpInc.SetField "PCode", "IQ200700" & .Item("pCode").Value
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
            
146     m_oW3Lyr.Params.Marker.Color = vbWhite
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
        Dim WRS As New adodb.Recordset
        Dim shpInc As XGIS_Shape
        Dim SymbolList As New XGIS_SymbolList
        
100     If m_oW3WHOLyr Is Nothing Then
            Exit Sub
        End If

102     With WRS
104         .Open "SELECT * FROM qryWho", m_Cnn, adOpenDynamic, adLockOptimistic

106         SafeMoveFirst WRS

108         If Not .BOF Then

110             Do While Not .EOF

112                 If Not .Fields.Item("latX").Value = vbNull Or Not .Fields.Item("longY").Value = vbNull Then
114                     Set shpInc = m_oW3WHOLyr.CreateShape(XgisShapeTypePoint)
            
116                     shpInc.Lock XgisLockExtent
118                     shpInc.AddPart
            
120                     With .Fields
122                         shpInc.AddPoint GisUtils.GISPoint(CDbl(Replace(.Item("longY").Value, ",", ".")), CDbl(Replace(.Item("latX").Value, ",", ".")))

                            '"PlaceType"
                            '"Cluster"
                            '"offname"

124                         shpInc.SetField "Organisation", .Item("name").Value
                            
126                         shpInc.SetField "PlaceType", .Item("PlaceType").Value
128                         shpInc.SetField "Cluster", .Item("Cluster").Value
130                         shpInc.SetField "offname", .Item("offname").Value
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
142         .Params.Marker.Color = vbWhite
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
    
116     If g_RSAppSettings.Fields.Item("SettingValue1").Value = "1" Then
    
118         Set m_oIncidentLyr = GisUtils.GisCreateLayer("OASIS Incidents", g_sAppPath & "\data\gis\temp\incidents.shp")
        
120         m_oIncidentLyr.AddField "ID", XgisFieldTypeString, 100, 0
122         m_oIncidentLyr.AddField "Name", XgisFieldTypeString, 100, 0
124         m_oIncidentLyr.AddField "Type", XgisFieldTypeString, 100, 0
126         m_oIncidentLyr.AddField "Target", XgisFieldTypeString, 100, 0
128         m_oIncidentLyr.AddField "Time", XgisFieldTypeString, 100, 0
130         m_oIncidentLyr.AddField "Description", XgisFieldTypeString, 255, 0
        End If
        
132     If g_RSAppSettings.Fields.Item("SettingValue2").Value = "1" Then
134         Set m_oW3Lyr = New XGIS_LayerVector
  
136         m_oW3Lyr.AddField "Name", XgisFieldTypeString, 100, 0
138         m_oW3Lyr.AddField "Type", XgisFieldTypeString, 100, 0
140         m_oW3Lyr.AddField "Description", XgisFieldTypeString, 255, 0
142         m_oW3Lyr.AddField "PCode", XgisFieldTypeString, 100, 0
    
144         m_oW3Lyr.Params.area.Color = RGB(0, 0, 255)
146         m_oW3Lyr.Params.Label.Field = "Name"
148         m_oW3Lyr.Params.Label.Visible = "YES"
            'Params.MarkerSymbol := SynbolList.Prepare( 'mysymbol.bmp?FALSE' );
            'm_oW3Lyr.Params.MarkerSymbol = "..\\\\..\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE"

            'm_oW3Lyr.Transparency = 50
150         m_oW3Lyr.Name = "Who What Where"
152         m_oW3Lyr.HideFromLegend = False

        End If

154     If g_RSAppSettings.Fields.Item("SettingValue3").Value = "1" Then

156         Set m_oW3WHOLyr = New XGIS_LayerVector
158         m_oW3WHOLyr.AddField "Organisation", XgisFieldTypeString, 100, 0
160         m_oW3WHOLyr.AddField "PlaceType", XgisFieldTypeString, 100, 0
162         m_oW3WHOLyr.AddField "Cluster", XgisFieldTypeString, 255, 0
164         m_oW3WHOLyr.AddField "offname", XgisFieldTypeString, 100, 0
    
166         m_oW3WHOLyr.Params.area.Color = RGB(0, 0, 255)
168         m_oW3WHOLyr.Params.Label.Field = "Organisation"
170         m_oW3WHOLyr.Params.Label.Visible = "YES"
            'Params.MarkerSymbol := SynbolList.Prepare( 'mysymbol.bmp?FALSE' );
            'm_oW3Lyr.Params.MarkerSymbol = "..\\\\..\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE"

            'm_oW3Lyr.Transparency = 50
172         m_oW3WHOLyr.Name = "W3-Who"
174         m_oW3WHOLyr.HideFromLegend = False
        End If

        CreateAddedScribbles

'176     Set m_oDrawLyr = New XGIS_LayerVector
'178     m_oDrawLyr.Params.area.Color = RGB(0, 0, 255)
'180     m_oDrawLyr.Transparency = 90
'182     m_oDrawLyr.Name = "Draw_Layer"
'184     m_oDrawLyr.HideFromLegend = True
    
186     Set m_oCosmeticLayer = New XGIS_LayerVector

188     With m_oCosmeticLayer
190         With .Params
    
192             .area.Color = RGB(0, 0, 255)
    
            End With
    
194         .Transparency = 50
196         .Name = "Cosmetic"
198         .HideFromLegend = True
200         .AddField "cosLabel", XgisFieldTypeString, 255, 0
202         .AddField "cosValue", XgisFieldTypeNumber, 16, 0
204         .AddField "cosText", XgisFieldTypeString, 255, 0
206         .AddField "cosDescription", XgisFieldTypeString, 100, 0

        End With
  
208     Set m_oBufferLyr = New XGIS_LayerVector
210     m_oBufferLyr.Params.area.Color = RGB(0, 0, 255)
212     m_oBufferLyr.Params.area.OutlineColor = RGB(0, 0, 255)
214     m_oBufferLyr.Transparency = 60
216     m_oBufferLyr.Name = "Buffers"
218     m_oBufferLyr.HideFromLegend = True

        On Error Resume Next
220     If g_RSAppSettings.Fields.Item("SettingValue1").Value = "1" Then m_oIncidentLyr.SaveAll
    
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

Public Sub LoadMapThread(sMap As Variant)
    On Error Resume Next
    GIS.Open sMap, False
End Sub

Public Sub LoadMapProducts(Optional sInitialMapName As String)
        '<EhHeader>
        On Error GoTo LoadMapProducts_Err
        '</EhHeader>


    '    Select Case variable
    
    '        Case ?
    
    '    End Select

100     If Not sInitialMapName = "" Then
            On Error Resume Next
102         GIS.Open sInitialMapName, False

           ' If Not pLoadMap.IsThreadRunning Then
           '     pLoadMap.CreateWin32Thread Me, "LoadMapThread", sInitialMapName
           ' End If
            'GIS.FullExtent
        Else
            Exit Sub
    '            Dim ll As XGIS_LayerSHP
    '            Dim lv As XGIS_LayerVector
    '            Dim lw As XGIS_LayerVector
    '
    '            ' add layers
    '106         Set ll = New XGIS_LayerSHP 'main map states
    '108         ll.Path = GisUtils.GisSamplesDataDir + "states.shp"
    '110         ll.Name = "states"
    '112         ll.UseConfig = False
    '114         ll.Params.area.Color = RGB(255, 255, 0) 'layers properties
    '116         ll.Params.area.OutlineColor = RGB(0, 0, 0)
    '118         ll.Params.area.OutlineWidth = 20
    '120         GIS.Add ll    'add to main map
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

Public Sub SetLayerFilter(oLayer As XGIS_LayerVector, sFieldName As String, sVal As String)
        '<EhHeader>
        On Error GoTo SetLayerFilter_Err
        '</EhHeader>
100     oLayer.MoveFirst oLayer.Extent, "", Nothing, "", True
    
102     Do While Not oLayer.EOF
    
           ' m_frmDebug.DebugPrint  oLayer.Shape.GetField(sFieldName)
104         If oLayer.Shape.GetField(sFieldName) = sVal Then
106             oLayer.Shape.IsHidden = True
            End If
        
108         oLayer.MoveNext
        Loop
    
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
            Dim ptg1 As XGIS_Point
            Dim ptg2 As XGIS_Point
            Dim ptg3 As XGIS_Point
            Dim ptg4 As XGIS_Point
            Dim ex As XGIS_Extent

    Exit Sub

100         If GIS.IsEmpty Then Exit Sub

102         Set ex = GIS.VisibleExtent
104         Set ptg1 = GisUtils.GISPoint(ex.XMin, ex.YMin)
106         Set ptg2 = GisUtils.GISPoint(ex.XMax, ex.YMin)
108         Set ptg3 = GisUtils.GISPoint(ex.XMax, ex.YMax)
110         Set ptg4 = GisUtils.GISPoint(ex.XMin, ex.YMax)

112         minishp.Reset
114         minishp.Lock XgisLockExtent
116         minishp.AddPart
118         minishp.AddPoint ptg1
120         minishp.AddPoint ptg2
122         minishp.AddPoint ptg3
124         minishp.AddPoint ptg4
126         minishp.Unlock

128         minishpo.Reset
130         minishpo.Lock XgisLockExtent
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
            Dim ptg As XGIS_Point
            Dim P1, p2, p3 As XGIS_Point
            Dim p4 As New XGIS_Point

100         Set ptg = m_frmOvMap.GISm.ScreenToMap(GisUtils.Point(x, y))
102         minishp.SetPosition miniRecalc(ptg), m_frmOvMap.GISm.get(MINIMAP_R_NAME), 5
104         m_frmOvMap.GISm.UpDate
106         fminiMove = False
108         Set P1 = minishp.GetPoint(0, 0)
110         Set p2 = minishp.GetPoint(0, 1)
112         Set p3 = minishp.GetPoint(0, 2)
114         p4.x = P1.x + (p2.x - P1.x) / 2
116         p4.y = P1.y + (p3.y - p2.y) / 2
118         GIS.Center = p4
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

' GIS.Open ("Storage=Native\nLAYER=lwaters\nDIALECT=MSJET\nADO=Provider=Microsoft.Jet.OLEDB.4.0;Data Source=gistest.mdb\n.ttkls")

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

Public Sub CloneLyrSettings(oLyrToCopyFrom As XGIS_LayerVector, oLyrToAssign As XGIS_LayerVector)
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

Public Sub SetMapExtent(oXtent As XGIS_Extent)
        '<EhHeader>
        On Error GoTo SetMapExtent_Err
        '</EhHeader>
100     GIS.VisibleExtent = oXtent
        '<EhFooter>
        Exit Sub

SetMapExtent_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.SetMapExtent " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadLayerAttrDataToGridInit(Optional bForceLoadAll As Boolean, _
                                       Optional bForceLoadLimited As Boolean)
        '<EhHeader>
        On Error GoTo LoadLayerAttrDataToGridInit_Err
        '</EhHeader>

        Dim GlobalGISGridLayerOriginal As XGIS_LayerVector
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
118     If Not g_RSGISGridTableSettings.EOF And Not g_RSGISGridTableSettings.BOF Then
        
120         Set GlobalGISGridLayerOriginal = GIS.get(g_RSGISGridTableSettings.Fields("Name").Value)
        Else
        
122         Set GlobalGISGridLayerOriginal = GIS.get(abGridTools.Tools.Item("comLyr").Text)
        End If
    
        'Layer not found?
124     If GlobalGISGridLayerOriginal Is Nothing Then
126         MsgBox "The Data You are trying to browse does not anymore exist in the map." & vbCrLf & "Either it has been removed or the data is corrupt." & vbCrLf & "Contact your OASIS adminstrator if this problem remains."
128         Set GlobalGISGridLayerOriginal = Nothing
            Exit Sub
        End If
    
130     GlobalGISGridLayerOriginal.Lock
132     Set GlobalGISGridLayer = New XGIS_LayerVector

134     If bForceLoadAll Then
        
            'Load all
136         GlobalGISGridLayerOriginal.ExportLayer GlobalGISGridLayer, GlobalGISGridLayerOriginal.Extent, XgisShapeTypeUnknown, GlobalGISGridLayerOriginal.scope, False
        
138     ElseIf bForceLoadLimited Then
        
            'Load limited
140         GlobalGISGridLayerOriginal.ExportLayer GlobalGISGridLayer, GIS.viewer.VisibleExtent, XgisShapeTypeUnknown, GlobalGISGridLayerOriginal.scope, False
        
        Else
        
            'Note that the MBR should be defined correctly for this to work good
142         If GIS.viewer.VisibleExtent.XMin > GlobalGISGridLayer.Extent.XMin Or GIS.viewer.VisibleExtent.XMax < GlobalGISGridLayer.Extent.XMax Or GIS.viewer.VisibleExtent.YMin > GlobalGISGridLayer.Extent.YMin Or GIS.viewer.VisibleExtent.YMax < GlobalGISGridLayer.Extent.YMax Then

                'not all records are visible
144             If chkOnlyVisible.Value = vbChecked Then 'MsgBox("Do you want to load only data which is visible in this map view?", vbYesNo, "Filter data by view") = vbYes Then
146                 GlobalGISGridLayerOriginal.ExportLayer GlobalGISGridLayer, GIS.viewer.VisibleExtent, XgisShapeTypeUnknown, GlobalGISGridLayerOriginal.scope, False
                Else
148                 GlobalGISGridLayerOriginal.ExportLayer GlobalGISGridLayer, GlobalGISGridLayerOriginal.Extent, XgisShapeTypeUnknown, GlobalGISGridLayerOriginal.scope, False
                End If

            Else
                'all records are visible
150             GlobalGISGridLayerOriginal.ExportLayer GlobalGISGridLayer, GlobalGISGridLayerOriginal.Extent, XgisShapeTypeUnknown, GlobalGISGridLayerOriginal.scope, False
            End If
        
        End If

152     GlobalGISGridLayerOriginal.Unlock
        sLayerName = GlobalGISGridLayerOriginal.Name
154     Set GlobalGISGridLayerOriginal = Nothing
    
156     With g_RSGISGridTableSettings
    
158         SafeMoveFirst g_RSGISGridTableSettings
160         .Find "Name = '" & sLayerName & "'"
            
162         If Not .BOF And Not .EOF Then
        
164             m_frmDebug.DebugPrint GlobalGISGridLayer.Items.Count
166             iTot = GlobalGISGridLayer.GetLastUid
        
                'Get list of excluded fields
            
168             If Not .EOF Then
170                 If Not .Fields.Item("excludedFlds").Value = vbNull Then
172                     sExcludedFldsString = .Fields.Item("excludedFlds").Value
                    End If
                End If
        
174             If Not .EOF Then
            
                    'Check count of records
176                 If .Fields.Item("datasetwarning").Value Then
                
                        'Too many records - ABORT!
178                     If iTot > .Fields.Item("MaxRec").Value Then
180                         MsgBox "Note! According to performace settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").Value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
182                         Set GlobalGISGridLayer = Nothing
                            Exit Sub
                        Else

                            'Many records - ask the question
184                         If iTot > .Fields.Item("warninglevel").Value Then
186                             If MsgBox("The data you are about to browse contains: " & iTot & " records." & vbCrLf & "This might take very long time. Would you like to continue?", vbYesNo) = vbNo Then
188                                 Set GlobalGISGridLayer = Nothing
                                    Exit Sub
                                End If
                            End If
                        End If
                    
                    Else

                        'Too many records - ABORT!
190                     If iTot > .Fields.Item("MaxRec").Value Then
192                         MsgBox "Note! According to performance settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").Value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
194                         Set GlobalGISGridLayer = Nothing
                            Exit Sub
                        End If
                    End If

                Else
                    'UserDefGridRecords
                    'the Layer is userdrawn, Use default values from AppSettings
196                 SafeMoveFirst g_RSAppSettings
198                 g_RSAppSettings.Find "SettingName = 'UserDefGridRecords'"
                
200                 If g_RSAppSettings.Fields.Item("SettingValue1").Value = "1" Then
                    
202                     If iTot > CLng(g_RSAppSettings.Fields.Item("SettingValue2").Value) Then
                        
                            'Too many records - ABORT!
204                         MsgBox "Note! According to performace settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").Value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
206                         Set GlobalGISGridLayer = Nothing
                            Exit Sub
                        
                        Else

                            'Too many records - ABORT!
208                         If iTot > CLng(g_RSAppSettings.Fields.Item("SettingValue3").Value) Then
210                             If MsgBox("The data you are about to browse contains:" & iTot & " records." & vbCrLf & "This might take very long time. Would you like to continue?", vbYesNo) = vbNo Then
212                                 Set GlobalGISGridLayer = Nothing
                                    Exit Sub
                                End If
                            End If
                        
                        End If
                    
                    End If
                
                End If
            
            End If
        
        End With
    
214     If bDebugMode Then
216         LoadLayerAttrDataToGrid sExcludedFldsString
        Else

218         If Not pLoadLyrAttrToGrdThread.IsThreadRunning Then
220             pLoadLyrAttrToGrdThread.CreateWin32Thread Me, "LoadLayerAttrDataToGrid", sExcludedFldsString
            End If
        End If

        On Error Resume Next
222     m_frmDebug.DebugPrint "GIS load operation took " & Round(((GetTickCount - lMilliSecsTimed) / 1000), 2) & " seconds"
     
        '<EhFooter>
        Exit Sub

LoadLayerAttrDataToGridInit_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.LoadLayerAttrDataToGridInit " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadLayerAttrDataToGrid(sExcludedFieldsPassed As Variant)
        '<EhHeader>
        On Error GoTo LoadLayerAttrDataToGrid_Err
        '</EhHeader>
        
        Dim vcm As Variant
        Dim i As Long
        Dim lngFldType As Long
        Dim lngDataLength As Long
        Dim flds() As String
        Dim col As DXDBGRIDLibCtl.dxGridColumn ' Variant
        Dim k As Long
    
        Dim j As Long
        Dim strVals As Variant
        Dim arVals As Variant
        Dim arVal As Variant
        Dim arFieldNames As Variant
        Dim varVal As Variant
        Dim fldPos() As Integer
        Dim bExclude As Boolean
        Dim iCountOfFields As Integer
        Dim lUID As Long
    
        Dim sExcludedFlds() As String

100     If ((Not sExcludedFlds) = -1&) Then
102         sExcludedFlds = Split(sExcludedFieldsPassed, ",")
        End If

104     With dxGISDataGrid
        
106         .Visible = False
108         .Columns.DestroyColumns
110         Set .DataSource = Nothing
112         Set GISGridRS = New adodb.Recordset
114         GlobalGISGridLayer.MoveFirst GlobalGISGridLayer.Extent, "", Nothing, "", True

116         Set dxProgressBar1.Container = dxGISDataGrid.Container
118         dxProgressBar1.Left = dxGISDataGrid.Left
120         dxProgressBar1.Top = dxGISDataGrid.Top
122         dxProgressBar1.Width = dxGISDataGrid.Width
124         dxProgressBar1.Height = dxGISDataGrid.Height
126         dxProgressBar1.MinPos = 0
128         dxProgressBar1.MaxPos = dxProgressBar1.MinPos + 1 + GlobalGISGridLayer.GetLastUid
130         dxProgressBar1.Pos = 0
132         dxProgressBar1.Step = 1
134         dxProgressBar1.Visible = True
136         dxProgressBar1.DoStep
            
            'LockWindowUpdate elGisAttr.Hwnd
            If mnuAutoSize.Checked Then
138             .Options.Set (egoAutoWidth)
            Else
                .Options.Unset (egoAutoWidth)
            End If
            
140         .Options.Set (egoShowGroupPanel)
142         .Options.Set (egoBandMoving)
144         .Options.Set (egoColumnMoving)
146         .Options.Set (egoMultiSort)
148         .Options.Set (egoShowFooter)
150         .Options.Set (egoAutoSort)
152         .Options.Set (egoShowButtons)
154         .Options.Set (egoShowRowFooter)
156         .Options.Set (egoAutoSearch)
158         .Options.Set (egoAutoExpandOnSearch)
160         .Options.Set (egoAnsiSort)
162         .Options.Set (egoLoadAllRecords)
164         .Options.Set (egoAutoSearch)
166         .Options.Unset (egoCanNavigation)
168         .Options.Unset (egoDynamicLoad)
170         .DatasetType = dtADODataset
172         .Filter.FilterActive = True
174         .Filter.FilterStatus = fsAlways
    
176         ReDim flds(0)
178         GISGridRS.Fields.Append "GIS_UID", adBigInt, 100
180         iCountOfFields = 1

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Set all the fieldnames in the RS
            '
182         For i = 0 To GlobalGISGridLayer.Fields.Count - 1

184             bExclude = False
                
186             If Not ((Not sExcludedFlds) = -1&) Then
188                 If SequentialSearchStringArray(sExcludedFlds, GlobalGISGridLayer.Fields.Item(i).Name) > -1 Then
190                     bExclude = True
                    End If
                End If
                
192             If Not bExclude Then

194                 iCountOfFields = iCountOfFields + 1
                
196                 ReDim Preserve flds(UBound(flds) + 1)
198                 ReDim Preserve fldPos(UBound(flds))
                    
200                 flds(UBound(flds)) = GlobalGISGridLayer.Fields.Item(i).Name '(i + 1) to Compensate for the GIS_UID
                    'Stop
202                 fldPos(UBound(flds)) = i
                    
204                 Select Case GlobalGISGridLayer.Fields.Item(i).FieldType
        
                        Case Is = XgisFieldTypeString '= 0,
206                         lngFldType = adVarChar 'xftString
208                         lngDataLength = GlobalGISGridLayer.Fields.Item(i).Width
        
210                     Case Is = XgisFieldTypeNumber ' = 1,
212                         lngFldType = adDouble 'xftFloat
214                         lngDataLength = GlobalGISGridLayer.Fields.Item(i).Width
        
216                     Case Is = XgisFieldTypeFloat '= 2,
218                         lngFldType = adDouble 'xftFloat
220                         lngDataLength = GlobalGISGridLayer.Fields.Item(i).Width
        
222                     Case Is = XgisFieldTypeBoolean '= 3,
224                         lngFldType = adBoolean 'xftBoolean
226                         lngDataLength = GlobalGISGridLayer.Fields.Item(i).Width
        
228                     Case Is = XgisFieldTypeDate '= 4
230                         lngFldType = adDate 'xftDate
232                         lngDataLength = GlobalGISGridLayer.Fields.Item(i).Width
                    End Select
                    
234                 GISGridRS.Fields.Append GlobalGISGridLayer.Fields.Item(i).Name, lngFldType, lngDataLength

                End If
                
            Next
    
236         GlobalGISGridLayer.MoveFirst GlobalGISGridLayer.Extent, "", Nothing, "", True
238         arVals = Split(strVals, ",")
240         arFieldNames = Split(strVals, ",")
242         strVals = ""
244         ReDim arVal(UBound(flds))

246         If Not GlobalGISGridLayer.Shape Is Nothing Then arVal(0) = GlobalGISGridLayer.Shape.uID
248         j = 0
250         GISGridRS.Open
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Populate the RS
            '
252         Do While Not GlobalGISGridLayer.EOF

                'LockWindowUpdate 0
254             dxProgressBar1.DoStep
256             elMap.Refresh
                'LockWindowUpdate elGisAttr.Hwnd
258             GISGridRS.AddNew
            
260             For i = 0 To iCountOfFields - 1 'UBound(flds)

262                 Select Case GlobalGISGridLayer.Fields.Item(fldPos(i)).FieldType 'i to compensate the GIS_UID on 0 position
    
                        Case Is = XgisFieldTypeString '= 0,
264                         varVal = CStr(GlobalGISGridLayer.Shape.GetField(flds(i)))
266                         arVal(i) = varVal
    
268                     Case Is = XgisFieldTypeNumber ' = 1,
270                         arVal(i) = CDbl(GlobalGISGridLayer.Shape.GetField(flds(i)))
    
272                     Case Is = XgisFieldTypeFloat '= 2,
274                         arVal(i) = CDbl(GlobalGISGridLayer.Shape.GetField(flds(i)))
    
276                     Case Is = XgisFieldTypeBoolean '= 3,
278                         arVal(i) = CBool(GlobalGISGridLayer.Shape.GetField(flds(i)))
    
280                     Case Is = XgisFieldTypeDate '= 4
282                         arVal(i) = CDate(GlobalGISGridLayer.Shape.GetField(flds(i)))
                    End Select
                    
284                 GISGridRS.Fields(i).Value = arVal(i)
                
                Next

286             j = j + 1

288             If Not GlobalGISGridLayer.Shape Is Nothing Then
290                 lUID = GlobalGISGridLayer.Shape.uID
292                 GISGridRS.Fields(0).Value = lUID
                End If

                ' arVal(0) = GlobalGISGridLayer.Shape.uID
294             GlobalGISGridLayer.MoveNext
                 
            Loop

296         If SafeMoveFirst(GISGridRS) Then
298             Set .DataSource = GISGridRS
300             .Columns.RetrieveFields
302             .Columns(0).Visible = False

304             If .Columns.Count > 1 Then
306                 If GISGridRS.Fields(1).Name = "GIS_UID" Or GISGridRS.Fields(1).Name = "UID" Or GISGridRS.Fields(1).Name = "ID" Then
308                     .Columns(1).Visible = False
                    End If
                End If
                
310             .KeyField = GISGridRS.Fields(0).Name
            End If
        
        End With
        
312     dxGISDataGrid.Visible = True
314     dxProgressBar1.Visible = False
        'LockWindowUpdate 0

        '<EhFooter>
        Exit Sub

LoadLayerAttrDataToGrid_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.LoadLayerAttrDataToGrid " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadLimitedAttr2Grid(oLayer As XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo LoadLimitedAttr2Grid_Err
        '</EhHeader>
        Dim vcm As Variant
        Dim i As Integer
        Dim lngFldType As Long
        Dim lngDataLength As Long
        Dim flds() As String
        Dim col As DXDBGRIDLibCtl.dxGridColumn ' Variant
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
116                 If Not .Fields.Item("excludedFlds").Value = vbNull Then
118                     sExcludedFlds = Split(.Fields.Item("excludedFlds").Value, ",")
                    End If
                End If
            End If
        
120         If Not .EOF Then
122             If .Fields.Item("datasetwarning").Value Then
124                 If iTot > .Fields.Item("MaxRec").Value Then
126                     MsgBox "Note! According to performace settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").Value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
                        Exit Sub
                    Else

128                     If iTot > .Fields.Item("warninglevel").Value Then
130                         If MsgBox("The data you are about to browse contains:" & iTot & " records." & vbCrLf & "This might take very long time. Would you like to continue?", vbYesNo) = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                    
                Else

132                 If iTot > .Fields.Item("MaxRec").Value Then
134                     MsgBox "Note! According to performace settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").Value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
                        Exit Sub
                    End If
                End If

            Else
                'UserDefGridRecords
                'the Layer is userdrawn, Use default values from AppSettings
136             SafeMoveFirst g_RSAppSettings
138             g_RSAppSettings.Find "SettingName = 'UserDefGridRecords'"
                
140             If g_RSAppSettings.Fields.Item("SettingValue1").Value = "1" Then
142                 If iTot > CLng(g_RSAppSettings.Fields.Item("SettingValue2").Value) Then
144                     MsgBox "Note! According to performace settings." & vbCrLf & "The OASIS System Administrator has limited the Max Record grid items to:" & .Fields.Item("MaxRec").Value & vbCrLf & "The data you are trying to browse is " & iTot & " Records.", vbInformation
                        Exit Sub
                    Else

146                     If iTot > CLng(g_RSAppSettings.Fields.Item("SettingValue3").Value) Then
148                         If MsgBox("The data you are about to browse contains:" & iTot & " records." & vbCrLf & "This might take very long time. Would you like to continue?", vbYesNo) = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        
        End With
    
150     oLayer.MoveFirst GIS.viewer.VisibleExtent, "", Nothing, "", True
    
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
208         Set col = .Columns.Add(gedTextEdit)
210         col.caption = "GIS_UID"
212         col.FieldName = "GIS_UID"
214         col.Visible = False
            
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
        
                        Case Is = XgisFieldTypeString '= 0,
236                         lngFldType = xftString
238                         lngDataLength = oLayer.Fields.Item(i).Width
        
240                     Case Is = XgisFieldTypeNumber ' = 1,
242                         lngFldType = xftInteger
244                         lngDataLength = oLayer.Fields.Item(i).Width
        
246                     Case Is = XgisFieldTypeFloat '= 2,
248                         lngFldType = xftFloat
250                         lngDataLength = oLayer.Fields.Item(i).Width
        
252                     Case Is = XgisFieldTypeBoolean '= 3,
254                         lngFldType = xftBoolean
256                         lngDataLength = oLayer.Fields.Item(i).Width
        
258                     Case Is = XgisFieldTypeDate '= 4
260                         lngFldType = xftDate
262                         lngDataLength = oLayer.Fields.Item(i).Width
                    End Select
                            
264                 .Dataset.MemoryDataset.AddField oLayer.Fields.Item(i).Name, lngFldType, lngDataLength
                    
266                 Set col = .Columns.Add(gedTextEdit)
268                 col.caption = oLayer.Fields.Item(i).Name
270                 col.FieldName = oLayer.Fields.Item(i).Name
                End If
                
            Next
    
272         oLayer.MoveFirst GIS.viewer.VisibleExtent, "", Nothing, "", True
274         .KeyField = "GIS_UID"
276         .Dataset.Open
    
            Dim j As Integer
            Dim strVals As Variant
            Dim arVals As Variant
            Dim arVal As Variant
            Dim varVal As Variant
        
278         arVals = Split(strVals, ",")
    
280         Do While Not oLayer.EOF
    
282             strVals = ""
    
284             ReDim arVal(UBound(flds))
286             arVal(0) = oLayer.Shape.uID

288             j = 0

290             For i = 1 To oLayer.Fields.Count - 1

292                 bExclude = False
                
294                 If Not ((Not sExcludedFlds) = -1&) Then
296                     If SequentialSearchStringArray(sExcludedFlds, oLayer.Fields.Item(i - 1).Name) > -1 Then
298                         bExclude = True
                        End If
                    End If
    
300                 If Not bExclude Then
                    
302                     m_frmDebug.DebugPrint oLayer.Fields.Item(i - 1).Name
                    
304                     Select Case oLayer.Fields.Item(i - 1).FieldType 'i to compensate the GIS_UID on 0 position
    
                            Case Is = XgisFieldTypeString '= 0,
306                             varVal = CStr(oLayer.Shape.GetField(flds(j + 1)))
308                             arVal(j + 1) = varVal
    
310                         Case Is = XgisFieldTypeNumber ' = 1,
                                'm_frmDebug.DebugPrint  oLayer.Shape.GetField(flds(j))
312                             arVal(j + 1) = CDbl(oLayer.Shape.GetField(flds(j + 1)))
    
314                         Case Is = XgisFieldTypeFloat '= 2,
316                             arVal(j + 1) = CDbl(oLayer.Shape.GetField(flds(j + 1)))
    
318                         Case Is = XgisFieldTypeBoolean '= 3,
320                             arVal(j + 1) = CBool(oLayer.Shape.GetField(flds(j + 1)))
    
322                         Case Is = XgisFieldTypeDate '= 4
324                             arVal(j + 1) = CDate(oLayer.Shape.GetField(flds(j + 1)))
                        End Select
    
326                     strVals = IIf(strVals <> "", strVals & ",", "") & oLayer.Shape.GetField(flds(j + 1))
                    
328                     j = j + 1
                    
                    End If
                
                Next
    
330             arVals = arVal
332             .Dataset.AppendRecord arVals
334             oLayer.MoveNext
            Loop
        
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

Private Sub abGridTools_ComboSelChange(ByVal Tool As ActiveBar3LibraryCtl.Tool)
        '<EhHeader>
        On Error GoTo abGridTools_ComboSelChange_Err

        Dim l As Long
        Dim lMAX As Long

        '</EhHeader>
100     Select Case Tool.Name

            Case Is = "comIncCategory"
102             CategorizeIncidents

104         Case Is = "comLyr"
            
                LoadLayerAttrDataToGridInit
                CreateSummaryGroup
116         Case Is = "comDateFrom"
                'FORMAT "dd.mm.yyyyy" Incident Date (ddmmyyyy)
118             LoadOASISIncidents
                'abGridTools.Tools.Item("comDateFrom").Text

120         Case Is = "s"
            
        End Select
        
        dxGISDataGrid.Option = egoAutoWidth
        dxGISDataGrid.OptionEnabled = True
        mnuAutoSize.Checked = True
        
        
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

Private Sub CreateSummaryGroup()
    Dim sumgroup As DXDBGRIDLibCtl.dxGridSummaryGroup
    Dim sumitem As DXDBGRIDLibCtl.dxGridSummaryItem
    Dim i As Integer
    Dim col As DXDBGRIDLibCtl.dxGridColumn

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
            Set col = .Columns(i)

            If col.Visible Then
                col.SummaryFooterType = cstCount
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
        Dim shp As XGIS_Shape
        Dim lL As XGIS_LayerVector
        Dim SymbolList As New XGIS_SymbolList
        Dim RS As New adodb.Recordset
        Dim sFont As String
        Dim sDBFieldName As String
        Dim sGISFieldName As String
  
100     SafeMoveFirst g_RSAppSettings
102     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"

        m_frmDebug.DebugPrint "OASIS Incident Layer Name:" & g_RSAppSettings.Fields.Item("SettingValue1").Value

104     Set lL = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)

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
    
106     Set RS = New adodb.Recordset
                
108     Set RS.ActiveConnection = m_Cnn
                
110     RS.CursorType = adOpenDynamic
112     RS.CursorLocation = adUseClient
114     RS.LockType = adLockReadOnly
                
116     Select Case m_frmMnuOperations.abGridTools.Tools.Item("comIncCategory").Text  'shp.GetField("Type")

            Case "--None--"
118             sDBFieldName = "Incident_Type_Name"
120             sGISFieldName = "Type"
122             RS.Open "SELECT * FROM IncTypeCategory ORDER BY [Incident_Type_Name]"
                m_frmDebug.DebugPrint " None Category"

124         Case "Target"
126             sDBFieldName = "Name"
128             sGISFieldName = "Type"
130             RS.Open "SELECT * FROM IncTarget ORDER BY [NAME]"
                m_frmDebug.DebugPrint " Incident Target Category"

132         Case "Time"
134             sDBFieldName = "Incident_Time_Name"
136             sGISFieldName = "TIME00"
138             RS.Open "SELECT * FROM IncTimeCategory ORDER BY [Incident_Time_Name]"
                m_frmDebug.DebugPrint " Time Category"

140         Case "Type"
142             sDBFieldName = "Incident_Type_Name ORDER BY [Incident_Type_Name]"
144             sGISFieldName = "Type"
146             RS.Open "SELECT * FROM IncTypeCategory ORDER BY [Incident_Type_Name]"
                m_frmDebug.DebugPrint " Type Category"
        End Select
        
148     Set shp = lL.FindFirst(GIS.viewer.Extent, "", Nothing, "", True)

150     While Not shp Is Nothing
            m_frmDebug.DebugPrint " INCIDENT Type:" & shp.GetField("Type")

152         SafeMoveFirst RS
154         RS.Find sDBFieldName & " = " & "'" & shp.GetField(sGISFieldName) & "'"
            m_frmDebug.DebugPrint sGISFieldName & " Value = " & shp.GetField(sGISFieldName)
            
156         If Not RS.BOF And Not RS.EOF Then
158             sFont = RS.Fields("Font_Name").Value
160             sFont = sFont & ":" & RS.Fields("Ascii").Value & ":NORMAL"
                m_frmDebug.DebugPrint " Font:" & sFont
            End If
                
162         With shp.Params.Marker
                
164             .Color = vbWhite
166             .OutlineColor = vbBlue
168             .Symbol = SymbolList.Prepare(sFont) 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
170             .Size = 640
                
            End With
        
172         shp.draw
        
174         Set shp = lL.FindNext()
        Wend
    
176     lL.Paint
    
        'GIS.Update

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
102     frddmmss.Move frmdddddd.Left, frmdddddd.Top
104     frddmmmm.Move frmdddddd.Left, frmdddddd.Top
106     frMGRS.Move frmdddddd.Left, frmdddddd.Top

108     With g_RSAppSettings
        
110         .Requery
112         SafeMoveFirst g_RSAppSettings
114         .Find "SettingName = 'AdminLevel0'"

116         If Not .Fields.Item("SettingValue1").Value = vbNullString Then m_sAdmVal1 = .Fields.Item("SettingValue1").Value
        
118         SafeMoveFirst g_RSAppSettings
120         .Find "SettingName = 'AdminLevel1'"

122         If Not .Fields.Item("SettingValue1").Value = vbNullString Then m_sAdmVal2 = .Fields.Item("SettingValue1").Value
        
124         SafeMoveFirst g_RSAppSettings
126         .Find "SettingName = 'AdminLocation'"

128         If Not .Fields.Item("SettingValue1").Value = vbNullString Then m_sAdmLoc = .Fields.Item("SettingValue1").Value
        
            Dim sModMenus() As String
            Dim i As Integer
    
130         SafeMoveFirst g_RSAppSettings
132         .Find "SettingName = 'VisibleMainModuleMenus'"
    
134         sModMenus = Split(.Fields.Item("SettingValue1").Value, ",")
        
136         For i = 0 To AB.Bands("bNavPane").ChildBands.Count - 1
138             AB.Bands("bNavPane").ChildBands(i).Visible = False
            Next
        
140         For i = 0 To UBound(sModMenus)
            
                '142             If sModMenus(i) = "cbProfile" Then
                '144                 .MoveFirst
                '146                 .Find "SettingName = 'InitURL'"
                '148                 WebBrowser1.Navigate2 .Fields.Item("SettingValue1").Value '"http:\\www.google.com"
                '                End If
            
150             AB.Bands("bNavPane").ChildBands(sModMenus(i)).Visible = True
152         Next i
        
154         CheckAvailableTools
        
156         SafeMoveFirst g_RSAppSettings
158         .Find "SettingName = 'UserDefGridRecords'"
        
160         If Not IsNull(.Fields.Item("SettingValue4").Value) Then
                
162             If Len(.Fields.Item("SettingValue4").Value) > 2 Then
164                 sGridOpt = Split(.Fields.Item("SettingValue4").Value, ",")
166                 chkFilterIn.Visible = IIf(sGridOpt(0) = "1", True, False)

                    If sGridOpt(0) <> "1" Then chkSelectIn.Left = chkFilterIn.Left
168                 chkSelectIn.Visible = IIf(sGridOpt(1) = "1", True, False)
                End If
        
            End If
        
170         SafeMoveFirst g_RSAppSettings
172         .Find "SettingName = 'CurrentActiveMeny'"
        
174         AB.Bands("bNavPane").ChildBands.CurrentChildBand = AB.Bands("bNavPane").ChildBands(.Fields.Item("SettingValue1").Value)

            If AB.Bands("bNavPane").ChildBands.CurrentChildBand.Name = "cbOperations" Then
                m_bMapInitialized = False
                ActivateOperations
            
                '(keith) Added this to fix non rendering of map when proile was set as default view
            ElseIf AB.Bands("bNavPane").ChildBands.Item("cbProfile").Visible Then
            
                InitMap
            End If

176         AB.RecalcLayout

            'Dim cn As New ADODB.Connection
            
178         If Not m_frmAddWhere.IsInitialized Then
                ' cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OASIS\F;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                ' m_frmAddWhere.Init cn
            End If

180         SafeMoveFirst g_RSAppSettings
182         .Find "SettingName = 'ShowExpandedToolBoxInGISWin'"
        
184         If CInt(.Fields("SettingValue1").Value) = 1 Then
186             C1TabFastFunction.Visible = True
188             C1TabFastFunction.TabsPerPage = CInt(.Fields("SettingValue2").Value)

                'GeoMarks
                If CInt(.Fields("SettingValue3").Value) = 1 Then
190                 C1TabFastFunction.TabVisible(1) = True
                    LoadGeoBookMarks
                Else
                    Themes.Visible = False
                    C1TabFastFunction.TabVisible(1) = False
                End If
                
                'Geo Convert
                If CInt(.Fields("SettingValue4").Value) = 1 Then
194                 C1TabFastFunction.TabVisible(2) = True
                Else
                    elGeoCalc.Visible = False
                    C1TabFastFunction.TabVisible(2) = False
                End If
                
                'Goto
                If CInt(.Fields("SettingValue5").Value) = 1 Then
196                 C1TabFastFunction.TabVisible(3) = True
                Else
                    C1TabFastFunction.TabVisible(3) = False
                End If

                'Settings
                If CInt(.Fields("SettingValue6").Value) = 1 Then
198                 C1TabFastFunction.TabVisible(4) = True
                Else
                    C1Elastic1.Visible = False
                    C1TabFastFunction.TabVisible(4) = False
                End If
                
                'Magnifier
                If CInt(.Fields("SettingValue7").Value) = 1 Then
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
208         m_frmMnuOperations.cmdSecurityAnalysis.Visible = IIf(.Fields.Item("SettingValue1").Value = "1", True, False)
210         m_frmMnuOperations.FraAnalysisLevel.Visible = IIf(.Fields.Item("SettingValue1").Value = "1", True, False)
212         m_frmMnuOperations.cmdCreateCharts.Visible = IIf(.Fields.Item("SettingValue2").Value = "1", True, False)
214         m_frmMnuOperations.cmdInsertIncident.Visible = IIf(.Fields.Item("SettingValue3").Value = "1", True, False)
216         m_frmMnuOperations.cmdSetScope.Visible = IIf(.Fields.Item("SettingValue4").Value = "1", True, False)
        
218         SafeMoveFirst g_RSAppSettings
220         .Find "SettingName = 'MineActionTools'"
222         m_frmMnuOperations.DxHUpdateData.Visible = IIf(.Fields.Item("SettingValue1").Value = "1", True, False)
        
224         SafeMoveFirst g_RSAppSettings
226         .Find "SettingName = 'InetConnectionSettings'"

228         tmrInternetCheck.Interval = IIf(1000 * CLng(.Fields.Item("SettingValue2").Value) > 65000, 65000, 1000 * CLng(g_RSAppSettings.Fields.Item("SettingValue2").Value))
230         tmrInternetCheck.Enabled = IIf(.Fields.Item("SettingValue1").Value = "1", True, False)

232         SafeMoveFirst g_RSAppSettings
234         .Find "SettingName = 'OpsTools'"
        
236         If Not .EOF Then
238             cmdCommand2.Visible = IIf(.Fields.Item("SettingValue1").Value = "1", True, False)
            Else
240             cmdCommand2.Visible = False
            End If

            SafeMoveFirst g_RSAppSettings
            .Find "SettingName = 'RangeSettingsColours'"
                        
            If Not .EOF Then
                                
                For i = 1 To 10
                    
                    If Not IsNull(.Fields.Item("SettingValue" & i).Value) Then
                        If IsNumeric(.Fields.Item("SettingValue" & i).Value) Then
                            ReDim Preserve arRangeColors(i - 1)
                            arRangeColors(i - 1) = CLng(.Fields.Item("SettingValue" & i).Value)
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
                        arRangeVals(i - 1) = CLng(.Fields.Item("SettingValue" & i).Value)
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
248             If Not IsNull(.Fields.Item("SettingValue2").Value) Then
250                 If Len(.Fields.Item("SettingValue2").Value) > 4 Then
                
                        Dim sScripts() As String
                        Dim k As Integer
                    
252                     sScripts = Split(.Fields.Item("SettingValue2").Value, ",")
                    
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
        Dim CatRS As New adodb.Recordset
        Dim bkrRS As New adodb.Recordset

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
    
130     If Not CatRS.BOF Then SafeMoveFirst CatRS
    
    
132     .AddNode (m_DefaultViewName), 0, 1
    
134     Do While Not CatRS.EOF
136         bkrRS.Open "SELECT * FROM GeoBookMarks WHERE BmkrID = " & CatRS.Fields.Item("ID").Value & " ORDER BY Name", m_Cnn
        
        
        
138         If Not bkrRS.BOF Then
140             SafeMoveFirst bkrRS
142             .AddNode (CatRS.Fields.Item("Name").Value), 0, 1
            End If
        
144         Do While Not bkrRS.EOF
                '.AddPictureNode bkrRS.Fields.item("Name").Value, 0, 2, 3, 0, 12
146             .AddNode (bkrRS.Fields.Item("Name").Value), 0, 2
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

Private Sub endEdit()
        '<EhHeader>
        On Error GoTo endEdit_Err
        '</EhHeader>

100   g_CurrentTool = m_prevTool
102   SetOASISTool

    '  If g_oEditLayer Is Nothing Then Exit Sub
    'On Error Resume Next
104     If Not m_oSQLOpsLyr Is Nothing Then
106            m_oSQLOpsLyr.SaveAll
        End If
      
      On Error Resume Next
      
108   GIS.Editor.endEdit
110   Set g_oEditLayer = Nothing
  
  
        '<EhFooter>
        Exit Sub

endEdit_Err:
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
106       GIS.Mode = XgisEdit
          'GIS.Editor.ShowTracking = True
108       g_CurrentFeatureType = Location
110       m_prevTool = g_CurrentTool
          'g_CurrentTool = oCreateLocationPolyline
        
112     m_frmAddSHPWiz.Show vbModal, Me
     
    '     GIS.Mode = XgisEdit
    '     GIS.Editor.Mode = XgisEditorModeAfterActivePoint
    '     Set g_oEditLayer = GIS.Get(comActiveLyr.List(comActiveLyr.ListIndex))
     
    '     GIS.Editor.ShowTracking = True

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
    Dim lL As New XGIS_LayerSHP
    Dim OASISLyr As String
    Dim AnalysisLyr As String


    'Dim sCriteria As String
    'Dim oRS As New ADODB.Recordset
    '
    'sCriteria = InputBox("Enter The Name of the Area you would like to filter on", Default:="Colombia City")
    '
    'oRS.Open "SELECT & FROM oIncidents WHERE Admin3 = '" & sCriteria & "'", cnn
    '
    'Set dxGISDataGrid.DataSource = oRS

    'dxGISDataGrid.Dataset.ADODataset

    


100     GIS.Lock

        If Not EventLayer Is Nothing Then
102         sPath = EventLayer.Path
104         sName = EventLayer.Name
        End If
        
106     SafeMoveFirst g_RSAppSettings
108     g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"
110     OASISLyr = g_RSAppSettings.Fields.Item("SettingValue1").Value
    
        'g_RSAppSettings.MoveFirst
        'g_RSAppSettings.Find "SettingName = 'AdminLevel0'"
        'AnalysisLyr = g_RSAppSettings.Fields.Item("SettingValue1").value
    
        'EventLayer.RereadConfig
    
112     CheckIncidentFrequency GIS.get(OASISLyr), GIS.get(m_sSecAnalysisLyrName) 'SaveAsShape("Provinces", "Analyzed")       '  GIS.Get("Provinces")

        'EventLayer.RereadConfig
    
114     Set EventLayer = Nothing
116     Set EventLayer = GIS.get(m_sSecAnalysisLyrName)
    
118     EventLayer.Params.Visible = True
    
120     If Not sName = "" Then
122         If m_bThematicsDone Then
124             GIS.get(sName).RereadConfig
            Else
126             m_bThematicsDone = True
            End If
        End If
    
128     GIS.Unlock
130     GIS.UpDate
    
        Exit Sub


132     GIS.Delete sName
    
    
134     Set EventLayer = Nothing
    
        'GIS.Update
    
136     lL.Path = sPath
138     lL.Name = sName
140     lL.Open
    
142     GIS.Add lL
    
144     Set EventLayer = GIS.get(sName)
    
146     GIS.UpDate
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
    Dim lL As New XGIS_LayerSHP
    Dim OASISLyr As String
    Dim AnalysisLyr As String
    
    

100     GIS.Lock
    
102     m_sSecAnalysisFieldName = "ScoringFld"
    
104     If sFreqOverlayLyr = sAnalysisLyr Then Exit Sub
    
106     CheckFrequency GIS.get(sFreqOverlayLyr), GIS.get(sAnalysisLyr), "ScoringFld" 'SaveAsShape("Provinces", "Analyzed")       '  GIS.Get("Provinces")

108     Set EventLayer = Nothing
110     m_sSecAnalysisFieldName = "ScoringFld"
112     Set EventLayer = GIS.get(sAnalysisLyr)
    
114     GIS.Unlock
116     GIS.UpDate

        '<EhFooter>
        Exit Sub

DoFrequencyAnalysis_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.DoFrequencyAnalysis " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateScoringField(oTargetLayer As XGIS_LayerVector, sFieldAnalysisName As String)
        '<EhHeader>
        On Error GoTo CreateScoringField_Err
        '</EhHeader>
        On Error GoTo ErrCode
    
        'If Not oTargetLayer.Shape.GetField(sFieldAnalysisName) = vbNull Then
        
        '    Exit Sub
        'Else
        On Error Resume Next
100      oTargetLayer.AddField sFieldAnalysisName, XgisFieldTypeNumber, 16, 0
        'End If
        Exit Sub
ErrCode:
        'oTargetLayer.AddField sFieldAnalysisName, XgisFieldTypeNumber, 16, 0
        '<EhFooter>
        Exit Sub

CreateScoringField_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.CreateScoringField " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckFrequency(oOverLayLayer As XGIS_LayerVector, _
                                  oTargetLayer As XGIS_LayerVector, sFieldAnalysisName As String, _
                                  Optional bResetPreviousScoring As Boolean = True)
        '<EhHeader>
        On Error GoTo CheckFrequency_Err
        '</EhHeader>
        Dim shp As XGIS_Shape
        Dim aShape As Object
        Dim oPolyShape As XGIS_ShapePolygon
        Dim iVal As Integer

100     CreateScoringField oTargetLayer, sFieldAnalysisName

102     If bResetPreviousScoring Then
104         ResetScoring oTargetLayer
        End If

106     oOverLayLayer.MoveFirst GIS.viewer.VisibleExtent, "", Nothing, "", True
    
108     Set shp = oOverLayLayer.FindFirst(GIS.viewer.VisibleExtent, "", Nothing, "", True)

110     While Not shp Is Nothing

112         Set aShape = oTargetLayer.FindFirst(shp.Extent, "", shp, GisUtils.GIS_RELATE_WITHIN, True)
        
114         If aShape Is Nothing Then
                'm_frmDebug.DebugPrint  "Hit! NUM:" & i & " INCIDENT NAME:" & shp.GetField("Name") & " ADM Name: Out of Bounds"
          
            Else

116             If Not aShape.GetField(sFieldAnalysisName) = vbNull Then
118                 Set oPolyShape = aShape.MakeEditable
120                 iVal = aShape.GetField(sFieldAnalysisName) + 1
122                 oPolyShape.SetField sFieldAnalysisName, CLng(iVal)  'CInt(aShape.GetField("Scoring") + 1)
                Else
                    'TODO CHECK IF THIS IS NEEDED!
                    'Set oPolyShape = aShape.MakeEditable
                    'oPolyShape.SetField "sFieldAnalysisName", 100
                End If
            
            End If

124         Set shp = oOverLayLayer.FindNext()
        Wend
    
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
        'If Node.Selected Then
            AutoZoom Node
        'End If
    End If

    'm_frmDebug.DebugPrint  node.Values(0) & node.Selected
End Sub

Private Sub dxGISDataGrid_OnFilterRecord(accept As Boolean)
        '<EhHeader>
        On Error GoTo dxGISDataGrid_OnFilterRecord_Err
        '</EhHeader>
100     accept = True
102     'm_frmDebug.DebugPrint "OnFilterRecord"
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
106         PopupMenu mnuGridAction, x:=ScaleX(ptPopUpPos.x, vbPixels, vbTwips) - Me.Left, y:=ScaleY(ptPopUpPos.y, vbPixels, vbTwips) - (Me.Top + 250)
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
'    m_frmDebug.DebugPrint  "MOUSEMOVE: " & X
End Sub

Private Sub dxGISDataGrid_OnSelectedCountChange()
        '<EhHeader>
        On Error GoTo dxGISDataGrid_OnSelectedCountChange_Err
        '</EhHeader>
        Dim i As Integer
        Dim Node As dxGridNode
        Dim bookmark As Variant

        'm_frmDebug.DebugPrint  "SEL COUNT CHANGE " & Now & dxGISDataGrid.ex.SelectedCount
    
100     If Not dxGISDataGrid.M.IsGridMode Then

102         For i = 0 To dxGISDataGrid.ex.SelectedCount - 1
104             Set Node = dxGISDataGrid.ex.SelectedNodes(i)
                'm_frmDebug.DebugPrint  node.Values(0) ' SelectedNodes(i)
            Next

        Else

106         For i = 0 To dxGISDataGrid.ex.SelectedCount - 1
108             bookmark = dxGISDataGrid.ex.SelectedRows(i)

110             If dxGISDataGrid.Dataset.BookmarkValid(bookmark) Then
                    'm_frmDebug.DebugPrint  dxGISDataGrid.Dataset.FieldValues(dxGISDataGrid.Dataset.FieldNameByNo(0))
112                 dxGISDataGrid.Dataset.GotoBookmark bookmark
                    'm_frmDebug.DebugPrint  dxGISDataGrid.Dataset.FieldValues(dxGISDataGrid.Dataset.FieldNameByNo(0))
                
                    '...
                End If

                'm_frmDebug.DebugPrint  node.Values(0) ' SelectedNodes(i)
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

        Dim lShape As XGIS_Shape
     
100     With dxGISDataGrid.ex
        
102         Set lShape = GlobalGISGridLayer.GetShape(Node.values(0))

104         If Not lShape Is Nothing And Not IsNull(lShape) Then
106             GIS.Lock
108             GIS.VisibleExtent = lShape.Extent
110             GIS.Unlock

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

Private Sub HideShowSelected(oLyr As XGIS_LayerVector, _
                             bShow As Boolean)
        '<EhHeader>
        On Error GoTo HideShowSelected_Err
        '</EhHeader>
    
        Dim shp As XGIS_Shape
        Dim xt As XGIS_Extent
            
100     Set shp = oLyr.FindFirst(oLyr.Extent, "GIS_Selected=True", Nothing, "", True)
            
102     GIS.Lock
            
104     Do While Not shp Is Nothing
106         Set shp = shp.MakeEditable
108         shp.IsHidden = Not bShow
            shp.Invalidate True
110         Set shp = oLyr.FindNext
        Loop
            
112     GIS.Unlock
    
114     'GIS.UpDate
    
        '<EhFooter>
        Exit Sub

HideShowSelected_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.HideShowSelected " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub elTatukGIS_RealignFrame()
        'abCOP.Width = elTatukGIS.Width
        'abCOP.Bands("bnMainMap").Width = elTatukGIS.Width
    'bnMainMap
       ' abCOP.Bands("bnMainMap").Width = elTatukGIS.Width
        '<EhHeader>
        On Error GoTo elTatukGIS_RealignFrame_Err
        '</EhHeader>
100     GIS.Width = elTatukGIS.Width '+ 1000
'102     GIS.Height = elTatukGIS.Height + 400 '+ 1000
    
    If MsgScroll.ListCount > 0 Then
        GIS.Height = elTatukGIS.Height + 150
        elScroller.Visible = True
        elScroller.Move 5, elTatukGIS.Height - MsgScroll.Height, elTatukGIS.Width
        elScroller.Move 5, elTatukGIS.Height - MsgScroll.Height, elTatukGIS.Width, 255
        elScroller.ZOrder
        cmdMSGScroller.ZOrder
    Else
        elScroller.Visible = False
        GIS.Height = elTatukGIS.Height + 400
    End If
    
    C1TabFastFunction.ZOrder 0

        '<EhFooter>
        Exit Sub

elTatukGIS_RealignFrame_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.elTatukGIS_RealignFrame " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateAddedScribbles()
        '<EhHeader>
        On Error GoTo CreateAddedScribbles_Err
        '</EhHeader>
        Dim xportLyr As XGIS_LayerSqlAdo
        Dim sAdoString As String
        Dim sLayerName As String
        Dim RS As adodb.Recordset
        Dim col As ADOx.Column

        '    If m_oDrawLyr Is Nothing Then Exit Sub
        '
        '    Set xportLyr = New XGIS_LayerSqlAdo

100     Set m_oDrawLyr = New XGIS_LayerSqlAdo
102     m_oDrawLyr.Params.area.Color = RGB(0, 0, 255)
104     m_oDrawLyr.Transparency = 70
106     m_oDrawLyr.Name = "Draw_Layer"
108     m_oDrawLyr.HideFromLegend = True

110     sAdoString = GetConnectionString(g_sAppPath & "\Data\Db\OASISClient.mdb")

112     sAdoString = "[TatukGIS Layer]" & vbCrLf & "Storage = Native" & vbCrLf & "layer=" & "Draw_Layer" & vbCrLf & "Dialect=MSJET" & vbCrLf & "ADO=" & sAdoString

114     If Not DoTableExists("Draw_Layer_GEO", m_Cnn) Then
116         Set RS = New adodb.Recordset

118         RS.Open "SELECT * FROM oincidents_GEO WHERE UID = 1000000", m_Cnn

120         CreateTable "Draw_Layer_GEO", RS, m_Cnn

124         Set RS = New adodb.Recordset
126         RS.Open "SELECT UID, NAME FROM oincidents_FEA WHERE UID = 1000000", m_Cnn
            '128         RS.Fields.Append "StringStyle", adVarChar, 255
128         CreateTable "Draw_Layer_FEA", RS, m_Cnn
               
            Set col = New ADOx.Column

130         With col
132             .Name = "StringStyle"
134             .Type = adVarChar
136             .DefinedSize = 255
            End With
                
            AddFieldToTable m_Cnn, "Draw_Layer_FEA", col
                
            Set col = Nothing
        End If

138     m_oDrawLyr.Path = sAdoString
        '116 oLyr.ExportLayer xportLyr, GisUtils.GisWholeWorld, XgisShapeTypeUnknown, "", True
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
        Dim RSUpdater As adodb.Recordset
        
        SavePrivateAppSettings
        
        If pLoadLyrAttrToGrdThread.IsThreadRunning Then
            pLoadLyrAttrToGrdThread.TerminateWin32Thread
        End If
        
        If Not m_oDrawLyr Is Nothing Then
        
            On Error Resume Next

            If g_ZoomToSettings.SaveOnExit Then m_oDrawLyr.SaveAll
        
        End If
        
100     OASISSynchFolderImporter.ScanAndProcessSynchFolder g_sAppPath & "\Data\Sync\import", g_sAppPath & "\Data\Db\OASISClient.mdb"
102     Set OASISSynchFolderImporter = Nothing

        If Not SilentHttpComms Is Nothing Then Set SilentHttpComms = Nothing
        If Not GISGridRS Is Nothing Then Set GISGridRS = Nothing
104     endEdit
    
106     If Not m_Cnn Is Nothing Then
108         If Not m_Cnn.State = adStateClosed Then
                      
                Set RSUpdater = New adodb.Recordset

                With RSUpdater

                    .Open "SELECT * FROM Personnell", m_Cnn, adOpenDynamic, adLockBatchOptimistic
                    .Find "Personnell_ID = " & g_CurrentUserID

                    If Not .EOF Then
                        .Fields("LatestViewX").Value = Replace(GIS.viewer.CenterPtg.x, ",", ".")
                        .Fields("LatestViewY").Value = Replace(GIS.viewer.CenterPtg.y, ",", ".")
                        .Fields("LatestViewZ").Value = Replace(GIS.viewer.zoom, ",", ".")
                        .Fields("LatestMapName").Value = GIS.Name
                        .UpdateBatch adAffectCurrent
                        .Close
                    End If

                End With

                Set RSUpdater = Nothing
            
            End If
        End If

        '    If Not GIS.MustSave Then
        '        Exit Sub
        '        On Error Resume Next
        '
        '        UnloadMenuFrames
        '
        '        Unload Me
        '        End
        '
        '    End If
112     SafeMoveFirst g_RSAppSettings
114     g_RSAppSettings.Find "SettingName = 'InitialOperationsTabNumber'"
                     
116     If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue2").Value) Then
118         If g_RSAppSettings.Fields.Item("SettingValue2").Value = "1" Then bPrompt = True
        End If

120     If bPrompt Then
122         If MsgBox("Do you want to save changes to the map?", vbYesNo, "OASIS Client") = vbYes Then
                'm_oSQLIncLyr.SaveAll
                On Error Resume Next
                
124             'Without these lines of code the incidents layer will not appear next load
                GIS.Delete m_oSQLIncLyr.Name
                GIS.Delete "Draw_Layer"
                GIS.SaveAll
                UpdateProjectFileToDB
            End If
        End If
    
        On Error Resume Next
        'OASISFolderMonitorImporter.StopMonitoring
        
        'Set OASISFolderMonitorImporter = Nothing
        
126     SaveStartUpParams

        g_clsHotKey.UnregisterKey "Language"
        g_clsHotKey.UnregisterKey "OPSV"
        g_clsHotKey.UnregisterKey "doPrint"
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
        g_clsHotKey.UnregisterKey "ResourcesFinder"
        g_clsHotKey.UnregisterKey "WVISitRep"
        
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
        
200     Set RSDataPacks = Nothing
202     Set mRSUGSettings = Nothing
    
204     UnloadMenuFrames
206     UnloadAllForms
    
208     m_Cnn.Close
    
210     Set m_Cnn = Nothing
    
212     Set g_PictureDialogLarge = Nothing
214     Set g_PictureDialogSmall = Nothing
216     Set g_PictureDialogLogo = Nothing
        
        oClientInterCom.UnregisterChannel
        Set oClientInterCom = Nothing
        
218     Set oSync = Nothing
220     Set m_InternetCheck = Nothing
222     Set m_oSQLLyrSynch = Nothing
        
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

'Private WithEvents EventLayer As XGIS_LayerVector
'
'Private WithEvents m_frmOvMap As frmOVMap
'Private WithEvents m_frmAddSHPWiz As frmAddPointWZ
'Private WithEvents m_fmrAddIncident As frmAddIncident
'Private WithEvents m_frmMAModule As frmMAModule
'Private WithEvents m_frmAddWhere As frmAddWhere
'Private WithEvents m_frmW3Wizard As frmW3Wizard
'Private WithEvents m_frmMnuOASISProfile As frmMnuOASISProfile
'Private WithEvents m_frmAddOns As frmAddons
'Private WithEvents m_frmChangeTracer As frmChangeTracer
'Private WithEvents m_frmFreqSettings As frmFreqSettings
'Private WithEvents m_frmIOMJOC As frmIOMJOC
'Private WithEvents m_frmMnuOperations As frmMnuOperations
'
'
'Private m_oCosmeticLayer As XGIS_LayerVector
'Private m_oDrawLyr As XGIS_LayerVector
'Private m_oBufferLyr As XGIS_LayerVector
'
'Private m_oIncidentLyr As XGIS_LayerVector
'Private m_oW3Lyr As XGIS_LayerVector
'Private m_oW3WHOLyr As XGIS_LayerVector
'Private m_oSQLIncLyr As XGIS_LayerSqlAdo
'Private m_oSQLOpsLyr As XGIS_LayerSqlAdo
'Private m_oIncRS As ADODB.Recordset
'Private g_oEditLayer As XGIS_LayerAbstract
'Private ProjectionList As New XGIS_ProjectionList
'
'Private m_oPtRubberStart As New XGIS_Point
'Private m_oLineRubber As New XGIS_ShapeArc
'Private m_oShpArc As XGIS_ShapeArc
'Private m_oShpPoint As XGIS_ShapePoint
'Private m_oShpPolygon As XGIS_ShapePolygon
'Private m_oShpUnknown As XGIS_Shape
'Private m_oIncShpPt As XGIS_ShapePoint
'
'
'Private m_navBar As NavigationBarExtension
'Private m_styleCombo As ToolbarStyleABCombo
'
'Private m_oXgisThemeWiz As New XGIS_ControlLegendVectorWiz
'Private m_cUploadThreads() As clsThreads
'
'Dim WithEvents FormThread As MThreadVB.Thread
'Dim WithEvents pThread As MThreadVB.Thread
'Dim WithEvents pInitThread As MThreadVB.Thread
'Dim WithEvents pSubmitGeoMarksThread As MThreadVB.Thread
'Dim WithEvents pLoadLyrAttrToGrdThread As MThreadVB.Thread
'Dim WithEvents pUpdateCheckerThread As MThreadVB.Thread
'
'Private RSDataPacks As New ADODB.Recordset
'Public mRSUGSettings As New ADODB.Recordset
'
'Private menuPos As XGIS_Point


End Sub

Private Sub ShowThemelegend(oLayer As XGIS_LayerVector)
'    If Not frmLegend.Visible Then
'        frmLegend.init GIS, oLayer
'        frmLegend.Show vbModeless, Me
'    Else
'        frmLegend.init GIS, oLayer
'        frmLegend.SetFocus
'    End If
    
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
102     Unload m_frmAddSHPWiz
104     Unload m_fmrAddIncident
106     Unload m_frmMAModule
108     Unload m_frmMnuOperations
110     Unload m_frmAddWhere
112     Unload m_frmW3Wizard
114     Unload m_frmMnuOASISProfile
116     Unload m_frmAddOns
118     Unload m_frmChangeTracer
120     Unload m_frmFreqSettings
122     Unload m_frmIOMJOC
        Unload m_frmAttributes
        Unload m_frmSearch
        Unload m_frmTextAnnoSettings
        Unload m_frmLocator
        Unload m_frmUpdateSettings
        Unload m_frmMainSettings
        Unload m_frmSelections
        Unload m_frmSelectorReports
        Unload m_frmSelectorSettings
        Unload m_frmDynamicContent
        
        Set m_frmDynamicContent = Nothing
        Set m_frmSelectorSettings = Nothing
        Set m_frmSelectorReports = Nothing
        Set m_frmSelections = Nothing
        Set m_frmMainSettings = Nothing
        Set m_frmAttributes = Nothing
        Set m_frmSearch = Nothing
124     Set m_frmIOMJOC = Nothing
126     Set m_frmFreqSettings = Nothing
128     Set m_frmChangeTracer = Nothing
130     Set m_frmAddOns = Nothing
132     Set m_frmMAModule = Nothing
134     Set m_frmOvMap = Nothing
136     Set m_frmAddSHPWiz = Nothing
138     Set m_fmrAddIncident = Nothing
140     Set m_frmMnuOperations = Nothing
142     Set m_frmAddWhere = Nothing
144     Set m_frmW3Wizard = Nothing
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

Private Sub GIS_OnAfterPaint(translated As Boolean)
        '<EhHeader>
        On Error GoTo GIS_OnAfterPaint_Err
        '</EhHeader>
            Dim ex As XGIS_Extent
            
            If Not m_bPrevActionUsed Then
                ReDim Preserve m_PrevExt(UBound(m_PrevExt) + 1)
                Set m_PrevExt(UBound(m_PrevExt)) = GIS.VisibleExtent
            Else
                m_bPrevActionUsed = False
            End If
            
100         translated = True
102         Set ex = GIS.VisibleExtent
104         Set lP1 = GisUtils.GISPoint(ex.XMin, ex.YMin)
106         Set lP2 = GisUtils.GISPoint(ex.XMax, ex.YMin)
108         Set lP3 = GisUtils.GISPoint(ex.XMax, ex.YMax)
110         Set lP4 = GisUtils.GISPoint(ex.XMin, ex.YMax)
    '        lblP1.Caption = "P1 : x: " + Str(lP1.X) + "   y: " + Str(lP1.Y)
    '        lblP2.Caption = "P2 : x: " + Str(lP2.X) + "   y: " + Str(lP2.Y)
    '        lblP3.Caption = "P3 : x: " + Str(lP3.X) + "   y: " + Str(lP3.Y)
    '        lblP4.Caption = "P4 : x: " + Str(lP4.X) + "   y: " + Str(lP4.Y)
112         miniMapRefresh
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
100     Set g_PrevExt = GIS.VisibleExtent
        '<EhFooter>
        Exit Sub

GIS_OnExtentChange_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS_OnExtentChange " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS_OnKeyDown(translated As Boolean, Key As Integer, ByVal Shift As TatukGIS_DK.XShiftState)
        '<EhHeader>
        On Error GoTo GIS_OnKeyDown_Err
        '</EhHeader>
100   If Key = VK_CONTROL Then
102     If Not vkControl Then ' avoid multiple call on key repeat
104       GIS.Mode = XgisSelect
106       vkControl = True
        End If
      End If
  
108   If Key = VK_DELETE Then
110     If GIS.Mode = XgisEdit Then
112       GIS.Editor.DeleteShape
114       GIS.Mode = XgisSelect
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

Public Sub GetAdmCode(ptg As XGIS_Point, _
                      sAdm1 As String, _
                      sAdm2 As String, _
                      sAdmloc As String)
        '<EhHeader>
        On Error GoTo GetAdmCode_Err
        '</EhHeader>

        Dim oVecLyr As XGIS_LayerVector
        Dim shp As XGIS_Shape
        Dim oShp As XGIS_Shape
        Dim j As Integer
    
100     Set oVecLyr = GIS.get(m_sAdmVal1)
                        
102     If Not oVecLyr Is Nothing Then
        
104         Set oShp = oVecLyr.Locate(ptg, 5 / GIS.zoom, True)
                        
106         If Not oShp Is Nothing Then
        
108             SafeMoveFirst g_RSAppSettings
110             g_RSAppSettings.Find "SettingName = 'AdminLevel0'"
112             sAdm1 = oShp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
    
                'For j = 0 To oVecLyr.Fields.Count - 1
                '    m_frmDebug.DebugPrint  oVecLyr.Name & "_____" & oVecLyr.Fields.Item(j).Name & ":" & oShp.GetField(oVecLyr.Fields.Item(j).Name)
                'Next

            End If
        End If

114     Set oVecLyr = GIS.get(m_sAdmVal2)
                        
116     If Not oVecLyr Is Nothing Then
118         Set oShp = oVecLyr.Locate(ptg, 5 / GIS.zoom, True)
                        
120         If Not oShp Is Nothing Then
            
122             SafeMoveFirst g_RSAppSettings
124             g_RSAppSettings.Find "SettingName = 'AdminLevel1'"
126             sAdm2 = oShp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)

                'For j = 0 To oVecLyr.Fields.Count - 1
                '    m_frmDebug.DebugPrint  oVecLyr.Name & "_____" & oVecLyr.Fields.Item(j).Name & ":" & oShp.GetField(oVecLyr.Fields.Item(j).Name)
                'Next

            End If
        End If

        If m_bUseDistrictOnly Then
        Dim x As Double
        Dim y As Double
            'x = oShp.Centroid.x
            'y = oShp.Centroid.y
            
            If Not oShp Is Nothing Then
                Set ptg = oShp.Centroid
                sAdmloc = "N/A"
            End If
        Else

128         SafeMoveFirst g_RSAppSettings
130         g_RSAppSettings.Find "SettingName = 'AdminLocation'"
            
132     Set oVecLyr = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
                        
134     If Not oVecLyr Is Nothing Then
136         Set oShp = oVecLyr.Locate(ptg, 20 / GIS.zoom, True)
                        
138         If Not oShp Is Nothing Then

140             sAdmloc = oShp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
    
                '            For j = 0 To oVecLyr.Fields.Count - 1
                '                m_frmDebug.DebugPrint  oVecLyr.Name & "_____" & oVecLyr.Fields.Item(j).Name & ":" & oShp.GetField(oVecLyr.Fields.Item(j).Name)
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

Public Function GetAdminNameFromPoint(ptg As XGIS_Point, _
                                      sAdmLayer As String, _
                                      sAdmField As String, _
                                      dPrecision As Double) As String

    Dim oVecLyr As XGIS_LayerVector
    Dim shp As XGIS_Shape
    Dim oShp As XGIS_Shape
    Dim j As Integer
    
    Set oVecLyr = GIS.get(sAdmLayer)
    GetAdminNameFromPoint = ""

    If Not oVecLyr Is Nothing Then
        
        Set oShp = oVecLyr.Locate(ptg, dPrecision, False)

        If Not oShp Is Nothing Then GetAdminNameFromPoint = oShp.GetField(sAdmField)

    End If
    
    Set oVecLyr = Nothing
    Set oShp = Nothing

End Function

Private Sub GIS_OnMouseDown(translated As Boolean, _
                            ByVal Button As TatukGIS_DK.XMouseButton, _
                            ByVal Shift As TatukGIS_DK.XShiftState, _
                            ByVal x As Long, _
                            ByVal y As Long)
        '<EhHeader>
        On Error GoTo GIS_OnMouseDown_Err
        '</EhHeader>
100     translated = True

        Dim sAdm1 As String
        Dim sAdm2 As String
        Dim sAdm3 As String
        Dim sAdm4 As String
        Dim sAdm5 As String
        Dim sAdmloc As String
        Dim lUID As Long
        Dim i As Integer
        Dim j As Integer
        Dim fdesc As String
        Dim oVecLyr As XGIS_LayerVector
        Dim ptg As XGIS_Point
        Dim shp As XGIS_Shape
        Dim oShp As XGIS_Shape
        Dim bRunUrl As Boolean
  
102     Debug.Print "----- **MouseDown Start** ----"
104     Debug.Print "gdpts:" & UBound(gdpts)
106     Debug.Print "ptsSel:" & UBound(ptsSel)
  
108     If Button = XmbRight Then
            Dim pt As POINTAPI
110         GetCursorPos pt
112         PopupMenu mnuMAPPopUp, , ScaleX(pt.x - 2, vbPixels, vbTwips), ScaleY(pt.y - 8, vbPixels, vbTwips)
            Exit Sub
        End If
  
114     Set ptg = GIS.ScreenToMap(GisUtils.Point(x, y))

        'ProjTest ptg.X, ptg.y

116     Select Case g_CurrentTool

            Case oSingleSelect
            
118             Set shp = GIS.Locate(ptg, 5 / GIS.zoom) ' 5 pixels precision
            
120             If Not shp Is Nothing Then
122                 shp.MakeEditable
124                 shp.Flash
126                 m_frmAttributes.ShowShape shp
128                 shp.IsSelected = True
130                 m_frmAttributes.ShowSelected shp.layer
                End If
    
132         Case oInfo
            
                ' m_frmDebug.DebugPrint  m_frmMnuOperations.Legend1.GIS_Layer.Name
            
                '128             Select Case m_frmMnuOperations.C1TabOpsView.CurrTab
                '
                '                    Case 0
                '
                '130                     If Not m_frmMnuOperations.Legend1.GIS_Layer Is Nothing Then
                '132                         Set oVecLyr = GIS.get(m_frmMnuOperations.Legend1.GIS_Layer.Name)
                '                        End If
                '
                '134                 Case 1
                '
                '136                     If Not m_frmMnuOperations.LgdSecurity.GIS_Layer Is Nothing Then
                '138                         Set oVecLyr = GIS.get(m_frmMnuOperations.LgdSecurity.GIS_Layer.Name)
                '                        End If
                '
                '140                 Case 2
                '
                '142                     If Not m_frmMnuOperations.lgdTest.GIS_Layer Is Nothing Then
                '144                         Set oVecLyr = GIS.get(m_frmMnuOperations.lgdTest.GIS_Layer.Name)
                '                        End If
                '
                '                End Select
            
134             If m_frmAttributes Is Nothing Then
136                 Set m_frmAttributes = New frmAttributes
                End If
                
138             If Not m_frmAttributes.Visible Then
140                 m_frmAttributes.Init GIS.viewer
                End If
                
142             m_frmAttributes.Show vbModeless, Me
            
144             Set oVecLyr = GIS.get(m_frmAttributes.LyrCol.Item(m_frmAttributes.ComLayer.List(m_frmAttributes.ComLayer.ListIndex)))
                'GIS.get(m_frmAttributes.ComLAyer.List(m_frmAttributes.ComLAyer.ListIndex))
            
146             If Not oVecLyr Is Nothing Then
                 
148                 Set oShp = oVecLyr.Locate(ptg, 15 / GIS.zoom, True)
150                 m_frmAttributes.ShowShape oShp
                    
152                 m_frmAttributes.caption = "Active layer: " & oVecLyr.caption

154                 If Not oShp Is Nothing Then

156                     If Not g_RSGISGridTableSettings.EOF Or Not g_RSGISGridTableSettings.BOF Then
                        
158                         SafeMoveFirst g_RSGISGridTableSettings
160                         g_RSGISGridTableSettings.Find "name = '" & oVecLyr.Name & "'"
                    
162                         If Not g_RSGISGridTableSettings.EOF Then
164                             If g_RSGISGridTableSettings.Fields.Item("isURLLayer").Value Then
166                                 If g_RSGISGridTableSettings.Fields.Item("autoRunUrls").Value Then
168                                     bRunUrl = True
                                    End If
                                End If
                            End If
                   
                            'For j = 0 To oVecLyr.Fields.Count - 1
                            '    fdesc = oVecLyr.Fields.Item(j).Name  'GIS.Field(oVecLyr.Name, 1, oVecLyr.Fields.Item(i).Name)
                            '    'm_frmDebug.DebugPrint  oVecLyr.Name & "_____" & fdesc & ":" & oShp.GetField(fdesc)
                            'Next
    
170                         If bRunUrl Then
172                             If oShp.GetField(g_RSGISGridTableSettings.Fields.Item("URLLayerField").Value) <> "" Then
174                                 ShellExecute Me.hWnd, vbNullString, oShp.GetField(g_RSGISGridTableSettings.Fields.Item("URLLayerField").Value), vbNullString, "C:\", 1
                                End If
                            End If
                            
                        End If

                        If Shift = XssShift Then
                            frmSendMess.lvwAttributes.ListItems.Add , , "X"
                            frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oShp.Centroid.x
                            
                            frmSendMess.lvwAttributes.ListItems.Add , , "Y"
                            frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oShp.Centroid.y
                            
                            For j = 0 To oVecLyr.Fields.Count - 1
                                frmSendMess.lvwAttributes.ListItems.Add , , oVecLyr.Fields.Item(j).Name
                                frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oShp.GetField(oVecLyr.Fields.Item(j).Name)
                            Next
                            
                            frmSendMess.Show vbModal, Me
                        End If
                        
                        'New Stuff
                        'frmAttributes.ShowSelected oVecLyr
                        
                    End If
                    
                Else
                
                    'If m_frmAttributes Is Nothing Then
                    '    Set m_frmAttributes = New frmAttributes
                    'End If
                
                    'If Not m_frmAttributes.Visible Then
                    '    m_frmAttributes.Init GIS.viewer
                    'End If

                    'm_frmAttributes.Show vbModeless, Me
                    
176                 Set oShp = GIS.Locate(ptg, 5 / GIS.zoom) ' 5 pixels precision
            
178                 If Not oShp Is Nothing Then m_frmAttributes.caption = "Active layer: " & oShp.layer.caption
                    
180                 m_frmAttributes.ShowShape oShp

                    If Shift = 17 Then
                        frmSendMess.lvwAttributes.ListItems.Add , , "X"
                        frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oShp.Centroid.x
                            
                        frmSendMess.lvwAttributes.ListItems.Add , , "Y"
                        frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oShp.Centroid.y
                            
                        Set oVecLyr = oShp.layer
                            
                        For j = 0 To oVecLyr.Fields.Count - 1
                            frmSendMess.lvwAttributes.ListItems.Add , , oVecLyr.Fields.Item(j).Name
                            frmSendMess.lvwAttributes.ListItems.Item(frmSendMess.lvwAttributes.ListItems.Count).SubItems(1) = oShp.GetField(oVecLyr.Fields.Item(j).Name)
                        Next
                            
                        frmSendMess.Show vbModal, Me
                    End If

                    'End If
            
                    '174                 For i = 0 To GIS.Items.Count - 1
                    '
                    '176                     If GisUtils.IsInherited(GIS.Items.Item(i), "XGIS_LayerVector") Then
                    '178                         Set oVecLyr = GIS.get(GIS.Items.Item(i).Name)
                    '180                         Set oShp = oVecLyr.Locate(ptg, 10 / GIS.Zoom, True)
                    '
                    '182                         If Not oShp Is Nothing Then
                    '
                    '                                'For j = 0 To oVecLyr.Fields.Count - 1
                    '                                '    fdesc = oVecLyr.Fields.Item(j).Name  'GIS.Field(oVecLyr.Name, 1, oVecLyr.Fields.Item(i).Name)
                    '                                '    'm_frmDebug.DebugPrint  oVecLyr.Name & "_____" & fdesc & ":" & oShp.GetField(fdesc)
                    '                                'Next
                    '
                    '184                             frmAttributes.ShowShape oShp
                    '186                             frmAttributes.Show vbModeless, Me
                    '                                Exit For
                    '                                'New Stuff
                    '                                'frmAttributes.ShowSelected oVecLyr
                    '
                    '                            End If
                    '
                    '                        End If
                    '
                    '                    Next

                End If

                'frmAttributes.Show vbModeless, Me
            
182         Case OASIS_TOOLS.oCreateLocationArea, oCreateLocationPolyline, oCreateLocationPoint, oCreateLocationLine, oCreateLocationMultipoint, oCreateLocationPoint

184             Select Case g_CurrentFeatureType
                
                    Case OASISFeatureTypes.Custom

                        If m_frmSpatialiseDD.Visible Then
                                 
                            Set oShp = mDDLayer.CreateShape(XgisShapeTypePoint)
                            oShp.Lock XgisLockExtent
                            oShp.AddPart
                            oShp.AddPoint ptg
                            UpdateSpatialiseDDList oShp

                        End If

186                     If Not m_frmSpatialize Is Nothing Then
188                         GetAdmCode ptg, sAdm1, sAdm2, sAdmloc
190                         m_frmSpatialize.OASISLocator1.SetCoordinateValue CStr(ptg.x), "Long:", CStr(ptg.y), "Lat:", "Lat/Long WGS84", Map1.MilitaryGridReferenceFromPoint(ptg.x, ptg.y)
192                         m_frmSpatialize.OASISLocator1.SetAdmValues sAdm1, sAdm2, sAdmloc
                        End If
                        
                        If FileExists(g_sAppPath & "\data\db\dynamicdata\WorldVision.mdb") Then
                        
                            If m_frmWVIIncidents.Visible Then
                        
                                With m_frmWVIIncidents
                                
                                    If .txtLocationAdmin1.Visible Then sAdm1 = GetAdminNameFromPoint(ptg, .txtLocationAdmin1.Tag, .lblLocationAdmin1.Tag, 200 / GIS.zoom)
                                    If .txtLocationAdmin2.Visible Then sAdm2 = GetAdminNameFromPoint(ptg, .txtLocationAdmin2.Tag, .lblLocationAdmin2.Tag, 200 / GIS.zoom)
                                    If .txtLocationAdmin3.Visible Then sAdm3 = GetAdminNameFromPoint(ptg, .txtLocationAdmin3.Tag, .lblLocationAdmin3.Tag, 200 / GIS.zoom)
                                    If .txtLocationAdmin4.Visible Then sAdm4 = GetAdminNameFromPoint(ptg, .txtLocationAdmin4.Tag, .lblLocationAdmin4.Tag, 200 / GIS.zoom)
                                    If .txtLocationAdmin5.Visible Then sAdm5 = GetAdminNameFromPoint(ptg, .txtLocationAdmin5.Tag, .lblLocationAdmin5.Tag, 200 / GIS.zoom)
                                
                                End With
                                
                                'GetAdmCode ptg, sAdm1, sAdm2, sAdmloc
                                m_frmWVIIncidents.SetLocationParams sAdm1, sAdm2, sAdm3, sAdm4, sAdm5, CStr(ptg.x), CStr(ptg.y), Round(GetNearestShape(ptg, "WV Offices", False), 2) & " km", Round(GetNearestShape(ptg, "WV Warehouses", False), 2) & " km", Round(GetNearestShape(ptg, "WV Op Areas", False), 2) & " km"
                            
                                'Round(oShp.distance(PassedPoint, 1000000000) / 1000, 0)
                            
                            End If
                        
                        End If
                    
194                 Case OASISFeatureTypes.GeoMark
                    
                    Case 666
                    
                        Dim sComms As String
                        Dim sRelief As String
                        Dim sShelter As String
                        Dim sTransport As String
                        Dim sWash As String
                        Dim oShapeX As XGIS_Shape
                        Dim oShapeX1 As New XGIS_ShapePolygon
                        Dim oTopology As New XGIS_Topology
                        Dim lNearestShape As Long
                        Dim lNearestShapeCurrent As Long
                        
                        lNearestShape = GetNearestShape(ptg, "Resources (Comms Available)")
                        
                        lNearestShapeCurrent = GetNearestShape(ptg, "Resources (Shelter Available)")

                        If lNearestShapeCurrent < lNearestShape Then lNearestShape = lNearestShapeCurrent
                        
                        lNearestShapeCurrent = GetNearestShape(ptg, "Resources (Relief Available)")

                        If lNearestShapeCurrent < lNearestShape Then lNearestShape = lNearestShapeCurrent
                        
                        lNearestShapeCurrent = GetNearestShape(ptg, "Resources (Wash Available)")

                        If lNearestShapeCurrent < lNearestShape Then lNearestShape = lNearestShapeCurrent
                        
                        lNearestShapeCurrent = GetNearestShape(ptg, "Resources (Transport Available)")

                        If lNearestShapeCurrent < lNearestShape Then lNearestShape = lNearestShapeCurrent
                        
                        sComms = GetDistanceToAllShapes(ptg, "Resources (Comms Available)", "WarehouseName")
                        sShelter = GetDistanceToAllShapes(ptg, "Resources (Shelter Available)", "WarehouseName")
                        sRelief = GetDistanceToAllShapes(ptg, "Resources (Relief Available)", "WarehouseName")
                        sWash = GetDistanceToAllShapes(ptg, "Resources (Wash Available)", "WarehouseName")
                        sTransport = GetDistanceToAllShapes(ptg, "Resources (Transport Available)", "WarehouseName")
                        
                        m_frmResourcesFinder.SetResources sComms, sRelief, sShelter, sTransport, sWash
                        Set oShapeX = m_oDrawLyr.Shape

                        If Not oShapeX Is Nothing Then m_oDrawLyr.Revert oShapeX.uID
                        m_oDrawLyr.Transparency = 35
                        Set oShapeX = m_oDrawLyr.CreateShape(XgisShapeTypeUnknown)
                        
                        oShapeX.AddPart
                        oShapeX.AddPoint ptg

                        Set oShapeX1 = oTopology.MakeBuffer(oShapeX, lNearestShape, 9, True)
                        m_oDrawLyr.Revert oShapeX.uID
                        m_oDrawLyr.AddShape oShapeX1
                        GIS.UpDate
                        Set oTopology = Nothing
                        Set oShapeX = Nothing
                        Set oShapeX1 = Nothing
                        
196                 Case OASISFeatureTypes.Incident
198                     GetAdmCode ptg, sAdm1, sAdm2, sAdmloc
                        'MsgBox GetNearestShape(ptg, "Office Locations")
                    
200                     If Not m_oIncShpPt Is Nothing Then
202                         m_oSQLIncLyr.Delete m_oIncShpPt.uID
                        End If
                    
204                     Set m_oIncShpPt = m_oSQLIncLyr.CreateShape(XgisShapeTypePoint)
                        
206                     m_oIncShpPt.Lock XgisLockExtent
208                     m_oIncShpPt.AddPart

210                     m_oIncShpPt.AddPoint ptg
212                     m_oIncShpPt.SetField "ID", Rnd(300000)
214                     m_oIncShpPt.SetField "Name", m_fmrAddIncident.txtEnteredBy.Text
216                     m_oIncShpPt.SetField "Type", m_fmrAddIncident.ComIncType.List(m_fmrAddIncident.ComIncType.ListIndex)
218                     m_oIncShpPt.SetField "Target", m_fmrAddIncident.ComIncTarget.List(m_fmrAddIncident.ComIncTarget.ListIndex)
220                     m_oIncShpPt.SetField "Incident_DATE", m_fmrAddIncident.MVIncident.Value
222                     m_oIncShpPt.SetField "Town", sAdmloc
224                     m_oIncShpPt.SetField "District", sAdm2
226                     m_oIncShpPt.SetField "Province", sAdm1
228                     m_oIncShpPt.SetField "Description", m_fmrAddIncident.txtIncidentDescription.Text & vbCrLf & m_fmrAddIncident.txtLocationDescription.Text
                        
                        'm_fmrAddIncident.SetNearestValues GetNearestShape(ptg, "Office Locations"), GetNearestShape(ptg, "Operational Areas")
230                     m_fmrAddIncident.SetCoordinateValue CStr(ptg.x), "Long:", CStr(ptg.y), "Lat:", "Lat/Long WGS84", Map1.MilitaryGridReferenceFromPoint(ptg.x, ptg.y)
                       
232                     m_fmrAddIncident.SetFocus
                    
234                     m_oIncShpPt.Unlock

                        'lUID = m_oIncShpPt.Uid
                    
                        'Set m_oIncShpPt = m_oSQLIncLyr.AddShape(m_oIncShpPt)
                        
                        'm_oSQLIncLyr.Delete lUID
                    
236                     m_fmrAddIncident.SetAdmValues sAdm1, sAdm2, sAdmloc
                                        
238                     GIS.UpDate

240                 Case OASISFeatureTypes.Location
                
242                     If Button = XmbLeft Then

                            '                If GIS.Mode = XgisEdit Then
244                         If g_oEditLayer Is Nothing Then
                                Exit Sub
                            Else
                    
                                '    GIS.Editor.CreateShape g_oEditLayer, ptg, XgisShapeTypePolygon

246                             Select Case g_CurrentTool
 
                                    Case OASIS_TOOLS.oCreateLocationArea
                                    
248                                     GIS.Editor.CreateShape g_oEditLayer, ptg, XgisShapeTypePolygon
250                                     Set g_oEditLayer = Nothing
                                        
                                        '                                    'GIS.Editor.CreateShape g_oEditLayer, ptg, XgisShapeTypePolygon
                                        '
                                        '                                    If m_oShpPolygon Is Nothing Then
                                        '
                                        '                                        Set m_oShpPolygon = g_oEditLayer.CreateShape(XgisShapeTypePolygon)
                                        '                                        m_oShpPolygon.Lock XgisLockExtent
                                        '                                        m_oShpPolygon.AddPart
                                        '                                        m_oShpPolygon.AddPoint ptg
                                        '                                        m_oShpPolygon.Unlock
                                        '                                        Set m_oPtRubberStart = ptg
                                        '                                        m_bShowRubberBand = True
                                        '                                    Else
                                        '                                        m_oShpPolygon.Lock XgisLockExtent
                                        '                                        m_oShpPolygon.AddPoint ptg
                                        '                                        m_oShpPolygon.Unlock
                                        '                                        Set m_oPtRubberStart = ptg
                                        '                                        GIS.Update
                                        '                                    End If

252                                 Case OASIS_TOOLS.oCreateLocationPolyline

254                                     GIS.Editor.CreateShape g_oEditLayer, ptg, XgisShapeTypeArc
256                                     Set g_oEditLayer = Nothing

                                        '                                    If m_oShpArc Is Nothing Then
                                        '
                                        '                                        Set m_oShpArc = g_oEditLayer.CreateShape(XgisShapeTypeArc)
                                        '                                        m_oShpArc.Lock XgisLockExtent
                                        '                                        m_oShpArc.AddPart
                                        '                                        m_oShpArc.AddPoint ptg
                                        '                                        m_oShpArc.Unlock
                                        '
                                        '                                    Else
                                        '                                        m_oShpArc.Lock XgisLockExtent
                                        '                                        m_oShpArc.AddPoint ptg
                                        '                                        m_oShpArc.Unlock
                                        '                                        GIS.Update
                                        '                                    End If
                            
                                        '                                GIS.Editor.CreateShape g_oEditLayer, ptg, XgisShapeTypeArc
                                        '                                If m_oShpArc Is Nothing Then
                                        '                                    Set m_oShpArc = GIS.Editor.CreateShape(g_oEditLayer, ptg, XgisShapeTypeArc)
                                        '                                Else
                                        '                                    GIS.Editor.AddPoint ptg
                                        '                                End If
                        
258                                 Case OASIS_TOOLS.oCreateLocationPoint
260                                     GIS.Editor.CreateShape g_oEditLayer, ptg, XgisShapeTypePoint
262                                     Set g_oEditLayer = Nothing

                                        'Set m_oShpPoint = GIS.Editor.CreateShape(g_oEditLayer, ptg, XgisShapeTypePoint)
264                                 Case oCreateLocationMultipoint
266                                     GIS.Editor.CreateShape g_oEditLayer, ptg, XgisShapeTypeMultiPoint
268                                     Set g_oEditLayer = Nothing
                                End Select
                    
                            End If

                            '                End If

                        Else

270                         Select Case g_CurrentTool
            
                                Case OASIS_TOOLS.oCreateLocationArea, oCreateLocationPolyline, oCreateLocationPoint, oCreateLocationLine, oCreateLocationMultipoint, oCreateLocationPoint

272                                 endEdit
                                    'Set m_oShpPolygon = m_oShpPolygon.AsPolygon
                                    'frmRegionStyle.Init m_oShpPolygon
                                    'frmRegionStyle.Show vbModal, Me
                                    'cmdCommand2_Click
274                                 GIS.UpDate

                                    ' Case OASIS_TOOLS.oCreateLocationPolyline
                                    '     Set m_oShpArc = m_oShpArc.AsArc
                            End Select

                        End If

                End Select

276         Case OASIS_TOOLS.oRadiusSelect, OASIS_TOOLS.oCircleSelect
            
278             Set oldPos = GisUtils.Point(x, y)
280             oldRadius = 0

282         Case oLineSelect, oPolyLineSelect, oPolySelect, oRectSelect, oMultiSelect, oAreaSelect
            
284             If g_CurrentTool = oLineSelect Then m_bPolyL = True
                If g_CurrentTool = oAreaSelect Then m_bPolyG = True
                
                '        If m_bPolyG Or m_bPolyL Then
                
286             GIS.PaintMinimum
                
                ' Private gdpts() As POINTAPI
            
288             Set ptsSel(UBound(ptsSel)) = GIS.ScreenToMap(GisUtils.Point(x, y))
290             gdpts(UBound(gdpts)).x = x: gdpts(UBound(gdpts)).y = y
292             GDI_Polyline GIS.hdc, gdpts
                
294             If Not m_bDrawFinished Then
296                 ReDim Preserve gdpts(UBound(gdpts) + 1)
298                 ReDim Preserve ptsSel(UBound(ptsSel) + 1)
                Else
300                 m_bDrawFinished = False
                End If
            
302             Set oldPos = GisUtils.Point(x, y)

304         Case OASIS_TOOLS.oFeatureSelect
            
306             Set oVecLyr = GIS.get(m_SelLyrCol.Item(ComFeatureLayer.List(ComFeatureLayer.ListIndex)))
            
308             If Not oVecLyr Is Nothing Then
                 
310                 Set oShp = oVecLyr.Locate(ptg, 15 / GIS.zoom, True)

312                 If Not oShp Is Nothing Then

314                     doFeatureSelect oShp
                    End If
                    
                End If
            
                '306             Set shp = GIS.Locate(ptg, 5 / GIS.Zoom) ' 5 pixels precision
                '
                '308             If Not shp Is Nothing Then
                '310                 shp.MakeEditable
                '312                 shp.Flash
                '314                 m_frmAttributes.ShowShape shp
                '316                 shp.IsSelected = True
                '318                 m_frmAttributes.ShowSelected shp.layer
                '                End If
            
316         Case OASIS_TOOLS.oCreateLocationText
318             AddAnnoText ptg
        End Select

320     Debug.Print "----- **MouseDown End** ----"
322     Debug.Print "gdpts:" & UBound(gdpts)
324     Debug.Print "ptsSel:" & UBound(ptsSel)
        
        '<EhFooter>
        Exit Sub

GIS_OnMouseDown_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GIS_OnMouseDown " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function doFeatureSelect(oShp As XGIS_Shape) As Long
        '<EhHeader>
        On Error GoTo doFeatureSelect_Err
        '</EhHeader>
        Dim i As Integer
        Dim ptg As XGIS_Point
       ' Dim oShp As XGIS_Shape
    
100     RemoveAlltabs
    
102     With m_oDrawLyr
104         .Lock
        
106         doFeatureSelect = oShp.uID

108         .draw
    
            Dim lL As XGIS_LayerVector
            Dim tmp As XGIS_Shape
        
            Dim buf1 As New XGIS_Shape 'TatukGIS_DK.IXGIS_Shape
            Dim buf2 As New XGIS_Shape 'TatukGIS_DK.IXGIS_Shape
        
            Dim sVals As String
            Dim tpl As New XGIS_Topology

110         m_oBufferLyr.RevertAll ' GIS.get("Buffers").RevertAll

           ' Set tpl = New XGIS_Topology
    
            ' 1 degree in km 111.2
112         oShp.MakeEditable
666         buf2.MakeEditable
667         Set buf2 = m_oBufferLyr.AddShape(tpl.MakeBuffer(oShp, m_udtSelectorSettings.dBuffeLevel / 111.2, 36, True))
113         'Set buf1 = New XGIS_ShapeAbstract
114         'buf1.MakeEditable

116        ' Set buf1 = tpl.MakeBuffer(oShp, m_udtSelectorSettings.dBuffeLevel / 111.2, 36, True)
118        ' Set buf2 = GIS.get("Buffers").AddShape(buf1)
        
120         m_lBufUID = buf2.uID
        
122         Set buf1 = Nothing
124         Set tpl = Nothing
    
126         Set lL = GIS.get(m_SelLyrCol.Item(ComSelLayer.List(ComSelLayer.ListIndex)))

128         If lL Is Nothing Then
130             .Unlock
132             GIS.UpDate
                Exit Function
            End If
    
134         lL.DeselectAll
    
            ' check all shapes
136         Set tmp = lL.FindFirst(buf2.Extent, "", buf2, m_udtSelectorSettings.sSpatialOperator, True)
    
138         While Not tmp Is Nothing
140             Set tmp = tmp.MakeEditable
                'Debug.Print tmp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
142             tmp.IsSelected = True

    '140             If m_frmSelections.Visible Then
144                 AddSelection tmp
    '                End If
            
146             Set tmp = lL.FindNext
            Wend
148         .Unlock
150         GIS.PaintMinimum

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
        Dim ptg As XGIS_Point
        Dim oShp As XGIS_Shape
    
100     GIS.Mode = XgisEdit
    
102     GIS.Editor.CreateShape m_oDrawLyr, ptsSel(0), XgisShapeTypePolygon
    
104     For i = LBound(ptsSel) + 1 To UBound(ptsSel)
106         GIS.Editor.AddPoint ptsSel(i)
        Next

108     CreatePolygon = GIS.Editor.uID

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

Private Sub UpdateSpatialiseDDList(sShp As XGIS_Shape)

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
    sShp.draw
    sShp.Flash
    m_oSpatialiseLayer.SaveAll
    m_frmSpatialiseDD.SetListValues sPoints
    endEdit
    
End Sub

Private Sub DynamicDataModule1_GetCurrentExtentCentroid(oPoint As XGIS_Point)

    oPoint.Prepare GIS.VisibleExtent.XMin + ((GIS.VisibleExtent.XMax - GIS.VisibleExtent.XMin) / 2), GIS.VisibleExtent.YMin + ((GIS.VisibleExtent.YMax - GIS.VisibleExtent.YMin) / 2)

End Sub

Private Function CreatePolygonEX() As Long
        '<EhHeader>
        On Error GoTo CreatePolygonEX_Err
        '</EhHeader>
        Dim i As Integer
        Dim ptg As XGIS_Point
        Dim oShp As XGIS_Shape
        
        Dim ipoints As Integer
        Dim iparts As Integer
        
        Dim oLayer As XGIS_LayerAbstract
        
        If Not m_frmSpatialiseDD.Visible Then
            Set oLayer = m_oDrawLyr
        Else

            'If Not m_oSpatialiseLayer Is Nothing Then m_oSpatialiseLayer.RevertAll
            Set oLayer = mDDLayer ' m_oSpatialiseLayer
        End If
    
100     With oLayer
102         .Lock
    
            If m_frmSpatialiseDD.Visible Then
                Set oShp = mDDLayer.CreateShape(mDDLayer.GetShape(mDDLayer.GetLastUid).ShapeType)
            Else
104             Set oShp = .CreateShape(XgisShapeTypePolygon)
            End If
    
106         oShp.Lock XgisLockExtent
108         oShp.AddPart
        
110         For i = LBound(ptsSel) To UBound(ptsSel)
112             oShp.AddPoint ptsSel(i)
            Next

            oShp.Unlock
            
            If m_frmSpatialiseDD.Visible Then
            
                endEdit
                UpdateSpatialiseDDList oShp
                
            End If

116         CreatePolygonEX = oShp.uID

118         .draw

            If Not m_frmSpatialiseDD.Visible Then
    
                Dim lL As XGIS_LayerVector
                Dim tmp As XGIS_Shape
        
            Dim buf1 As TatukGIS_DK.IXGIS_Shape
            Dim buf2 As TatukGIS_DK.IXGIS_Shape
        
            Dim sVals As String
            Dim tpl As XGIS_Topology

120         GIS.get("Buffers").RevertAll

122         Set tpl = New XGIS_Topology
    
            ' 1 degree in km 111.2
    
124         Set buf1 = tpl.MakeBuffer(oShp, m_udtSelectorSettings.dBuffeLevel / 111.2, 36, True)
126         Set buf2 = GIS.get("Buffers").AddShape(buf1)
        
128         m_lBufUID = buf2.uID
        
130         Set buf1 = Nothing
132         Set tpl = Nothing
    
134         Set lL = GIS.get(m_SelLyrCol.Item(ComSelLayer.List(ComSelLayer.ListIndex)))

136         If lL Is Nothing Then
138             .Unlock
140             GIS.UpDate
                Exit Function
            End If
    
142         lL.DeselectAll
    
            ' check all shapes
144         Set tmp = lL.FindFirst(buf2.Extent, "", buf2, m_udtSelectorSettings.sSpatialOperator, True)
    
146         While Not tmp Is Nothing
148             Set tmp = tmp.MakeEditable
                'Debug.Print tmp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
150             tmp.IsSelected = True

'152             If m_frmSelections.Visible Then
154                 AddSelection tmp
'                End If
            
156             Set tmp = lL.FindNext
            Wend
158         .Unlock
            .RevertAll
160         GIS.PaintMinimum

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
        Dim ptg As XGIS_Point
        Dim oShp As XGIS_Shape
        
        Dim oLayer As XGIS_LayerAbstract
        
        If Not m_frmSpatialiseDD.Visible Then
            Set oLayer = m_oDrawLyr
        Else
            Set oLayer = mDDLayer
        End If
        
100     With oLayer
102         .Lock
    
            If m_frmSpatialiseDD.Visible Then
                Set oShp = mDDLayer.GetShape(mDDLayer.GetLastUid)   ' mDDShape 'm_frmSpatialiseDD.GetShape
            Else
104             Set oShp = .CreateShape(XgisShapeTypeArc)
            End If
    
106         oShp.Lock XgisLockExtent
108         oShp.AddPart
        
110         For i = LBound(ptsSel) To UBound(ptsSel)
112             oShp.AddPoint ptsSel(i)
            Next
            
            oShp.Unlock
        
            If m_frmSpatialiseDD.Visible Then
            
                endEdit
                UpdateSpatialiseDDList oShp
                
            End If
            
116         CreatePolyLine = oShp.uID

118         .draw

            If Not m_frmSpatialiseDD.Visible Then
    
                Dim lL As XGIS_LayerVector
                Dim tmp As XGIS_Shape
        
            Dim buf1 As TatukGIS_DK.IXGIS_Shape
            Dim buf2 As TatukGIS_DK.IXGIS_Shape
        
            Dim sVals As String
            Dim tpl As XGIS_Topology

120         GIS.get("Buffers").RevertAll

122         Set tpl = New XGIS_Topology
    
            ' 1 degree in km 111.2
    
124         Set buf1 = tpl.MakeBuffer(oShp, m_udtSelectorSettings.dBuffeLevel / 111.2, 36, True)
126         Set buf2 = GIS.get("Buffers").AddShape(buf1)
        
128         m_lBufUID = buf2.uID
        
130         Set buf1 = Nothing
132         Set tpl = Nothing
    
134         Set lL = GIS.get(m_SelLyrCol.Item(ComSelLayer.List(ComSelLayer.ListIndex)))

136         If lL Is Nothing Then
138             .Unlock
140             GIS.UpDate
                Exit Function
            End If
    
142         lL.DeselectAll
    
            ' check all shapes
144         Set tmp = lL.FindFirst(buf2.Extent, "", buf2, m_udtSelectorSettings.sSpatialOperator, True)
    
146         While Not tmp Is Nothing
148             Set tmp = tmp.MakeEditable
                'Debug.Print tmp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
150             tmp.IsSelected = True

'152             If m_frmSelections.Visible Then
154                 AddSelection tmp
'                End If
            
156             Set tmp = lL.FindNext
            Wend
158         .Unlock
            .RevertAll
160         GIS.PaintMinimum

            End If

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
    Dim oLyr As New XGIS_LayerVector
    Dim oLine As XGIS_Shape
    Dim oPol As XGIS_Shape
    Dim oPt As XGIS_Shape
    Dim pt As XGIS_Point
        
    oLyr.Open
    
    oLyr.Params.area.Pattern = GIS.SelectionPattern
    oLyr.Params.Transparancy = GIS.SelectionTransparency
    oLyr.Params.area.Pattern = GIS.SelectionColor
    oLyr.Params.area.Pattern = GIS.SelectionOutlineOnly
    oLyr.Params.area.Pattern = GIS.SelectionWidth
    
    '    GIS.SelectionColor = vbRed
'    GIS.SelectionPattern = XgisLockExtent
'    GIS.SelectionTransparency = 50
    
    Set oLine = oLyr.CreateShape(XgisShapeTypeArc)
        
    Dim ptg As XGIS_Point
    Dim oShp As XGIS_Shape
    
    oLyr.Lock
    
    oLine.Lock XgisLockExtent
    oLine.AddPart
        
    For i = 0 To 2
        Set pt = New XGIS_Point
        pt.x = i
        pt.y = i
        oLine.AddPoint pt
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

Private Sub GIS_OnMouseMove(translated As Boolean, _
                            ByVal Shift As TatukGIS_DK.XShiftState, _
                            ByVal x As Long, _
                            ByVal y As Long)
        '<EhHeader>
        On Error GoTo GIS_OnMouseMove_Err
        '</EhHeader>
        Dim ptg As XGIS_Point
        Dim shp As XGIS_Shape
        Dim i As Integer
        Dim sMGRS As String
            
        Dim j As Integer
        Dim oVecLyr As XGIS_LayerVector
        Dim oShp As XGIS_Shape
        Dim bRunUrl As Boolean
  
        'Set ptg = GIS.ScreenToMap(GisUtils.Point(X, Y))
        '  StatusBar.SimpleText = "x: " + Str(ptg.X) + "   y: " + Str(ptg.Y)
  
        If m_bScrib Then
            GIS.PaintMinimum
            '        pts(UBound(pts)).X = ScaleX(X, vbPixels, vbTwips): pts(UBound(pts)).Y = ScaleY(Y, vbPixels, vbTwips)
            gdpts(UBound(gdpts)).x = x: gdpts(UBound(gdpts)).y = y
        
            GDI_Polyline GIS.hdc, gdpts
            ReDim Preserve gdpts(UBound(gdpts) + 1)
        
        End If

        '        If m_bPolyL Then
        '            GIS.PaintMinimum
        '            'Me.Cls
        '            '        pts(UBound(pts)).X = ScaleX(X, vbTwips, vbPixels): pts(UBound(pts)).Y = ScaleY(Y, vbTwips, vbPixels)
        '            '        pts(UBound(pts)).X = ScaleX(X, vbPixels, vbTwips): pts(UBound(pts)).Y = ScaleY(Y, vbPixels, vbTwips)
        '            gdpts(UBound(gdpts)).x = x: gdpts(UBound(gdpts)).y = y
        '
        '            GDI_Polyline GIS.hdc, gdpts
        '            'ReDim Preserve pts(UBound(pts) + 1)
        '
        '        End If

        If m_bPolyG Or m_bPolyL Then
            GIS.PaintMinimum
            Set ptsSel(UBound(ptsSel)) = GIS.ScreenToMap(GisUtils.Point(x, y))
            gdpts(UBound(gdpts)).x = x: gdpts(UBound(gdpts)).y = y
            GDI_Polyline GIS.hdc, gdpts
        End If

        If m_bRECT Then
            GIS.PaintMinimum

            If UBound(gdpts) > 0 Then
                gdpts(1).x = x: gdpts(1).y = y
                GDI_Rectangle GIS.hdc, gdpts(0).x, gdpts(0).y, gdpts(1).y, gdpts(1).y
            End If
        
        End If
  
100     Set ptg = GIS.ScreenToMap(GisUtils.Point(x, y))

        If oMapTipSetting.Enabled And Not oMapTipSetting.MapTipLayer = "" Then
            If oMapTipSetting.MapTipLayer = "--All--" Then
                Set m_oToolTipSHP = GIS.Locate(ptg, 5 / GIS.zoom) ' 5 pixels precision
            Else
                Set m_oToolTipSHP = GIS.get(m_LyrCol.Item(oMapTipSetting.MapTipLayer)).Locate(ptg, 5 / GIS.zoom)
            End If
        End If
        
        If g_CurrentTool = oInfo Then
        
134         If Not m_frmAttributes Is Nothing Then
                
138             If m_frmAttributes.Visible Then
                    
                    If m_frmAttributes.chkDynamicInfo = vbChecked Then
            
144                     Set oVecLyr = GIS.get(m_frmAttributes.LyrCol.Item(m_frmAttributes.ComLayer.List(m_frmAttributes.ComLayer.ListIndex)))
            
146                     If Not oVecLyr Is Nothing Then
                 
148                         Set oShp = oVecLyr.Locate(ptg, 15 / GIS.zoom, True)
150                         m_frmAttributes.ShowShape oShp
                    
152                         m_frmAttributes.caption = "Active layer: " & oVecLyr.caption

154                         If Not oShp Is Nothing Then

156                             If Not g_RSGISGridTableSettings.EOF Or Not g_RSGISGridTableSettings.BOF Then
                        
158                                 SafeMoveFirst g_RSGISGridTableSettings
160                                 g_RSGISGridTableSettings.Find "name = '" & oVecLyr.Name & "'"
                    
162                                 If Not g_RSGISGridTableSettings.EOF Then
164                                     If g_RSGISGridTableSettings.Fields.Item("isURLLayer").Value Then
166                                         If g_RSGISGridTableSettings.Fields.Item("autoRunUrls").Value Then
168                                             bRunUrl = True
                                            End If
                                        End If
                                    End If
    
170                                 If bRunUrl Then
172                                     If oShp.GetField(g_RSGISGridTableSettings.Fields.Item("URLLayerField").Value) <> "" Then
174                                         ShellExecute Me.hWnd, vbNullString, oShp.GetField(g_RSGISGridTableSettings.Fields.Item("URLLayerField").Value), vbNullString, "C:\", 1
                                        End If
                                    End If
                                End If
                        
                            End If
                    
                        Else

176                         Set oShp = GIS.Locate(ptg, 5 / GIS.zoom) ' 5 pixels precision
            
178                         If Not oShp Is Nothing Then m_frmAttributes.caption = "Active layer: " & oShp.layer.caption
                    
180                         m_frmAttributes.ShowShape oShp

                        End If
                    End If
                    
                End If
            End If
        
        End If
        
102     sMGRS = " MGRS:" & Map1.MilitaryGridReferenceFromPoint(ptg.x, ptg.y) & " Zoom:" & Round(GIS.zoom, 2) & " Scale:" & GIS.ScaleAsText
        
104     AB.Bands.Item("bStatus").Tools.Item("lblCoords").Text = "X:" & Round(ptg.x, 8) & " Y:" & Round(ptg.y, 8) & sMGRS
106     AB.Bands.Item("bStatus").Tools.Item("lblCoords").caption = "X:" & Round(ptg.x, 8) & " Y:" & Round(ptg.y, 8) & sMGRS
        
108     If m_oToolTipSHP Is Nothing Then
            'Me.caption = "X:" & ptg.X & " Y:" & ptg.Y
            
            '  For i = 0 To ProjectionList.Count - 1
            '      m_frmDebug.DebugPrint  ProjectionList
            '  Next
            
            'GisUtils.GisLatitudeToStr ptg.x
            'Dim o As XGIS_ProjectionList
            'Dim io As XGIS_ProjectionAbstract
            'Set io = o.FindEx("MGRS")
            ptTip.x = 0
            ptTip.y = 0
            picToolTip.Visible = False
        Else
            GetCursorPos ptTip
            'StatusBar.SimpleText = m_oToolTipSHP.GetField("name")
        End If

110     If g_CurrentTool = oRadiusSelect Or g_CurrentTool = oCircleSelect Then
112         If Not (Shift = XssLeft) Then Exit Sub

114         SetROP2 GIS.hdc, R2_XORPEN

116         If oldRadius <> 0 Then
118             Ellipse GIS.hdc, oldPos.x - oldRadius, oldPos.y - oldRadius, oldPos.x + oldRadius, oldPos.y + oldRadius
            End If

120         oldRadius = Round(Sqr(((oldPos.x - x) * (oldPos.x - x)) + ((oldPos.y - y) * (oldPos.y - y))))

122         Ellipse GIS.hdc, oldPos.x - oldRadius, oldPos.y - oldRadius, oldPos.x + oldRadius, oldPos.y + oldRadius

        End If

        If g_CurrentTool = oLineSelect Then
        
        End If
        
        If g_CurrentTool = oPolyLineSelect Then
        
        End If
        
        If g_CurrentTool = oPolySelect Then
        
        End If
        
        If g_CurrentTool = oRectSelect Then
        
        End If

        If g_CurrentTool = oSingleSelect Then
        
        End If
        
        If g_CurrentTool = oMultiSelect Then
        
        End If
        
        '
        '        If m_bShowRubberBand Then
        '            Dim part_no As Integer
        '
        '            Set ptg = m_oLineRubber.GetLastPoint
        '
        '            Set shp_in = m_oCosmeticLayer.Shape
        '            shp_in.Lock XgisLockExtent
        '            Set shp_out = m_oCosmeticLayer.CreateShape(shp_in.ShapeType)
        '            shp_out.Lock XgisLockProjection
        '
        '            For part_no = 0 To shp_in.GetNumParts - 1
        '                shp_out.AddPart
        '
        '                For point_no = 0 To shp_in.GetPartSize(part_no) - 1
        '                    shp_out.AddPoint shp_in.GetPoint(part_no, point_no)
        '                Next point_no
        '            Next part_no
        '
        '            shp_out.Unlock
        '            shp_in.Unlock
        '        End If
    
        On Error Resume Next
124     SSC.Run "OASISGis_MouseMove", sMGRS, ptg.x, ptg.y
    
126     translated = True
        '<EhFooter>
        Exit Sub

GIS_OnMouseMove_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GIS_OnMouseMove " & "at line " & Erl
       
        '</EhFooter>
End Sub

Private Sub DumpAllGISProperties()
        Dim oLayer As XGIS_LayerAbstract
        Dim i As Integer

        On Error Resume Next

        With GIS
        
            m_frmDebug.DebugPrint "gis properties"
            m_frmDebug.DebugPrint .Align
            m_frmDebug.DebugPrint .AutoCenter
        
            With .BigExtent
                m_frmDebug.DebugPrint "bIG eXTENT"
                m_frmDebug.DebugPrint .XMax
                m_frmDebug.DebugPrint .XMin
                m_frmDebug.DebugPrint .YMax
                m_frmDebug.DebugPrint .YMin
        
            End With

            m_frmDebug.DebugPrint .BigExtentMargin
            m_frmDebug.DebugPrint .BorderStyle
            m_frmDebug.DebugPrint .BusyText
            m_frmDebug.DebugPrint .CachedPaint
            m_frmDebug.DebugPrint .CausesValidation
        
            With .Center
                m_frmDebug.DebugPrint "Center"
                m_frmDebug.DebugPrint .x
                m_frmDebug.DebugPrint .y
        
            End With
        
            With .CenterPtg
                m_frmDebug.DebugPrint .x
                m_frmDebug.DebugPrint .y
        
            End With
        
            m_frmDebug.DebugPrint "Charset"
            m_frmDebug.DebugPrint .Charset
            m_frmDebug.DebugPrint .ChartLabelCache
            m_frmDebug.DebugPrint .CodePage
            m_frmDebug.DebugPrint .Color
            m_frmDebug.DebugPrint .Ctl3D
            m_frmDebug.DebugPrint .Cursor
            
            m_frmDebug.DebugPrint "DrawMode"
            m_frmDebug.DebugPrint .DrawMode
            m_frmDebug.DebugPrint .FullPaint
            m_frmDebug.DebugPrint .IncrementalPaint
            m_frmDebug.DebugPrint .InDirectPaint
            m_frmDebug.DebugPrint .MinZoomSize
            m_frmDebug.DebugPrint .OutCodePage
            m_frmDebug.DebugPrint .ProjectFile
        
        End With
        
        m_frmDebug.DebugPrint "Starting With ALyers"
        
        For Each oLayer In GIS.Items
            
100         With oLayer

                m_frmDebug.DebugPrint "|||||||||||||||||||||||||||"
                m_frmDebug.DebugPrint .Name
                m_frmDebug.DebugPrint .Active
                m_frmDebug.DebugPrint .Age
                m_frmDebug.DebugPrint .CachedPaint
                m_frmDebug.DebugPrint .caption
                m_frmDebug.DebugPrint .Charset
                m_frmDebug.DebugPrint .CodePage
                m_frmDebug.DebugPrint .Collapsed
                m_frmDebug.DebugPrint .Comments
                
                m_frmDebug.DebugPrint "CONFIG NAME"
                m_frmDebug.DebugPrint .ConfigFile
                m_frmDebug.DebugPrint .ConfigName
                m_frmDebug.DebugPrint .DirectMode
                m_frmDebug.DebugPrint .DormantMode
                m_frmDebug.DebugPrint .FileInfo
                m_frmDebug.DebugPrint .HideFromLegend
                m_frmDebug.DebugPrint .IncrementalPaint
                m_frmDebug.DebugPrint .IsLocked
                m_frmDebug.DebugPrint .IsOpened
                
                m_frmDebug.DebugPrint "IS OPEN"
                m_frmDebug.DebugPrint .LabelsOnTop
                m_frmDebug.DebugPrint .OutCodePage
                m_frmDebug.DebugPrint .Path
                m_frmDebug.DebugPrint .StoreParamsInProject
                m_frmDebug.DebugPrint .SubType
                m_frmDebug.DebugPrint .Transparency
                m_frmDebug.DebugPrint .UseConfig
                m_frmDebug.DebugPrint .UseFileParams
                m_frmDebug.DebugPrint .ZOrder
                m_frmDebug.DebugPrint .ZOrderEx
                
                
                
102             With .Params
                    m_frmDebug.DebugPrint "Strating With Params"
                    m_frmDebug.DebugPrint .IsAssigned
                    m_frmDebug.DebugPrint .MaxScale
                    m_frmDebug.DebugPrint .MaxZoom
                    m_frmDebug.DebugPrint .MinScale
                    m_frmDebug.DebugPrint .MinZoom
                    
                    m_frmDebug.DebugPrint .Serial
                    m_frmDebug.DebugPrint .Style
                    m_frmDebug.DebugPrint .Visible

104                 With .area
                        m_frmDebug.DebugPrint "AREA PARAMS"
106                     m_frmDebug.DebugPrint .Color
                        m_frmDebug.DebugPrint .Symbol.FontName
                        m_frmDebug.DebugPrint .SymbolSize
                        m_frmDebug.DebugPrint .SymbolGap
                        m_frmDebug.DebugPrint .SymbolRotate
                        m_frmDebug.DebugPrint .BITMAP
                        m_frmDebug.DebugPrint .Pattern
                        m_frmDebug.DebugPrint .OutlineColor
                        m_frmDebug.DebugPrint .OutlineWidth
                        m_frmDebug.DebugPrint .OutlineStyle
                        m_frmDebug.DebugPrint .OutlineSymbol
                        m_frmDebug.DebugPrint .OutlineSymbolGap
                        m_frmDebug.DebugPrint .OutlineSymbolRotate
                        m_frmDebug.DebugPrint .OutlineBitmap
                        m_frmDebug.DebugPrint .OutlinePattern
                        m_frmDebug.DebugPrint .SmartSize
                        m_frmDebug.DebugPrint .SmartSizeField
            
                    End With
           
110                 With .Marker
                        m_frmDebug.DebugPrint "MARKER PARAMS"
112                     m_frmDebug.DebugPrint .Color
                        m_frmDebug.DebugPrint .Marker.Symbol
                        m_frmDebug.DebugPrint .SymbolSize
                        m_frmDebug.DebugPrint .SymbolGap
                        m_frmDebug.DebugPrint .SymbolRotate
                        m_frmDebug.DebugPrint .BITMAP
                        m_frmDebug.DebugPrint .Pattern
                        m_frmDebug.DebugPrint .OutlineColor
                        m_frmDebug.DebugPrint .OutlineWidth
                        m_frmDebug.DebugPrint .OutlineStyle
                        m_frmDebug.DebugPrint .OutlineSymbol
                        m_frmDebug.DebugPrint .OutlineSymbolGap
                        m_frmDebug.DebugPrint .OutlineSymbolRotate
                        m_frmDebug.DebugPrint .OutlineBitmap
                        m_frmDebug.DebugPrint .OutlinePattern
                        m_frmDebug.DebugPrint .SmartSize
                        m_frmDebug.DebugPrint .SmartSizeField
                    End With
                    
114                 With .Line
                        m_frmDebug.DebugPrint "LINE:nnnn"
116                     m_frmDebug.DebugPrint .Color
                        m_frmDebug.DebugPrint .Marker.Symbol
                
                        m_frmDebug.DebugPrint .SymbolSize
                        m_frmDebug.DebugPrint .SymbolGap
                        m_frmDebug.DebugPrint .SymbolRotate
                        m_frmDebug.DebugPrint .BITMAP
                        m_frmDebug.DebugPrint .Pattern
                        m_frmDebug.DebugPrint .OutlineColor
                        'm_frmDebug.DebugPrint  .OutlineWidth
                        'm_frmDebug.DebugPrint  .OutlineStyle
                        'm_frmDebug.DebugPrint  .OutlineSymbol
                        'm_frmDebug.DebugPrint  .OutlineSymbolGap
                        'm_frmDebug.DebugPrint  .OutlineSymbolRotate
                        'm_frmDebug.DebugPrint  .OutlineBitmap
                        'm_frmDebug.DebugPrint  .OutlinePattern
                        'm_frmDebug.DebugPrint  .SmartSize
                        'm_frmDebug.DebugPrint  .SmartSizeField
                    End With
           
                End With

            End With

        Next

End Sub

Private Sub GIS_OnMouseUp(translated As Boolean, _
                          ByVal Button As TatukGIS_DK.XMouseButton, _
                          ByVal Shift As TatukGIS_DK.XShiftState, _
                          ByVal x As Long, _
                          ByVal y As Long)
        '<EhHeader>
        On Error GoTo GIS_OnMouseUp_Err
        '</EhHeader>
                          
        Dim tpl As TatukGIS_DK.IXGIS_Topology
        Dim lL As TatukGIS_DK.IXGIS_LayerVector
        Dim tmp As TatukGIS_DK.IXGIS_Shape
        Dim buf1 As TatukGIS_DK.IXGIS_Shape
        Dim buf2 As TatukGIS_DK.IXGIS_Shape
        Dim ptg As TatukGIS_DK.IXGIS_Point
        Dim ptg1 As TatukGIS_DK.IXGIS_Point
        Dim distance As Double
        Dim sVals As String
        Dim colPlaces As New Collection
        
100     Debug.Print "----- **MouseUp Start** ----"
102     Debug.Print "gdpts:" & UBound(gdpts)
104     Debug.Print "ptsSel:" & UBound(ptsSel)
        
106     If g_CurrentTool = oRadiusSelect Or g_CurrentTool = oCircleSelect Then
108         If oldRadius = 0 Then
                Exit Sub
            End If
    
110         Set ptg = GIS.ScreenToMap(oldPos)
112         m_oDrawLyr.Lock
114         Set tmp = m_oDrawLyr.CreateShape(XgisShapeTypePoint)
116         tmp.Params.Marker.Size = 0
118         tmp.Lock XgisLockExtent
120         tmp.AddPart
122         tmp.AddPoint ptg
124         tmp.Unlock
126         GIS.get("Buffers").RevertAll
    
128         m_oDrawLyr.Unlock
130         Set tpl = New XGIS_Topology
            'distance recalc
132         Set ptg1 = GIS.ScreenToMap(GisUtils.Point(oldPos.x + oldRadius, y))
134         distance = ptg1.x - ptg.x
    
136         Set buf1 = tpl.MakeBuffer(tmp, distance, 36, True)
138         Set buf2 = GIS.get("Buffers").AddShape(buf1)
140         Set buf1 = Nothing
142         Set tpl = Nothing
            
144         If g_CurrentTool = oRadiusSelect Then
146             SafeMoveFirst g_RSAppSettings
148             g_RSAppSettings.Find "SettingName = 'PCodeAdminLocation'"
    
150             Set lL = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)

152             If lL Is Nothing Then
154                 GIS.UpDate
                    Exit Sub
                End If
    
156             lL.DeselectAll
    
                ' check all shapes
158             Set tmp = lL.FindFirst(buf2.Extent, "", buf2, RELATE_INTERSECT, True)
    
160             While Not tmp Is Nothing
                    ' if any has a common point with buffer mark it
162                 Set tmp = tmp.MakeEditable
                    'sVals = sVals & vbCrLf & tmp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value) '("CITY")
164                 colPlaces.Add tmp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").Value)
166                 tmp.IsSelected = True
168                 Set tmp = lL.FindNext
                Wend
    
172             frmVillageAvailable.AddPlaces colPlaces
174             frmVillageAvailable.Show vbModeless, Me
            Else
176             Set lL = GIS.get(m_SelLyrCol.Item(ComSelLayer.List(ComSelLayer.ListIndex)))
                
                RemoveAlltabs
                
178             'lL.Lock
                
180             lL.DeselectAll
                
182             Set tmp = lL.FindFirst(buf2.Extent, "", buf2, RELATE_INTERSECT, True)
184             While Not tmp Is Nothing
                    
186                 Set tmp = tmp.MakeEditable
188                 tmp.IsSelected = True

'                    If m_frmSelections.Visible Then
                        AddSelection tmp
'                    End If

190                 Set tmp = lL.FindNext
                Wend
                
192             'lL.Unlock

            End If
            m_oDrawLyr.RevertAll
            GIS.UpDate
196         translated = True
        End If
                          
198     If g_CurrentTool = oLineSelect Then
            'DoLineSelect
        End If
        
200     If g_CurrentTool = oPolyLineSelect Then
        
        End If
        
202     If g_CurrentTool = oPolySelect Then
        
        End If
        
204     If g_CurrentTool = oRectSelect Then
        
        End If

206     If g_CurrentTool = oSingleSelect Then
        
        End If
        
208     If g_CurrentTool = oMultiSelect Then
        
        End If
                  
        '    translated = True
        '
        '    Dim ptg As XGIS_Point
        '    Dim shp As XGIS_Shape
        '
        '    Set ptg = GIS.ScreenToMap(GisUtils.Point(x, y))
        '    Set shp = GIS.Locate(ptg, 5 / GIS.Zoom) ' 5 pixels precision
        '
        '    Select Case g_CurrentTool
        '
        '        Case oSingleSelect
        '
        '            If Not shp Is Nothing Then
        '                shp.Flash
        '                frmAttributes.ShowShape shp
        '                shp.IsSelected = True
        '                frmAttributes.ShowSelected shp.Layer
        '            End If
        '
        '        Case OASIS_TOOLS.oCreateLocationArea, oCreateLocationPolyline, oCreateLocationPoint
        '
        '            If Button = XmbLeft Then
        '                If GIS.Mode = XgisEdit Then
        '                    If g_oEditLayer Is Nothing Then
        '                        Exit Sub
        '                    Else
        '
        '
        '                    '    GIS.Editor.CreateShape g_oEditLayer, ptg, XgisShapeTypePolygon
        '
        '                        Select Case g_CurrentTool
        '
        '                            Case OASIS_TOOLS.oCreateLocationArea
        '                                GIS.Editor.CreateShape g_oEditLayer, ptg, XgisShapeTypePolygon
        '
        '                                'If m_oShpPolygon Is Nothing Then
        '                                '    Set m_oShpPolygon = GIS.Editor.CreateShape(g_oEditLayer, ptg, XgisShapeTypePolygon)
        '                                'Else
        '                                '    GIS.Editor.AddPoint ptg
        '                                'End If
        '                            Case OASIS_TOOLS.oCreateLocationPolyline
        '                                GIS.Editor.CreateShape g_oEditLayer, ptg, XgisShapeTypeArc
        ''                                If m_oShpArc Is Nothing Then
        ''                                    Set m_oShpArc = GIS.Editor.CreateShape(g_oEditLayer, ptg, XgisShapeTypeArc)
        ''                                Else
        ''                                    GIS.Editor.AddPoint ptg
        ''                                End If
        '
        '                            Case OASIS_TOOLS.oCreateLocationPoint
        '
        '                                'Set m_oShpPoint = GIS.Editor.CreateShape(g_oEditLayer, ptg, XgisShapeTypePoint)
        '
        '
        '                        End Select
        '
        '
        '
        '                    End If
        '                End If
        '            Else
        '                If GIS.Mode = XgisEdit Then
        '                    endEdit
        '                End If
        '            End If
        '
        '
        '    End Select

210     miniMapRefresh

212     Debug.Print "----- **MouseUP End** ----"
214     Debug.Print "gdpts:" & UBound(gdpts)
216     Debug.Print "ptsSel:" & UBound(ptsSel)

        '<EhFooter>
        Exit Sub

GIS_OnMouseUp_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GIS_OnMouseUp " & "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub GIS_OnPaint(translated As Boolean)
        '<EhHeader>
        On Error GoTo GIS_OnPaint_Err
        '</EhHeader>
100     '  Me.caption = "zoom: " & Format(GIS.ZoomEx, "0.0000")
        '<EhFooter>
        Exit Sub

GIS_OnPaint_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.GIS_OnPaint " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmbSnap_Click()
        '<EhHeader>
        On Error GoTo cmbSnap_Click_Err
        '</EhHeader>
100   If cmbSnap.ListIndex > 0 Then
102     GIS.Editor.SnapLayer = GIS.get(cmbSnap.List(cmbSnap.ListIndex))
      Else
104     GIS.Editor.SnapLayer = Nothing
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

Private Sub GIS_OnVisibleExtentChange(translated As Boolean)
        '<EhHeader>
        On Error GoTo GIS_OnVisibleExtentChange_Err
        '</EhHeader>
100     Set g_PrevExt = GIS.VisibleExtent
    
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
        
100     'm_frmDebug.DebugPrint  "INC Visible = " & m_oSQLIncLyr.Params.Visible & "  " & GIS.Zoom
        
102     translated = True
104     miniMapRefresh
106     Set g_PrevExt = GIS.VisibleExtent
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
102     GIS.Mode = XgisZoomEx

        '    m_oIncShpPt.MakeEditable
        '    m_oIncShpPt.Delete
    
104     If Not m_oIncShpPt Is Nothing Then
106         m_oSQLIncLyr.Delete m_oIncShpPt.uID
        End If

108     Set m_oIncShpPt = Nothing
    
110     m_oSQLIncLyr.SaveData
    
112     GIS.UpDate
    
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
        Dim ptg As New XGIS_Point
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
     
114     GIS.CenterViewport ptg
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
        
        m_bUseDistrictOnly = IIf(m_fmrAddIncident.chkUseDistrict.Value = vbChecked, True, False)
        
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

Private Sub m_frmAddSHPWiz_AddShape(dtShapeType As TatukGIS_DK.XGIS_ShapeType, dtFeatureType As OASISLocationType)
        '<EhHeader>
        On Error GoTo m_frmAddSHPWiz_AddShape_Err
        '</EhHeader>
    
100     m_prevTool = g_CurrentTool
    
102     Select Case dtShapeType
    
            Case TatukGIS_DK.XGIS_ShapeType.XgisShapeTypeArc
104             g_CurrentTool = oCreateLocationPolyline
106         Case TatukGIS_DK.XGIS_ShapeType.XgisShapeTypePoint
108             g_CurrentTool = oCreateLocationPoint
110         Case TatukGIS_DK.XGIS_ShapeType.XgisShapeTypePolygon
112             g_CurrentTool = oCreateLocationArea
114         Case TatukGIS_DK.XGIS_ShapeType.XgisShapeTypeMultiPoint
116             g_CurrentTool = oCreateLocationMultipoint
118         Case TatukGIS_DK.XGIS_ShapeType.XgisShapeTypeUnknown
120             Err.Raise 666, , "NOT IMPLEMENTED YET!"
        End Select
    
122     Select Case dtFeatureType
    
            Case OASISLocationType.Dynamic
124             g_CurrentLocationType = Dynamic
126         Case OASISLocationType.Permanent
128             g_CurrentLocationType = Permanent
130         Case OASISLocationType.Temporary
132             g_CurrentLocationType = Temporary
        End Select
    
134     g_CurrentFeatureType = Location
    
136     SetOASISTool
        
138     m_frmAddSHPWiz.Hide
    
        '<EhFooter>
        Exit Sub

m_frmAddSHPWiz_AddShape_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmAddSHPWiz_AddShape " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetOASISTool()
        '<EhHeader>
        On Error GoTo SetOASISTool_Err
        '</EhHeader>

    '    Dim os As Variant
    '
    '    os = GIS.Mode
    '
    '    GIS.CursorPrepare 2048, LoadCursorFromFile("C:\Program Files\Microsoft Visual Studio\Common\Graphics\Cursors\4WAY04.CUR")
    '    Cursor.Current = 2048
    '    GIS.CursorForZoom = 2048
    '    GIS.Mode = XgisZoom
    '    GIS.Mode

100     Select Case g_CurrentTool
    
            Case oCreateLocationArea
102             GIS.Mode = XgisEdit
               ' Set m_oShpPolygon = Nothing
               ' GIS.Mode = XgisUserDefined 'XgisEdit
                'GIS.Cursor
                '  GIS.CursorPrepare( 2048, LoadCursorFormFile( "mycursor.cur" ) ) ;

104         Case oCreateLocationPoint
106             GIS.Mode = XgisEdit
                'Set m_oShpPoint = Nothing
                'GIS.Mode = XgisUserDefined 'XgisEdit

108         Case oCreateLocationPolyline
110             GIS.Mode = XgisEdit
                'Set m_oShpArc = Nothing
                'GIS.Mode = XgisUserDefined 'XgisEdit

112         Case OASIS_TOOLS.oZoomEx
114             GIS.Mode = XgisZoomEx

116         Case OASIS_TOOLS.oZoom
118             GIS.Mode = XgisZoom

120         Case OASIS_TOOLS.oSingleSelect
122             GIS.Mode = XgisSelect

124         Case OASIS_TOOLS.oPan
126             GIS.Mode = XgisDrag
    
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

Private Sub m_frmMAModule_LoadMALyr(oLyr As TatukGIS_DK.XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo m_frmMAModule_LoadMALyr_Err
        '</EhHeader>
100     AddLayer oLyr
        '<EhFooter>
        Exit Sub

m_frmMAModule_LoadMALyr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMAModule_LoadMALyr " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmMAModule_UnloadMALyr(sLyrName As String)
        '<EhHeader>
        On Error GoTo m_frmMAModule_UnloadMALyr_Err
        '</EhHeader>
100     GIS.Delete sLyrName
102     FillCOPValues
104     GIS.UpDate
        '<EhFooter>
        Exit Sub

m_frmMAModule_UnloadMALyr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMAModule_UnloadMALyr " & _
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
104     WebBrowser1.Navigate2 g_RSAppSettings.Fields.Item("SettingValue1").Value

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
104     WebBrowser1.Navigate2 g_RSAppSettings.Fields.Item("SettingValue1").Value

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
    
104     If g_RSAppSettings.Fields.Item("SettingValue1").Value <> "" Then
106         WebBrowser1.Navigate2 g_RSAppSettings.Fields.Item("SettingValue1").Value
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
104         GIS.UpDate
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
102     GIS.UpDate
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
102         GIS.get(m_sSecAnalysisLyrName).Params.Visible = False
        End If
    
104     SafeMoveFirst g_RSAppSettings
    
106     Select Case m_frmMnuOperations.SecAnalysisevel
    
            Case 0
108             g_RSAppSettings.Find "SettingName = 'AdmProvSec'"
            
110             m_sSecAnalysisFieldName = g_RSAppSettings.Fields.Item("SettingValue3").Value
112             m_sSecAnalysisLyrName = g_RSAppSettings.Fields.Item("SettingValue1").Value

114         Case 1
116             g_RSAppSettings.Find "SettingName = 'AdmDistSec'"
            
118             m_sSecAnalysisFieldName = g_RSAppSettings.Fields.Item("SettingValue3").Value
120             m_sSecAnalysisLyrName = g_RSAppSettings.Fields.Item("SettingValue1").Value
        
122         Case Else
124             g_RSAppSettings.Find "SettingName = 'SecurityGridZoomLevels'"
            
126             Select Case CLng(Mid(GIS.ScaleAsText, 3))
            
                    Case Is > CLng(g_RSAppSettings.Fields.Item("SettingValue1").Value)
128                     SafeMoveFirst g_RSAppSettings
130                     g_RSAppSettings.Find "SettingName = 'SecGrid1'"
132                 Case Is > CLng(g_RSAppSettings.Fields.Item("SettingValue2").Value)
134                     SafeMoveFirst g_RSAppSettings
136                     g_RSAppSettings.Find "SettingName = 'SecGrid2'"
138                 Case Else
140                     SafeMoveFirst g_RSAppSettings
142                     g_RSAppSettings.Find "SettingName = 'SecGrid3'"
                End Select
            
144             m_sSecAnalysisFieldName = g_RSAppSettings.Fields.Item("SettingValue3").Value
146             m_sSecAnalysisLyrName = g_RSAppSettings.Fields.Item("SettingValue1").Value
    
        End Select

148     If m_sSecAnalysisLyrName = "" Then Exit Sub

150     GIS.get(m_sSecAnalysisLyrName).Params.Visible = True

152     Command1_Click
154     GIS.get(m_sSecAnalysisLyrName).draw
156     GIS.viewer.PrintClipboard
    
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
100     If Not m_fmrAddIncident.Visible Then
102         m_fmrAddIncident.Init m_Cnn, "Enter Your Name"
104         m_fmrAddIncident.Show vbModeless, Me
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
100     GIS.get(sLayerName).Active = bActivated
102     GIS.UpDate
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

Private Sub m_frmMnuOperations_ZoomToLoc(sName As String, sID As String)
        '<EhHeader>
        On Error GoTo m_frmMnuOperations_ZoomToLoc_Err
        '</EhHeader>
    Dim lL As XGIS_LayerVector
    Dim shp As XGIS_Shape

     On Error Resume Next
100         SafeMoveFirst g_RSAppSettings
102         g_RSAppSettings.Find "SettingName = 'MAPcodeLayer'"
        
104         Set lL = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value)
106         Set shp = lL.FindFirst(GIS.Extent, g_RSAppSettings.Fields.Item("SettingValue5").Value & " = '" & sID & "'", Nothing, "", True)
        
        
            'm_frmDebug.DebugPrint  shp.GetField(g_RSAppSettings.Fields.Item("SettingValue2").value)
108         If Not shp Is Nothing Then
110             GIS.VisibleExtent = shp.Extent
                'GIS.Update
                'shp.IsSelected = True
112             shp.MakeEditable
114             shp.Flash
            End If

    
        '<EhFooter>
        Exit Sub

m_frmMnuOperations_ZoomToLoc_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmMnuOperations_ZoomToLoc " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmOvMap_MapOnMouseMove(translated As Boolean, ByVal Shift As TatukGIS_DK.XShiftState, ByVal x As Long, ByVal y As Long)
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

Private Sub m_frmOvMap_MapOnMouseUp(translated As Boolean, ByVal Button As TatukGIS_DK.XMouseButton, ByVal Shift As TatukGIS_DK.XShiftState, ByVal x As Long, ByVal y As Long)
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

110     Set m_frmAddSHPWiz = New frmAddPointWZ
112     Set m_fmrAddIncident = New frmAddIncident
114     Set m_frmOvMap = New frmOVMap
116     Set m_frmMAModule = New frmMAModule
118     Set m_frmIOMJOC = New frmIOMJOC
120     Set m_frmLocator = New frmLocator
122     Set m_frmOASISCharts = New frmOASISCharts
        Set m_frmTextAnnoSettings = New frmTextAnnoSettings
        Set m_frmAttributes = New frmAttributes
        Set m_frmSearch = New frmSearch
        Set m_frmUpdateSettings = New frmUpdateSettings
        Set m_frmMainSettings = New frmMainSettings
        Set m_frmSelections = New frmSelections
        Set m_frmSelectorReports = New frmSelectorReports
        Set m_frmSelectorSettings = New frmSelectorSettings
        Set m_frmSpatialiseDD = New frmSpatialiseDD
        Set m_frmResourcesFinder = New frmResourcesFinder
        Set m_frmWVISitrepGenerator = New frmWVISitrepGenerator
        Set m_frmWVIIncidents = New frmWVIIncidents
        Set m_frmWVIControlPanel = New frmWVIControlPanel
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
        
        'This screws things up in runtime so disable for now:

        '134     If Not bDebugMode Then
        '136         pThread.CreateWin32Thread Me, "CheckDataPacks", 0
        '
        '138         If pThread.IsThreadRunning = True Then
        '140             pThread.ThreadPriority = THREAD_PRIORITY_NORMAL
        '            End If
        '
        '        End If
        
        On Error Resume Next
        
        If g_udtSynchUpdateOptions.AutoUpdate Then
142         ShellExecute Me.hWnd, vbNullString, g_sAppPath & "\OASIS_SynchNG_Client.exe", "CheckBackground", "C:\", 1
144         ShellExecute Me.hWnd, vbNullString, g_sAppPath & "\AUClient.exe", "CheckBackground", "C:\", 1
        End If
        
        'Call ShellExecute(Me.Hwnd, vbNullString, "C:\Program Files\Globalsat\TR Management Center\OASIS_Tracker.exe", "", "c:\", 2)
    
146     Set oINIReader = Nothing

        CheckOVBFiles
        
        OpenLocalAppSettings
        
'        With oCoordTransSettings
'            .Inverse_Flattening = 6378137
'            .Semi_Major_Axis = CDbl("298.2572236")
'        End With
'
'        With oMapTipSetting
'            .Enabled = False
'            .MapTipLayer = "--All--"
'            .MapTipField = "UID"
'            .TextColor = &H80000012
'            .TipColor = &HC0FFFF
'            .TipDelay = 4
'            .TipBorder = False
'            tmrToolTip.Interval = 4000
'            tmrToolTip.Enabled = True
'        End With
'
'        With oSelectionStyle
'            .Color = GIS.SelectionColor
'            .OutLineOnly = GIS.SelectionOutlineOnly
'            .Transparency = GIS.SelectionTransparency
'            .Width = GIS.SelectionWidth
'        End With
'
'        oLocatorSettings.Level1 = "Intersect Interior Interior" 'GisUtils.GIS_RELATE_INTERSECT_INTERIOR_INTERIOR
'        oLocatorSettings.Level2 = "Contains" 'GisUtils.GIS_RELATE_CONTAINS
        
        PrepareConvertorUTM_LL
   
        '            AB.Bands("bNavPane").ChildBands.CurrentChildBand = AB.Bands("bNavPane").ChildBands("cbProfile")
            
        '            AB.RecalcLayout
        
        '        EventLayer.CachedPaint = False
        
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

        Dim oRS As New adodb.Recordset
        Dim oRSSQLLayer As New adodb.Recordset
        Dim sMapProj As String
        Dim sText As String
        Dim oStream As adodb.Stream
        Dim bInUse As String
        
        Dim oSQLLyr As XGIS_LayerSqlAdo
        Dim oSQLLyrAbs As XGIS_LayerAbstract
        
        Dim sINI As String
        'Dim oStream As ADODB.Stream
        Dim oIni As New XGIS_Ini
        Dim lNumLayers As Long
     
100     If UCase(GIS.ProjectName) = UCase(g_sAppPath & "\data\user\Exports\tempproj.ttkgp") Then
                
102         GIS.UpDate

104         DoEvents
            m_Cnn.Execute "DELETE from [ttkGISLayerSQLInProject]"
             
106         With oRSSQLLayer
            
                lNumLayers = GIS.Items.Count
            
                Do Until lNumLayers < 1
            
                    Kill g_sAppPath & "\data\user\Exports\tempini.ini"
                    Set oSQLLyrAbs = GIS.Items.Item(lNumLayers - 1)
            
108                 .Open "SELECT * from [ttkGISLayerSQLInProject] WHERE [LayerCaption] = '" & oSQLLyrAbs.Name & "'", m_Cnn, adOpenDynamic, adLockBatchOptimistic
            
114                 Set oSQLLyr = GIS.get(oSQLLyrAbs.Name)
                    
116                 If Not oSQLLyr Is Nothing Then
                    
                        If .EOF Then
                            .AddNew
                            .Fields("LayerName").Value = oSQLLyr.Table
                            .Fields("LayerCaption").Value = oSQLLyr.Name
                        End If
                        
118                     Set oStream = New adodb.Stream
120                     Set oIni = New XGIS_Ini
                        
124                     oIni.Create_ oSQLLyr, g_sAppPath & "\data\user\Exports\tempini.ini"
126                     oSQLLyr.ParamsList.SaveToIni oIni
128                     oIni.Save
                        
130                     oStream.Open
132                     oStream.Type = 2    ' type = text
134                     oStream.Charset = "ascii"
136                     oStream.LoadFromFile g_sAppPath & "\data\user\Exports\tempini.ini"
138                     .Fields("INISettings").Value = oStream.ReadText
140                     .Fields("Transparency").Value = oSQLLyr.Transparency
142                     .Fields("IsVisible").Value = oSQLLyr.Active
144                     .Fields("Sequence").Value = lNumLayers - 1 'oSQLLyr.ZOrder
146                     .Fields("IsExpanded").Value = oSQLLyr.Collapsed
                        .Fields("Dialect").Value = oSQLLyr.SQLParameter("DIALECT")
                        .Fields("ADO").Value = Replace(oSQLLyr.SQLParameter("ADO"), g_sAppPath, "CLIENTDBPATH")
                        .Fields("XMIN").Value = oSQLLyr.Extent.XMin
                        .Fields("XMAX").Value = oSQLLyr.Extent.XMax
                        .Fields("YMIN").Value = oSQLLyr.Extent.YMin
                        .Fields("YMAX").Value = oSQLLyr.Extent.YMax

148                     'If .Fields("Sequence").Value < iLowestZOrder Then iLowestZOrder = .Fields("Sequence").Value
                        'If .Fields("Sequence").Value = 0 Then Stop
150                     .UpdateBatch adAffectCurrent
152                     oStream.Close
154                     Set oSQLLyr = Nothing
156                     Set oStream = Nothing
158                     Set oIni = Nothing

                    End If
                    
                    .Close
160                 lNumLayers = lNumLayers - 1
                Loop
                
162             '.MoveFirst
                
164             Do Until 5 = 5 '.EOF

                    'If GIS.get(frmLayerSelection.GetItem).FileInfo = "TatukGIS SQL Vector Coverage (TTKLS)" Then
                    If GIS.get(.Fields("LayerCaption").Value) Is Nothing Then
                        m_Cnn.Execute "DELETE FROM [ttkGISLayerSQLInProject] WHERE [LayerCaption] = '" & .Fields("LayerCaption").Value & "'"
                    Else
                        'If iLowestZOrder = 1 Then .Fields("Sequence").Value = .Fields("Sequence").Value - 1
                        
                        .Fields("Sequence").Value = GIS.get(.Fields("LayerCaption").Value).ZOrderEx
                        GIS.Delete .Fields("LayerCaption").Value
                        .UpdateBatch adAffectCurrent
                    End If
    
172                 .MoveNext
                Loop
                
174             ' .Close
176             GIS.SaveAll
            
            End With
                
178         Set oRSSQLLayer = Nothing
                
180         GIS.SaveProjectAs g_sAppPath & "\data\user\Exports\tempproj.ttkgp", True
    
182         Set oStream = New adodb.Stream
184         oStream.Open
186         oStream.Type = 2    ' type = text
188         oStream.Charset = "ascii"
190         oStream.LoadFromFile (g_sAppPath & "\data\user\Exports\tempproj.ttkgp")
192         sText = oStream.ReadText
194         oStream.Close
196         Set oStream = Nothing
    
198         If Len(sText) > 10 Then
        
200             With oRS
                
202                 .Open "SELECT * from [ttkGISProjectDef]", m_Cnn, adOpenDynamic, adLockBatchOptimistic

204                 If Not .State = adStateClosed Then
                
206                     If Not .EOF Then
208                         .Delete adAffectCurrent
210                         .UpdateBatch
                        End If
                    
212                     .AddNew
214                     .Fields("MapData").Value = sText
216                     .Fields("sGUID").Value = GUIDGen
218                     .Fields("InUse").Value = True
220                     .Fields("XMin").Value = GIS.VisibleExtent.XMin
222                     .Fields("XMax").Value = GIS.VisibleExtent.XMax
224                     .Fields("YMin").Value = GIS.VisibleExtent.YMin
226                     .Fields("YMax").Value = GIS.VisibleExtent.YMax

228                     .UpdateBatch adAffectCurrent
230                     .Close
                    End If

                End With
        
            End If
        
        Else
        
232         With oRSSQLLayer
            
234             .Open "SELECT * from [ttkGISLayerSQLInProject] ORDER BY [Sequence] DESC", m_Cnn, adOpenDynamic, adLockBatchOptimistic

236             Do Until .EOF
                    
238                 If Not GIS.get(.Fields("LayerCaption").Value) Is Nothing Then GIS.get(.Fields("LayerCaption").Value).ZOrder = 0
240                 .MoveNext
                Loop
         
242             .Close
244             GIS.SaveAll
            
            End With
            
        End If

246     Set oRS = Nothing
248     Set oRSSQLLayer = Nothing

        '<EhFooter>
        Exit Sub

UpdateProjectFileToDB_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.UpdateProjectFileToDB " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function AddSQLLayersToProject()
        '<EhHeader>
        On Error GoTo AddSQLLayersToProject_Err
        '</EhHeader>

      Dim oRS As New adodb.Recordset
      Dim oRsDD As adodb.Recordset
      Dim oConnectionDD As adodb.Connection
  
      Dim oSQLLyr As XGIS_LayerSqlAdo
      Dim oExtent As XGIS_Extent
     ' Dim bExtentDone As Boolean
  
      Dim sINI As String
      Dim oStream As adodb.Stream
      Dim oIni As New XGIS_Ini
  
      Dim sName As String
  
      'Exit Function
  
100   With oRS
 
102       .Open "SELECT * from [ttkGISLayerSQLInProject] ORDER BY [Sequence] DESC", m_Cnn, adOpenDynamic, adLockBatchOptimistic

104       Do Until .EOF

              sName = .Fields("LayerCaption").Value
134           Set oRsDD = New adodb.Recordset
136           Set oConnectionDD = New adodb.Connection
138           oConnectionDD.Open Replace(.Fields("ADO").Value, "CLIENTDBPATH", g_sAppPath)
140           oRsDD.Open "SELECT * FROM [ttkGISLayerSQL] WHERE [name] = '" & .Fields("LayerName").Value & "'", oConnectionDD, adOpenDynamic, adLockBatchOptimistic
142           If oRsDD.EOF Then oRsDD.AddNew
144           oRsDD.Fields("name").Value = .Fields("LayerName").Value
146           oRsDD.Fields("xmin").Value = .Fields("xmin").Value
148           oRsDD.Fields("xmax").Value = .Fields("xmax").Value
150           oRsDD.Fields("ymin").Value = .Fields("ymin").Value
152           oRsDD.Fields("ymax").Value = .Fields("ymax").Value
154           oRsDD.Fields("shapetype").Value = .Fields("Shapetype").Value
156           oRsDD.UpdateBatch
158           oRsDD.Close
160           oConnectionDD.Close
162           Set oRsDD = Nothing
164           Set oConnectionDD = Nothing
          
106           Set oSQLLyr = New XGIS_LayerSqlAdo
108           Set oExtent = New XGIS_Extent
        
110
112           oSQLLyr.Name = .Fields("LayerCaption").Value
114           oSQLLyr.SQLParameter("LAYER") = .Fields("LayerName").Value
116           oSQLLyr.SQLParameter("DIALECT") = .Fields("Dialect").Value
118           oSQLLyr.SQLParameter("ADO") = Replace(.Fields("ADO").Value, "CLIENTDBPATH", g_sAppPath)
120           oSQLLyr.HideFromLegend = False
122           oSQLLyr.Params.Visible = True
124           oExtent.Prepare .Fields("XMIN").Value, .Fields("YMIN").Value, .Fields("XMAX").Value, .Fields("YMAX").Value
126           oSQLLyr.Extent = oExtent
128           Set oStream = New adodb.Stream
130           Set oIni = New XGIS_Ini

132           GIS.Add oSQLLyr
          

          
166           oSQLLyr.Transparency = .Fields("Transparency").Value
168           oSQLLyr.Collapsed = .Fields("IsExpanded").Value
170           oSQLLyr.Active = .Fields("IsVisible").Value
          
172           If Not IsNull(.Fields("INISettings").Value) Then
              
174               sINI = .Fields("INISettings").Value
              
176               oStream.Open
178               oStream.Type = 2    ' type = text
180               oStream.Charset = "ascii"
182               Kill g_sAppPath & "\data\user\Exports\tempini.ini"
184               oStream.WriteText sINI, adWriteChar
186               oStream.SaveToFile g_sAppPath & "\data\user\Exports\tempini.ini"
              
188               oIni.Create_ oSQLLyr, g_sAppPath & "\data\user\Exports\tempini.ini"
                  'oSQLLyr.Params.LoadFromIni oIni
190               oSQLLyr.ParamsList.LoadFromIni oIni
192               oStream.Close
              
              End If

194           Set oStream = Nothing
196           Set oIni = Nothing
198           .MoveNext

          Loop
      
200       If Not .EOF Or Not .BOF Then .MoveFirst

202       Do Until .EOF
204           Set oSQLLyr = GIS.get(.Fields("LayerCaption").Value)
206           oSQLLyr.ZOrder = .Fields("Sequence").Value
208           .MoveNext
          Loop

210       .Close
 
      End With
  
      'If GIS.VisibleExtent.XMin = 0 And GIS.VisibleExtent.YMax = 0 Then
          'GIS.VisibleExtent = GIS.Extent
      'End If

        '<EhFooter>
        Exit Function

AddSQLLayersToProject_Err:
  m_frmDebug.DebugPrint "Error loading SQL Layer: " & sName
  m_frmDebug.DebugPrint Err.Description & vbCrLf & "in OASISClient.frmMain.AddSQLLayersToProject " & "at line " & Erl
  m_frmDebug.DebugPrint "resuming next...."
  'MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.AddSQLLayersToProject " & "at line " & Erl
  'Stop
  Resume Next
  '</EhFooter>
End Function


Private Function PrepareLocalAppSettingValue(sSetting As String) As Boolean
    SafeMoveFirst g_RSLocalAppSettings
    
    With g_RSLocalAppSettings
        .Find "SettingName = '" & sSetting & "'"
    
        If .EOF Then
            .AddNew
            .Fields("SettingName").Value = sSetting
            .UpDate
        Else
            PrepareLocalAppSettingValue = True
        End If
    
    End With
 End Function

Private Sub OpenLocalAppSettings()

    With oMapSettings
        .AlwaysSaveMapStateOnExit = CBool(IIf(PrepareLocalAppSettingValue("oMapSettings.AlwaysSaveMapStateOnExit"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .AutoScroll = CBool(IIf(PrepareLocalAppSettingValue("oMapSettings.AutoScroll"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .iMAPUnits = CInt(IIf(PrepareLocalAppSettingValue("oMapSettings.iMAPUnits"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 0))
        .MapRotation = CInt(IIf(PrepareLocalAppSettingValue("oMapSettings.MapRotation"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 0))
        .ScrollBars = CBool(IIf(PrepareLocalAppSettingValue("oMapSettings.ScrollBars"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .StoreLayerParamsInProject = CBool(IIf(PrepareLocalAppSettingValue("oMapSettings.StoreLayerParamsInProject"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
    End With

    With oIncidentLayerSettings
        .CachedPaint = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.CachedPaint"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .ConfigFilePAth = IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.ConfigFilePAth"), g_RSLocalAppSettings.Fields("SettingValue1").Value, "")
        .HideFromLegend = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.HideFromLegend"), g_RSLocalAppSettings.Fields("SettingValue1").Value, True))
        .IgnoreShapeParams = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.IgnoreShapeParams"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .IncrementalPaint = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.IncrementalPaint"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .UseConfig = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.UseConfig"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .UseFileParams = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.UseFileParams"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .VisibleFromStart = CBool(IIf(PrepareLocalAppSettingValue("oIncidentLayerSettings.VisibleFromStart"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
    End With


    With oCoordTransSettings
        .False_Easting = CLng(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.False_Easting"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 0))
        .False_Northing = CLng(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.False_Northing"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 0))
        .Lat_Of_Origin = CLng(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.Lat_Of_Origin"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 0))
        .Long_Of_Origin = CLng(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.Long_Of_Origin"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 0))
        .Sphere = CBool(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.Sphere"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .Inverse_Flattening = IIf(PrepareLocalAppSettingValue("oCoordTransSettings.Inverse_Flattening"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 6378137)
        .Semi_Major_Axis = CDbl(IIf(PrepareLocalAppSettingValue("oCoordTransSettings.Semi_Major_Axis"), g_RSLocalAppSettings.Fields("SettingValue1").Value, "298.2572236"))
    End With
        
    With oUrlLayerSettings
        .AutoShutTime = CInt(IIf(PrepareLocalAppSettingValue("oUrlLayerSettings.AutoShutTime"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 0))
        .AutoShutWin = CBool(IIf(PrepareLocalAppSettingValue("oUrlLayerSettings.AutoShutWin"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .UseExtendedInfoWin = CBool(IIf(PrepareLocalAppSettingValue("oUrlLayerSettings.UseExtendedInfoWin"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .WinHeight = CInt(IIf(PrepareLocalAppSettingValue("oUrlLayerSettings.WinHeight"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 5235))
        .WinWidth = CInt(IIf(PrepareLocalAppSettingValue("oUrlLayerSettings.WinWidth"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 4800))
    End With
                
    With oMapTipSetting
        .Enabled = CBool(IIf(PrepareLocalAppSettingValue("oMapTipSetting.Enabled"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .MapTipLayer = IIf(PrepareLocalAppSettingValue("oMapTipSetting.MapTipLayer"), g_RSLocalAppSettings.Fields("SettingValue1").Value, "--All--")
        .MapTipField = IIf(PrepareLocalAppSettingValue("oMapTipSetting.MapTipField"), g_RSLocalAppSettings.Fields("SettingValue1").Value, "UID")
        .TextColor = CLng(IIf(PrepareLocalAppSettingValue("oMapTipSetting.TextColor"), g_RSLocalAppSettings.Fields("SettingValue1").Value, &H80000012))
        .TipColor = CLng(IIf(PrepareLocalAppSettingValue("oMapTipSetting.TipColor"), g_RSLocalAppSettings.Fields("SettingValue1").Value, &HC0FFFF))
        .TipDelay = CInt(IIf(PrepareLocalAppSettingValue("oMapTipSetting.TipDelay"), g_RSLocalAppSettings.Fields("SettingValue1").Value, 4))
        .TipBorder = CBool(IIf(PrepareLocalAppSettingValue("oMapTipSetting.TipBorder"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        tmrToolTip.Interval = .TipDelay * 1000
        tmrToolTip.Enabled = .Enabled
    End With
        
    With oSelectionStyle
        .Color = CLng(IIf(PrepareLocalAppSettingValue("oSelectionStyle.Color"), g_RSLocalAppSettings.Fields("SettingValue1").Value, GIS.SelectionColor))
        .OutLineOnly = CBool(IIf(PrepareLocalAppSettingValue("oSelectionStyle.OutLineOnly"), g_RSLocalAppSettings.Fields("SettingValue1").Value, GIS.SelectionOutlineOnly))
        .Transparency = CInt(IIf(PrepareLocalAppSettingValue("oSelectionStyle.Transparency"), g_RSLocalAppSettings.Fields("SettingValue1").Value, GIS.SelectionTransparency))
        .Width = CInt(IIf(PrepareLocalAppSettingValue("oSelectionStyle.Width"), g_RSLocalAppSettings.Fields("SettingValue1").Value, GIS.SelectionWidth))
    End With
        
    oLocatorSettings.Level1 = IIf(PrepareLocalAppSettingValue("oLocatorSettings.Level1"), g_RSLocalAppSettings.Fields("SettingValue1").Value, "Intersect Interior Interior") 'GisUtils.GIS_RELATE_INTERSECT_INTERIOR_INTERIOR
    oLocatorSettings.Level2 = IIf(PrepareLocalAppSettingValue("oLocatorSettings.Level2"), g_RSLocalAppSettings.Fields("SettingValue1").Value, "Contains")
        
    With g_ZoomToSettings
        .SaveOnExit = CBool(IIf(PrepareLocalAppSettingValue("g_ZoomToSettings.SaveOnExit"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
        .UseMultiple = CBool(IIf(PrepareLocalAppSettingValue("g_ZoomToSettings.UseMultiple"), g_RSLocalAppSettings.Fields("SettingValue1").Value, False))
    End With
        
End Sub

Private Sub SavePrivateAppSettings()
        '<EhHeader>
        On Error GoTo SavePrivateAppSettings_Err
        '</EhHeader>
        
        With oMapSettings
'            .AlwaysSaveMapStateOnExit
'            .AutoScroll
'            .iMAPUnits
'            .MapRotation
'            .ScrollBars
'            .StoreLayerParamsInProject
        End With


100     With oIncidentLayerSettings
102         PrepareLocalAppSettingValue "oIncidentLayerSettings.CachedPaint"
104         g_RSLocalAppSettings.Fields("SettingValue1").Value = .CachedPaint
106         g_RSLocalAppSettings.UpDate
        
108         PrepareLocalAppSettingValue "oIncidentLayerSettings.ConfigFilePAth"
110         g_RSLocalAppSettings.Fields("SettingValue1").Value = IIf(Len(.ConfigFilePAth) > 0, .ConfigFilePAth, " ")
112         g_RSLocalAppSettings.UpDate
        
114         PrepareLocalAppSettingValue "oIncidentLayerSettings.HideFromLegend"
116         g_RSLocalAppSettings.Fields("SettingValue1").Value = .HideFromLegend
118         g_RSLocalAppSettings.UpDate
        
120         PrepareLocalAppSettingValue "oIncidentLayerSettings.IgnoreShapeParams"
122         g_RSLocalAppSettings.Fields("SettingValue1").Value = .IgnoreShapeParams
124         g_RSLocalAppSettings.UpDate
        
126         PrepareLocalAppSettingValue "oIncidentLayerSettings.IncrementalPaint"
128         g_RSLocalAppSettings.Fields("SettingValue1").Value = .IncrementalPaint
130         g_RSLocalAppSettings.UpDate
        
132         PrepareLocalAppSettingValue "oIncidentLayerSettings.UseConfig"
134         g_RSLocalAppSettings.Fields("SettingValue1").Value = .UseConfig
136         g_RSLocalAppSettings.UpDate
        
138         PrepareLocalAppSettingValue "oIncidentLayerSettings.UseFileParams"
140         g_RSLocalAppSettings.Fields("SettingValue1").Value = .UseFileParams
142         g_RSLocalAppSettings.UpDate
        
144         PrepareLocalAppSettingValue "oIncidentLayerSettings.VisibleFromStart"
146         g_RSLocalAppSettings.Fields("SettingValue1").Value = .VisibleFromStart
148         g_RSLocalAppSettings.UpDate

        End With

150     With oCoordTransSettings

152         PrepareLocalAppSettingValue "oCoordTransSettings.False_Northing"
154         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Inverse_Flattening
156         g_RSLocalAppSettings.UpDate
        
158         PrepareLocalAppSettingValue "oCoordTransSettings.False_Easting"
160         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Inverse_Flattening
162         g_RSLocalAppSettings.UpDate
        
164         PrepareLocalAppSettingValue "oCoordTransSettings.Sphere"
166         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Inverse_Flattening
168         g_RSLocalAppSettings.UpDate
        
170         PrepareLocalAppSettingValue "oCoordTransSettings.Long_Of_Origin"
172         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Inverse_Flattening
174         g_RSLocalAppSettings.UpDate
        
176         PrepareLocalAppSettingValue "oCoordTransSettings.Lat_Of_Origin"
178         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Inverse_Flattening
180         g_RSLocalAppSettings.UpDate
        
182         PrepareLocalAppSettingValue "oCoordTransSettings.Inverse_Flattening"
184         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Inverse_Flattening
186         g_RSLocalAppSettings.UpDate
        
188         PrepareLocalAppSettingValue "oCoordTransSettings.Semi_Major_Axis"
190         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Semi_Major_Axis
192         g_RSLocalAppSettings.UpDate
        End With
        
194     With oUrlLayerSettings
196         PrepareLocalAppSettingValue "oUrlLayerSettings.AutoShutTime"
198         g_RSLocalAppSettings.Fields("SettingValue1").Value = .AutoShutTime
200         g_RSLocalAppSettings.UpDate
        
202         PrepareLocalAppSettingValue "oUrlLayerSettings.AutoShutWin"
204         g_RSLocalAppSettings.Fields("SettingValue1").Value = .AutoShutWin
206         g_RSLocalAppSettings.UpDate
        
208         PrepareLocalAppSettingValue "oUrlLayerSettings.UseExtendedInfoWin"
210         g_RSLocalAppSettings.Fields("SettingValue1").Value = .UseExtendedInfoWin
212         g_RSLocalAppSettings.UpDate
        
214         PrepareLocalAppSettingValue "oUrlLayerSettings.WinHeight"
216         g_RSLocalAppSettings.Fields("SettingValue1").Value = .WinHeight
218         g_RSLocalAppSettings.UpDate
        
220         PrepareLocalAppSettingValue "oUrlLayerSettings.WinWidth"
222         g_RSLocalAppSettings.Fields("SettingValue1").Value = .WinWidth
224         g_RSLocalAppSettings.UpDate
        
        End With
    
226     With oMapTipSetting
228         PrepareLocalAppSettingValue "oMapTipSetting.Enabled"
230         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Enabled
232         g_RSLocalAppSettings.UpDate
        
234         PrepareLocalAppSettingValue "oMapTipSetting.MapTipLayer"
236         g_RSLocalAppSettings.Fields("SettingValue1").Value = IIf(Len(.MapTipLayer) > 0, .MapTipLayer, " ")
238         g_RSLocalAppSettings.UpDate
        
240         PrepareLocalAppSettingValue "oMapTipSetting.MapTipField"
242         g_RSLocalAppSettings.Fields("SettingValue1").Value = IIf(Len(.MapTipField) > 0, .MapTipField, " ")
244         g_RSLocalAppSettings.UpDate
        
246         PrepareLocalAppSettingValue "oMapTipSetting.TextColor"
248         g_RSLocalAppSettings.Fields("SettingValue1").Value = .TextColor
250         g_RSLocalAppSettings.UpDate
        
252         PrepareLocalAppSettingValue "oMapTipSetting.TipColor"
254         g_RSLocalAppSettings.Fields("SettingValue1").Value = .TipColor
256         g_RSLocalAppSettings.UpDate
        
258         PrepareLocalAppSettingValue "oMapTipSetting.TipDelay"
260         g_RSLocalAppSettings.Fields("SettingValue1").Value = .TipDelay
262         g_RSLocalAppSettings.UpDate
        
264         PrepareLocalAppSettingValue "oMapTipSetting.TipBorder"
266         g_RSLocalAppSettings.Fields("SettingValue1").Value = .TipBorder
268         g_RSLocalAppSettings.UpDate
           
        End With
        
270     With oSelectionStyle
272         PrepareLocalAppSettingValue "oSelectionStyle.Color"
274         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Color
276         g_RSLocalAppSettings.UpDate
        
278         PrepareLocalAppSettingValue "oSelectionStyle.OutLineOnly"
280         g_RSLocalAppSettings.Fields("SettingValue1").Value = .OutLineOnly
282         g_RSLocalAppSettings.UpDate
        
284         PrepareLocalAppSettingValue "oSelectionStyle.Transparency"
286         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Transparency
288         g_RSLocalAppSettings.UpDate
        
290         PrepareLocalAppSettingValue "oSelectionStyle.Width"
292         g_RSLocalAppSettings.Fields("SettingValue1").Value = .Width
294         g_RSLocalAppSettings.UpDate
        End With

296     With g_ZoomToSettings
298         PrepareLocalAppSettingValue "g_ZoomToSettings.SaveOnExit"
300         g_RSLocalAppSettings.Fields("SettingValue1").Value = .SaveOnExit
302         g_RSLocalAppSettings.UpDate
304         PrepareLocalAppSettingValue "g_ZoomToSettings.UseMultiple"
306         g_RSLocalAppSettings.Fields("SettingValue1").Value = .UseMultiple
308         g_RSLocalAppSettings.UpDate
        End With

310     PrepareLocalAppSettingValue "oLocatorSettings.Level1"
312     g_RSLocalAppSettings.Fields("SettingValue1").Value = IIf(Len(oLocatorSettings.Level1) > 0, oLocatorSettings.Level1, " ")
314     g_RSLocalAppSettings.UpDate
        'GisUtils.GIS_RELATE_INTERSECT_INTERIOR_INTERIOR
    
316     PrepareLocalAppSettingValue "oLocatorSettings.Level2"
318     g_RSLocalAppSettings.Fields("SettingValue1").Value = IIf(Len(oLocatorSettings.Level2) > 0, oLocatorSettings.Level2, " ")
320     g_RSLocalAppSettings.UpDate
     
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
        Dim rsSynchDS As New adodb.Recordset
        On Error GoTo Hell
    
100     rsSynchDS.CursorLocation = adUseClient
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
    '113     LoadOASISIncidents


104     If Not m_oDrawLyr Is Nothing Then
106         GIS.Add m_oDrawLyr
        End If
        
        If Not m_oSpatialiseLayer Is Nothing Then
             GIS.Add m_oSpatialiseLayer
        End If
        
        
        
108     If Not m_oBufferLyr Is Nothing Then
110         GIS.Add m_oBufferLyr
        End If
        
112     'If Not m_oW3Lyr Is Nothing Then
114     '    GIS.Add m_oW3Lyr
        'End If
        
116     'If Not m_oW3WHOLyr Is Nothing Then
118     '    GIS.Add m_oW3WHOLyr
        'End If
        
120     SafeMoveFirst g_RSAppSettings
122     g_RSAppSettings.Find "SettingName = 'EventLayerName'"

124     Set EventLayer = GIS.get(g_RSAppSettings.Fields.Item("SettingValue1").Value) '("Provinces_region")
        'GIS.CachedPaint = False

        '<EhFooter>
        Exit Sub

AddAllAdditionalStandardLayers_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.AddAllAdditionalStandardLayers " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub InitMap(Optional sMapPath As String)
        '<EhHeader>
        On Error GoTo InitMap_Err
        '</EhHeader>

        'Picture1.ZOrder
        ' and make it transparent
        'TrimPicture Picture1, &HFFFFFF
        
        If Len(sMapPath) < 5 Then
            If LoadProjectFileFromDB Then
                m_sInitialMapName = g_sAppPath & "\data\user\Exports\tempproj.ttkgp"
            End If
        End If
        
100     GIS.Lock
    
102     m_bLOADING = True

104     m_bMapInitialized = True
    
106     PrepareOverViewMap IIf(Len(sMapPath) > 0, sMapPath, m_sInitialMapName)
    
108     LoadMapProducts IIf(Len(sMapPath) > 0, sMapPath, m_sInitialMapName)
        
110     LoadAvailableThematics

        AddSQLLayersToProject
112     AddAllAdditionalStandardLayers
        
114     ReDim m_oSQLGenericLyrs(0)
        
116     SafeMoveFirst g_RSAppSettings
118     g_RSAppSettings.Find "SettingName = 'W3Settings'"
        
        'Temporary removed
120     'If g_RSAppSettings.Fields.Item("SettingValue1").Value = "1" Then CreateW3Layers
        
122     'If g_RSAppSettings.Fields.Item("SettingValue2").Value = "1" Then CreateW3WhoLayers

        'LoadLayerAttrDataToGrid m_oIncidentLyr
124     FillCOPValues

126     Set g_PrevExt = GIS.Extent
    
        Dim RS As New adodb.Recordset
    
128     Set RS = m_Cnn.Execute("SELECT DefaultViewName,DefaultViewX,DefaultViewY,DefaultViewZ,LatestViewX,LatestViewY,LatestViewZ,LatestMapName FROM Personnell WHERE Personnell_ID = " & g_CurrentUserID)
        
130     If Not RS.BOF Then SafeMoveFirst RS
        
132     m_DefaultViewName = RS.Fields.Item("DefaultViewName").Value
134     m_DefaultViewx = RS.Fields.Item("DefaultViewX").Value
136     m_DefaultViewy = RS.Fields.Item("DefaultViewY").Value
138     m_DefaultViewz = RS.Fields.Item("DefaultViewZ").Value

        '    If Not RS.Fields.Item("LatestViewX").value = vbNull Then
        '        If MsgBox("Do You want to start where you ended your last session?", vbYesNo, "OASIS COP") = vbYes Then
        '            GIS.Viewer.SetViewport RS.Fields.Item("LatestViewX").value, RS.Fields.Item("LatestViewY").value
        '            GIS.Viewer.Zoom = RS.Fields.Item("LatestViewZ").value
        '            GIS.Update
        '        Else
        '          '  GIS.Viewer.SetViewport m_DefaultViewx, m_DefaultViewy
        '          '  GIS.Viewer.Zoom = m_DefaultViewz
        '          '  GIS.Update
        '
        '        End If
        '
        '    Else
        '        'Map1.ZoomTo m_DefaultViewz, m_DefaultViewx, m_DefaultViewy
        '    End If
    
140     ComThemes.ListIndex = 0
142     C1TabFastFunction.Width = 340

144     If GIS.Items.Count > 0 Then
146         SafeMoveFirst g_RSAppSettings
148         g_RSAppSettings.Find "SettingName = 'MADataSetFiles'"
    
            Dim sLayers() As String
            Dim i As Integer
            Dim oLyr As XGIS_LayerAbstract
            
150         sLayers = Split(g_RSAppSettings.Fields.Item("SettingValue1").Value, ",")
    
152         For i = 0 To UBound(sLayers)
154             Set oLyr = GIS.get(sLayers(i))

156             If Not oLyr Is Nothing Then oLyr.HideFromLegend = True
            Next
    
158         SafeMoveFirst g_RSAppSettings
160         g_RSAppSettings.Find "SettingName = 'HiddenLayers'"
    
            If Not IsNull(g_RSAppSettings.Fields.Item("SettingValue1").Value) Then
    
162             sLayers = Split(g_RSAppSettings.Fields.Item("SettingValue1").Value, ",")
        
164             For i = 0 To UBound(sLayers)
166                 Set oLyr = GIS.get(sLayers(i))
    
168                 If Not oLyr Is Nothing Then oLyr.HideFromLegend = True
                Next
            
            End If

        End If
        
170     m_frmMnuOperations.Init GIS.viewer, m_Cnn
        
        'AddSQLLayersToProject
172     AB.RecalcLayout
174     m_bLOADING = False
        udRotationAngle.Value = 0
        GIS.RotationPoint = GisUtils.GisCenterPoint(GIS.Extent)
        
        If Not g_DatabaseSpecExtent Is Nothing Then
            GIS.VisibleExtent = g_DatabaseSpecExtent
            Set g_DatabaseSpecExtent = Nothing
        End If
176     GIS.Unlock
        
        SetDefCoordSys
        
        ReDim m_PrevExt(0)
        Set m_PrevExt(0) = GIS.VisibleExtent

        InitSelector

        '<EhFooter>
        Exit Sub

InitMap_Err:
        MsgBox Err.Description '& vbCrLf & _
        '       "in OASISClient.frmMain.InitMap " & _
        '       "at line " & Erl
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
 
        '        For i = 0 To GIS.Items.Count - 1
        '            If GisUtils.IsInherited(GIS.Items.Item(i), "XGIS_LayerVector") Then
        '                cmbSnap.AddItem GIS.Items.Item(i).Name
        '                comActiveLyr.AddItem GIS.Items.Item(i).Name
        '                'm_frmDebug.DebugPrint  GIS.Items.Item(i).Activated
        '            End If
        '        Next
        
        'If Len(g_RSGISGridTableSettings.Filter) > 0 Then g_RSGISGridTableSettings.Filter = ""
        
        If Not g_RSGISGridTableSettings.BOF Then SafeMoveFirst g_RSGISGridTableSettings
        
        abGridTools.Tools.Item("comLyr").CBAddItem "---Nothing---"
        
        If Not g_RSGISGridTableSettings.BOF And Not g_RSGISGridTableSettings.EOF Then
112         'g_RSGISGridTableSettings.MoveFirst
        
114         Do While Not g_RSGISGridTableSettings.EOF
            
116             If Not GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value) Is Nothing Then
118                 If Not g_RSGISGridTableSettings.Fields.Item("alias").Value = vbNull Then
120                     abGridTools.Tools.Item("comLyr").CBAddItem g_RSGISGridTableSettings.Fields.Item("alias").Value
                    End If
                End If
            
122             g_RSGISGridTableSettings.MoveNext
            Loop

        End If
        
124     If m_oColUserLayers.Count > 0 Then

126         For i = 1 To m_oColUserLayers.Count
128             abGridTools.Tools.Item("comLyr").CBAddItem m_oColUserLayers.Item(i)
            Next

        End If

        'OLD Method
        '102     For i = 0 To GIS.Items.Count - 1
        '
        '104         If GisUtils.IsInherited(GIS.Items.Item(i), "XGIS_LayerVector") Then
        '106             abGridTools.Tools.Item("comLyr").CBAddItem GIS.Items.Item(i).Name
        '                cmbSnap.AddItem GIS.Items.Item(i).Name
        '                comActiveLyr.AddItem GIS.Items.Item(i).Name
        ''                m_frmDebug.DebugPrint  GIS.Items.Item(i).Activated
        '            End If
        '
        '108         If GisUtils.IsInherited(GIS.Items.Item(i), "XGIS_LayerPixel") Then
        '                'do something
        '            End If
        '
        '        Next

130     If abGridTools.Tools.Item("comLyr").CBListCount > 0 Then abGridTools.Tools.Item("comLyr").CBListIndex = 0
            
132     If cmbSnap.ListCount > 0 Then
134         cmbSnap.ListIndex = 0
136         comActiveLyr.ListIndex = 0
        End If

        'abGridTools.Tools.Item("comLyr").CBList.NewIndex = 2
        'End If
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
        
       ' g_RSAppSettings.Open "SELECT * FROM AppSettings", m_Cnn, adOpenDynamic, adLockReadOnly

102     g_RSStyles.Open "SELECT * FROM Style", m_Cnn, adOpenDynamic, adLockReadOnly

104     g_RSRender.Open "SELECT * FROM Render", m_Cnn, adOpenDynamic, adLockReadOnly
  
    '        'SetRightPanel
    '        Dim sFnt As New StdFont
    '104     sFnt.Name = "Verdana"
    '106     sFnt.Size = 11
    '108     sFnt.Italic = True
    '
    '110     With vbalExplorerBarCtl2
    '112         .Redraw = False
    '            '.UseExplorerStyle = False
    '
    '            'OASIS MAIN MODULES
    '114         .ImageList = vbalImageListOasisMain.hIml
    '116         .BarTitleImageList = ilsTitleIcons.hIml
    '
    '118         Set cBar = .Bars.Add(, "SPECIAL", "&OASIS MODULES")
    '120         cBar.IsSpecial = True
    '122         cBar.ToolTipText = "Recent & Urgent Security Updates & Events."
    '124         cBar.IconIndex = 6
    '
    '126         Set sFnt = New StdFont
    '128         sFnt.Name = "Arial"
    '130         sFnt.Size = 14
    '132         sFnt.Italic = True
    '134         sFnt.Bold = True
    '
    '136         g_RSAppSettings.MoveFirst
    '138         g_RSAppSettings.Find "SettingName = 'HomeTitle'"
    '            '
    '140         Set cItem = cBar.Items.Add(, "home", g_RSAppSettings.Fields.Item("SettingValue1").value, 13)
    '
    '142         g_RSAppSettings.MoveFirst
    '144         g_RSAppSettings.Find "SettingName = 'HomeTitleToolTipText'"
    '
    '146         cItem.ToolTipText = g_RSAppSettings.Fields.Item("SettingValue1").value
    '148         Set cItem.Font = sFnt
    '
    '150         g_RSAppSettings.MoveFirst
    '152         g_RSAppSettings.Find "SettingName = 'HomeSubTitle'"
    '
    '154         Set cItem = cBar.Items.Add(, "sdfhjfsjfhdsfhjhf", g_RSAppSettings.Fields.Item("SettingValue1").value, , eItemText)
    '156         cItem.ToolTipText = "Takes you to VVAF's Home page"
    '
    '158         Set cItem = cBar.Items.Add(, "COP", "OASIS COP - Common Operations Picture", 10)
    '160         cItem.ToolTipText = "Activates the OASIS COP module."
    '162         Set cItem.Font = sFnt
    '164         Set cItem = cBar.Items.Add(, "SITwDESCkhkkh", "OASIS COP module, tools for Common operations security analysis and reports", , eItemText)
    '166         cItem.ToolTipText = "OASIS COP module, tools for Common operations security analysis and reports"
    '
    '168         If Not g_CurrentUserID = 3 Then
    '170             If Not g_CurrentUserID = 2 Then
    '172                 Set cItem = cBar.Items.Add(, "URGEwNCIES", "OASIS Weather Satellite Viewer", 5)
    '174                 cItem.ToolTipText = "Activates OASIS Weather Satellite Viewer."
    '176                 Set cItem.Font = sFnt
    '178                 Set cItem = cBar.Items.Add(, "SITwDESC", "VVAFs weather satellite imagery viewer", , eItemText)
    '180                 cItem.ToolTipText = "Displays OASIS Weather Satellite Viewer."
    '                End If
    '            End If
    '
    '182         Set cItem = cBar.Items.Add(, "CNN", "OASIS Dynamic Content Update", 11)
    '184         cItem.ToolTipText = "Activates the dynamic content modules."
    '186         Set cItem.Font = sFnt
    '188         Set cItem = cBar.Items.Add(, "INCIDENTSDEkjSC1", "Latest feeds from OASIS server", , eItemText)
    '190         cItem.ToolTipText = "Displays the latest feeds from OASIS server"
    '
    '192         If g_CurrentUserID = 2 Then
    '                'EXTENDED MODULES
    '194             Set cBar = .Bars.Add(, "SPECIALEXT", "&GENERAL MODULES")
    '196             cBar.IsSpecial = True
    '198             cBar.ToolTipText = "General Custom Modules"
    '200             cBar.IconIndex = 6
    '202             cBar.State = eBarCollapsed
    '
    '204             Set sFnt = New StdFont
    '206             sFnt.Name = "Arial"
    '208             sFnt.Size = 14
    '210             sFnt.Italic = True
    '212             sFnt.Bold = True
    '214             Set cItem = cBar.Items.Add(, "INCIDzExxNkojkTS", "Files And Archives", 11)
    '216             cItem.ToolTipText = "Explore Files and folders."
    '218             Set cItem.Font = sFnt
    '220             Set cItem = cBar.Items.Add(, "INCIDENTzxzxSDESC1", "Document library and archive", , eItemText)
    '222             cItem.ToolTipText = "Document Library And Archive."
    '
    '224             Set cItem = cBar.Items.Add(, "INCIDzEzxcNTS", "HAOST", 11)
    '226             cItem.ToolTipText = "JHAOST"
    '228             Set cItem.Font = sFnt
    '230             Set cItem = cBar.Items.Add(, "INCIDENTSDzxzzESC1", "Hazardous and abandoned ordnance survey tool", , eItemText)
    '232             cItem.ToolTipText = "HAOST"
    '
    '234             Set cItem = cBar.Items.Add(, "COMMUNICAzxwTION", "OPEX", 6)
    '236             cItem.ToolTipText = "OPEX"
    '238             Set cItem.Font = sFnt
    '240             Set cItem = cBar.Items.Add(, "IPmGDwEasdSC", "Mine Action Operations Explorer", , eItemText)
    '242             cItem.ToolTipText = "OPEX"
    '
    '244             Set cItem = cBar.Items.Add(, "LOGIzSTICS1", "HIS", 12)
    '246             cItem.ToolTipText = "Health Information System"
    '248             Set cItem.Font = sFnt
    '250             Set cItem = cBar.Items.Add(, "FISzDESC", "Health Information System", , eItemText)
    '252             cItem.ToolTipText = "Health Information System"
    '            End If
    '
    '254         .Redraw = True
    '        End With
    '
    '        'SEARCH BAR
    '
    '256     With vbalExplorerBarCtlSearch
    '258         .Redraw = False
    '
    '            '            '.UseExplorerStyle = False
    '260         If g_CurrentUserID = 2 Then
    '
    '                'OASIS MAIN MODULES
    '
    '267             .ImageList = vbalImageListOasisMain.hIml
    '269             .BarTitleImageList = ilsTitleIcons.hIml
    '
    '270             Set cBar = .Bars.Add(, "SPECIAL", "OASIS Wizards & Tools")
    '272             cBar.IsSpecial = True
    '274             cBar.ToolTipText = "Activate OASIS tools & wizards."
    '276             cBar.CanExpand = False
    '
    '278             Set cItem = cBar.Items.Add(, "LOGISklTICS1kh", "", , eItemControlPlaceHolder)
    '280             pnlSubMenu.Visible = True 'pnlSearch.Visible = True
    '282             cItem.Control = pnlSubMenu 'pnlSearch
    '284             .Visible = True
    '                'C1Navigator.Grid(gsRowHeight, 1) = (3600)
    '286             C1Navigator.Grid(gsRowHeight, 1) = (3600)
    '
    '262         ElseIf g_CurrentUserID = 1 Then
    '                '                'OASIS MAIN MODULES
    '                '290             .ImageList = vbalImageListOasisMain.hIml
    '                '292             .BarTitleImageList = ilsTitleIcons.hIml
    '                '
    '                '294             Set cBar = .Bars.Add(, "SPECIAL", "OASIS Time Utility")
    '                '296             cBar.IsSpecial = True
    '                '298             cBar.ToolTipText = "Recent & Urgent Updates & Events."
    '                '300             cBar.CanExpand = False
    '                '
    '                '302             Set cItem = cBar.Items.Add(, "LOGISTICS1kh", "", , eItemControlPlaceHolder)
    '                '304             pnlTime.Visible = True
    '                '306             cItem.Control = pnlTime
    '                '308             .Visible = True
    '                '                'C1Navigator.Grid(gsRowHeight, 1) = (3600)
    '                '310             C1Navigator.Grid(gsRowHeight, 1) = (3600)
    '                '312         ElseIf g_CurrentUserID = 3 Then
    '                '                'OASIS MAIN MODULES
    '                '314             .ImageList = vbalImageListOasisMain.hIml
    '                '316             .BarTitleImageList = ilsTitleIcons.hIml
    '                '
    '                '318             Set cBar = .Bars.Add(, "SPECIAL", "OASIS Todays Events")
    '                '320             cBar.IsSpecial = True
    '                '322             cBar.ToolTipText = "Recent & Urgent Updates & Events."
    '                '324             cBar.CanExpand = False
    '                '
    '                '326             Set cItem = cBar.Items.Add(, "LOGISTICS1kh", "", , eItemControlPlaceHolder)
    '                '328             pnlTodaysEvents.Visible = True
    '                '330             cItem.Control = pnlTodaysEvents
    '                '332             .Visible = True
    '                '                'C1Navigator.Grid(gsRowHeight, 1) = (3600)
    '                '334             C1Navigator.Grid(gsRowHeight, 1) = (3600)
    '            Else
    '264             .Visible = False
    '
    '            End If
    '
    '266         .Redraw = True
    '        End With
    '
    '        '        '  SetParent C1Browser.hWnd, C1AppHolder.hWnd
    '        '340     C1Init.Visible = True
    '        '342     SetParent C1Init.hWnd, C1AppHolder.hWnd
    '        '344     OASISInit1.Init
    '        '        'SetParent C1Map.hwnd, C1AppHolder.hwnd
    '        '346     OasisSubMenu1.SetBackColor cBar.BackColor

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
        Dim llm As XGIS_LayerSHP

            Exit Sub

        If sMapProj = "" Then
            ' add layers
100         Set llm = New XGIS_LayerSHP  'minimap states
102         llm.Path = GisUtils.GisSamplesDataDir + "states.shp"
104         llm.Name = "states"
106         llm.UseConfig = False
108         llm.Params.area.Color = RGB(255, 255, 255)
110         llm.Params.area.OutlineColor = RGB(&HC0&, &HC0&, &HC0&)
        
112         m_frmOvMap.GISm.Add llm  'add to minimap
        Else
            m_frmOvMap.GISm.Open sMapProj, False
        End If
        
        If Not m_frmOvMap.Visible Then
            m_frmOvMap.Show vbModeless, Me
        End If
        
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
    
    '100     With abCOP
    '
    '102         With .Bands.Item("bnMainMap")
    '
    '104          .Tools("axGIS").Custom = GIS 'm_frmOvMap
    '
    '            End With
    '
    '106         With .Bands.Item("bnOVmap")
    '
    '108          .Tools("frmOVMap1").Custom = m_frmOvMap
    '
    '            End With
    '
    '110          With .Bands.Item("bnMapLegend")
    '
    '112          .Tools("frmLegend").Custom = frmLegend
    '114          frmLegend.Init GIS.Viewer
    '            End With
    '
    '111          With .Bands.Item("bnMapAttributes")
    '
    '115          .Tools("frmAttributes").Custom = frmAttributes
    '
    '            End With
    '
    '116         .RecalcLayout
    '        End With
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
        ReDim m_shpProps(0)
        ReDim gdpts(0)
        ReDim ptsSel(0)
100     ReDim m_PrevExt(0)
102     Set m_PrevExt(0) = GIS.VisibleExtent

        If g_bOnlineCheckedAtLogin Then
            SetIMess "Online"
        Else
            SetIMess "Offline"
        End If

104     Set g_clsHotKey = New cRegHotKey
        
105        Set oClientInterCom = New IClient
106   oClientInterCom.RegisterDataChannel 1, Me.hWnd
        
107     Set oSync = New SynchWorker
108     Set m_InternetCheck = New SynchWorker
109        Set m_oSQLLyrSynch = New SynchWorker
        
110     Set g_ColAlerts = New Collection
        
112     Set ClipBoard_xt = New cCustomClipboard
114     Set m_oClipboardViewer = New cClipboardViewer
116     m_oClipboardViewer.InitClipboardChangeNotification Me.hWnd

118     g_clsHotKey.Attach Me.hWnd
120     g_clsHotKey.RegisterKey "Language", vbKeyL, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "OPSV", vbKeyP, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "doPrint", vbKeyP, MOD_CONTROL '
122     g_clsHotKey.RegisterKey "Admin", vbKeyA, MOD_ALT + MOD_CONTROL
124     g_clsHotKey.RegisterKey "Dude", vbKeyD, MOD_ALT + MOD_CONTROL
126     g_clsHotKey.RegisterKey "Script1", vbKey1, MOD_ALT + MOD_CONTROL
128     g_clsHotKey.RegisterKey "Script2", vbKey2, MOD_ALT + MOD_CONTROL
130     g_clsHotKey.RegisterKey "Script3", vbKey3, MOD_ALT + MOD_CONTROL
132     g_clsHotKey.RegisterKey "Script4", vbKey4, MOD_ALT + MOD_CONTROL
134     g_clsHotKey.RegisterKey "AppState", vbKeyS, MOD_ALT + MOD_CONTROL
136     g_clsHotKey.RegisterKey "Folders", vbKeyF, MOD_ALT + MOD_CONTROL
138     g_clsHotKey.RegisterKey "GPS", vbKeyG, MOD_ALT + MOD_CONTROL
140     g_clsHotKey.RegisterKey "OVB", vbKeyO, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "CompInfo", vbKeyC, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "Ticker", vbKeyQ, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "ForceIni", vbKeyI, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "Berserk", vbKeyB, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "Tracker", vbKeyX, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "LyrSearch", vbKeyZ, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "ResourcesFinder", vbKeyR, MOD_ALT + MOD_CONTROL
        g_clsHotKey.RegisterKey "WVISitRep", vbKeyW, MOD_ALT + MOD_CONTROL


        'Instantiate a reference to the multithreader library

142     Set pUpdateCheckerThread = New Thread
144     Set pLoadLyrAttrToGrdThread = New Thread
146     Set pSubmitGeoMarksThread = New Thread
148     Set pInitThread = New Thread
150     Set pThread = New Thread
152     Set FormThread = New Thread
154     Set pSynchThread = New Thread
156     Set pCheckSynchThread = New Thread
158     Set incThread = New Thread
160     Set pGetIncThread = New Thread
        Set pLoadMap = New Thread
        Set pCheckInet = New Thread
    
        
162     Me.caption = "OASIS Client Version: " & App.major & "." & App.minor
    
        'On Error Resume Next
164     SSC.Language = "VBScript"
166     SSC.AllowUI = True
 
168     With SSC
170         .AddObject "OASISGis", GIS, True
172         .AddObject "OASISGisUtils", GisUtils, True
174         .AddObject "RSAppSettings", g_RSAppSettings, True
176         .AddObject "RSGISGridTableSettings", g_RSGISGridTableSettings, True
178         .AddObject "CN", m_Cnn, True
180         .AddObject "OASISGISDataGrid", dxGISDataGrid, True
182         .AddObject "OASISToolBar", AB, True
184         .AddObject "OASISCharting", frmSecChart, True
186         .AddObject "OASISMum", Me, True
        End With
                
188     DE9IM.Add "Contains", "T*****FF"
190     DE9IM.Add "Cross", "T*T "
192     DE9IM.Add "Cross Line", "0 "
194     DE9IM.Add "Disjoint", "FF*FF"
196     DE9IM.Add "Equality", "T*F**FF*"
198     DE9IM.Add "Intersect", "T "
200     DE9IM.Add "Intersect Boundary Boundary", "****T "
202     DE9IM.Add "Intersect Boundary Interior", "***T "
204     DE9IM.Add "Intersect Interior Boundary", "*T "
206     DE9IM.Add "Intersect Interior Interior", "T"
208     DE9IM.Add "Intersect1", "*T"
210     DE9IM.Add "Intersect2", "***T"
212     DE9IM.Add "Intersect3", "****T"
214     DE9IM.Add "Line Cross Line", "0"
216     DE9IM.Add "Line Cross Polygon", "T*T"
218     DE9IM.Add "Line Travers Polygon", "T**F"
220     DE9IM.Add "Overlap", "T*T***T"
222     DE9IM.Add "Overlap Line", "1*T***T"
224     DE9IM.Add "Polygon Crossed By Line", "T*****T"
226     DE9IM.Add "Polygon CrossTraversed By Line", "TF**F*T"
228     DE9IM.Add "Polygon Traversed By Line", "TF"
230     DE9IM.Add "Touch", "F***T "
232     DE9IM.Add "Touch Boundary Boundary", "F***T"
234     DE9IM.Add "Touch Boundary Interior", "F**T"
236     DE9IM.Add "Touch Interior", "F**T "
238     DE9IM.Add "Touch Interior Boundary", "FT"
240     DE9IM.Add "Within", "T*F**F"
          
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.Form_Load " & "at line " & Erl
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
        
106             sSuffix = UCase$(Right$(oFile.Name, 3))
        
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

Private Sub m_frmW3Wizard_MyActivities()
        '<EhHeader>
        On Error GoTo m_frmW3Wizard_MyActivities_Err
        '</EhHeader>
100     ctrWhat1.Visible = True
102     ctrWhere1.Visible = False
104     strWho1.Visible = False
106     lblLblW3.caption = "WHAT?"
        '<EhFooter>
        Exit Sub

m_frmW3Wizard_MyActivities_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmW3Wizard_MyActivities " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmW3Wizard_MyLocations()
        '<EhHeader>
        On Error GoTo m_frmW3Wizard_MyLocations_Err
        '</EhHeader>
100     ctrWhat1.Visible = False
102     ctrWhere1.Visible = True
104     strWho1.Visible = False
106     lblLblW3.caption = "WHERE?"
        '<EhFooter>
        Exit Sub

m_frmW3Wizard_MyLocations_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmW3Wizard_MyLocations " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmW3Wizard_MyOrganization()
        '<EhHeader>
        On Error GoTo m_frmW3Wizard_MyOrganization_Err
        '</EhHeader>
100     ctrWhat1.Visible = False
102     ctrWhere1.Visible = False
104     strWho1.Visible = True
106     lblLblW3.caption = "WHO?"
        '<EhFooter>
        Exit Sub

m_frmW3Wizard_MyOrganization_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.m_frmW3Wizard_MyOrganization " & _
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
100     m_frmDebug.DebugPrint sID
    
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
        '<EhHeader>
        On Error GoTo mnuFlash_Click_Err
        '</EhHeader>
        Dim i As Integer
        Dim xx As Double
        Dim yy As Double
        Dim sSQL As String
    
        Dim lLayer As XGIS_LayerVector
        Dim lShape As XGIS_Shape
        Dim lExtent As XGIS_Extent

100     SafeMoveFirst g_RSGISGridTableSettings
102     g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"

104     If Not g_RSGISGridTableSettings.EOF Then
106         Set lLayer = GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
        Else
108         Set lLayer = GIS.get(abGridTools.Tools.Item("comLyr").Text)
        End If

110     With dxGISDataGrid.ex
        
112         lLayer.Lock

114         For i = 0 To .SelectedCount - 1

116             Set lExtent = New XGIS_Extent
118             xx = GlobalGISGridLayer.GetShape(.SelectedNodes(i).values(0)).PointOnShape.x
120             yy = GlobalGISGridLayer.GetShape(.SelectedNodes(i).values(0)).PointOnShape.y
122             lExtent.Prepare xx, yy, xx, yy

124             If GISGridRS.Fields.Count > 1 Then
                        
126                 'If GISGridRS.Fields(1).Type = adNumeric Then
                    If dxGISDataGrid.Columns(1).FieldType = xftDecimal Or dxGISDataGrid.Columns(1).FieldType = xftFloat Or dxGISDataGrid.Columns(1).FieldType = xftInteger Then

                        'sSQL = GISGridRS.Fields(1).Name & " = " & .SelectedNodes(i).values(1)
128                     sSQL = dxGISDataGrid.Columns(1).FieldName & " = " & .SelectedNodes(i).values(1)
                    Else
                        'sSQL = GISGridRS.Fields(1).Name & " = '" & .SelectedNodes(i).values(1) & "'"
130                     sSQL = dxGISDataGrid.Columns(1).FieldName & " = '" & .SelectedNodes(i).values(1) & "'"
                    End If
                    
                    
                    'override for flash issue on 16-sept-10
                    If Not lLayer.FindFirst(lExtent, "", Nothing, "", True) Is Nothing Then
                        sSQL = "GIS_UID = " & lLayer.FindFirst(lExtent, "", Nothing, "", True).uID
                    End If
                
                Else
132                 sSQL = ""
                End If
                
                
134             Set lShape = lLayer.FindFirst(lLayer.Extent, sSQL, lShape, "", True)

136             If Not IsNull(lShape) And Not lShape Is Nothing Then
138                 lShape.Flash
140                 lShape.Invalidate
                End If
                
142             Set lShape = Nothing
                
            Next

144         lLayer.Unlock
        
        End With



        '<EhFooter>
        Exit Sub

mnuFlash_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.mnuFlash_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuSelectInMap_Click()
        '<EhHeader>
        On Error GoTo mnuSelectInMap_Click_Err
        '</EhHeader>

        Dim i As Integer
        Dim xx As Double
        Dim yy As Double
        Dim sSQL As String
    
        Dim lLayer As XGIS_LayerVector
        Dim lShape As XGIS_Shape
        Dim lExtent As XGIS_Extent

100     SafeMoveFirst g_RSGISGridTableSettings
102     g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"

104     If Not g_RSGISGridTableSettings.EOF Then
106         Set lLayer = GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
        Else
108         Set lLayer = GIS.get(abGridTools.Tools.Item("comLyr").Text)
        End If

110     With dxGISDataGrid.ex
        
112         lLayer.Lock

114         For i = 0 To .SelectedCount - 1

116             Set lExtent = New XGIS_Extent
118             xx = GlobalGISGridLayer.GetShape(.SelectedNodes(i).values(0)).PointOnShape.x
120             yy = GlobalGISGridLayer.GetShape(.SelectedNodes(i).values(0)).PointOnShape.y
122             lExtent.Prepare xx, yy, xx, yy

124             If GISGridRS.Fields.Count > 1 Then

126                 If GISGridRS.Fields(1).Type = adNumeric Then
128                     sSQL = GISGridRS.Fields(1).Name & " = " & .SelectedNodes(i).values(1)
                    Else
130                     sSQL = GISGridRS.Fields(1).Name & " = '" & .SelectedNodes(i).values(1) & "'"
                    End If
                
                Else
132                 sSQL = ""
                End If
                
134             Set lShape = lLayer.FindFirst(lExtent, sSQL, lShape, "", True)

136             If Not IsNull(lShape) And Not lShape Is Nothing Then

138                 Set lShape = lShape.MakeEditable
140                 lShape.IsSelected = True
142                 lShape.Invalidate

                End If
                
            Next

144         lLayer.Unlock
        
        End With
        
146     Set lLayer = Nothing
148     Set lShape = Nothing
150     Set lExtent = Nothing

        'GIS.UpDate

        '<EhFooter>
        Exit Sub

mnuSelectInMap_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.mnuSelectInMap_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuClearSelections_Click()
    
    Dim i As Integer
    Dim xx As Double
    Dim yy As Double
    Dim sSQL As String
    
    Dim lLayer As XGIS_LayerVector
    Dim lShape As XGIS_Shape
    Dim lExtent As XGIS_Extent

    SafeMoveFirst g_RSGISGridTableSettings
    g_RSGISGridTableSettings.Find "alias = '" & abGridTools.Tools.Item("comLyr").Text & "'"

    If Not g_RSGISGridTableSettings.EOF Then
        Set lLayer = GIS.get(g_RSGISGridTableSettings.Fields.Item("name").Value)
    Else
        Set lLayer = GIS.get(abGridTools.Tools.Item("comLyr").Text)
    End If
    
    lLayer.RevertAll
    GIS.UpDate
    Exit Sub

    With dxGISDataGrid.ex
        
        lLayer.Lock

        For i = 0 To .SelectedCount - 1

            Set lExtent = New XGIS_Extent
            xx = GlobalGISGridLayer.GetShape(.SelectedNodes(i).values(0)).PointOnShape.x
            yy = GlobalGISGridLayer.GetShape(.SelectedNodes(i).values(0)).PointOnShape.y
            lExtent.Prepare xx, yy, xx, yy

            If GISGridRS.Fields.Count > 1 Then

                If GISGridRS.Fields(1).Type = adNumeric Then
                    sSQL = GISGridRS.Fields(1).Name & " = " & .SelectedNodes(i).values(1)
                Else
                    sSQL = GISGridRS.Fields(1).Name & " = '" & .SelectedNodes(i).values(1) & "'"
                End If
                
            Else
                sSQL = ""
            End If
                
            Set lShape = lLayer.FindFirst(lExtent, sSQL, lShape, "", True)

            If Not IsNull(lShape) And Not lShape Is Nothing Then

                Set lShape = lShape.MakeEditable
                lShape.IsSelected = True
                lShape.Invalidate

            End If
                
        Next

        lLayer.Unlock
        
    End With
        
    Set lLayer = Nothing
    Set lShape = Nothing
    Set lExtent = Nothing

End Sub

Private Sub mnuZoomTo_Click()
        '<EhHeader>
        On Error GoTo mnuZoomTo_Click_Err
        '</EhHeader>
        Dim i As Integer
        Dim lShape As XGIS_Shape
        Dim lExtent As XGIS_Extent
        
100     With dxGISDataGrid.ex
        
102         Set lShape = GlobalGISGridLayer.GetShape(.SelectedNodes(i).values(0))
    
104         For i = 0 To .SelectedCount - 1
                
106             If lExtent Is Nothing Then
108                 Set lExtent = lShape.Extent
                End If
                
110             Set lExtent = GisUtils.GisMaxExtent(lExtent, lShape.Extent)
            Next
    
112         If .SelectedCount > 0 Then
114             GIS.Lock
116             GIS.VisibleExtent = lExtent
118             GIS.Unlock
            End If
    
        End With
        
120     Set lExtent = Nothing
122     Set lShape = Nothing

    
        '<EhFooter>
        Exit Sub

mnuZoomTo_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.mnuZoomTo_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
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
100     m_frmDebug.DebugPrint ctTreeBookmrks.NodeText(nIndex) & " Level:" & ctTreeBookmrks.NodeLevel(nIndex)
   
        Exit Sub
    
        Dim RS As New adodb.Recordset
   
102     If ctTreeBookmrks.NodeText(nIndex) = m_DefaultViewName Then
104         GIS.viewer.SetViewport m_DefaultViewx, m_DefaultViewy
106         GIS.viewer.zoom = m_DefaultViewz
108         GIS.UpDate
            'Map1.ZoomTo m_DefaultViewz, m_DefaultViewx, m_DefaultViewy
        Else

110         If ctTreeBookmrks.NodeLevel(nIndex) = 2 Then
112             RS.Open "SELECT X,Y,Z FROM GeoBookMarks WHERE Name ='" & ctTreeBookmrks.NodeText(nIndex) & "'", m_Cnn, adOpenDynamic, adLockReadOnly

114             If Not RS.BOF Then SafeMoveFirst RS
            
                Dim ptg As New XGIS_Point
            
            
                'Set ptg = gis.Center  .CreateShape(XgisShapeTypePoint)
            
116             ptg.Prepare CDbl(RS.Fields.Item("X")), CDbl(RS.Fields.Item("Y"))
            
                'GIS.Lock
                'GIS.CenterPtg.Prepare CDbl(RS.Fields.Item("X")), CDbl(RS.Fields.Item("Y"))
            
            
118             GIS.zoom = RS.Fields.Item("Z")
120             GIS.CenterViewport ptg
                'GIS.Unlock
122             GIS.UpDate
            
                'GIS.Viewer.SetViewport RS.Fields.Item("X"), RS.Fields.Item("Y")
                'GIS.Update
                'GIS.Viewer.Zoom = RS.Fields.Item("Z")
                'GIS.Update
                'Map1.ZoomTo RS.Fields.Item("Z"), RS.Fields.Item("X"), RS.Fields.Item("Y")
     
            End If
        End If

        'm_frmDebug.DebugPrint  ctTreeBookmrks.NodeLevel(nIndex)
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

100     FrmAddBookMark.Init m_Cnn, GIS.viewer.CenterPtg.x, GIS.viewer.CenterPtg.y, GIS.viewer.zoom, GIS.Name
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
138          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & IIf(.OptIncViolence(0).Value, "Yes", IIf(.OptIncViolence(2).Value, "Unknown", "No")) & "</strong></td></tr>"  'Violent"

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
162          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & IIf(.chkUnknown.Value = vbChecked, "Unknown", "Known") & "</strong></td></tr>" 'Cualties Known"

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
214          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & IIf(.chkSendMe.Value = vbChecked, "Yes", "No") & "</strong></td>" 'Verify When published"

216          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>Allow Other Users To Contact me:</strong></td>" 'Allow Contact"
218          sHTML = sHTML & vbCrLf & "<td width=""""25%""""><strong>" & IIf(.chkAllowOther.Value = vbChecked, "Yes", "No") & "</strong></td></tr>" 'Allow Contact"
    
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
  edRotationAngle.Text = udRotationAngle.Value
  GIS.RotationAngle = DegToRad(udRotationAngle.Value)
  GIS.UpDate
End Sub

Private Sub ListAllAnnotationShps()
        '<EhHeader>
        On Error GoTo ListAllShps_Err
        '</EhHeader>

        m_frmTextAnnoSettings.lstTexts.Clear

100     If m_oDrawLyr Is Nothing Then Exit Sub

102     With m_oDrawLyr
104         .MoveFirst .Extent, "", Nothing, "", True
        
108         Do While Not .EOF

                If Not g_lPinUID = .Shape.uID Then
110                 m_frmTextAnnoSettings.lstTexts.AddItem .Shape.Params.labels.Value
112                 m_frmTextAnnoSettings.lstTexts.ItemData(m_frmTextAnnoSettings.lstTexts.ListCount - 1) = .Shape.uID
                    
                    On Error Resume Next
                    m_frmTextAnnoSettings.panColorBack.BackColor = .Color
                    m_frmTextAnnoSettings.panColorFore.BackColor = .FontColor
                    On Error GoTo ListAllShps_Err
                End If

114             .MoveNext
            Loop
    
        End With
    
        If m_frmTextAnnoSettings.lstTexts.ListCount > 0 Then m_frmTextAnnoSettings.lstTexts.ListIndex = 0
    
        '<EhFooter>
        Exit Sub

ListAllShps_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.ListAllShps " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub RemoveAllShps(oLyr As XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo RemoveAllShps_Err
        '</EhHeader>
100     oLyr.MoveFirst oLyr.Extent, "", Nothing, "", False
    
102     Do While Not oLyr.EOF
            If Not g_lPinUID = oLyr.Shape.uID Then
104             oLyr.Delete oLyr.Shape.GetField("GIS_UID")
            End If
106         oLyr.MoveNext
        Loop
    
        '<EhFooter>
        Exit Sub

RemoveAllShps_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMain.RemoveAllShps " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub AddAnnoText(ptg As XGIS_Point)
        '<EhHeader>
        On Error GoTo AddAnnoText_Err
        '</EhHeader>
        Dim oShp As XGIS_Shape

        If Not m_frmTextAnnoSettings.chkMultipleText.Value = vbChecked Then
            RemoveAllShps m_oDrawLyr
           m_frmTextAnnoSettings.lstTexts.Clear
        End If
        
100     Set oShp = m_oDrawLyr.CreateShape(XgisShapeTypePoint)
                        
102     If oShp Is Nothing Then Exit Sub
                        
104     With oShp
106         .Lock XgisLockExtent
108         .AddPart
110         .AddPoint ptg

            With .Params.labels
                .rotate = m_frmTextAnnoSettings.txtRotation.Text
                .Alignment = XgisLabelAlignmentCenter
                .FontColor = m_frmTextAnnoSettings.panColorFore.BackColor
                .Color = m_frmTextAnnoSettings.panColorBack.BackColor
112             .Allocator = False
114             .Duplicates = True
                
                If IsNumeric(m_frmTextAnnoSettings.comFontSize.Text) Then
116                 .Font.Size = m_frmTextAnnoSettings.comFontSize.Text
                Else
                    .Font.Size = 12
                End If
                
120             .OutlineWidth = 0
122             .Pattern = XbsClear
124             .Position = XgisLabelPositionMiddleCenter
126             .Value = m_frmTextAnnoSettings.txtAnnoText.Text
            End With
                
              m_frmTextAnnoSettings.lstTexts.AddItem m_frmTextAnnoSettings.txtAnnoText.Text
127           m_frmTextAnnoSettings.lstTexts.ItemData(m_frmTextAnnoSettings.lstTexts.ListCount - 1) = .uID
              m_frmTextAnnoSettings.lstTexts.ListIndex = m_frmTextAnnoSettings.lstTexts.ListCount - 1
                
128         .Params.Marker.Size = 1
            .Unlock
        End With

130     GIS.UpDate

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

'102     If chkAutomaticClear.Value = vbChecked Then ClearAll

        m_udtSelectorSettings.bAutoClear = True
104     m_udtSelectorSettings.dBuffeLevel = CDbl(1 / 2)
        
106     If ComSelLayer.ListCount > 0 Then
        
108         sCurrItem = ComSelLayer.List(ComSelLayer.ListIndex)
        
        End If
        
110     If ComFeatureLayer.ListCount > 0 Then
112         sCurFeatItm = ComFeatureLayer.List(ComFeatureLayer.ListIndex)
        End If
        
'114     txtSpatialOperation.Clear
'116     keyArray = DE9IM.Keys

'118     For Each element In keyArray
'            Debug.Print element
'120         txtSpatialOperation.AddItem element
'        Next
        m_udtSelectorSettings.sSpatialOperator = DE9IM.Item("Intersect")
'122     txtSpatialOperation.ListIndex = 5
        
124     ComSelLayer.Clear
126     ComFeatureLayer.Clear
        
128     For i = 0 To GIS.Items.Count - 1
            On Error Resume Next
130         If GisUtils.IsInherited(GIS.Items.Item(i), "XGIS_LayerVector") Then
132             m_SelLyrCol.Add GIS.Items.Item(i).Name, GIS.Items.Item(i).caption
134             ComSelLayer.AddItem GIS.Items.Item(i).caption 'Name
136             ComFeatureLayer.AddItem GIS.Items.Item(i).caption
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

'154     mnuLineSelect_Click

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
        Debug.Print Err.Description & vbCrLf & _
               "in OASISClient.frmMain.RemoveAlltabs " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetGISSelectionStyle()
'    GIS.SelectionColor = vbRed
'    GIS.SelectionPattern = XgisLockExtent
'    GIS.SelectionTransparency = 50
End Sub

Private Sub AddTabs(sCaption As String, _
                    Optional oShp As XGIS_Shape, _
                    Optional bShowGeo As Boolean = True)
        '<EhHeader>
        On Error GoTo AddTabs_Err
        '</EhHeader>
        
100     Set SelAttributes1(SelAttributes1.UBound).Container = elDynHolder(elDynHolder.UBound)
102     SelAttributes1(SelAttributes1.UBound).Visible = True
        SelAttributes1(SelAttributes1.UBound).ReadOnly = Not m_udtSelectorSettings.bEdit
    
104     If Not oShp Is Nothing Then
        
106         SelAttributes1(SelAttributes1.UBound).AllowRestructure = True
108         SelAttributes1(SelAttributes1.UBound).ShowShape oShp
110         GEO1.ShowSelected oShp.layer
    
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

    Dim prj As XGIS_ProjectionAbstract
    Dim scoords As New XGIS_Coordinate
    Dim ucoords As XGIS_Coordinate
    Dim ocoords As XGIS_Coordinate
    Dim xParts() As String
    Dim yParts() As String
    Dim sLongt As String
    Dim sLat As String
   
    If Not bIsUTM Then
        Set prj = ProjectionList.FindEx("UTM")

        With oCoordTransSettings
            prj.SetUp .Long_Of_Origin, .Lat_Of_Origin, .False_Easting, .False_Northing, 0, 0, .Zone, 0, 0, 0, 0, 0, 0, 0, 0
        
            txtUTMX.Text = x
            txtUTMY.Text = y
            Set ucoords = New XGIS_Coordinate
            ucoords.x = x
            ucoords.y = y
            
            Set ocoords = prj.Unprojected(ucoords)
            txtDDX.Text = RadToDeg(x)
            txtDDY.Text = RadToDeg(y)
        End With
        
    Else
        'txtUTMX.Text
    
    End If

'    scoords.Y = GisUtils.GisStrToLatitude("1 ° 17 ' 35 N")
'    scoords.X = GisUtils.GisStrToLongitude("103° 51' 21 E")
'
'    ' convert to UTM
'    Set ucoords = prj.Projected(scoords)
'
'    ucoords.X = X
'    ucoords.Y = Y
'
'    ' and back to geographic ccordinates – result is in radians!!!!
'    Set ocoords = prj.Unprojected(ucoords)
'
'    Dim lLat As Double
'    Dim llong As Double
'
'    UTM_2_LL llong, lLat, X, Y, ""
'
'    'Convert Radians To DD
'    'Debug.Print GisUtils.GisLongitudeToStr(ocoords.X) * 180 / Pi
'    'Debug.Print GisUtils.GisLatitudeToStr(ocoords.Y) * 180 / Pi
'
'    sLongt = GisUtils.GisLongitudeToStr(ocoords.X)
'    sLat = GisUtils.GisLatitudeToStr(ocoords.Y)
'
'    xParts = Split(sLongt, " ")
'    yParts = Split(sLat, " ")
'
'    'xpart(0) = Replace(Trim$(xpart(0)), "°", "")
'    'xpart(1) = Replace(Trim$(xpart(1)), "'", "")
'    'xpart(2) = Replace(Trim$(xpart(2)), """", "")
'
'    'yParts(0) = Replace(Trim$(yParts(0)), "°", "")
'    'yParts(1) = Replace(Trim$(yParts(1)), "'", "")
'    'yParts(2) = Replace(Trim$(yParts(2)), """", "")
'
'    txtDDX.Text = sLongt
'    txtDDY.Text = sLat
'
'    Debug.Print GisUtils.GisLongitudeToStr(ocoords.X) + " " + GisUtils.GisLatitudeToStr(ocoords.Y)
'    '104° 45' 19.57" E 41° 43' 51.76" N
'
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
    Dim pt As POINTAPI
    Dim mWnd As Long, WR As RECT

    picToolTip.Visible = False

    With oMapTipSetting

        If .Enabled Then

            'First Make Sure the PTTip is Initiated
            If ptTip.x + ptTip.y > 0 Then
                'Get Current Cursor Position to see if has changed since last time
                GetCursorPos pt
         
                If pt.x - ptTip.x = 0 And pt.y - ptTip.y = 0 Then
         
                    If Not m_oToolTipSHP Is Nothing Then
                        mWnd = WindowFromPoint(ptTip.x, ptTip.y)
                        'Get the window's position in Pixels
                        GetWindowRect mWnd, WR
                        
                        If .MapTipField = "UID" Then
                            lblToolTip.caption = m_oToolTipSHP.uID 'GetField("Name")
                        Else
                            lblToolTip.caption = m_oToolTipSHP.GetField(.MapTipField)
                        End If
                        
                        picToolTip.Move elMap.Left + ScaleX(ptTip.x - WR.Left + 7, vbPixels, vbTwips), ScaleY(ptTip.y - WR.Top - 17, vbPixels, vbTwips), lblToolTip.Width + 140, lblToolTip.Height + 80

                        'Debug.Print "Pic:" & ScaleX(WR.Left + ptTip.X - 2, vbPixels, vbTwips) & " Win:" & Me.Left & " Mouse:" & ptTip.X & " Rect:" & WR.Left
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


