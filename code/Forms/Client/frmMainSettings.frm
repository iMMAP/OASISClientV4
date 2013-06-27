VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TatukGIS_XDK10_WIN32.ocx"
Begin VB.Form frmMainSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "OASIS Settings"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmSubmit 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   104
      Top             =   3900
      Width           =   3735
      Begin VB.CommandButton cmdDoCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   107
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDoOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   1380
         TabIndex        =   106
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDoApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   2640
         TabIndex        =   105
         Top             =   0
         Width           =   1095
      End
   End
   Begin C1SizerLibCtl.C1Tab c1Settings 
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   3795
      _cx             =   6694
      _cy             =   7752
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
      Caption         =   "General|Security|Grid|Coordinate Catcher"
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
      Flags(1)        =   2
      Flags(2)        =   2
      Flags(3)        =   2
      Begin C1SizerLibCtl.C1Elastic elCoordCatcher 
         Height          =   4020
         Left            =   5040
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   330
         Width           =   3705
         _cx             =   6535
         _cy             =   7091
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
         Begin VB.ComboBox comZone 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   60
            Width           =   795
         End
         Begin VB.Frame fraEllipsoid 
            Caption         =   "Ellipsoid Parameters"
            Height          =   1395
            Left            =   60
            TabIndex        =   15
            Top             =   540
            Width           =   3555
            Begin VB.TextBox txtAxis 
               Height          =   285
               Left            =   1560
               TabIndex        =   18
               Text            =   "6378137"
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtFInv 
               Height          =   285
               Left            =   1560
               TabIndex        =   17
               Text            =   "298.2572236"
               Top             =   720
               Width           =   1575
            End
            Begin VB.CheckBox chkSphere 
               Caption         =   "Sphere"
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   1020
               Width           =   975
            End
            Begin VB.Label lblAxis 
               Caption         =   "Semi-major Axis"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblFlat 
               Caption         =   "Inverse Flattening"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   780
               Width           =   1395
            End
         End
         Begin VB.Frame fraProjection 
            Caption         =   "Projection Parameters"
            Height          =   1935
            Left            =   60
            TabIndex        =   6
            Top             =   1980
            Width           =   3555
            Begin VB.TextBox txtFE 
               Height          =   285
               Left            =   1740
               TabIndex        =   10
               Text            =   "0"
               Top             =   1140
               Width           =   1575
            End
            Begin VB.TextBox txtFN 
               Height          =   285
               Left            =   1740
               TabIndex        =   9
               Text            =   "0"
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtLatOrig 
               Height          =   285
               Left            =   1740
               TabIndex        =   8
               Text            =   "0"
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txtLonOrig 
               Height          =   285
               Left            =   1740
               TabIndex        =   7
               Text            =   "0"
               Top             =   1500
               Width           =   1575
            End
            Begin VB.Label lblFE 
               Caption         =   "False Easting"
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   1140
               Width           =   1155
            End
            Begin VB.Label lblFN 
               Caption         =   "False Northing"
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblLatOrig 
               Caption         =   "Latitude of Origin"
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label lblLonOrig 
               Caption         =   "Longitude of Origin"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   1500
               Width           =   1455
            End
         End
         Begin VB.Label lblUTMZone 
            Caption         =   "UTM Zone:"
            Height          =   255
            Left            =   60
            TabIndex        =   22
            Top             =   180
            Width           =   855
         End
      End
      Begin C1SizerLibCtl.C1Elastic elGrid 
         Height          =   4020
         Left            =   4740
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   330
         Width           =   3705
         _cx             =   6535
         _cy             =   7091
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
      End
      Begin C1SizerLibCtl.C1Elastic elSec 
         Height          =   4020
         Left            =   4440
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   330
         Width           =   3705
         _cx             =   6535
         _cy             =   7091
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
         _GridInfo       =   $"frmMainSettings.frx":0000
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   525
            Left            =   0
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   3495
            Width           =   3705
            _cx             =   6535
            _cy             =   926
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
            Begin VB.Frame FraActivate 
               Caption         =   "Activate:"
               Height          =   510
               Left            =   0
               TabIndex        =   61
               Top             =   0
               Width           =   2295
               Begin VB.CheckBox chkChkActivate 
                  Caption         =   "Type"
                  Height          =   285
                  Index           =   0
                  Left            =   60
                  TabIndex        =   64
                  Top             =   180
                  Value           =   1  'Checked
                  Width           =   705
               End
               Begin VB.CheckBox chkChkActivate 
                  Caption         =   "Target"
                  Height          =   285
                  Index           =   1
                  Left            =   780
                  TabIndex        =   63
                  Top             =   180
                  Value           =   1  'Checked
                  Width           =   780
               End
               Begin VB.CheckBox chkChkActivate 
                  Caption         =   "Time"
                  Height          =   285
                  Index           =   2
                  Left            =   1560
                  TabIndex        =   62
                  Top             =   180
                  Value           =   1  'Checked
                  Width           =   690
               End
            End
            Begin VB.CommandButton cmdApply 
               Caption         =   "Apply"
               Height          =   375
               Left            =   2400
               TabIndex        =   60
               Top             =   120
               Width           =   1305
            End
         End
         Begin C1SizerLibCtl.C1Tab c1Tab 
            Height          =   3495
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   3705
            _cx             =   6535
            _cy             =   6165
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
            Caption         =   "Type|Target|Time"
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
            Begin DXDBGRIDLibCtl.dxDBGrid dxTypeScoring 
               Height          =   3150
               Left            =   45
               OleObjectBlob   =   "frmMainSettings.frx":003E
               TabIndex        =   56
               Top             =   330
               Width           =   825
            End
            Begin DXDBGRIDLibCtl.dxDBGrid dxTargetScoring 
               Height          =   3150
               Left            =   1560
               OleObjectBlob   =   "frmMainSettings.frx":17A3
               TabIndex        =   57
               Top             =   330
               Width           =   825
            End
            Begin DXDBGRIDLibCtl.dxDBGrid dxTimeScoring 
               Height          =   3150
               Left            =   1860
               OleObjectBlob   =   "frmMainSettings.frx":2EF0
               TabIndex        =   58
               Top             =   330
               Width           =   825
            End
         End
         Begin VB.CommandButton cmdScoringSettings 
            Caption         =   "Scoring Settings"
            Height          =   3495
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   3705
         End
      End
      Begin C1SizerLibCtl.C1Elastic elGen 
         Height          =   4020
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   3705
         _cx             =   6535
         _cy             =   7091
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
         Begin C1SizerLibCtl.C1Tab C1TabSet 
            Height          =   3495
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   3675
            _cx             =   6482
            _cy             =   6165
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
            Caption         =   "Misc|Maptip|Selection Style|Locator|Map"
            Align           =   0
            CurrTab         =   4
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
            Flags(0)        =   2
            Flags(2)        =   2
            Flags(3)        =   2
            Begin C1SizerLibCtl.C1Elastic elMapSet 
               Height          =   3120
               Left            =   45
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   330
               Width           =   3585
               _cx             =   6324
               _cy             =   5503
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
               Begin VB.Frame FraWaterMark 
                  Caption         =   "Water mark:"
                  Height          =   975
                  Left            =   60
                  TabIndex        =   114
                  Top             =   2100
                  Width           =   3495
                  Begin TatukGIS_XDK10.XGIS_ControlNorthArrow WaterMark 
                     Height          =   795
                     Left            =   2400
                     TabIndex        =   116
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   975
                     Symbol          =   0
                     Transparent     =   0   'False
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
                  Begin VB.CheckBox chkShowWater 
                     Caption         =   "Show Water Mark:"
                     Height          =   495
                     Left            =   60
                     TabIndex        =   115
                     Top             =   180
                     Width           =   1695
                  End
               End
               Begin VB.TextBox edRotationAngle 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   1320
                  Locked          =   -1  'True
                  TabIndex        =   108
                  Text            =   "0"
                  Top             =   480
                  Width           =   435
               End
               Begin MSComCtl2.UpDown udRotationAngle 
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   109
                  Top             =   480
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   503
                  _Version        =   393216
                  Alignment       =   0
                  AutoBuddy       =   -1  'True
                  BuddyControl    =   "edRotationAngle"
                  BuddyDispid     =   196636
                  OrigLeft        =   1680
                  OrigRight       =   1920
                  OrigBottom      =   255
                  Increment       =   5
                  Max             =   180
                  Min             =   -180
                  SyncBuddy       =   -1  'True
                  Wrap            =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.CheckBox chkEnableMap 
                  Caption         =   "Enable map state saving"
                  Height          =   255
                  Left            =   60
                  TabIndex        =   113
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   3075
               End
               Begin VB.Frame FraNorthArrow 
                  Caption         =   "North Arrow"
                  Height          =   1275
                  Left            =   60
                  TabIndex        =   98
                  Top             =   840
                  Width           =   3495
                  Begin TatukGIS_XDK10.XGIS_ControlNorthArrow Arrow 
                     Height          =   915
                     Left            =   2400
                     TabIndex        =   110
                     Top             =   180
                     Width           =   1035
                     Symbol          =   0
                     Transparent     =   0   'False
                     Path            =   ""
                     Align           =   0
                     BevelInner      =   0
                     BevelOuter      =   0
                     BorderStyle     =   0
                     Color           =   -16777201
                     Ctl3D           =   0   'False
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
                     Color2          =   65535
                     Color1          =   255
                  End
                  Begin VB.CheckBox chkArrowTransparent 
                     Caption         =   "Transparent"
                     Height          =   255
                     Left            =   60
                     TabIndex        =   102
                     Top             =   900
                     Width           =   1275
                  End
                  Begin VB.CheckBox chkUseCustom 
                     Caption         =   "Custom Type"
                     Height          =   195
                     Left            =   60
                     TabIndex        =   101
                     Top             =   600
                     Width           =   1275
                  End
                  Begin VB.ComboBox cboArrow 
                     Height          =   315
                     Left            =   1320
                     Style           =   2  'Dropdown List
                     TabIndex        =   100
                     Top             =   840
                     Width           =   1035
                  End
                  Begin VB.CheckBox chkUseNorth 
                     Caption         =   "Show Arrow"
                     Height          =   255
                     Left            =   60
                     TabIndex        =   99
                     Top             =   240
                     Width           =   1275
                  End
                  Begin OASISClient.ColorPicker cpArrow 
                     Height          =   315
                     Left            =   1800
                     TabIndex        =   111
                     Top             =   180
                     Width           =   555
                     _ExtentX        =   979
                     _ExtentY        =   556
                  End
                  Begin VB.Label lblColor 
                     AutoSize        =   -1  'True
                     Caption         =   "Color:"
                     Height          =   195
                     Left            =   1320
                     TabIndex        =   112
                     Top             =   240
                     Width           =   405
                  End
                  Begin VB.Label lblType 
                     AutoSize        =   -1  'True
                     Caption         =   "Type:"
                     Height          =   195
                     Left            =   1320
                     TabIndex        =   103
                     Top             =   600
                     Width           =   405
                  End
               End
               Begin VB.ComboBox ComScroll 
                  Height          =   315
                  ItemData        =   "frmMainSettings.frx":467F
                  Left            =   2520
                  List            =   "frmMainSettings.frx":468F
                  Style           =   2  'Dropdown List
                  TabIndex        =   84
                  Top             =   480
                  Width           =   1085
               End
               Begin VB.CheckBox chkAutoScroll 
                  Caption         =   "Auto Scroll"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   82
                  Top             =   2100
                  Visible         =   0   'False
                  Width           =   1455
               End
               Begin VB.ComboBox ComMapUnits 
                  Height          =   315
                  ItemData        =   "frmMainSettings.frx":46B5
                  Left            =   1020
                  List            =   "frmMainSettings.frx":46BC
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   80
                  Top             =   2760
                  Visible         =   0   'False
                  Width           =   2475
               End
               Begin VB.CheckBox chkStoreLayer 
                  Caption         =   "Store Layer Params in Map project File"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   78
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   3315
               End
               Begin VB.CheckBox chkAlwaysSave 
                  Caption         =   "Remember My Map Settings On Exit"
                  Height          =   255
                  Left            =   60
                  TabIndex        =   77
                  Top             =   0
                  Width           =   3135
               End
               Begin VB.Label lblScrollBars 
                  AutoSize        =   -1  'True
                  Caption         =   "Scroll bar:"
                  Height          =   195
                  Left            =   1800
                  TabIndex        =   83
                  Top             =   540
                  Width           =   705
               End
               Begin VB.Label lblMapRotation 
                  AutoSize        =   -1  'True
                  Caption         =   "Map rotation:"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   81
                  Top             =   540
                  Width           =   930
               End
               Begin VB.Label lblMapUnits 
                  AutoSize        =   -1  'True
                  Caption         =   "Map Units:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   79
                  Top             =   2760
                  Visible         =   0   'False
                  Width           =   765
               End
            End
            Begin C1SizerLibCtl.C1Elastic elLocator 
               Height          =   3120
               Left            =   -4230
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   330
               Width           =   3585
               _cx             =   6324
               _cy             =   5503
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
               Begin VB.ComboBox comLevel2SpatialOps 
                  Height          =   315
                  ItemData        =   "frmMainSettings.frx":46C5
                  Left            =   60
                  List            =   "frmMainSettings.frx":46CC
                  Style           =   2  'Dropdown List
                  TabIndex        =   44
                  Top             =   1020
                  Width           =   3315
               End
               Begin VB.ComboBox comLevel1SpatialOps 
                  Height          =   315
                  ItemData        =   "frmMainSettings.frx":46DB
                  Left            =   60
                  List            =   "frmMainSettings.frx":46E2
                  Style           =   2  'Dropdown List
                  TabIndex        =   42
                  Top             =   300
                  Width           =   3315
               End
               Begin VB.Label lblSpatialOp 
                  Caption         =   "Level 2 Spatial Operation (DE-9IM)"
                  Height          =   255
                  Index           =   1
                  Left            =   60
                  TabIndex        =   45
                  Top             =   780
                  Width           =   2475
               End
               Begin VB.Label lblSpatialOp 
                  Caption         =   "Level 1 Spatial Operation (DE-9IM)"
                  Height          =   255
                  Index           =   0
                  Left            =   60
                  TabIndex        =   43
                  Top             =   60
                  Width           =   2655
               End
            End
            Begin C1SizerLibCtl.C1Elastic elMisc 
               Height          =   3120
               Left            =   -4530
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   330
               Width           =   3585
               _cx             =   6324
               _cy             =   5503
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
               Begin OASISClient.ColorPicker clSelection 
                  Height          =   315
                  Left            =   2220
                  TabIndex        =   75
                  Top             =   780
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   556
               End
               Begin VB.TextBox txtSelectionWidth 
                  Height          =   315
                  Left            =   2220
                  TabIndex        =   50
                  Top             =   1140
                  Width           =   1035
               End
               Begin VB.Frame FraPreview 
                  Caption         =   "Preview"
                  Height          =   1755
                  Left            =   180
                  TabIndex        =   47
                  Top             =   600
                  Width           =   3255
                  Begin TatukGIS_XDK10.XGIS_ViewerWnd GISPoly 
                     Height          =   975
                     Left            =   2160
                     TabIndex        =   97
                     Top             =   600
                     Width           =   1035
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
                     SelectionPattern=   "frmMainSettings.frx":46F1
                     SelectionTransparency=   100
                     SelectionWidth  =   100
                     SelectionOutlineOnly=   0   'False
                     OldCachedPaint  =   0   'False
                     PrinterModeDraft=   0   'False
                     PrinterModeForceBitmap=   0   'False
                     GDIType         =   1
                     ScaleAsFloat    =   1
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
                     Ctl3D           =   0   'False
                     Object.Visible         =   -1  'True
                     Cursor          =   16
                     DoubleBuffered  =   0   'False
                     ModeMouseButton =   0
                     CursorForUserDefined=   0
                     View3D          =   0   'False
                  End
                  Begin TatukGIS_XDK10.XGIS_ViewerWnd GISLine 
                     Height          =   975
                     Left            =   1105
                     TabIndex        =   96
                     Top             =   600
                     Width           =   1035
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
                     SelectionPattern=   "frmMainSettings.frx":12663
                     SelectionTransparency=   100
                     SelectionWidth  =   100
                     SelectionOutlineOnly=   0   'False
                     OldCachedPaint  =   0   'False
                     PrinterModeDraft=   0   'False
                     PrinterModeForceBitmap=   0   'False
                     GDIType         =   1
                     ScaleAsFloat    =   1
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
                     Ctl3D           =   0   'False
                     Object.Visible         =   -1  'True
                     Cursor          =   16
                     DoubleBuffered  =   0   'False
                     ModeMouseButton =   0
                     CursorForUserDefined=   0
                     View3D          =   0   'False
                  End
                  Begin TatukGIS_XDK10.XGIS_ViewerWnd GISPoint 
                     Height          =   975
                     Left            =   50
                     TabIndex        =   95
                     Top             =   600
                     Width           =   1035
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
                     SelectionPattern=   "frmMainSettings.frx":205D5
                     SelectionTransparency=   100
                     SelectionWidth  =   100
                     SelectionOutlineOnly=   0   'False
                     OldCachedPaint  =   0   'False
                     PrinterModeDraft=   0   'False
                     PrinterModeForceBitmap=   0   'False
                     GDIType         =   1
                     ScaleAsFloat    =   1
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
                     Ctl3D           =   0   'False
                     Object.Visible         =   -1  'True
                     Cursor          =   16
                     DoubleBuffered  =   0   'False
                     ModeMouseButton =   0
                     CursorForUserDefined=   0
                     View3D          =   0   'False
                  End
                  Begin VB.Label lblPoint 
                     AutoSize        =   -1  'True
                     Caption         =   "Polygon"
                     Height          =   195
                     Index           =   2
                     Left            =   2340
                     TabIndex        =   54
                     Top             =   300
                     Width           =   570
                  End
                  Begin VB.Label lblPoint 
                     AutoSize        =   -1  'True
                     Caption         =   "Line"
                     Height          =   195
                     Index           =   1
                     Left            =   1380
                     TabIndex        =   53
                     Top             =   300
                     Width           =   300
                  End
                  Begin VB.Label lblPoint 
                     AutoSize        =   -1  'True
                     Caption         =   "Point"
                     Height          =   195
                     Index           =   0
                     Left            =   300
                     TabIndex        =   52
                     Top             =   300
                     Width           =   360
                  End
               End
               Begin VB.CheckBox chkOutlineOnly 
                  Caption         =   "Outline Only"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   46
                  Top             =   60
                  Width           =   1215
               End
               Begin VB.ComboBox ComSelectionStyle 
                  Height          =   315
                  ItemData        =   "frmMainSettings.frx":2E547
                  Left            =   2220
                  List            =   "frmMainSettings.frx":2E549
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   40
                  Top             =   360
                  Width           =   1035
               End
               Begin VB.Label lbl 
                  AutoSize        =   -1  'True
                  Caption         =   "%"
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
                  Left            =   3240
                  TabIndex        =   51
                  Top             =   420
                  Width           =   210
               End
               Begin VB.Label lblSelectio 
                  AutoSize        =   -1  'True
                  Caption         =   "Selection Width:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   49
                  Top             =   1260
                  Width           =   1170
               End
               Begin VB.Label lblSelectionColor 
                  AutoSize        =   -1  'True
                  Caption         =   "Selection Color:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   48
                  Top             =   870
                  Width           =   1110
               End
               Begin VB.Label lblSelectionStyle 
                  AutoSize        =   -1  'True
                  Caption         =   "Selection Transparency:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   39
                  Top             =   480
                  Width           =   1725
               End
            End
            Begin C1SizerLibCtl.C1Elastic elMapTip 
               Height          =   3120
               Left            =   -4830
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   330
               Width           =   3585
               _cx             =   6324
               _cy             =   5503
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
               Begin VB.ComboBox ComTipField 
                  Height          =   315
                  Left            =   180
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   37
                  Top             =   1500
                  Width           =   3135
               End
               Begin VB.CheckBox chkTipBorder 
                  Caption         =   "Border"
                  Height          =   315
                  Left            =   1140
                  TabIndex        =   35
                  Top             =   120
                  Width           =   855
               End
               Begin OASISClient.ColorPicker ClTextColor 
                  Height          =   315
                  Left            =   2820
                  TabIndex        =   33
                  Top             =   2100
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
               End
               Begin OASISClient.ColorPicker ClTipColor 
                  Height          =   315
                  Left            =   960
                  TabIndex        =   31
                  Top             =   2100
                  Width           =   555
                  _ExtentX        =   979
                  _ExtentY        =   556
               End
               Begin VB.ComboBox ComTipLyr 
                  Height          =   315
                  Left            =   180
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   30
                  Top             =   720
                  Width           =   3135
               End
               Begin VB.ComboBox ComTipDelay 
                  Height          =   315
                  Left            =   2700
                  Style           =   2  'Dropdown List
                  TabIndex        =   28
                  Top             =   60
                  Width           =   735
               End
               Begin VB.CheckBox chkTIPEnabled 
                  Caption         =   "Enabled"
                  Height          =   315
                  Left            =   180
                  TabIndex        =   26
                  Top             =   120
                  Value           =   1  'Checked
                  Width           =   915
               End
               Begin VB.Label lblTipField 
                  Caption         =   "Tip Field:"
                  Height          =   255
                  Left            =   180
                  TabIndex        =   36
                  Top             =   1260
                  Width           =   1155
               End
               Begin VB.Label lblTipText 
                  Caption         =   "Tip Text Color:"
                  Height          =   195
                  Left            =   1680
                  TabIndex        =   34
                  Top             =   2160
                  Width           =   1095
               End
               Begin VB.Label lblTipColor 
                  Caption         =   "Tip Color:"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   32
                  Top             =   2160
                  Width           =   795
               End
               Begin VB.Label lblMapTip 
                  AutoSize        =   -1  'True
                  Caption         =   "Map Tip Layer:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   29
                  Top             =   480
                  Width           =   1065
               End
               Begin VB.Label lblTipDelay 
                  AutoSize        =   -1  'True
                  Caption         =   "Delay (S):"
                  Height          =   195
                  Left            =   1980
                  TabIndex        =   27
                  Top             =   180
                  Width           =   690
               End
            End
            Begin C1SizerLibCtl.C1Elastic elGene 
               Height          =   3120
               Left            =   -5130
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   330
               Width           =   3585
               _cx             =   6324
               _cy             =   5503
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
               Begin VB.Frame FraUrlLayer 
                  Caption         =   "Url Layer:"
                  Height          =   1215
                  Left            =   60
                  TabIndex        =   85
                  Top             =   60
                  Width           =   3315
                  Begin VB.TextBox txtHeight 
                     Height          =   285
                     Left            =   2160
                     TabIndex        =   94
                     Top             =   780
                     Width           =   555
                  End
                  Begin VB.TextBox txtWidth 
                     Height          =   285
                     Left            =   660
                     TabIndex        =   93
                     Top             =   780
                     Width           =   555
                  End
                  Begin VB.TextBox txtTimeOut 
                     Height          =   285
                     Left            =   2160
                     TabIndex        =   89
                     Top             =   480
                     Width           =   435
                  End
                  Begin VB.CheckBox chkAutoClose 
                     Caption         =   "Auto Close"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   87
                     Top             =   540
                     Width           =   1215
                  End
                  Begin VB.CheckBox chkUsePreview4URLLayers 
                     Caption         =   "Use Preview for URL layers"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   86
                     Top             =   240
                     Width           =   2355
                  End
                  Begin VB.Label lblTimeOut 
                     AutoSize        =   -1  'True
                     Caption         =   "Height:"
                     Height          =   195
                     Index           =   3
                     Left            =   1500
                     TabIndex        =   92
                     Top             =   840
                     Width           =   510
                  End
                  Begin VB.Label lblTimeOut 
                     AutoSize        =   -1  'True
                     Caption         =   "Width:"
                     Height          =   195
                     Index           =   2
                     Left            =   120
                     TabIndex        =   91
                     Top             =   840
                     Width           =   465
                  End
                  Begin VB.Label lblTimeOut 
                     AutoSize        =   -1  'True
                     Caption         =   "sec"
                     Height          =   195
                     Index           =   1
                     Left            =   2640
                     TabIndex        =   90
                     Top             =   540
                     Width           =   255
                  End
                  Begin VB.Label lblTimeOut 
                     AutoSize        =   -1  'True
                     Caption         =   "Time out:"
                     Height          =   195
                     Index           =   0
                     Left            =   1380
                     TabIndex        =   88
                     Top             =   540
                     Width           =   660
                  End
               End
               Begin VB.Frame FraIncidentLayer 
                  Caption         =   "Incident Layer Settings:"
                  Height          =   1935
                  Left            =   0
                  TabIndex        =   65
                  Top             =   360
                  Width           =   3315
                  Begin VB.CommandButton cmdOpenConfigPath 
                     Caption         =   "..."
                     Height          =   255
                     Left            =   2880
                     TabIndex        =   74
                     Top             =   1500
                     Width           =   315
                  End
                  Begin VB.TextBox txtConfigPath 
                     Height          =   255
                     Left            =   1560
                     TabIndex        =   73
                     Top             =   1500
                     Width           =   1275
                  End
                  Begin VB.CheckBox chkVisibleFrom 
                     Caption         =   "Visible From Start"
                     Height          =   375
                     Left            =   180
                     TabIndex        =   72
                     Top             =   1440
                     Width           =   1335
                  End
                  Begin VB.CheckBox chkIncludeIn 
                     Caption         =   "Include in main legend"
                     Height          =   435
                     Left            =   1560
                     TabIndex        =   71
                     Top             =   1020
                     Width           =   1515
                  End
                  Begin VB.CheckBox chkUseFile 
                     Caption         =   "Use File Params"
                     Height          =   435
                     Left            =   180
                     TabIndex        =   70
                     Top             =   1020
                     Width           =   1215
                  End
                  Begin VB.CheckBox chkUseConfig 
                     Caption         =   "Use Config File"
                     Height          =   435
                     Left            =   1560
                     TabIndex        =   69
                     Top             =   630
                     Width           =   1515
                  End
                  Begin VB.CheckBox chkIncrementalPaint 
                     Caption         =   "Incremental Paint"
                     Height          =   435
                     Left            =   180
                     TabIndex        =   68
                     Top             =   630
                     Width           =   1515
                  End
                  Begin VB.CheckBox chkIgnoreShape 
                     Caption         =   "Ignore Shape Params"
                     Height          =   435
                     Left            =   1560
                     TabIndex        =   67
                     Top             =   240
                     Width           =   1515
                  End
                  Begin VB.CheckBox chkCashedPaint 
                     Caption         =   "Cached Paint"
                     Height          =   435
                     Left            =   180
                     TabIndex        =   66
                     Top             =   240
                     Width           =   1515
                  End
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmMainSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ScoringSettings()
Public Event GetFields(sName As String)
Public Event NorthArrowChanges(bCustom As Boolean, Iconid As Integer, sArrowPAth As String, bPersistent As Boolean)
Public Event refreshArrow()
Public Event DoApply()
Public Event doOK()
Public Event DoCancel()
Public bShowNArrow As Boolean


Private m_bLOADING As Boolean

Private Sub cboArrow_Click()
   ' RaiseEvent NorthArrowChanges(
   
   If cboArrow.Text = "iMMAP" Then
   Arrow.Path = g_sAppPath & "\data\user\maps\immaparrow.gif"
   Else
   Arrow.Path = ""
    Arrow.Symbol = cboArrow.ListIndex
    End If
End Sub

Private Sub chkOutlineOnly_Click()
    transferSelectionStyle
End Sub

Private Sub chkShowWater_Click()
    Dim c As New cCommonDialog
    Dim sName As String
    
    If m_bLOADING Then Exit Sub
    
    m_bLOADING = True
    
    If chkShowWater.value = vbChecked Then
        With c
            DebugPrint GisUtils.GisSupportedFiles(XGIS_FileType.XgisFileTypeAll, True)
            .Filter = "JPEG File Interchange Format (*.jpg;*.jpeg)|*.JPG;*.JPEG|GIF File Format (*.gif)|*.gif|Scalable Vector Graphics (*.svg)|*.SVG|Portable Network Graphic (*.png)|*.PNG|CGM File (*.cgm)|*.CGM|Window Bitmap (*.bmp)|*.BMP"
            .CancelError = True
            .DialogTitle = "Add Custom water mark"
            .InitDir = g_sAppPath
            .ShowOpen
                
            If Not Len(.Filename) < 3 Then
                oMapObjects.WaterMarkPath = .Filename
                chkShowWater.Tag = .Filename
                watermark.Path = .Filename
                watermark.Visible = True
                watermark.UpDate
            Else
                oMapObjects.WaterMarkPath = ""
                chkShowWater.Tag = ""
                watermark.Path = ""
                watermark.Visible = False
                chkShowWater.value = vbUnchecked
            End If
    
        End With
    Else
        oMapObjects.WaterMarkPath = ""
        chkShowWater.Tag = ""
        watermark.Path = ""
        watermark.Visible = False
    End If
    
    m_bLOADING = False
    
End Sub

Private Sub chkUseCustom_Click()
    Dim c As New cCommonDialog
    Dim sName As String
    
    If m_bLOADING Then Exit Sub
    
    m_bLOADING = True
    
    If chkUseCustom.value = vbChecked Then
    
        If Len(chkUseCustom.Tag) > 3 Then
            m_bLOADING = False
            Exit Sub
        End If
       
        
        With c
            DebugPrint GisUtils.GisSupportedFiles(XGIS_FileType.XgisFileTypeAll, True)
            .Filter = "JPEG File Interchange Format (*.jpg;*.jpeg)|*.JPG;*.JPEG|Scalable Vector Graphics (*.svg)|*.SVG|Portable Network Graphic (*.png)|*.PNG|CGM File (*.cgm)|*.CGM|Window Bitmap (*.bmp)|*.BMP"
            .CancelError = True
            .DialogTitle = "Add Custom North Arrow"
            .InitDir = g_sAppPath
            .ShowOpen
                
            If Not Len(.Filename) < 3 Then
                chkUseCustom.Tag = .Filename
                Arrow.Path = .Filename
            Else
                chkUseCustom.Tag = ""
                Arrow.Path = ""
                Arrow.Symbol = cboArrow.ListIndex
                chkUseCustom.value = vbUnchecked
            End If
    
        End With
    Else
        chkUseCustom.Tag = ""
        Arrow.Path = ""
        Arrow.Symbol = cboArrow.ListIndex
    End If
    m_bLOADING = False
    
End Sub

Private Sub chkUseNorth_Click()
    bShowNArrow = Not bShowNArrow
    RaiseEvent refreshArrow
End Sub

Private Sub clSelection_Click()
    transferSelectionStyle
End Sub


Private Sub cmdApply_Click()
    RaiseEvent ScoringSettings
End Sub

Private Sub cmdDoApply_Click()
    GetUpdatedValues
    RaiseEvent DoApply
End Sub

Private Sub cmdDoCancel_Click()
    RaiseEvent DoCancel
    Me.Hide
End Sub

Private Sub cmdDoOK_Click()
    GetUpdatedValues
    RaiseEvent doOK
    Me.Hide
End Sub
 
Private Sub SaveStuff()
  '  Dim bSaveOnExit As Boolean
  '  Dim dMapRotation As Long
  '  Dim bScrollH As Boolean
  '  Dim bScrollV As Boolean
  '  Dim bUseNA
  '  Dim bUseCNA
  '  Dim sCAPath
    
    
    'chkAlwaysSave.Value
    
End Sub

Private Sub cmdScoringSettings_Click()

100      frmScoring.Init
102      frmScoring.Show vbModal
     
104      If frmScoring.m_bApply Then
            RaiseEvent ScoringSettings
108         frmScoring.m_bApply = False
         End If
    
End Sub

Private Sub SetupScoring()
'102     With dxTypeScoring
'104         .Dataset.ADODataset.ConnectionString = m_Cnn.ConnectionString
'106         .Dataset.Open
'108         .Dataset.Active = True
'        End With
'
'110     With dxTargetScoring
'112         .Dataset.ADODataset.ConnectionString = m_Cnn.ConnectionString
'114         .Dataset.Open
'116         .Dataset.Active = True
'        End With
'
'118     With dxTimeScoring
'120         .Dataset.ADODataset.ConnectionString = m_Cnn.ConnectionString
'122         .Dataset.Open
'124         .Dataset.Active = True
'        End With

End Sub


Private Sub ComSelectionStyle_Click()
    transferSelectionStyle
End Sub

Private Sub ComTipLyr_Click()
    If Not ComTipLyr.List(ComTipLyr.ListIndex) = "--All--" Then
        RaiseEvent GetFields(ComTipLyr.List(ComTipLyr.ListIndex))
    Else
        ComTipField.Clear
        ComTipField.AddItem "UID"
        ComTipField.ListIndex = 0
    End If
End Sub

Private Sub comZone_Click()
   oCoordTransSettings.Zone = comZone.List(comZone.ListIndex)
End Sub

Private Sub cpArrow_Click()
    Arrow.Color1 = cpArrow.color
    Arrow.Color2 = cpArrow.color
End Sub

Private Sub Form_Load()
    LoadTheSettings
End Sub

Public Sub LoadTheSettings()
    Dim i As Integer
    Dim keyArray() As Variant
    Dim element As Variant
       
    If m_bLOADING Then Exit Sub
       
    m_bLOADING = True
    
    GISPoint.color = &H8000000F
    GISPoly.color = &H8000000F
    GISLine.color = &H8000000F
    SetStylePreview
    
    With cpArrow
        .ShowDefault = True
        .ShowCustomColors = True
        .ShowMoreColors = True
        .ShowToolTips = True
        .ShowSysColorButton = True
    End With
    
    For i = 0 To 11
        cboArrow.AddItem i
    Next
    
    If FileExists(g_sAppPath & "\data\user\maps\immaparrow.gif") Then
        'NArrow.Path = g_sAppPath & "\data\user\maps\immaparrow.gif"
        cboArrow.AddItem "iMMAP"
        cboArrow.Text = "iMMAP"
    Else
        cboArrow.ListIndex = 0
    End If
    
    For i = 0 To 100
        comZone.AddItem i
        ComSelectionStyle.AddItem i
    Next
    
    For i = 1 To 10
        ComTipDelay.AddItem i
    Next
    
    ComTipDelay.ListIndex = 0
    comZone.ListIndex = 0
    ComSelectionStyle.ListIndex = 0
    
    ClTextColor.ShowDefault = True
    ClTextColor.ShowCustomColors = True
    ClTextColor.ShowMoreColors = True
    ClTextColor.ShowToolTips = True
    ClTextColor.ShowSysColorButton = True
    
    With oMapTipSetting
        chkTIPEnabled.value = IIf(.Enabled, vbChecked, vbUnchecked)
        FindIndexStrEx ComTipLyr, .MapTipLayer
        ClTextColor.color = .TextColor
        ClTipColor.color = .TipColor
        chkTipBorder.value = IIf(.TipBorder, vbChecked, vbUnchecked)
        FindIndexStrEx ComTipDelay, CStr(.TipDelay)
    End With
    
    With oSelectionStyle
        clSelection.color = .color
        chkOutlineOnly.value = IIf(.OutLineOnly, vbChecked, vbUnchecked)
        FindIndexStrEx ComSelectionStyle, CStr(.Transparency)
        txtSelectionWidth.Text = .Width
    End With
    
    With oIncidentLayerSettings
        chkCashedPaint.value = IIf(.CachedPaint, vbChecked, vbUnchecked)
        txtConfigPath.Text = .ConfigFilePAth
        chkIncludeIn.value = IIf(.HideFromLegend, vbUnchecked, vbChecked)
        chkIgnoreShape.value = IIf(.IgnoreShapeParams, vbChecked, vbUnchecked)
        chkIncrementalPaint.value = IIf(.IncrementalPaint, vbChecked, vbUnchecked)
        chkUseConfig.value = IIf(.UseConfig, vbChecked, vbUnchecked)
        chkUseFile.value = IIf(.UseFileParams, vbChecked, vbUnchecked)
        chkVisibleFrom.value = IIf(.VisibleFromStart, vbChecked, vbUnchecked)
    End With
    
    With oMapSettings
        chkAlwaysSave.value = IIf(.AlwaysSaveMapStateOnExit, vbChecked, vbUnchecked)
        chkAutoScroll.value = IIf(.AutoScroll, vbChecked, vbUnchecked)
        chkStoreLayer.value = IIf(.StoreLayerParamsInProject, vbChecked, vbUnchecked)
        ComScroll.ListIndex = .ScrollBars
        udRotationAngle.value = .MapRotation
        'txtMapRotation.Text = 0
    End With
    
    With oUrlLayerSettings
        txtTimeOut.Text = .AutoShutTime
        chkAutoClose.value = IIf(.AutoShutWin, vbChecked, vbUnchecked)
        chkUsePreview4URLLayers.value = IIf(.UseExtendedInfoWin, vbChecked, vbUnchecked)
        TxtHeight.Text = .WinHeight
        TxtWidth.Text = .WinWidth
    End With
    
    With oMapObjects
        chkUseNorth.value = IIf(.UseNorthArrow, vbChecked, vbUnchecked)
        cpArrow.color = .NorthArrowColor
        Arrow.Path = .NorthArrowPicture
        cboArrow.ListIndex = .NorthArrowType
        chkArrowTransparent.value = IIf(.NorthArrowTransparency, vbChecked, vbUnchecked)
        '.UseScaleBar
        If .UseWaterMark Then
            chkShowWater.value = vbChecked
            watermark.Path = .WaterMarkPath
        End If
    End With
    
    comLevel1SpatialOps.Clear
    comLevel2SpatialOps.Clear
    keyArray = DE9IM.Keys

    For Each element In keyArray
        comLevel1SpatialOps.AddItem element
        comLevel2SpatialOps.AddItem element
    Next

    comLevel1SpatialOps.ListIndex = 0
    comLevel2SpatialOps.ListIndex = 0
    
    FindIndexStrEx comLevel1SpatialOps, oLocatorSettings.Level1
    FindIndexStrEx comLevel2SpatialOps, oLocatorSettings.Level2
    
    SetupScoring
    
    m_bLOADING = False


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If Not Validate Then Cancel = True
    
    GetUpdatedValues
    
End Sub

Private Sub GetUpdatedValues()

    With oCoordTransSettings
        .Inverse_Flattening = 6378137
        .Semi_Major_Axis = CDbl("298.2572236")
    End With
   
    With oMapTipSetting
        .Enabled = IIf(chkTIPEnabled.value = vbChecked, True, False)
        .MapTipLayer = ComTipLyr.List(ComTipLyr.ListIndex)
        .TextColor = ClTextColor.color
        .TipBorder = IIf(chkTipBorder.value = vbChecked, True, False)
        .TipColor = ClTipColor.color
        .TipDelay = ComTipDelay.List(ComTipDelay.ListIndex)
        .MapTipField = ComTipField.List(ComTipField.ListIndex)
    End With
   
    With oIncidentLayerSettings
        .CachedPaint = IIf(chkCashedPaint.value = vbChecked, True, False)
        .ConfigFilePAth = txtConfigPath.Text
        .HideFromLegend = IIf(chkIncludeIn.value = vbChecked, False, True)
        .IgnoreShapeParams = IIf(chkIgnoreShape.value = vbChecked, True, False)
        .IncrementalPaint = IIf(chkIncrementalPaint.value = vbChecked, True, False)
        .UseConfig = IIf(chkUseConfig.value = vbChecked, True, False)
        .UseFileParams = IIf(chkUseFile.value = vbChecked, True, False)
        .VisibleFromStart = IIf(chkVisibleFrom.value = vbChecked, True, False)
    End With

    With oMapSettings
        .AlwaysSaveMapStateOnExit = IIf(chkAlwaysSave.value = vbChecked, True, False)
        .AutoScroll = IIf(chkAutoScroll.value = vbChecked, True, False)
        .iMAPUnits = 0
        .ScrollBars = ComScroll.ListIndex
        .MapRotation = udRotationAngle.value
        .StoreLayerParamsInProject = IIf(chkStoreLayer.value = vbChecked, True, False)
    End With

    With oUrlLayerSettings
        .AutoShutTime = IIf(IsNumeric(txtTimeOut.Text), CLng(txtTimeOut.Text) * 1000, 30000)
        .AutoShutWin = IIf(chkAutoClose.value = vbChecked, True, False)
        .UseExtendedInfoWin = IIf(chkUsePreview4URLLayers.value = vbChecked, True, False)
        .WinHeight = IIf(IsNumeric(TxtHeight.Text), CLng(TxtHeight.Text), 3300)
        .WinWidth = IIf(IsNumeric(TxtWidth.Text), CLng(TxtWidth.Text), 3300)
    End With
    
    With oMapObjects
        
        .UseNorthArrow = IIf(chkUseNorth.value = vbChecked, True, False)
        .NorthArrowColor = cpArrow.color
        .NorthArrowPicture = Arrow.Path
        .NorthArrowType = cboArrow.ListIndex
        .NorthArrowTransparency = IIf(chkArrowTransparent.value = vbChecked, True, False)
        '.UseScaleBar
        .UseWaterMark = IIf(chkShowWater.value = vbChecked, True, False)
        .WaterMarkPath = watermark.Path
    End With
End Sub

Private Sub transferSelectionStyle()

If m_bLOADING Then Exit Sub

    With oSelectionStyle
        .color = clSelection.color
        .OutLineOnly = IIf(chkOutlineOnly.value = vbChecked, True, False)
        .Transparency = CInt(ComSelectionStyle.List(ComSelectionStyle.ListIndex))
        .Width = CInt(IIf(IsNumeric(txtSelectionWidth.Text), txtSelectionWidth.Text, 100))
    End With
    CommitSelStyle
End Sub

Public Function Validate() As Boolean
    Dim dblA As Double 'Semi-major axis of ellipsoid or radius of sphere
    Dim dblFinv As Double 'Inverse Flattening
    
    Dim dblFN As Double 'False Northing
    Dim dblFE As Double 'False Easting
    Dim dblLatOrig As Double 'Latitude of Origin
    Dim dblLonOrig As Double 'Longitude of Origin
    
    Dim strResult As String

    With oCoordTransSettings

        If IsNumeric(txtAxis.Text) Then
            dblA = CDbl(txtAxis.Text)
            .Semi_Major_Axis = dblA
        Else
            txtAxis.SetFocus
            MsgBox "Please enter a valid Semi-Major Axis value in metres!", vbExclamation, "Mercator"
            Validate = False
            Exit Function
        End If

        If IsNumeric(txtFInv.Text) Then
            dblFinv = CDbl(txtFInv.Text)
            .Inverse_Flattening = dblFinv
        Else
            MsgBox "Please enter a valid Inverse Flattening value!", vbExclamation, "Mercator"
            Validate = False
            Exit Function
        End If

        If IsNumeric(txtFN.Text) Then
            dblFN = CDbl(txtFN.Text)
            .False_Northing = dblFN
        Else
            MsgBox "Please enter a valid False Northing value in metres!", vbExclamation, "Mercator"
            Validate = False
            Exit Function
        End If

        If IsNumeric(txtFE.Text) Then
            dblFE = CDbl(txtFE.Text)
            .False_Easting = dblFE
        Else
            MsgBox "Please enter a valid False Easting value in metres!", vbExclamation, "Mercator"
            Validate = False
            Exit Function
        End If

        If IsNumeric(txtLatOrig.Text) Then
            dblLatOrig = DegToRad(CDbl(txtLatOrig.Text))
            .Lat_Of_Origin = dblLatOrig
        Else
            MsgBox "Please enter a valid Latitude of Origin value in metres!", vbExclamation, "Mercator"
            Validate = False
            Exit Function
        End If

        If IsNumeric(txtLonOrig.Text) Then
            dblLonOrig = DegToRad(CDbl(txtLonOrig))
            .Long_Of_Origin = dblLonOrig
        Else
            MsgBox "Please enter a valid Longitude of Origin value in metres!", vbExclamation, "Mercator"
            Validate = False
            Exit Function
        End If
    
        Validate = True
    
'        'Initialise the ellipsoid variables in the GeoMercator class module
'        strResult = objProj.SetEllipsoid(dblA, dblFinv, CBool(chkSphere))
'
'        'If strResult is not a zero-length string, it means there was an error with the
'        'initialisation. The error message will be contained in strResult, so display the error message.
'        If Len(strResult) > 0 Then
'            MsgBox strResult, vbExclamation, "Mercator"
'            Validate = False
'            Exit Function
'        End If
'
'        'Initialise the projection variables in the GeoMercator class module
'        strResult = objProj.SetProjection(dblFN, dblFE, dblLatOrig, dblLonOrig)
'
'        'If strResult is not a zero-length string, it means there was an error with the
'        'initialisation. The error message will be contained in strResult, so display the error message.
'        If Len(strResult) > 0 Then
'            MsgBox strResult, vbExclamation, "Mercator"
'            Validate = False
'        Else
'            Validate = True
'        End If
    
    End With
    
End Function

Private Sub CommitSelStyle()
Dim zoom As Double
    With oSelectionStyle
        
        GISPoint.SelectionColor = .color
        GISPoint.SelectionOutlineOnly = .OutLineOnly
        GISPoint.SelectionTransparency = .Transparency
        GISPoint.SelectionWidth = .Width
        zoom = GISPoint.zoom
        GISPoint.zoom = zoom * 2
        GISPoint.zoom = zoom
        
        GISPoly.SelectionColor = .color
        GISPoly.SelectionOutlineOnly = .OutLineOnly
        GISPoly.SelectionTransparency = .Transparency
        GISPoly.SelectionWidth = .Width
        zoom = GISPoly.zoom
        GISPoly.zoom = zoom * 2
        GISPoly.zoom = zoom
        
        GISLine.SelectionColor = .color
        GISLine.SelectionOutlineOnly = .OutLineOnly
        GISLine.SelectionTransparency = .Transparency
        GISLine.SelectionWidth = .Width
        zoom = GISLine.zoom
        GISLine.zoom = zoom * 2
        GISLine.zoom = zoom
        
        'txtSelectionWidth.Text = GISPoint.SelectionWidth
    End With

End Sub

Private Sub SetStylePreview()
    Dim i As Integer
    Dim oLyr As New TatukGIS_XDK10.XGIS_LayerVector
    Dim optLyr As New TatukGIS_XDK10.XGIS_LayerVector
    Dim opolLyr As New TatukGIS_XDK10.XGIS_LayerVector
    
    Dim oLine As TatukGIS_XDK10.XGIS_Shape
    Dim oPol As TatukGIS_XDK10.XGIS_Shape
    Dim oPt As TatukGIS_XDK10.XGIS_Shape
    Dim Pt As TatukGIS_XDK10.XGIS_Point
    Dim ptg As TatukGIS_XDK10.XGIS_Point
    Dim oSHP As TatukGIS_XDK10.XGIS_Shape
    
    CommitSelStyle
     
    GISLine.Add oLyr
    
    Set oLine = oLyr.CreateShape(XgisShapeTypeArc)
    
    oLyr.Lock
    
    oLine.Lock TatukGIS_XDK10.XgisLockExtent
    oLine.AddPart
        
    For i = 0 To 2
        Set Pt = New TatukGIS_XDK10.XGIS_Point
        Pt.x = i
        Pt.y = i
        oLine.AddPoint Pt
    Next
    
    oLine.IsSelected = True
    oLine.Unlock
    oLyr.Unlock
    GISLine.VisibleExtent = oLine.Extent
    
    GISPoint.Add optLyr
    
    Set oPt = optLyr.CreateShape(XgisShapeTypePoint)
    
    optLyr.Lock
    
    oPt.Lock TatukGIS_XDK10.XgisLockExtent
    oPt.AddPart
        
    Set Pt = New TatukGIS_XDK10.XGIS_Point
    Pt.x = 1
    Pt.y = 1
    oPt.AddPoint Pt
    
    oPt.IsSelected = True
    oPt.Unlock
    optLyr.Unlock
    GISPoint.VisibleExtent = oPt.Extent
    
    GISPoly.Add opolLyr
    
    Set oPol = opolLyr.CreateShape(XgisShapeTypePolygon)
    
    opolLyr.Lock
    
    oPol.Lock TatukGIS_XDK10.XgisLockExtent
    oPol.AddPart
    
    Set Pt = New TatukGIS_XDK10.XGIS_Point
    Pt.x = 1
    Pt.y = 0
    oPol.AddPoint Pt
   
    Set Pt = New TatukGIS_XDK10.XGIS_Point
    Pt.x = 2
    Pt.y = 0
    oPol.AddPoint Pt
    
    Set Pt = New TatukGIS_XDK10.XGIS_Point
    Pt.x = 2.5
    Pt.y = 2.5
    oPol.AddPoint Pt
        
    Set Pt = New TatukGIS_XDK10.XGIS_Point
    Pt.x = 2
    Pt.y = 3
    oPol.AddPoint Pt
    
    
    Set Pt = New TatukGIS_XDK10.XGIS_Point
    Pt.x = 1
    Pt.y = 3
    oPol.AddPoint Pt
    
    Set Pt = New TatukGIS_XDK10.XGIS_Point
    Pt.x = 1 / 2
    Pt.y = 5 / 2
    oPol.AddPoint Pt
    
    oPol.IsSelected = True
    oPol.Unlock
    opolLyr.Unlock
    GISPoly.VisibleExtent = oPol.Extent
End Sub

Private Sub txtSelectionWidth_Change()
    transferSelectionStyle
End Sub

Private Sub udRotationAngle_Change()
    'edRotationAngle.Text = udRotationAngle.Value
End Sub
