VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TatukGIS_XDK10_WIN32.ocx"
Begin VB.Form frmAttributes 
   Caption         =   "Info"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   Icon            =   "frmAttributes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAttributes.frx":6852
   ScaleHeight     =   5985
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin TatukGIS_XDK10.XGIS_ControlAttributes GEO1 
      Height          =   5115
      Left            =   60
      TabIndex        =   30
      Top             =   -60
      Visible         =   0   'False
      Width           =   4035
      ReadOnly        =   0   'False
      AllowRestructure=   -1  'True
      ColorHeader     =   -16777201
      ColorGrid       =   -16777211
      Align           =   0
      BevelInner      =   0
      BevelOuter      =   2
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
   Begin C1SizerLibCtl.C1Elastic elAttr 
      Height          =   5985
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4275
      _cx             =   7541
      _cy             =   10557
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
      _GridInfo       =   $"frmAttributes.frx":6C4D
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic elTop 
         Height          =   555
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   4275
         _cx             =   7541
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
         GridRows        =   2
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"frmAttributes.frx":6C8B
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.ComboBox ComLAyer 
            Height          =   315
            Left            =   30
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   285
            Width           =   4215
         End
         Begin VB.Label lblActiveInfo 
            AutoSize        =   -1  'True
            Caption         =   "Active Info Layer:"
            Height          =   255
            Left            =   30
            TabIndex        =   8
            Top             =   30
            Width           =   4215
         End
      End
      Begin C1SizerLibCtl.C1Tab tbShapAttr 
         Height          =   5430
         Left            =   0
         TabIndex        =   1
         Top             =   555
         Width           =   4275
         _cx             =   7541
         _cy             =   9578
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
         Caption         =   "Attributes|Report|Geo|Style"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   5055
            Left            =   45
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   330
            Width           =   4185
            _cx             =   7382
            _cy             =   8916
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
            _GridInfo       =   $"frmAttributes.frx":6CC9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Tab C1TTab1Tab2 
               Height          =   4605
               Left            =   30
               TabIndex        =   13
               Top             =   30
               Width           =   4125
               _cx             =   7276
               _cy             =   8123
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
               Begin TatukGIS_XDK10.XGIS_ControlAttributes Attributes1 
                  Height          =   4575
                  Left            =   15
                  TabIndex        =   29
                  Top             =   15
                  Width           =   4095
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
            Begin VB.Frame Frame2 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   30
               TabIndex        =   10
               Top             =   4665
               Width           =   4125
               Begin VB.CheckBox chkDynamicInfo 
                  Caption         =   "Dynamic Info"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   14
                  Top             =   60
                  Width           =   1455
               End
               Begin VB.CheckBox chkMultiInfo 
                  Caption         =   "Multi Info"
                  Height          =   255
                  Left            =   150
                  TabIndex        =   12
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   1275
               End
               Begin VB.CommandButton cmdClear 
                  Caption         =   "Clear"
                  Height          =   285
                  Left            =   2220
                  TabIndex        =   11
                  Top             =   30
                  Width           =   975
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   5055
            Left            =   4920
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   330
            Width           =   4185
            _cx             =   7382
            _cy             =   8916
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
            GridRows        =   6
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmAttributes.frx":6D0B
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic el4 
               Height          =   315
               Left            =   30
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   30
               Width           =   4125
               _cx             =   7276
               _cy             =   556
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
               GridRows        =   1
               GridCols        =   2
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmAttributes.frx":6D7B
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.TextBox txtTitle 
                  Height          =   315
                  Left            =   735
                  TabIndex        =   27
                  Top             =   0
                  Width           =   3390
               End
               Begin VB.Label lblTitle 
                  AutoSize        =   -1  'True
                  Caption         =   "Title:"
                  Height          =   315
                  Left            =   0
                  TabIndex        =   28
                  Top             =   0
                  Width           =   735
               End
            End
            Begin C1SizerLibCtl.C1Elastic el3 
               Height          =   315
               Left            =   30
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   705
               Width           =   4125
               _cx             =   7276
               _cy             =   556
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
               GridRows        =   1
               GridCols        =   2
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmAttributes.frx":6DB5
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.TextBox txtMapTitle 
                  Height          =   315
                  Left            =   735
                  TabIndex        =   24
                  Top             =   0
                  Width           =   3390
               End
               Begin VB.Label lblMapTitle 
                  AutoSize        =   -1  'True
                  Caption         =   "Map Title:"
                  Height          =   315
                  Left            =   0
                  TabIndex        =   25
                  Top             =   0
                  Width           =   735
               End
            End
            Begin C1SizerLibCtl.C1Elastic el2 
               Height          =   1050
               Left            =   30
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   1050
               Width           =   4125
               _cx             =   7276
               _cy             =   1852
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
               Begin VB.CheckBox chkIncludeGeo 
                  Caption         =   "Include Geo ID"
                  Height          =   285
                  Left            =   0
                  TabIndex        =   22
                  Top             =   0
                  Width           =   2235
               End
               Begin VB.CheckBox chkIncludeArea 
                  Caption         =   "Include Area"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   21
                  Top             =   270
                  Width           =   1725
               End
               Begin VB.CheckBox chkIncludeLength 
                  Caption         =   "Include Length"
                  Height          =   225
                  Left            =   0
                  TabIndex        =   20
                  Top             =   540
                  Width           =   1605
               End
               Begin VB.CheckBox chkIncludeCentroid 
                  Caption         =   "Include Centroid"
                  Height          =   285
                  Left            =   0
                  TabIndex        =   19
                  Top             =   780
                  Width           =   1485
               End
            End
            Begin C1SizerLibCtl.C1Elastic el1 
               Height          =   2505
               Left            =   30
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   2130
               Width           =   4125
               _cx             =   7276
               _cy             =   4419
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
               _GridInfo       =   $"frmAttributes.frx":6DEF
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin MSComctlLib.ListView lvAttributes 
                  Height          =   4080
                  Left            =   0
                  TabIndex        =   16
                  Top             =   210
                  Width           =   7290
                  _ExtentX        =   12859
                  _ExtentY        =   7197
                  View            =   2
                  Arrange         =   1
                  LabelEdit       =   1
                  MultiSelect     =   -1  'True
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  Checkboxes      =   -1  'True
                  FullRowSelect   =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   0
               End
               Begin VB.Label lblReportFields 
                  Caption         =   "Report Fields:"
                  Height          =   210
                  Left            =   0
                  TabIndex        =   17
                  Top             =   0
                  Width           =   7290
               End
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "Print"
               Height          =   360
               Left            =   30
               TabIndex        =   5
               Top             =   4665
               Width           =   4125
            End
            Begin VB.CheckBox chkIncludeMap 
               Caption         =   "Include Map"
               Height          =   300
               Left            =   30
               TabIndex        =   4
               Top             =   375
               Width           =   4125
            End
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   5055
            Left            =   5220
            TabIndex        =   2
            Top             =   330
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   8916
            _Version        =   393217
            Style           =   7
            Appearance      =   1
         End
      End
   End
End
Attribute VB_Name = "frmAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oLLs() As TatukGIS_XDK10.XGIS_LayerVector
Private m_oShps() As TatukGIS_XDK10.XGIS_Shape
Private m_oShp As TatukGIS_XDK10.XGIS_Shape
Public LyrCol As Collection
Public Event ExitTool()

Public Sub ShowMultiSelected(oLL As Variant)
        '<EhHeader>
        On Error GoTo ShowMultiSelected_Err
        '</EhHeader>
100     m_oLLs = oLL
        '<EhFooter>
        Exit Sub

ShowMultiSelected_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAttributes.ShowMultiSelected " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub ShowMultiShape(oSHP As Variant)
        '<EhHeader>
        On Error GoTo ShowMultiShape_Err
        '</EhHeader>
100     m_oShps = oSHP
        '<EhFooter>
        Exit Sub

ShowMultiShape_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAttributes.ShowMultiShape " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub ShowSelected(lL As TatukGIS_XDK10.XGIS_LayerVector) 'XGIS_LayerVector
        Dim i As Long

        '<EhHeader>
        On Error GoTo ShowSelected_Err
        '</EhHeader>
100     GEO1.ShowSelected lL
        
        With lL.Fields
        
            lvAttributes.ListItems.Clear
        
            For i = 0 To .Count
                lvAttributes.ListItems.Add , , .Item(i).Name
                lvAttributes.ListItems.Item(i).Checked = True
                DebugPrint .Item(i).Name
            Next
        
        End With
        
        '<EhFooter>
        Exit Sub

ShowSelected_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmAttributes.ShowSelected " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub ShowShape(shp As TatukGIS_XDK10.XGIS_Shape)
        '<EhHeader>
        On Error GoTo ShowShape_Err
        '</EhHeader>
        Dim oLyrVector As TatukGIS_XDK10.XGIS_LayerVector
        
100     'Attributes1(0).ShowShape shp
        Attributes1.ShowShape shp
102     If shp Is Nothing Then Exit Sub
        
104     Set m_oShp = shp
        
        Dim i As Long
        
106     lvAttributes.ListItems.Clear
108     Set oLyrVector = m_oShp.layer
        
        
110     For i = 1 To m_oShp.layer.Fields.Count - 1
112         DebugPrint oLyrVector.Fields.Item(i).Name
114         lvAttributes.ListItems.Add Key:=oLyrVector.Fields.Item(i).Name, Text:=oLyrVector.Fields.Item(i).Name
116         lvAttributes.ListItems.Item(i).Checked = True
            
        Next
        
118     txtTitle.Text = "Layer:" & shp.layer.caption & " GEO ID:" & shp.uID


        '<EhFooter>
        Exit Sub

ShowShape_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAttributes.ShowShape " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub cmdClear_Click()
        '<EhHeader>
        On Error GoTo cmdClear_Click_Err
        '</EhHeader>
100     Attributes1.Clear '(0).Clear
        '<EhFooter>
        Exit Sub

cmdClear_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAttributes.cmdClear_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdPrint_Click()
        '<EhHeader>
        On Error GoTo cmdPrint_Click_Err
        '</EhHeader>
        Dim oLyr As TatukGIS_XDK10.XGIS_LayerVector
        Dim i As Integer
        Dim oRS As New ADODB.Recordset
        Dim iType As Integer
        Dim iFldCount As Integer

100     If Not m_oShp Is Nothing Then
102         Set oLyr = m_oShp.layer
         
104         If Not oLyr Is Nothing Then
                                       
106             For i = 1 To oLyr.Fields.Count - 1
            
108                 If lvAttributes.ListItems.Item(oLyr.Fields.Item(i).Name).Checked Then
                        ' DebugPrint oLyr.FieldInfo(i).Name
                        'DebugPrint oLyr.FieldInfo(i).Width
                            
110                     Select Case oLyr.FieldInfo(i).FieldType
           
                            Case TatukGIS_XDK10.XgisFieldTypeBoolean
112                             iType = adBoolean

114                         Case TatukGIS_XDK10.XgisFieldTypeDate
116                             iType = adDate

118                         Case TatukGIS_XDK10.XgisFieldTypeFloat
120                             iType = adDouble

122                         Case TatukGIS_XDK10.XgisFieldTypeNumber
124                             iType = adDouble

126                         Case TatukGIS_XDK10.XgisFieldTypeString
128                             iType = adVarChar
                        End Select

                        'oRS.Open
130                     oRS.Fields.Append oLyr.Fields.Item(i).Name, iType, oLyr.FieldInfo(i).Width ', , m_oSHP.GetField(oLyr.Fields.Item(i).Name)
                    End If
                
                Next

132             iFldCount = oRS.Fields.Count - 1

134             If chkIncludeArea.value = vbChecked Then
136                 oRS.Fields.Append "Geo_AREA", adDouble
                End If
            
138             If chkIncludeGeo.value = vbChecked Then
140                 oRS.Fields.Append "GEO_ID", adDouble
                End If
            
142             If chkIncludeLength.value = vbChecked Then
144                 oRS.Fields.Append "GEO_Length", adDouble
                End If
            
146             If chkIncludeCentroid.value = vbChecked Then
148                 oRS.Fields.Append "GEO_CenterX", adDouble
150                 oRS.Fields.Append "GEO_CenterY", adDouble
                End If
            
152             oRS.Open , , adOpenDynamic, adLockBatchOptimistic
154             oRS.AddNew
            
156             For i = 0 To oRS.Fields.Count - 1

                    '                If oRS.Fields.Item(i).Name = "UID" Then
                    '                    oRS.Fields.Item(i).Value = m_oSHP.uID
                    '                Else
158                 DebugPrint oRS.Fields.Item(i).Type
160                 DebugPrint oRS.Fields.Item(i).Name
162                 oRS.Fields.Item(i).value = m_oShp.GetField(oRS.Fields.Item(i).Name)
                    '                End If

                Next
                
164             If chkIncludeArea.value = vbChecked Then
166                 oRS.Fields.Item("Geo_AREA").value = m_oShp.area
                End If
                
168             If chkIncludeGeo.value = vbChecked Then
170                 oRS.Fields.Item("GEO_ID").value = m_oShp.uID
                End If
            
172             If chkIncludeLength.value = vbChecked Then
174                 oRS.Fields.Item("GEO_Length").value = m_oShp.Length
                End If
            
176             If chkIncludeCentroid.value = vbChecked Then
178                 oRS.Fields.Item("GEO_CenterX").value = m_oShp.Centroid.X
180                 oRS.Fields.Item("GEO_CenterY").value = m_oShp.Centroid.Y
                End If
            
182             If chkIncludeMap.value = vbChecked Then
184                 Clipboard.Clear
186                 oLyr.viewer.PrintClipboard
188                 frmReportsFromRS.SetReportRS txtTitle.Text, oRS, "", Clipboard.GetData(vbCFEMetafile), txtMapTitle.Text, ""
                Else
190                 Clipboard.Clear
192                 frmReportsFromRS.SetReportRS txtTitle.Text, oRS, ""
                End If
            
194             frmReportsFromRS.ShowReport
196             frmReportsFromRS.Show vbModal, Me
            
            End If
        End If

        '<EhFooter>
        Exit Sub

cmdPrint_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAttributes.cmdPrint_Click " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub Init(oGIS As TatukGIS_XDK10.XGIS_Viewer)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
    
        Dim i As Integer
        Dim sCurrItem As String
        
        Set LyrCol = New Collection
        
        If ComLayer.ListCount > 0 Then
        
            sCurrItem = ComLayer.List(ComLayer.ListIndex)
        
        End If
        
100     ComLayer.Clear
    
102     ComLayer.AddItem "--All--"
        LyrCol.Add "--All--", "--All--"
        
104     For i = 0 To oGIS.items.Count - 1
            On Error Resume Next
106         If GisUtils.IsInherited(oGIS.items.Item(i), "XGIS_LayerVector") Then
                LyrCol.Add oGIS.items.Item(i).Name, oGIS.items.Item(i).caption
108             ComLayer.AddItem oGIS.items.Item(i).caption 'Name
            End If
        
        Next
       
110     If ComLayer.ListCount > 0 Then
            
            If Len(sCurrItem) > 0 Then
                FindIndexStrEx ComLayer, sCurrItem
            Else
                FindIndexStrEx ComLayer, "--All--"
            End If
            
        End If
       
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmAttributes.Init " & "at line " & Erl
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
        
        Frame2.ZOrder 0

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAttributes.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
RaiseEvent ExitTool
End Sub
