VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OASIS Selector Tool"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "frmSelector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic 
      Height          =   5100
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5340
      _cx             =   9419
      _cy             =   8996
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
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "frmSelector.frx":076A
      Caption         =   ""
      Align           =   5
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
      PicturePos      =   6
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   8
      GridCols        =   8
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmSelector.frx":15B1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CheckBox chkAutoSelect 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Select"
         Height          =   330
         Left            =   2700
         TabIndex        =   29
         Top             =   4680
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chkLoadTo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Load to grid"
         Height          =   330
         Left            =   1395
         TabIndex        =   28
         Top             =   4680
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   330
         Left            =   4005
         TabIndex        =   2
         Top             =   4680
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   90
         TabIndex        =   1
         Top             =   4680
         Width           =   1245
      End
      Begin C1SizerLibCtl.C1Elastic C1ESPATIALDENSITY 
         Height          =   1230
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   90
         Width           =   3210
         _cx             =   5662
         _cy             =   2170
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   64
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "SELECTOR TOOL"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   0
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   2
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
      Begin C1SizerLibCtl.C1Tab C1TOverlaySelection 
         Height          =   3240
         Left            =   90
         TabIndex        =   4
         Top             =   1380
         Width           =   5160
         _cx             =   9102
         _cy             =   5715
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
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483645
         BackTabColor    =   16777215
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "Layer|Fields|Method"
         Align           =   0
         CurrTab         =   2
         FirstTab        =   0
         Style           =   5
         Position        =   5
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   7
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   3210
            Index           =   1
            Left            =   -5295
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   15
            Width           =   4380
            _cx             =   7726
            _cy             =   5662
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
            BackColor       =   16777215
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
            GridRows        =   9
            GridCols        =   8
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmSelector.frx":1689
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic C1EPleaseSelect 
               Height          =   795
               Index           =   0
               Left            =   225
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   225
               Width           =   4155
               _cx             =   7329
               _cy             =   1402
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
               BackColor       =   16777215
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   $"frmSelector.frx":1767
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
            Begin VB.ListBox lstLayer 
               Height          =   1425
               Left            =   225
               TabIndex        =   7
               Top             =   1410
               Width           =   3930
            End
            Begin VB.ListBox lstLayerName 
               Height          =   1035
               Left            =   3225
               TabIndex        =   6
               Top             =   225
               Visible         =   0   'False
               Width           =   1155
            End
            Begin C1SizerLibCtl.C1Elastic C1EOverlayLayer 
               Height          =   390
               Index           =   3
               Left            =   225
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   1020
               Width           =   1965
               _cx             =   3466
               _cy             =   688
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
               BackColor       =   16777215
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Selection layer:"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   3210
            Index           =   0
            Left            =   -4995
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   15
            Width           =   4380
            _cx             =   7726
            _cy             =   5662
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
            BackColor       =   16777215
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
            GridRows        =   9
            GridCols        =   8
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmSelector.frx":1805
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic C1EOverlayLayer 
               Height          =   375
               Index           =   0
               Left            =   225
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   975
               Width           =   1965
               _cx             =   3466
               _cy             =   661
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
               BackColor       =   16777215
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Layer Fields:"
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
            Begin C1SizerLibCtl.C1Elastic C1EPleaseSelect 
               Height          =   750
               Index           =   1
               Left            =   225
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   225
               Width           =   4155
               _cx             =   7329
               _cy             =   1323
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
               BackColor       =   16777215
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Please select the fields which are to be included in the resulting layer with selections"
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
            Begin MSComctlLib.ListView lstFields 
               Height          =   1860
               Left            =   225
               TabIndex        =   14
               Top             =   1350
               Width           =   3930
               _ExtentX        =   6932
               _ExtentY        =   3281
               View            =   2
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   3210
            Index           =   2
            Left            =   765
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   15
            Width           =   4380
            _cx             =   7726
            _cy             =   5662
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
            BackColor       =   16777215
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
            GridRows        =   9
            GridCols        =   8
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmSelector.frx":18E3
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame FraSelectionMethod 
               BackColor       =   &H00FFFFFF&
               Height          =   2220
               Left            =   225
               TabIndex        =   17
               Top             =   990
               Width           =   3930
               Begin VB.ComboBox cmbOtherLayerName 
                  Height          =   315
                  Left            =   1800
                  TabIndex        =   27
                  Text            =   "Combo1"
                  Top             =   1140
                  Visible         =   0   'False
                  Width           =   1695
               End
               Begin VB.CheckBox chkMeasureDistance 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Measure distance"
                  Height          =   330
                  Left            =   240
                  TabIndex        =   26
                  Top             =   1800
                  Value           =   1  'Checked
                  Width           =   1890
               End
               Begin VB.ComboBox cmbOtherLayer 
                  Height          =   315
                  Left            =   1680
                  TabIndex        =   23
                  Text            =   "Combo1"
                  Top             =   1500
                  Visible         =   0   'False
                  Width           =   2055
               End
               Begin VB.OptionButton OptAnotherLayer 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Another layer"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   22
                  Top             =   1560
                  Width           =   1275
               End
               Begin VB.OptionButton OptPolygon 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Polygon"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   21
                  Top             =   1260
                  Width           =   1815
               End
               Begin VB.OptionButton OptPolyline 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Polyline"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   20
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.OptionButton OptCircleSpecify 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Circle (specify km)"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   19
                  Top             =   660
                  Width           =   1755
               End
               Begin VB.OptionButton OptCircleDraw 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Circle (draw)"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   18
                  Top             =   360
                  Width           =   1755
               End
               Begin XpressEditorsLibCtl.dxTextEdit txtBufferDistance 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "#,##0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  Height          =   315
                  Left            =   2820
                  OleObjectBlob   =   "frmSelector.frx":19C1
                  TabIndex        =   24
                  Top             =   600
                  Width           =   885
               End
               Begin C1SizerLibCtl.C1Elastic C1EBufferKm 
                  Height          =   345
                  Left            =   2760
                  TabIndex        =   25
                  TabStop         =   0   'False
                  Top             =   300
                  Width           =   990
                  _cx             =   1746
                  _cy             =   609
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
                  BackColor       =   16777215
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "  Buffer (km)"
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
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1EPleaseSelect 
               Height          =   765
               Index           =   3
               Left            =   225
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   225
               Width           =   3930
               _cx             =   6932
               _cy             =   1349
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
               BackColor       =   16777215
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   $"frmSelector.frx":1A31
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
      End
      Begin CONTROLSLibCtl.dxProgressBar dxProgressBar1 
         Height          =   330
         Left            =   1395
         TabIndex        =   10
         Top             =   4680
         Visible         =   0   'False
         Width           =   2550
         _Version        =   65536
         _cx             =   4498
         _cy             =   582
         ForeColor       =   0
         BackColor       =   15790320
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MinPos          =   0
         MaxPos          =   100
         Pos             =   0
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
End
Attribute VB_Name = "frmSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event ChangeSelector(sTools As String)
Public Event DisableSelections(sLayerName As String)
Public Event GetLayers(sLayers As String)
Public Event UseOtherLayer(sOtherLayer As String)
Public Event MergeOtherSelections(sOtherLayer As String, bSuccess As Boolean)
Public Event ChangeActiveLayer(sName As String, sExcluded As String)
Public Event GetLayerSelectedInLegend(sLayerName As String)
 
Private m_PErcentMark As Long
Private m_PErcentJump As Long
Private m_LayerName As String

Public Sub AddNewLayer(sName As String, _
                       sCaption As String)

    If Not sName = "selection_layer" Then
        lstLayer.AddItem sCaption
        lstLayerName.AddItem sName
        cmbOtherLayer.AddItem sCaption
        cmbOtherLayerName.AddItem sName
    End If
    
End Sub

Public Function GetActiveLayer() As String
    GetActiveLayer = lstLayerName.List(lstLayer.ListIndex)
End Function

Public Function DistanceEnabled() As Boolean
    DistanceEnabled = IIf(chkMeasureDistance.value = vbChecked, True, False)
End Function

Public Sub RenewSelection()
    OptCircleSpecify.value = False
    OptAnotherLayer.value = False
    OptCircleDraw.value = False
    OptPolygon.value = False
    OptPolyline.value = False
    cmdApply.Visible = False
    cmbOtherLayer.Visible = False
    txtBufferDistance.Enabled = False
    'Call OptAnotherLayer_Click

End Sub

Public Function GetFields() As String
    
    Dim i As Long
    i = 0
    
    Do Until i = lstFields.ListItems.Count

        If lstFields.ListItems(i + 1).Checked Then
            GetFields = GetFields & ";" & lstFields.ListItems(i + 1).Text
        End If

        i = i + 1
    Loop
    
End Function

Public Function GetTool() As OASIS_TOOLS

    Select Case True

        Case OptCircleSpecify.value
        
            GetTool = oPointBuffer

        Case OptPolyline.value
            
            GetTool = oLineSelect
        
        Case OptCircleDraw.value
            
            GetTool = oCircleSelect

        Case OptPolygon.value
           
            GetTool = oAreaSelect
           
        Case Else
            
            GetTool = 0
            
    End Select
    
End Function

Public Sub InitProgressBar(lCount As Long)

    dxProgressBar1.Step = 1
    dxProgressBar1.MaxPos = lCount + 1
    dxProgressBar1.MinPos = 0
    dxProgressBar1.Pos = 0
    m_PErcentMark = Round((lCount + 1) / 100, 0)
    m_PErcentJump = m_PErcentMark
End Sub

Public Sub ProgressStep()
    dxProgressBar1.DoStep

    If dxProgressBar1.Pos = m_PErcentMark Then

        DoEvents
        m_PErcentMark = m_PErcentMark + m_PErcentJump
    End If

End Sub

Public Sub SetLayer(sLayerName As String, _
                    sCaption As String, _
                    sFields As String)
    Dim sFieldNames() As String
    Dim i As Long

    If Len(sLayerName) > 0 Then
        
        i = 1
        FraSelectionMethod.Enabled = True
        sFieldNames = Split(sFields, ",")
        lstFields.ListItems.Clear
        
        Do Until i > UBound(sFieldNames)
            lstFields.ListItems.Add , , sFieldNames(i)
            lstFields.ListItems(i).Checked = True
            i = i + 1
        Loop

        lstFields.Enabled = True
        
    End If

End Sub


Private Sub cmbOtherLayer_Click()

    If Len(cmbOtherLayer.Text) > 0 Then
        cmdApply.Visible = True
    Else
        cmdApply.Visible = False
    End If
    
End Sub

Private Sub cmdApply_Click()

    Dim sSelectionLayer As String
    Dim sOtherLayer  As String
    Dim bSuccess As Boolean
    
    sSelectionLayer = cmbOtherLayerName.List(cmbOtherLayer.ListIndex)
    
    If sSelectionLayer <> "" Then
    
        RaiseEvent MergeOtherSelections(sSelectionLayer, bSuccess)
    
        If bSuccess Then
            sOtherLayer = lstLayerName.List(lstLayer.ListIndex)
    
            If sOtherLayer <> "" Then
    
                If "selection_layer" <> sOtherLayer Then
                    RaiseEvent UseOtherLayer("selection_layer")
                    'OptCircleDraw.value = True
                Else
                    MsgBox "You need to select a new layer in the listbox to analyse"
                End If
            
            End If
        End If

    Else
        MsgBox "You have no layer selected in the drop-down box!"
    End If

End Sub

Private Sub SelectionOptionChange()

    Dim sLayers As String
    Dim i As Long
    Dim sLayersArray() As String
    Dim sLayerName As String
    'Dim bSuccess As Boolean
    cmbOtherLayer.Visible = False
    Dim sSelectionLayer As String
    sSelectionLayer = lstLayerName.List(lstLayer.ListIndex)
        
    If Len(sSelectionLayer) > 0 Then
        
        Select Case True

            Case OptCircleSpecify.value
        
                RaiseEvent ChangeSelector(OptCircleSpecify.caption)

            Case OptPolyline.value
            
                RaiseEvent ChangeSelector(OptPolyline.caption)
        
            Case OptCircleDraw.value
            
                RaiseEvent ChangeSelector(OptCircleDraw.caption)

            Case OptPolygon.value
           
                RaiseEvent ChangeSelector(OptPolygon.caption)
           
            Case Else
            
                RaiseEvent ChangeSelector("Pan")
                cmbOtherLayer.Visible = True
            
        End Select

        If OptCircleSpecify.value Or OptPolyline.value Then
            txtBufferDistance.Enabled = True
        Else
            txtBufferDistance.Enabled = False
        End If
 
        If OptAnotherLayer.value Then
            cmdApply.Visible = True
        Else
            cmdApply.Visible = False
        End If
    
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    FraSelectionMethod.Enabled = False
    txtBufferDistance = 0
    cmbOtherLayer.Clear
    cmbOtherLayer.Visible = False
    RenewSelection
    C1TOverlaySelection.CurrTab = 0
End Sub

Private Sub OptAnotherLayer_Click()
    SelectionOptionChange
End Sub

Private Sub OptCircleDraw_Click()
    SelectionOptionChange
End Sub

Private Sub OptCircleSpecify_Click()
    SelectionOptionChange
End Sub

Private Sub OptPolygon_Click()
    SelectionOptionChange
End Sub

Private Sub OptPolyline_Click()
    SelectionOptionChange
End Sub

Public Function GetBuffer() As Double

    GetBuffer = 0

    If Not txtBufferDistance.Enabled Then
        GetBuffer = -1
    ElseIf Len(txtBufferDistance) > 0 Then
        GetBuffer = CDbl(txtBufferDistance) / 100

    End If

End Function

Private Sub lstLayer_Click()
    Dim sLayerName As String
    Dim sExcluded As String

    SafeMoveFirst g_RSGISGridTableSettings
    g_RSGISGridTableSettings.Find "alias = '" & lstLayer.Text & "'"
    
    If Not g_RSGISGridTableSettings.EOF And Not g_RSGISGridTableSettings.Bof Then
        sLayerName = g_RSGISGridTableSettings.Fields("Name").value
        sExcluded = IIf(IsNull(g_RSGISGridTableSettings.Fields("excludedFlds").value), "", g_RSGISGridTableSettings.Fields("excludedFlds").value)
    Else
        sLayerName = lstLayer.Text 'COMLayer
    End If

    sLayerName = lstLayerName.List(lstLayer.ListIndex)
            
    FraSelectionMethod.Enabled = True
    lstFields.ListItems.Clear
    lstFields.Enabled = False
      
    RaiseEvent ChangeActiveLayer(sLayerName, sExcluded)
    
End Sub
