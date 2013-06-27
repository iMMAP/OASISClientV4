VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFreqAnalysis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OASIS Spatial Density Calculator"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "frmFreqAnalysis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7215
      _cx             =   12726
      _cy             =   9975
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
      Picture         =   "frmFreqAnalysis.frx":038A
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
      _GridInfo       =   $"frmFreqAnalysis.frx":11D1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CheckBox chkShowLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Labels"
         Height          =   330
         Left            =   3645
         TabIndex        =   33
         Top             =   5235
         Value           =   1  'Checked
         Width           =   1710
      End
      Begin VB.CheckBox chkAdvancedMode 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Advanced Mode"
         Height          =   330
         Left            =   1860
         TabIndex        =   31
         Top             =   5235
         Width           =   1725
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   90
         TabIndex        =   3
         Top             =   5235
         Width           =   1710
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Enabled         =   0   'False
         Height          =   330
         Left            =   5415
         TabIndex        =   2
         Top             =   5235
         Width           =   1710
      End
      Begin C1SizerLibCtl.C1Elastic C1ESECURITYINCIDENTS 
         Height          =   1560
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   90
         Width           =   4380
         _cx             =   7726
         _cy             =   2752
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
         Caption         =   "SPATIAL DENSITY    CALCULATOR"
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
      Begin C1SizerLibCtl.C1Tab C1TDetailLocation 
         Height          =   3465
         Left            =   90
         TabIndex        =   4
         Top             =   1710
         Width           =   7035
         _cx             =   12409
         _cy             =   6112
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
         Caption         =   "Overlay|Selection|Colour|Weighting *|Colour *|Summary"
         Align           =   0
         CurrTab         =   4
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   3435
            Left            =   -6510
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   15
            Width           =   5895
            _cx             =   10398
            _cy             =   6059
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
            GridCols        =   7
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmFreqAnalysis.frx":12AD
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame Frame1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   1830
               TabIndex        =   42
               Top             =   75
               Width           =   3840
               Begin VB.OptionButton optOverlay 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Overlay Layer"
                  Height          =   195
                  Left            =   2160
                  TabIndex        =   44
                  Top             =   120
                  Width           =   1335
               End
               Begin VB.OptionButton optSelection 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Selection Layer"
                  Height          =   195
                  Left            =   240
                  TabIndex        =   43
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1815
               End
            End
            Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
               Height          =   1260
               Left            =   225
               OleObjectBlob   =   "frmFreqAnalysis.frx":1380
               TabIndex        =   23
               Top             =   1950
               Width           =   5445
            End
            Begin C1SizerLibCtl.C1Elastic C1EWeightingFields 
               Height          =   375
               Index           =   2
               Left            =   225
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   75
               Width           =   1605
               _cx             =   2831
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
               Caption         =   "Weighting fields"
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
            Begin C1SizerLibCtl.C1Elastic lblFieldValues 
               Height          =   375
               Left            =   225
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1575
               Width           =   5445
               _cx             =   9604
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
               Caption         =   "Field value weights"
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
            Begin MSComctlLib.ListView listFieldLayer 
               Height          =   1125
               Left            =   225
               TabIndex        =   32
               Top             =   450
               Width           =   5445
               _ExtentX        =   9604
               _ExtentY        =   1984
               View            =   2
               LabelEdit       =   1
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3435
            Index           =   0
            Left            =   -7110
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   15
            Width           =   5895
            _cx             =   10398
            _cy             =   6059
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
            _GridInfo       =   $"frmFreqAnalysis.frx":2A55
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.ListBox listFeatureLayerName 
               Height          =   1230
               Left            =   4170
               TabIndex        =   29
               Top             =   225
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VB.ListBox listFeatureLayer 
               Height          =   1620
               Left            =   225
               TabIndex        =   13
               Top             =   1500
               Width           =   5445
            End
            Begin C1SizerLibCtl.C1Elastic C1EFeatureLayer 
               Height          =   420
               Index           =   2
               Left            =   225
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   1080
               Width           =   2820
               _cx             =   4974
               _cy             =   741
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
               Caption         =   "Feature layer:"
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
            Begin C1SizerLibCtl.C1Elastic C1EInThis 
               Height          =   855
               Index           =   1
               Left            =   225
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   225
               Width           =   5445
               _cx             =   9604
               _cy             =   1508
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
               Caption         =   $"frmFreqAnalysis.frx":2B36
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
            Height          =   3435
            Index           =   1
            Left            =   -7410
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   15
            Width           =   5895
            _cx             =   10398
            _cy             =   6059
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
            _GridInfo       =   $"frmFreqAnalysis.frx":2C02
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.ListBox listOverlayLayerName 
               Height          =   1230
               Left            =   4395
               TabIndex        =   28
               Top             =   225
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.ListBox listOverlayLayer 
               Height          =   1620
               Left            =   225
               TabIndex        =   10
               Top             =   1455
               Width           =   5445
            End
            Begin C1SizerLibCtl.C1Elastic C1EFeatureLayer 
               Height          =   405
               Index           =   3
               Left            =   225
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   1050
               Width           =   2730
               _cx             =   4815
               _cy             =   714
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
               Caption         =   "Overlay layer:"
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
            Begin C1SizerLibCtl.C1Elastic C1EInThis 
               Height          =   825
               Index           =   0
               Left            =   225
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   225
               Width           =   5445
               _cx             =   9604
               _cy             =   1455
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
               Caption         =   $"frmFreqAnalysis.frx":2CE0
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic28 
            Height          =   3435
            Left            =   -6810
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   15
            Width           =   5895
            _cx             =   10398
            _cy             =   6059
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
            GridCols        =   7
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmFreqAnalysis.frx":2D9B
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   390
               Left            =   225
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   3045
               Width           =   5445
               _cx             =   9604
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
               Caption         =   ""
               Align           =   0
               AutoSizeChildren=   8
               BorderWidth     =   2
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
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   0
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmFreqAnalysis.frx":2E71
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin XpressEditorsLibCtl.dxSpinEdit dxLabelSize 
                  Height          =   315
                  Left            =   4110
                  OleObjectBlob   =   "frmFreqAnalysis.frx":2EC3
                  TabIndex        =   41
                  Top             =   30
                  Width           =   1305
               End
               Begin C1SizerLibCtl.C1Elastic C1ELabelColour 
                  Height          =   330
                  Index           =   0
                  Left            =   30
                  TabIndex        =   38
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   1305
                  _cx             =   2302
                  _cy             =   582
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
                  Caption         =   "  Label colour:"
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
               Begin XpressEditorsLibCtl.dxColorEdit dxColorLabel 
                  Height          =   315
                  Left            =   1395
                  OleObjectBlob   =   "frmFreqAnalysis.frx":2FBD
                  TabIndex        =   39
                  Top             =   30
                  Width           =   1305
               End
               Begin C1SizerLibCtl.C1Elastic C1ELabelSize 
                  Height          =   330
                  Index           =   1
                  Left            =   2760
                  TabIndex        =   40
                  TabStop         =   0   'False
                  Top             =   30
                  Width           =   1290
                  _cx             =   2275
                  _cy             =   582
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
                  Caption         =   "  Label size:"
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
            Begin VB.CommandButton cmdUpdate2 
               Caption         =   "Apply"
               Height          =   855
               Left            =   5115
               TabIndex        =   35
               Top             =   285
               Width           =   555
            End
            Begin C1SizerLibCtl.C1Elastic C1EWeightingFields 
               Height          =   375
               Index           =   1
               Left            =   225
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   1140
               Width           =   2595
               _cx             =   4577
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
               Caption         =   "Start colour:"
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
            Begin C1SizerLibCtl.C1Elastic C1EOtherApplicable 
               Height          =   375
               Left            =   225
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   1905
               Width           =   2595
               _cx             =   4577
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
               Caption         =   "Colour range:"
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
            Begin XpressEditorsLibCtl.dxColorEdit dxColorStart 
               Height          =   315
               Left            =   225
               OleObjectBlob   =   "frmFreqAnalysis.frx":30CC
               TabIndex        =   17
               Top             =   1515
               Width           =   2595
            End
            Begin C1SizerLibCtl.C1Elastic C1EInThis 
               Height          =   855
               Index           =   2
               Left            =   225
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   285
               Width           =   4890
               _cx             =   8625
               _cy             =   1508
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
               Caption         =   $"frmFreqAnalysis.frx":31DB
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
            Begin C1SizerLibCtl.C1Elastic C1EEndColour 
               Height          =   375
               Index           =   1
               Left            =   2985
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   1140
               Width           =   2130
               _cx             =   3757
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
               Caption         =   "End colour:"
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
            Begin XpressEditorsLibCtl.dxColorEdit dxColorEnd 
               Height          =   315
               Left            =   2985
               OleObjectBlob   =   "frmFreqAnalysis.frx":32A4
               TabIndex        =   21
               Top             =   1515
               Width           =   2685
            End
            Begin CONTROLSLibCtl.dxProgressBar dxProgressBar1 
               Height          =   765
               Left            =   225
               TabIndex        =   18
               Top             =   2280
               Width           =   5445
               _Version        =   65536
               _cx             =   9604
               _cy             =   1349
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
               Pos             =   100
               Step            =   10
               ShowText        =   0   'False
               Orientation     =   0
               StartColor      =   16711680
               EndColor        =   16777215
               DrawBorderStyle =   1
               ShowTextStyle   =   0
               DrawBarStyle    =   2
               DrawBarBorderStyle=   2
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3435
            Index           =   2
            Left            =   1125
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   15
            Width           =   5895
            _cx             =   10398
            _cy             =   6059
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
            _GridInfo       =   $"frmFreqAnalysis.frx":3397
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin OASISClient.OASISThemeRangePicker OASISThemeRangePicker1 
               Height          =   2565
               Left            =   0
               TabIndex        =   36
               Top             =   645
               Width           =   5895
               _ExtentX        =   10398
               _ExtentY        =   4524
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Apply"
               Height          =   420
               Left            =   4680
               TabIndex        =   34
               Top             =   225
               Width           =   990
            End
            Begin C1SizerLibCtl.C1Elastic C1EInThis 
               Height          =   420
               Index           =   3
               Left            =   225
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   225
               Width           =   4455
               _cx             =   7858
               _cy             =   741
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
               Caption         =   "Contact an OASIS admin for more info on this feature"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3435
            Index           =   3
            Left            =   8760
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   15
            Width           =   5895
            _cx             =   10398
            _cy             =   6059
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
            _GridInfo       =   $"frmFreqAnalysis.frx":3478
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
      End
      Begin CONTROLSLibCtl.dxProgressBar dxProgressBar2 
         Height          =   330
         Left            =   1860
         TabIndex        =   22
         Top             =   5235
         Visible         =   0   'False
         Width           =   3495
         _Version        =   65536
         _cx             =   6165
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
Attribute VB_Name = "frmFreqAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event SetFields(sLayer As String, sFieldsSelected As String)
Public Event CheckZoom(sLayerName As String, lZoom As Long, bZoomCorrect As Boolean)
Public Event UpdateSymbology(lStartCol As Long, lEndCol As Long, lMax As Long)
Public Event CountFrequency(sOverLayLayer As String, sTargetLayer As String, lStartCol As Long, lEndCol As Long, bAdvanced As Boolean, oRSScoring As ADODB.Recordset, sFieldsOverlay As String, sFieldSelection As String)
Public Event SetSocring(sLayer As String, sFieldName As String, oRSScoring As ADODB.Recordset)
'Private mRS As ADODB.Recordset
'Private mRSScoring() As ADODB.Recordset
Public m_bCancelFlag As Boolean
Private RSScoring As ADODB.Recordset

Private sFieldNamesOverlay As String
Private sFieldNamesSelection As String


Private Sub chkAdvancedMode_Click()

    C1TDetailLocation.CurrTab = 0
    C1TDetailLocation.TabVisible(5) = False

    If chkAdvancedMode.value = vbChecked Then

        C1TDetailLocation.TabVisible(2) = False
        C1TDetailLocation.TabVisible(3) = True
        C1TDetailLocation.TabVisible(4) = True
    Else
        C1TDetailLocation.TabVisible(2) = True
        C1TDetailLocation.TabVisible(3) = False
        C1TDetailLocation.TabVisible(4) = False
    
    End If
    
End Sub

Public Sub SetFieldValueRecordset(oRS As ADODB.Recordset)

    Set RSScoring = oRS
    Set dxDBGrid1.DataSource = RSScoring
    dxDBGrid1.Columns.DestroyColumns
    dxDBGrid1.Columns.RetrieveFields
    dxDBGrid1.Columns(0).Visible = False
    dxDBGrid1.Columns(1).Visible = False
    dxDBGrid1.Columns(2).Visible = True
    dxDBGrid1.Columns(2).Width = 3 * dxDBGrid1.Width / 4
    dxDBGrid1.Columns(2).ReadOnly = True
    dxDBGrid1.Columns(3).Visible = True
    dxDBGrid1.Columns(3).ColumnType = gedSpinEdit
    dxDBGrid1.Columns(3).SpinColumn.minValue = 1
    dxDBGrid1.Columns(3).SpinColumn.maxValue = 10
    
    dxDBGrid1.Columns(3).caption = "Weighting"
    dxDBGrid1.Columns(3).Width = dxDBGrid1.Width / 4
    dxDBGrid1.KeyField = "FieldValue" '
    
End Sub

Private Sub cmdCancel_Click()
    m_bCancelFlag = True
    Unload Me
End Sub

Private Sub cmdApply_Click()
        '<EhHeader>
        On Error GoTo cmdApply_Click_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
        Dim i As Long
        'Dim sFieldNames As String
100     chkAdvancedMode.Visible = False
102     dxProgressBar2.Visible = True
104     chkShowLabels.Visible = False
        dxDBGrid1.Visible = False
        cmdApply.Enabled = False
    
106     If chkAdvancedMode.value = vbChecked Then
RSScoring.UpdateBatch
            RSScoring.Filter = "LayerName = '" & listFeatureLayerName.Text & "' OR LayerName = '" & listOverlayLayerName.Text & "'"
            
            i = 0
            
            If optOverlay.value = True Then
        
                sFieldNamesOverlay = ""
                Do Until i = listFieldLayer.ListItems.Count
                    If listFieldLayer.ListItems(i + 1).Checked Then sFieldNamesOverlay = sFieldNamesOverlay & "," & listFieldLayer.ListItems.Item(i + 1).Text
                    i = i + 1
                Loop

            Else
    
                sFieldNamesSelection = ""
                Do Until i = listFieldLayer.ListItems.Count
                    If listFieldLayer.ListItems(i + 1).Checked Then sFieldNamesSelection = sFieldNamesSelection & "," & listFieldLayer.ListItems.Item(i + 1).Text
                    i = i + 1
                Loop

            End If

142         RaiseEvent CountFrequency(listOverlayLayerName.Text, listFeatureLayerName.Text, dxColorStart, dxColorEnd, True, RSScoring, sFieldNamesOverlay, sFieldNamesSelection)
            Call cmdUpdate_Click
        Else
144         RaiseEvent CountFrequency(listOverlayLayerName.Text, listFeatureLayerName.Text, dxColorStart, dxColorEnd, False, Nothing, "", "")
            Call cmdUpdate_Click
        End If
        
        lblFieldValues.caption = "Field value weights"
        Set dxDBGrid1.DataSource = Nothing
146     dxProgressBar2.Visible = False
148     chkAdvancedMode.Visible = True
150     chkShowLabels.Visible = True
152     Set oRS = Nothing
        dxDBGrid1.Visible = True
        cmdApply.Enabled = True
        '<EhFooter>
        Exit Sub

cmdApply_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmFreqAnalysis.cmdApply_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdUpdate_Click()
    RaiseEvent UpdateSymbology(dxColorStart, dxColorEnd, -1)
End Sub

Private Sub cmdUpdate2_Click()
    RaiseEvent UpdateSymbology(dxColorStart, dxColorEnd, -1)
End Sub

Private Sub dxColorEnd_KeyDown(KeyCode As Integer, _
                               Shift As Integer)
    dxProgressBar1.EndColor = dxColorEnd
End Sub

Private Sub dxColorStart_Change()
    dxProgressBar1.StartColor = dxColorStart
End Sub



Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
        
100     If Not DoesTableExist(m_Cnn.ConnectionString, "ScoringV2") Then
102         oRS.Fields.Append "LayerName", adVarWChar, 255
104         oRS.Fields.Append "FieldName", adVarWChar, 255
106         oRS.Fields.Append "FieldValue", adVarWChar, 255
108         oRS.Fields.Append "Scoring", adInteger
110         oRS.Open
112         CreateTable "ScoringV2", oRS, m_Cnn
114         oRS.Close
        End If
    
116     Set RSScoring = New ADODB.Recordset
118     RSScoring.Open "SELECT * FROM ScoringV2", m_Cnn, adOpenDynamic, adLockBatchOptimistic

120     Call chkAdvancedMode_Click
122     C1TDetailLocation.CurrTab = 0
124     dxProgressBar1.EndColor = dxColorEnd
126     dxProgressBar1.StartColor = dxColorStart
128     dxDBGrid1.KeyField = "UID"
130     dxColorEnd = vbRed
132     dxColorStart = vbGreen
    
        OASISThemeRangePicker1.Init
134     OASISThemeRangePicker1.SetNumberOfIntervals 1
136     OASISThemeRangePicker1.SetInterval 1, dxColorEnd, 255
138     OASISThemeRangePicker1.SetThemeStartColor dxColorStart
140     OASISThemeRangePicker1.Render
142     Set oRS = Nothing
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmFreqAnalysis.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RSScoring = Nothing
End Sub

Private Sub listFeatureLayer_Click()
    RSScoring.UpdateBatch
    listFeatureLayerName.ListIndex = listFeatureLayer.ListIndex

    If optSelection.value = True Then
        RaiseEvent SetFields(listFeatureLayerName.Text, "")
        sFieldNamesSelection = ""
    End If

    If Len(listOverlayLayer.Text) > 0 And Len(listFeatureLayer.Text) > 0 Then cmdApply.Enabled = True
    Set dxDBGrid1.DataSource = Nothing
    ReDim mRSScoring(listFieldLayer.ListItems.Count)
End Sub

Private Sub listFieldLayer_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    Item.Selected = True
    
    If optOverlay.value = False Then
    
        RaiseEvent SetSocring(listFeatureLayerName.Text, listFieldLayer.SelectedItem.Text, RSScoring)
        lblFieldValues.caption = "Field value weighted VALUES for field [" & listFieldLayer.SelectedItem.Text & "]:"
    Else
    
        RaiseEvent SetSocring(listOverlayLayerName.Text, listFieldLayer.SelectedItem.Text, RSScoring)
        lblFieldValues.caption = "Field value weighted MULTIPLIERS for field [" & listFieldLayer.SelectedItem.Text & "]:"
        
    End If
    
    On Error Resume Next
    
End Sub

Private Sub listFieldLayer_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Item.Selected = Not Item.Selected = True
    Item.Checked = Not Item.Checked
    Call listFieldLayer_ItemCheck(Item)
    
End Sub

Private Sub listOverlayLayer_Click()

    Dim bZoomCorrect As Boolean
    listOverlayLayerName.ListIndex = listOverlayLayer.ListIndex
    
    If optOverlay.value = True Then
        RaiseEvent SetFields(listOverlayLayerName.Text, "")
        sFieldNamesOverlay = ""
    End If
    
    If Len(listOverlayLayer.Text) > 0 Then
        RaiseEvent CheckZoom(listOverlayLayerName.Text, listOverlayLayerName.ItemData(listOverlayLayerName.ListIndex), bZoomCorrect)

        If bZoomCorrect Then

            If Len(listOverlayLayer.Text) > 0 And Len(listFeatureLayer.Text) > 0 Then cmdApply.Enabled = True
        
        Else
            cmdApply.Enabled = False
            listOverlayLayer.ListIndex = -1
        End If
        
    Else
        cmdApply.Enabled = False
        listOverlayLayer.ListIndex = -1
    End If
    
End Sub

Private Sub optOverlay_Click()

    Dim i As Long

    If optSelection.value = True Then
        
        sFieldNamesOverlay = ""

        Do Until i = listFieldLayer.ListItems.Count

            If listFieldLayer.ListItems(i + 1).Checked Then sFieldNamesOverlay = sFieldNamesOverlay & "," & listFieldLayer.ListItems.Item(i + 1).Text
            i = i + 1
        Loop

        RaiseEvent SetFields(listFeatureLayerName.Text, sFieldNamesSelection)
        
    Else
    
        sFieldNamesSelection = ""

        Do Until i = listFieldLayer.ListItems.Count

            If listFieldLayer.ListItems(i + 1).Checked Then sFieldNamesSelection = sFieldNamesSelection & "," & listFieldLayer.ListItems.Item(i + 1).Text
            i = i + 1
        Loop

        RaiseEvent SetFields(listOverlayLayerName.Text, sFieldNamesOverlay)
    End If
    
    Set dxDBGrid1.DataSource = Nothing
    lblFieldValues.caption = "Field value weights"

End Sub

Private Sub optSelection_Click()
 Call optOverlay_Click
End Sub
