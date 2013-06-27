VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmWVIIncidents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "World Vision Security Events"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9810
   Icon            =   "frmWVIIncidents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic 
      Height          =   5040
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9810
      _cx             =   17304
      _cy             =   8890
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
      Picture         =   "frmWVIIncidents.frx":6852
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
      _GridInfo       =   $"frmWVIIncidents.frx":77EA
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1ESecurityEvent 
         Height          =   945
         Left            =   90
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   90
         Width           =   6000
         _cx             =   10583
         _cy             =   1667
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
         ForeColor       =   33023
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   " SECURITY EVENT DATA ENTRY"
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
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   330
         Left            =   6150
         TabIndex        =   80
         Top             =   4620
         Width           =   3570
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   90
         TabIndex        =   79
         Top             =   4620
         Width           =   3570
      End
      Begin C1SizerLibCtl.C1Tab C1TTab1Tab 
         Height          =   3465
         Left            =   90
         TabIndex        =   1
         Top             =   1095
         Width           =   9630
         _cx             =   16986
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
         Caption         =   "Detail|Location|Event|Casualities|Subjects|Victims|Outcome"
         Align           =   0
         CurrTab         =   1
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic23 
            Height          =   3435
            Left            =   12480
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   15
            Width           =   8565
            _cx             =   15108
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
            BackColor       =   -2147483645
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
            GridRows        =   10
            GridCols        =   6
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmWVIIncidents.frx":78C3
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.ComboBox comOutcomeMaterial 
               Height          =   315
               Left            =   225
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   78
               Top             =   600
               Width           =   2625
            End
            Begin VB.ComboBox comOutcomeImpact 
               Height          =   315
               Left            =   225
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Top             =   2085
               Width           =   2625
            End
            Begin VB.ComboBox comOutcomeTrend 
               Height          =   315
               Left            =   225
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   76
               Top             =   1350
               Width           =   2625
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic26 
               Height          =   360
               Left            =   225
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   1725
               Width           =   2625
               _cx             =   4630
               _cy             =   635
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Impact:"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic25 
               Height          =   375
               Left            =   225
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   975
               Width           =   2625
               _cx             =   4630
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Trend:"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic24 
               Height          =   375
               Left            =   225
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   225
               Width           =   2625
               _cx             =   4630
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Material Recovery:"
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
            Begin C1SizerLibCtl.C1Elastic lblOutcomeMaterial 
               Height          =   750
               Left            =   2850
               TabIndex        =   100
               TabStop         =   0   'False
               Top             =   225
               Width           =   5490
               _cx             =   9684
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic lblOutcomeTrend 
               Height          =   750
               Left            =   2850
               TabIndex        =   101
               TabStop         =   0   'False
               Top             =   975
               Width           =   5490
               _cx             =   9684
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic lblOutcomeImpact 
               Height          =   735
               Left            =   2850
               TabIndex        =   102
               TabStop         =   0   'False
               Top             =   1725
               Width           =   5490
               _cx             =   9684
               _cy             =   1296
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
               BackColor       =   -2147483645
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
            Begin RichTextLib.RichTextBox txtOutcomeDescription 
               Height          =   750
               Left            =   225
               TabIndex        =   103
               Top             =   2460
               Width           =   8115
               _ExtentX        =   14314
               _ExtentY        =   1323
               _Version        =   393217
               TextRTF         =   $"frmWVIIncidents.frx":7999
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   3435
            Left            =   12180
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   15
            Width           =   8565
            _cx             =   15108
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   3435
               Left            =   0
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   0
               Width           =   8565
               _cx             =   15108
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
               BackColor       =   -2147483645
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
               GridRows        =   8
               GridCols        =   7
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmWVIIncidents.frx":7A49
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin XpressEditorsLibCtl.dxSpinEdit txtVictimsAge 
                  Height          =   315
                  Left            =   5790
                  OleObjectBlob   =   "frmWVIIncidents.frx":7B11
                  TabIndex        =   104
                  Top             =   225
                  Width           =   2550
               End
               Begin VB.TextBox txtVictimsName 
                  Height          =   390
                  Left            =   1350
                  TabIndex        =   65
                  Top             =   225
                  Width           =   2850
               End
               Begin VB.ComboBox comVictimsNationality 
                  Height          =   315
                  Left            =   1350
                  Sorted          =   -1  'True
                  TabIndex        =   64
                  Top             =   1005
                  Width           =   2850
               End
               Begin VB.ComboBox comVictimsGender 
                  Height          =   315
                  Left            =   5790
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   63
                  Top             =   615
                  Width           =   2550
               End
               Begin VB.ComboBox comVictimsAff 
                  Height          =   315
                  Left            =   5790
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   62
                  Top             =   1005
                  Width           =   2550
               End
               Begin VB.ComboBox comVictimsType 
                  Height          =   315
                  Left            =   1350
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   61
                  Top             =   615
                  Width           =   2850
               End
               Begin VB.CommandButton cmdVictimsRemove 
                  Caption         =   "Remove"
                  Height          =   285
                  Left            =   225
                  TabIndex        =   60
                  Top             =   2925
                  Width           =   3975
               End
               Begin VB.CommandButton cmdVictimsAdd 
                  Caption         =   "Add"
                  Height          =   285
                  Left            =   4365
                  TabIndex        =   59
                  Top             =   2925
                  Width           =   3975
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic17 
                  Height          =   390
                  Left            =   225
                  TabIndex        =   66
                  TabStop         =   0   'False
                  Top             =   615
                  Width           =   1125
                  _cx             =   1984
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
                  BackColor       =   -2147483645
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "Type:"
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic18 
                  Height          =   390
                  Left            =   4365
                  TabIndex        =   67
                  TabStop         =   0   'False
                  Top             =   1005
                  Width           =   1425
                  _cx             =   2514
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
                  BackColor       =   -2147483645
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "WV Affiliation:"
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic19 
                  Height          =   390
                  Left            =   4365
                  TabIndex        =   68
                  TabStop         =   0   'False
                  Top             =   615
                  Width           =   1425
                  _cx             =   2514
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
                  BackColor       =   -2147483645
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "Gender:"
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic20 
                  Height          =   390
                  Left            =   4365
                  TabIndex        =   69
                  TabStop         =   0   'False
                  Top             =   225
                  Width           =   1425
                  _cx             =   2514
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
                  BackColor       =   -2147483645
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "Age:"
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic21 
                  Height          =   390
                  Left            =   225
                  TabIndex        =   70
                  TabStop         =   0   'False
                  Top             =   1005
                  Width           =   1125
                  _cx             =   1984
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
                  BackColor       =   -2147483645
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "Nationality:"
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic22 
                  Height          =   390
                  Left            =   225
                  TabIndex        =   71
                  TabStop         =   0   'False
                  Top             =   225
                  Width           =   1125
                  _cx             =   1984
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
                  BackColor       =   -2147483645
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   "Name:"
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
               Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridVictims 
                  Height          =   1260
                  Left            =   225
                  OleObjectBlob   =   "frmWVIIncidents.frx":7C0B
                  TabIndex        =   111
                  Top             =   1665
                  Width           =   8115
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   3435
            Left            =   11880
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   15
            Width           =   8565
            _cx             =   15108
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
            BackColor       =   -2147483645
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
            GridRows        =   8
            GridCols        =   7
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmWVIIncidents.frx":88B3
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin DXDBGRIDLibCtl.dxDBGrid dxDBGridSubjects 
               Height          =   1320
               Left            =   225
               OleObjectBlob   =   "frmWVIIncidents.frx":897B
               TabIndex        =   110
               Top             =   1620
               Width           =   8115
            End
            Begin XpressEditorsLibCtl.dxSpinEdit txtSubjectsAge 
               Height          =   315
               Left            =   5805
               OleObjectBlob   =   "frmWVIIncidents.frx":9623
               TabIndex        =   105
               Top             =   225
               Width           =   2535
            End
            Begin VB.CommandButton cmdSubjectsAdd 
               Caption         =   "Add"
               Height          =   270
               Left            =   4380
               TabIndex        =   57
               Top             =   2940
               Width           =   3960
            End
            Begin VB.CommandButton cmdSubjectsRemove 
               Caption         =   "Remove"
               Height          =   270
               Left            =   225
               TabIndex        =   56
               Top             =   2940
               Width           =   3990
            End
            Begin VB.ComboBox comSubjectsType 
               Height          =   315
               Left            =   1350
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Top             =   600
               Width           =   2865
            End
            Begin VB.ComboBox comSubjectsAff 
               Height          =   315
               Left            =   5805
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   54
               Top             =   975
               Width           =   2535
            End
            Begin VB.ComboBox comSubjectsGender 
               Height          =   315
               Left            =   5805
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Top             =   600
               Width           =   2535
            End
            Begin VB.ComboBox comSubjectsNationality 
               Height          =   315
               Left            =   1350
               Sorted          =   -1  'True
               TabIndex        =   52
               Top             =   975
               Width           =   2865
            End
            Begin VB.TextBox txtSubjectsName 
               Height          =   375
               Left            =   1350
               TabIndex        =   51
               Top             =   225
               Width           =   2865
            End
            Begin C1SizerLibCtl.C1Elastic C1ESector 
               Height          =   375
               Left            =   225
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   600
               Width           =   1125
               _cx             =   1984
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Type:"
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
            Begin C1SizerLibCtl.C1Elastic C1EAffiliation 
               Height          =   375
               Left            =   4380
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   975
               Width           =   1425
               _cx             =   2514
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "WV Affiliation:"
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
            Begin C1SizerLibCtl.C1Elastic C1EGender 
               Height          =   375
               Left            =   4380
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   600
               Width           =   1425
               _cx             =   2514
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Gender:"
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
            Begin C1SizerLibCtl.C1Elastic C1EAge 
               Height          =   375
               Left            =   4380
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   225
               Width           =   1425
               _cx             =   2514
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Age:"
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
            Begin C1SizerLibCtl.C1Elastic C1ENationality 
               Height          =   375
               Left            =   225
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   975
               Width           =   1125
               _cx             =   1984
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Nationality:"
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
            Begin C1SizerLibCtl.C1Elastic C1EName 
               Height          =   375
               Left            =   225
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   225
               Width           =   1125
               _cx             =   1984
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Name:"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   3435
            Left            =   11280
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   15
            Width           =   8565
            _cx             =   15108
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
            BackColor       =   -2147483645
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
            _GridInfo       =   $"frmWVIIncidents.frx":971D
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.ComboBox comEventSubject 
               Height          =   315
               Left            =   4215
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   107
               Top             =   450
               Width           =   4125
            End
            Begin VB.CheckBox chkPropertyDamagedWV 
               BackColor       =   &H80000003&
               Caption         =   "of World Vision?"
               Enabled         =   0   'False
               Height          =   375
               Left            =   6285
               TabIndex        =   99
               Top             =   1575
               Width           =   2055
            End
            Begin VB.CheckBox chkPropertyDamaged 
               BackColor       =   &H80000003&
               Caption         =   "Property damaged?"
               Height          =   375
               Left            =   4215
               TabIndex        =   98
               Top             =   1575
               Width           =   2070
            End
            Begin VB.ListBox listEventCategories 
               BackColor       =   &H80000003&
               Columns         =   1
               Enabled         =   0   'False
               Height          =   1635
               ItemData        =   "frmWVIIncidents.frx":97F2
               Left            =   225
               List            =   "frmWVIIncidents.frx":9802
               Style           =   1  'Checkbox
               TabIndex        =   96
               Top             =   1575
               Width           =   3825
            End
            Begin RichTextLib.RichTextBox txtEventDesc 
               Height          =   885
               Left            =   4215
               TabIndex        =   44
               Top             =   2325
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   1561
               _Version        =   393217
               TextRTF         =   $"frmWVIIncidents.frx":9812
            End
            Begin VB.TextBox txtEventPropCost 
               BackColor       =   &H80000003&
               Height          =   375
               Left            =   6285
               TabIndex        =   43
               Top             =   1950
               Width           =   2055
            End
            Begin VB.ComboBox comEventVictim 
               Height          =   315
               Left            =   4215
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   42
               Top             =   1200
               Width           =   4125
            End
            Begin VB.ComboBox comEventType 
               Height          =   315
               Left            =   225
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   450
               Width           =   3825
            End
            Begin VB.ComboBox comEventCategory 
               BackColor       =   &H80000003&
               Enabled         =   0   'False
               Height          =   315
               Left            =   1830
               Style           =   2  'Dropdown List
               TabIndex        =   40
               Top             =   825
               Width           =   2220
            End
            Begin C1SizerLibCtl.C1Elastic C1ECostOf 
               Height          =   375
               Left            =   4215
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   1950
               Width           =   2070
               _cx             =   3651
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Cost of Property Damage:"
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
            Begin C1SizerLibCtl.C1Elastic C1EEventTarget 
               Height          =   375
               Left            =   4215
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   825
               Width           =   4125
               _cx             =   7276
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Event Victim:"
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
            Begin C1SizerLibCtl.C1Elastic C1EEventType 
               Height          =   375
               Left            =   225
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   75
               Width           =   3825
               _cx             =   6747
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Event Type:"
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
            Begin C1SizerLibCtl.C1Elastic C1ERiskCategory 
               Height          =   375
               Left            =   225
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   825
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   " Risk Category:"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic30 
               Height          =   375
               Left            =   225
               TabIndex        =   97
               TabStop         =   0   'False
               Top             =   1200
               Width           =   3825
               _cx             =   6747
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Other Applicable Risk Categories:"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic27 
               Height          =   375
               Left            =   4215
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   75
               Width           =   4125
               _cx             =   7276
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Event Subject:"
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
            Index           =   0
            Left            =   1050
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   15
            Width           =   8565
            _cx             =   15108
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
            BackColor       =   -2147483645
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
            _GridInfo       =   $"frmWVIIncidents.frx":98BA
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic C1EDistancesTo 
               Height          =   435
               Left            =   4230
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   645
               Width           =   4110
               _cx             =   7250
               _cy             =   767
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
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
               BackColor       =   -2147483646
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "DISTANCE TO NEAREST:"
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
            Begin C1SizerLibCtl.C1Elastic txtLocationDistToWH 
               Height          =   420
               Left            =   5355
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   1935
               Width           =   2985
               _cx             =   5265
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic txtLocationDistToOpArea 
               Height          =   435
               Left            =   5355
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   1500
               Width           =   2985
               _cx             =   5265
               _cy             =   767
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic txtLocationDistToOffice 
               Height          =   420
               Left            =   5355
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   1080
               Width           =   2985
               _cx             =   5265
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic14 
               Height          =   420
               Left            =   4230
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   1935
               Width           =   1125
               _cx             =   1984
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Warehouse:"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic13 
               Height          =   435
               Left            =   4230
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   1500
               Width           =   1125
               _cx             =   1984
               _cy             =   767
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Op Area:"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic12 
               Height          =   420
               Left            =   4230
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1125
               _cx             =   1984
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Office:"
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
            Begin RichTextLib.RichTextBox txtLocationDesc 
               Height          =   855
               Left            =   4230
               TabIndex        =   28
               Top             =   2355
               Width           =   4110
               _ExtentX        =   7250
               _ExtentY        =   1508
               _Version        =   393217
               TextRTF         =   $"frmWVIIncidents.frx":999C
            End
            Begin C1SizerLibCtl.C1Elastic txtLocationAdmin2 
               Height          =   435
               Left            =   1875
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   1500
               Width           =   2355
               _cx             =   4154
               _cy             =   767
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic txtLocationAdmin3 
               Height          =   420
               Left            =   1875
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   1935
               Width           =   2355
               _cx             =   4154
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic txtLocationAdmin1 
               Height          =   420
               Left            =   1875
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1080
               Width           =   2355
               _cx             =   4154
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic txtLocationX 
               Height          =   420
               Left            =   1875
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   225
               Width           =   2355
               _cx             =   4154
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic lblLocationAdmin3 
               Height          =   420
               Left            =   225
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   1935
               Width           =   1650
               _cx             =   2910
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Tehsil:"
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
            Begin C1SizerLibCtl.C1Elastic lblLocationAdmin2 
               Height          =   435
               Left            =   225
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   1500
               Width           =   1650
               _cx             =   2910
               _cy             =   767
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "District:"
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
            Begin C1SizerLibCtl.C1Elastic lblLocationAdmin1 
               Height          =   420
               Left            =   225
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1650
               _cx             =   2910
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "UNProvince:"
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
            Begin C1SizerLibCtl.C1Elastic C1ECoordinates 
               Height          =   420
               Left            =   225
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   225
               Width           =   1650
               _cx             =   2910
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "X:"
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
            Begin VB.CommandButton cmdLocationSet 
               Caption         =   "-- CLICK HERE TO SET LOCATION --"
               Height          =   420
               Left            =   4230
               TabIndex        =   19
               Top             =   225
               Width           =   4110
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic15 
               Height          =   435
               Left            =   225
               TabIndex        =   108
               TabStop         =   0   'False
               Top             =   645
               Width           =   1650
               _cx             =   2910
               _cy             =   767
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Y:"
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
            Begin C1SizerLibCtl.C1Elastic txtLocationY 
               Height          =   435
               Left            =   1875
               TabIndex        =   109
               TabStop         =   0   'False
               Top             =   645
               Width           =   2355
               _cx             =   4154
               _cy             =   767
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic lblLocationAdmin4 
               Height          =   435
               Left            =   225
               TabIndex        =   112
               TabStop         =   0   'False
               Top             =   2355
               Width           =   1650
               _cx             =   2910
               _cy             =   767
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Union Council:"
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
            Begin C1SizerLibCtl.C1Elastic lblLocationAdmin5 
               Height          =   420
               Left            =   225
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   2790
               Width           =   1650
               _cx             =   2910
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Nearest Settlement:"
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
            Begin C1SizerLibCtl.C1Elastic txtLocationAdmin4 
               Height          =   435
               Left            =   1875
               TabIndex        =   114
               TabStop         =   0   'False
               Top             =   2355
               Width           =   2355
               _cx             =   4154
               _cy             =   767
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
               BackColor       =   -2147483645
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
            Begin C1SizerLibCtl.C1Elastic txtLocationAdmin5 
               Height          =   420
               Left            =   1875
               TabIndex        =   115
               TabStop         =   0   'False
               Top             =   2790
               Width           =   2355
               _cx             =   4154
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
               BackColor       =   -2147483645
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   3435
            Left            =   -9180
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   15
            Width           =   8565
            _cx             =   15108
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
            BackColor       =   -2147483645
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
            _GridInfo       =   $"frmWVIIncidents.frx":9A54
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin XpressEditorsLibCtl.dxTimeEdit txtDetailTime 
               Height          =   315
               Left            =   4275
               OleObjectBlob   =   "frmWVIIncidents.frx":9B35
               TabIndex        =   18
               Top             =   1050
               Width           =   4065
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic11 
               Height          =   405
               Left            =   225
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   2760
               Width           =   4050
               _cx             =   7144
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Responsible office:"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic10 
               Height          =   405
               Left            =   225
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   2355
               Width           =   4050
               _cx             =   7144
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Reported by (source):"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic9 
               Height          =   420
               Left            =   225
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   1935
               Width           =   4050
               _cx             =   7144
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Entered by:"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   405
               Left            =   225
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   1050
               Width           =   4050
               _cx             =   7144
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Time of security event:"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   420
               Left            =   225
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   630
               Width           =   4050
               _cx             =   7144
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Date of security event:"
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
            Begin C1SizerLibCtl.C1Elastic C1EDateOf 
               Height          =   405
               Left            =   225
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   225
               Width           =   4050
               _cx             =   7144
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Date of report:"
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
            Begin VB.ComboBox comDetailResponsibleOffice 
               Height          =   315
               Left            =   4275
               Sorted          =   -1  'True
               TabIndex        =   11
               Text            =   "Combo3"
               Top             =   2760
               Width           =   4065
            End
            Begin VB.ComboBox comDetailReportedBy 
               Height          =   315
               Left            =   4275
               Sorted          =   -1  'True
               TabIndex        =   10
               Text            =   "Combo2"
               Top             =   2355
               Width           =   4065
            End
            Begin VB.ComboBox comDetailEnteredBy 
               Height          =   315
               Left            =   4275
               TabIndex        =   9
               Text            =   "Combo1"
               Top             =   1935
               Width           =   4065
            End
            Begin XpressEditorsLibCtl.dxDateEdit txtDetailDateEvent 
               Height          =   315
               Left            =   4275
               OleObjectBlob   =   "frmWVIIncidents.frx":9C7B
               TabIndex        =   8
               Top             =   630
               Width           =   4065
            End
            Begin XpressEditorsLibCtl.dxDateEdit txtDetailDateReport 
               Height          =   315
               Left            =   4275
               OleObjectBlob   =   "frmWVIIncidents.frx":9D1B
               TabIndex        =   7
               Top             =   225
               Width           =   4065
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic28 
            Height          =   3435
            Left            =   11580
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   15
            Width           =   8565
            _cx             =   15108
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
            BackColor       =   -2147483645
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
            _GridInfo       =   $"frmWVIIncidents.frx":9DBB
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.TextBox txtCasInjuries 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               Height          =   390
               Left            =   3030
               TabIndex        =   89
               Text            =   "txtCasInjuries"
               Top             =   750
               Width           =   1680
            End
            Begin VB.TextBox txtCasDeaths 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               Height          =   390
               Left            =   3030
               TabIndex        =   88
               Text            =   "txtCasDeaths"
               Top             =   1515
               Width           =   1680
            End
            Begin VB.TextBox txtCasInjuriesWV 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               Height          =   375
               Left            =   3030
               TabIndex        =   87
               Text            =   "txtCasInjuriesWV"
               Top             =   1140
               Width           =   1680
            End
            Begin VB.TextBox txtCasDeathsWV 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               Height          =   375
               Left            =   3030
               TabIndex        =   86
               Text            =   "txtCasDeathsWV"
               Top             =   1905
               Width           =   1680
            End
            Begin VB.CheckBox chkCasualties 
               BackColor       =   &H80000003&
               Caption         =   "THE NUMBER OF CASUALTIES ARE KNOWN"
               Height          =   465
               Left            =   225
               TabIndex        =   85
               Top             =   285
               Width           =   8115
            End
            Begin VB.TextBox txtCasCapturedWV 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               Height          =   375
               Left            =   3030
               TabIndex        =   84
               Text            =   "txtCasCapturedWV"
               Top             =   2670
               Width           =   1680
            End
            Begin VB.TextBox txtCasCaptured 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               Height          =   390
               Left            =   3030
               TabIndex        =   83
               Text            =   "txtCasCaptured"
               Top             =   2280
               Width           =   1680
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic29 
               Height          =   375
               Left            =   225
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   1140
               Width           =   2805
               _cx             =   4948
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Injuries (WV):"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic31 
               Height          =   375
               Left            =   225
               TabIndex        =   91
               TabStop         =   0   'False
               Top             =   1905
               Width           =   2805
               _cx             =   4948
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Deaths (WV):"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic32 
               Height          =   390
               Left            =   225
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   1515
               Width           =   2805
               _cx             =   4948
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Deaths (total):"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic39 
               Height          =   390
               Left            =   225
               TabIndex        =   93
               TabStop         =   0   'False
               Top             =   2280
               Width           =   2805
               _cx             =   4948
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Captured (total):"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   375
               Index           =   2
               Left            =   225
               TabIndex        =   94
               TabStop         =   0   'False
               Top             =   2670
               Width           =   2805
               _cx             =   4948
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Captured (WV):"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic16 
               Height          =   390
               Left            =   225
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   750
               Width           =   2805
               _cx             =   4948
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Injuries (total):"
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
   End
End
Attribute VB_Name = "frmWVIIncidents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event GetLocationOnMap()
Private mRSMaster As New adodb.Recordset
Private mRSSubjects As New adodb.Recordset
Private mRSVictims As New adodb.Recordset
Private mCN As adodb.Connection
Private bSaving As Boolean

Private Sub PopulateCombo(Combo1 As ComboBox, _
                          sTableName As String, _
                          Optional sFieldName As String = "option")
        '<EhHeader>
        On Error GoTo PopulateCombo_Err
        '</EhHeader>

        Dim sSQL As String
        Dim oRS As New adodb.Recordset

100     Combo1.Clear
102     sSQL = "SELECT DISTINCT [" & sFieldName & "] FROM [" & sTableName & "] ORDER BY [" & sTableName & "].[" & sFieldName & "]"
    
104     With oRS
    
106         .Open sSQL, mCN, adOpenDynamic, adLockBatchOptimistic
        
108         Do Until .EOF
        
110             If Not IsNull(.Fields(0).Value) Then
112                 Combo1.AddItem .Fields(0).Value
                End If

114             .MoveNext
        
            Loop
     
116         .Close
    
        End With
    
118     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

PopulateCombo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.PopulateCombo " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GuidGUIDForValue(sValue As String, _
                                  sTable As String) As String
        '<EhHeader>
        On Error GoTo GuidGUIDForValue_Err
        '</EhHeader>

        Dim oRS As New adodb.Recordset

100     oRS.Open "SELECT [GUID1] FROM [" & sTable & "] WHERE [option] = '" & sValue & "'", mCN, adOpenDynamic, adLockBatchOptimistic

102     If Not oRS.EOF Then
104         GuidGUIDForValue = oRS.Fields(0).Value
        Else
106         GuidGUIDForValue = ""
        End If
    
108     oRS.Close
110     Set oRS = Nothing

        '<EhFooter>
        Exit Function

GuidGUIDForValue_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.GuidGUIDForValue " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub PopulateList(List1 As ListBox, _
                         sTableName As String, _
                         Optional sFieldName As String = "option")
        '<EhHeader>
        On Error GoTo PopulateList_Err
        '</EhHeader>

        Dim sSQL As String
        Dim oRS As New adodb.Recordset
    
100     List1.Clear
102     sSQL = "SELECT DISTINCT [" & sFieldName & "] FROM [" & sTableName & "] ORDER BY [" & sTableName & "].[" & sFieldName & "]"
    
104     With oRS
    
106         .Open sSQL, mCN, adOpenDynamic, adLockBatchOptimistic
        
108         Do Until .EOF

110             If Not IsNull(.Fields(0).Value) Then
112                 List1.AddItem .Fields(0).Value
                End If

114             .MoveNext
            Loop
     
116         .Close
    
        End With
    
118     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

PopulateList_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.PopulateList " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckDeathFields()
        '<EhHeader>
        On Error GoTo CheckDeathFields_Err
        '</EhHeader>

100     txtCasDeaths = IIf(chkCasualties.Value = vbUnchecked, "", 0)
102     txtCasDeathsWV = IIf(chkCasualties.Value = vbUnchecked, "", 0)
104     txtCasDeaths.Enabled = IIf(chkCasualties.Value = vbUnchecked, False, True)
106     txtCasDeathsWV.Enabled = IIf(chkCasualties.Value = vbUnchecked, False, True)
108     txtCasDeaths.BackColor = IIf(chkCasualties.Value = vbUnchecked, chkCasualties.BackColor, vbWhite)
110     txtCasDeathsWV.BackColor = IIf(chkCasualties.Value = vbUnchecked, chkCasualties.BackColor, vbWhite)

        '<EhFooter>
        Exit Sub

CheckDeathFields_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.CheckDeathFields " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckInjuryFields()
        '<EhHeader>
        On Error GoTo CheckInjuryFields_Err
        '</EhHeader>

100     txtCasInjuries = IIf(chkCasualties.Value = vbUnchecked, "", 0)
102     txtCasInjuriesWV = IIf(chkCasualties.Value = vbUnchecked, "", 0)
104     txtCasInjuries.Enabled = IIf(chkCasualties.Value = vbUnchecked, False, True)
106     txtCasInjuriesWV.Enabled = IIf(chkCasualties.Value = vbUnchecked, False, True)
108     txtCasInjuries.BackColor = IIf(chkCasualties.Value = vbUnchecked, chkCasualties.BackColor, vbWhite)
110     txtCasInjuriesWV.BackColor = IIf(chkCasualties.Value = vbUnchecked, chkCasualties.BackColor, vbWhite)
    
        '<EhFooter>
        Exit Sub

CheckInjuryFields_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.CheckInjuryFields " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckCapturedFields()
        '<EhHeader>
        On Error GoTo CheckCapturedFields_Err
        '</EhHeader>
    
100     txtCasCaptured = IIf(chkCasualties.Value = vbUnchecked, "", 0)
102     txtCasCapturedWV = IIf(chkCasualties.Value = vbUnchecked, "", 0)
104     txtCasCaptured.Enabled = IIf(chkCasualties.Value = vbUnchecked, False, True)
106     txtCasCapturedWV.Enabled = IIf(chkCasualties.Value = vbUnchecked, False, True)
108     txtCasCaptured.BackColor = IIf(chkCasualties.Value = vbUnchecked, chkCasualties.BackColor, vbWhite)
110     txtCasCapturedWV.BackColor = IIf(chkCasualties.Value = vbUnchecked, chkCasualties.BackColor, vbWhite)
    
        '<EhFooter>
        Exit Sub

CheckCapturedFields_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.CheckCapturedFields " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkCasualties_Click()
        '<EhHeader>
        On Error GoTo chkCasualties_Click_Err
        '</EhHeader>
100     Call CheckInjuryFields
102     Call CheckCapturedFields
104     Call CheckDeathFields
        '<EhFooter>
        Exit Sub

chkCasualties_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.chkCasualties_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkPropertyDamaged_Click()
        '<EhHeader>
        On Error GoTo chkPropertyDamaged_Click_Err
        '</EhHeader>

100     If chkPropertyDamaged.Value = vbChecked Then

102         chkPropertyDamagedWV.Enabled = True

        Else
        
104         chkPropertyDamagedWV.Value = vbUnchecked
106         chkPropertyDamagedWV.Enabled = False

        End If
    
108     Call chkPropertyDamagedWV_Click

        '<EhFooter>
        Exit Sub

chkPropertyDamaged_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.chkPropertyDamaged_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkPropertyDamagedWV_Click()
        '<EhHeader>
        On Error GoTo chkPropertyDamagedWV_Click_Err
        '</EhHeader>

100     If chkPropertyDamagedWV.Value = vbChecked Then

102         txtEventPropCost.BackColor = vbWhite
104         txtEventPropCost.Enabled = True
106         txtEventPropCost.Text = 0

        Else

108         txtEventPropCost.BackColor = lblOutcomeImpact.BackColor
110         txtEventPropCost.Enabled = False
112         txtEventPropCost.Text = ""

        End If

        '<EhFooter>
        Exit Sub

chkPropertyDamagedWV_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.chkPropertyDamagedWV_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     Unload Me
 
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdLocationSet_Click()
        '<EhHeader>
        On Error GoTo cmdLocationSet_Click_Err
        '</EhHeader>

100     RaiseEvent GetLocationOnMap

        '<EhFooter>
        Exit Sub

cmdLocationSet_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.cmdLocationSet_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetLocationParams(sAdmin1 As String, _
                             sAdmin2 As String, _
                             sAdmin3 As String, _
                             sAdmin4 As String, _
                             sAdmin5 As String, _
                             sX As String, _
                             sY As String, _
                             sOffice As String, _
                             sWarehouse As String, _
                             sOpArea As String)
        '<EhHeader>
        On Error GoTo SetLocationParams_Err
        '</EhHeader>

100     txtLocationAdmin1.caption = sAdmin1
102     txtLocationAdmin2.caption = sAdmin2
104     txtLocationAdmin3.caption = sAdmin3
106     txtLocationAdmin4.caption = sAdmin4
108     txtLocationAdmin5.caption = sAdmin5
    
110     txtLocationX.caption = sX
112     txtLocationY.caption = sY
114     txtLocationDistToOffice = sOffice
116     txtLocationDistToOpArea = sOpArea
118     txtLocationDistToWH = sWarehouse

        '<EhFooter>
        Exit Sub

SetLocationParams_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.SetLocationParams " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CheckIfListboxTicked(sValue As String, _
                                      List1 As ListBox) As Boolean
        '<EhHeader>
        On Error GoTo CheckIfListboxTicked_Err
        '</EhHeader>

        Dim i As Long
100     i = 0
102     CheckIfListboxTicked = False
    
104     Do Until i = List1.ListCount Or CheckIfListboxTicked = True
    
106         If listEventCategories.List(i) = sValue Then
        
108             If listEventCategories.Selected(i) = True Then CheckIfListboxTicked = True
        
            End If
    
110         i = i + 1
        Loop

        '<EhFooter>
        Exit Function

CheckIfListboxTicked_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.CheckIfListboxTicked " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>

        Dim i As Long
100     bSaving = True

102     With mRSMaster
    
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Page 1: Detail
            '
104         .Fields("DetailReportDate").Value = txtDetailDateReport
106         .Fields("DetailEventDate").Value = txtDetailDateEvent
108         .Fields("DetailTimeOfEvent").Value = CStr(Format(txtDetailTime, "hh:mm"))
110         .Fields("DetailEnteredBy").Value = comDetailEnteredBy.Text
112         .Fields("DetailReportedBy").Value = comDetailReportedBy.Text
114         .Fields("DetailResponsibleOffice").Value = comDetailResponsibleOffice.Text
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Page 2: Location
            '
116         .Fields("Longitude").Value = "0" & txtLocationX
118         .Fields("Latitude").Value = "0" & txtLocationY
120         .Fields("LocationAdmin1").Value = txtLocationAdmin1
122         .Fields("LocationAdmin2").Value = txtLocationAdmin2
124         .Fields("LocationAdmin3").Value = txtLocationAdmin3
126         .Fields("LocationAdmin4").Value = txtLocationAdmin4
128         .Fields("LocationAdmin5").Value = txtLocationAdmin5
130         .Fields("LocationNearestOffice").Value = txtLocationDistToOffice
132         .Fields("LocationNearestOpArea").Value = txtLocationDistToOpArea
134         .Fields("LocationNearestWH").Value = txtLocationDistToWH
136         .Fields("LocationDescription").Value = txtLocationDesc.Text
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Page 3: Event
            '
138         .Fields("EventRiskCategorySP").Value = CheckIfListboxTicked(comEventCategory.List(0), listEventCategories)
140         .Fields("EventRiskCategoryCS").Value = CheckIfListboxTicked(comEventCategory.List(1), listEventCategories)
142         .Fields("EventRiskCategoryCN").Value = CheckIfListboxTicked(comEventCategory.List(2), listEventCategories)
144         .Fields("EventRiskCategoryTR").Value = CheckIfListboxTicked(comEventCategory.List(3), listEventCategories)
146         .Fields("EventRiskCategoryKI").Value = CheckIfListboxTicked(comEventCategory.List(4), listEventCategories)
148         .Fields("EventRiskCategoryHU").Value = CheckIfListboxTicked(comEventCategory.List(5), listEventCategories)
150         .Fields("EventRiskCategoryIN").Value = CheckIfListboxTicked(comEventCategory.List(6), listEventCategories)
        
152         Select Case comEventCategory.Text
        
                Case "Social & Political"
154                 .Fields("EventRiskCategorySP").Value = True

156             Case "Crime & Security"
158                 .Fields("EventRiskCategoryCS").Value = True

160             Case "Conflict"
162                 .Fields("EventRiskCategoryCN").Value = True

164             Case "Terrorism"
166                 .Fields("EventRiskCategoryTR").Value = True

168             Case "Kidnapping"
170                 .Fields("EventRiskCategoryKI").Value = True

172             Case "Humanitarian Space"
174                 .Fields("EventRiskCategoryHU").Value = True

176             Case "Infrastructure"
178                 .Fields("EventRiskCategoryIN").Value = True
        
            End Select
        
180         .Fields("EventPropDamaged").Value = IIf(chkPropertyDamaged.Value = vbChecked, True, False)
182         .Fields("EventPropWV").Value = IIf(chkPropertyDamagedWV.Value = vbChecked, True, False)
184         .Fields("EventPropCost").Value = "0" & txtEventPropCost
186         .Fields("EventEventDesc").Value = txtEventDesc.Text
        
188         .Fields("dd_WVISec_ddEventType").Value = GuidGUIDForValue(comEventType.Text, "dd_WVISec_ddEventType")
190         .Fields("dd_WVISec_ddEventCategory").Value = GuidGUIDForValue(comEventCategory.Text, "dd_WVISec_ddEventCategory")
192         .Fields("dd_WVISec_ddEventSubject").Value = GuidGUIDForValue(comEventSubject.Text, "dd_WVISec_ddEventSubject")
194         .Fields("dd_WVISec_ddEventVictim").Value = GuidGUIDForValue(comEventVictim.Text, "dd_WVISec_ddEventVictim")
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Page 4: Casualities
            '
196         .Fields("Casualties").Value = IIf(chkCasualties.Value = vbChecked, True, False)
198         .Fields("CasualtiesInjuries").Value = "0" & txtCasInjuries
200         .Fields("CasualtiesInjuriesWV").Value = "0" & txtCasInjuriesWV
202         .Fields("CasualtiesDeaths").Value = "0" & txtCasDeaths
204         .Fields("CasualtiesDeathsWV").Value = "0" & txtCasDeathsWV
206         .Fields("CasualtiesCaptured").Value = "0" & txtCasCaptured
208         .Fields("CasualtiesCapturedWV").Value = "0" & txtCasCapturedWV
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Page 7: Outcome
            '
210         .Fields("dd_WVISec_ddOutcomeImp").Value = GuidGUIDForValue(comOutcomeImpact.Text, "dd_WVISec_ddOutcomeImp")
212         .Fields("dd_WVISec_ddOutcomeMat").Value = GuidGUIDForValue(comOutcomeMaterial.Text, "dd_WVISec_ddOutcomeMat")
214         .Fields("dd_WVISec_ddOutcomeTre").Value = GuidGUIDForValue(comOutcomeTrend.Text, "dd_WVISec_ddOutcomeTre")
216         .Fields("OutcomeDesc").Value = txtOutcomeDescription.Text

218         SynchHistoryAddNew mCN, GUIDGen, .Fields("GUID1").Value, "WVI Incident Module", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, "dd_WVISec_mastertable", False, "false", "dd_WVISec_"
220         .UpdateBatch adAffectCurrent
            
            'SynchHistoryAddNew GUIDGen, mDDRSGuid, "DD ddDef Add", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, DDDefCurrent.Prefix & DDTableCurrent.TableName, DDTableCurrent.IsGEOTable, "false", DDDefCurrent.Prefix
        End With
    
222     GetGUIDsForRS mRSSubjects, "dd_WVISec_ddGender"
224     GetGUIDsForRS mRSSubjects, "dd_WVISec_ddEventSubject"
226     GetGUIDsForRS mRSSubjects, "dd_WVISec_ddAffiliation"
228     GetGUIDsForRS mRSVictims, "dd_WVISec_ddGender"
230     GetGUIDsForRS mRSVictims, "dd_WVISec_ddEventVictim"
232     GetGUIDsForRS mRSVictims, "dd_WVISec_ddAffiliation"

234     SafeMoveFirst mRSSubjects
236     SafeMoveFirst mRSVictims
        
238     Do Until mRSSubjects.EOF
240         SynchHistoryAddNew mCN, GUIDGen, mRSSubjects.Fields("GUID2").Value, "WVI Incident Module", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, "dd_WVISec_linkSubjects", False, "false", "dd_WVISec_"
242         mRSSubjects.MoveNext
        Loop
        
244     Do Until mRSVictims.EOF
246         SynchHistoryAddNew mCN, GUIDGen, mRSVictims.Fields("GUID2").Value, "WVI Incident Module", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, "dd_WVISec_linkVictims", False, "false", "dd_WVISec_"
248         mRSVictims.MoveNext
        Loop
        
250     SafeMoveFirst mRSSubjects
252     SafeMoveFirst mRSVictims

254     mRSSubjects.UpdateBatch adAffectAllChapters
256     mRSVictims.UpdateBatch adAffectAllChapters
    
258     MsgBox "Saved", vbInformation, "Incident Added"
260     Unload Me

        '<EhFooter>
        Exit Sub

cmdSave_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmWVIIncidents.cmdSave_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GetGUIDsForRS(oRSPassed As adodb.Recordset, _
                          sFieldName As String)
        '<EhHeader>
        On Error GoTo GetGUIDsForRS_Err
        '</EhHeader>

        Dim sSQL As String
        Dim oRSTemp As New adodb.Recordset
    
100     sSQL = "SELECT [GUID1], [option] FROM [" & sFieldName & "]"
102     oRSTemp.Open sSQL, mCN, adOpenDynamic, adLockBatchOptimistic

104     If Not oRSPassed.EOF Then oRSPassed.MoveFirst
    
106     Do Until oRSPassed.EOF
    
108         With oRSTemp
    
110             If Not .EOF Then
    
112                 .MoveFirst
114                 .Find "[option] = '" & oRSPassed.Fields(sFieldName).Value & "'"

116                 If Not .EOF Then
                    
118                     oRSPassed.Fields(sFieldName).Value = .Fields("GUID1").Value
                    
                    End If
                
                End If
    
            End With

120         oRSPassed.MoveNext
        Loop

        '<EhFooter>
        Exit Sub

GetGUIDsForRS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.GetGUIDsForRS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSubjectsAdd_Click()
        '<EhHeader>
        On Error GoTo cmdSubjectsAdd_Click_Err
        '</EhHeader>

100     If Len(txtSubjectsName) > 0 And Len(comSubjectsType.Text) > 0 Then
    
102         With mRSSubjects
        
104             .AddNew
106             .Fields(0).Value = mRSMaster.Fields(0).Value
108             .Fields(1).Value = GUIDGen
110             .Fields(2).Value = txtSubjectsName
112             .Fields(3).Value = comSubjectsType.Text
114             .Fields(4).Value = comSubjectsNationality.Text
116             .Fields(5).Value = txtSubjectsAge
118             .Fields(6).Value = comSubjectsGender.Text
120             .Fields(7).Value = comSubjectsAff.Text
        
            End With
        
        Else
    
122         MsgBox "You must at least specify a name and subject type!", vbInformation, "Not enough data"
        
        End If

        '<EhFooter>
        Exit Sub

cmdSubjectsAdd_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.cmdSubjectsAdd_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSubjectsRemove_Click()
        '<EhHeader>
        On Error GoTo cmdSubjectsRemove_Click_Err
        '</EhHeader>

100     If Not mRSSubjects.RecordCount = 0 And Not mRSSubjects.EOF And Not mRSSubjects.BOF Then mRSSubjects.Delete adAffectCurrent
    
        '<EhFooter>
        Exit Sub

cmdSubjectsRemove_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.cmdSubjectsRemove_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdVictimsAdd_Click()
        '<EhHeader>
        On Error GoTo cmdVictimsAdd_Click_Err
        '</EhHeader>

100     If Len(txtVictimsName) > 0 And Len(comVictimsType.Text) > 0 Then
    
102         With mRSVictims
        
104             .AddNew
106             .Fields(0).Value = mRSMaster.Fields(0).Value
108             .Fields(1).Value = GUIDGen
110             .Fields(2).Value = txtVictimsName
112             .Fields(3).Value = comVictimsType.Text
114             .Fields(4).Value = comVictimsNationality.Text
116             .Fields(5).Value = txtVictimsAge
118             .Fields(6).Value = comVictimsGender.Text
120             .Fields(7).Value = comVictimsAff.Text
        
            End With
        
        Else
    
122         MsgBox "You must at least specify a name and victim type!", vbInformation, "Not enough data"
        
        End If
    
        '<EhFooter>
        Exit Sub

cmdVictimsAdd_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.cmdVictimsAdd_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdVictimsRemove_Click()
        '<EhHeader>
        On Error GoTo cmdVictimsRemove_Click_Err
        '</EhHeader>
    
100     If Not mRSVictims.RecordCount = 0 And Not mRSVictims.EOF And Not mRSVictims.BOF Then mRSVictims.Delete adAffectCurrent

        '<EhFooter>
        Exit Sub

cmdVictimsRemove_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.cmdVictimsRemove_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub comEventCategory_Click()
        '<EhHeader>
        On Error GoTo comEventCategory_Click_Err
        '</EhHeader>

        Dim i As Long
100     i = listEventCategories.ListCount - 1
102     PopulateList listEventCategories, "dd_WVISec_ddEventCategory"

104     Do Until i < 0

106         If listEventCategories.List(i) = comEventCategory.Text Then
108             listEventCategories.RemoveItem i
            End If

110         i = i - 1
        Loop

        '<EhFooter>
        Exit Sub

comEventCategory_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.comEventCategory_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetTextRetValue(sSQL As String)
        '<EhHeader>
        On Error GoTo GetTextRetValue_Err
        '</EhHeader>

        Dim oRS As New adodb.Recordset
    
100     With oRS
    
102         .Open sSQL, mCN, adOpenDynamic, adLockBatchOptimistic

104         If Not .EOF Then GetTextRetValue = .Fields(0).Value
106         .Close
    
        End With
    
108     Set oRS = Nothing
    
        '<EhFooter>
        Exit Function

GetTextRetValue_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.GetTextRetValue " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub comEventType_Click()
        '<EhHeader>
        On Error GoTo comEventType_Click_Err
        '</EhHeader>

100     comEventCategory.Text = GetTextRetValue("SELECT [RiskCategory] FROM [dd_WVISec_ddEventType] WHERE [option] = '" & comEventType & "'")
102     Call comEventCategory_Click
    
104     comEventCategory.Enabled = True
106     listEventCategories.Enabled = True
108     comEventCategory.BackColor = vbWhite
110     listEventCategories.BackColor = vbWhite
   
        '<EhFooter>
        Exit Sub

comEventType_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.comEventType_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub comOutcomeImpact_Click()
        '<EhHeader>
        On Error GoTo comOutcomeImpact_Click_Err
        '</EhHeader>
100     lblOutcomeImpact.caption = GetTextRetValue("SELECT [comment] FROM [dd_WVISec_ddOutcomeImp] WHERE [option] = '" & comOutcomeImpact & "'")
        '<EhFooter>
        Exit Sub

comOutcomeImpact_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.comOutcomeImpact_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub comOutcomeMaterial_Click()
        '<EhHeader>
        On Error GoTo comOutcomeMaterial_Click_Err
        '</EhHeader>
100     lblOutcomeMaterial.caption = GetTextRetValue("SELECT [comment] FROM [dd_WVISec_ddOutcomeMat] WHERE [option] = '" & comOutcomeMaterial & "'")
        '<EhFooter>
        Exit Sub

comOutcomeMaterial_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.comOutcomeMaterial_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub comOutcomeTrend_Click()
        '<EhHeader>
        On Error GoTo comOutcomeTrend_Click_Err
        '</EhHeader>
100     lblOutcomeTrend.caption = GetTextRetValue("SELECT [comment] FROM [dd_WVISec_ddOutcomeTre] WHERE [option] = '" & comOutcomeTrend & "'")
        '<EhFooter>
        Exit Sub

comOutcomeTrend_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.comOutcomeTrend_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGridSubjects_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, _
                                          ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
        '<EhHeader>
        On Error GoTo dxDBGridSubjects_OnChangeNode_Err
        '</EhHeader>

100     If Not bSaving Then
    
102         txtSubjectsName = mRSSubjects.Fields(2).Value
104         comSubjectsType.Text = mRSSubjects.Fields(3).Value
106         comSubjectsNationality.Text = mRSSubjects.Fields(4).Value
108         txtSubjectsAge = mRSSubjects.Fields(5).Value

110         If Not mRSSubjects.Fields(6).Value = "" Then comSubjectsGender.Text = mRSSubjects.Fields(6).Value
112         If Not mRSSubjects.Fields(7).Value = "" Then comSubjectsAff.Text = mRSSubjects.Fields(7).Value
    
        End If

        '<EhFooter>
        Exit Sub

dxDBGridSubjects_OnChangeNode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.dxDBGridSubjects_OnChangeNode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGridVictims_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, _
                                         ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
        '<EhHeader>
        On Error GoTo dxDBGridVictims_OnChangeNode_Err
        '</EhHeader>

100     If Not bSaving Then
    
102         txtVictimsName = mRSVictims.Fields(2).Value
104         comVictimsType.Text = mRSVictims.Fields(3).Value
106         comVictimsNationality.Text = mRSVictims.Fields(4).Value
108         txtVictimsAge = mRSVictims.Fields(5).Value

110         If Not mRSVictims.Fields(6).Value = "" Then comVictimsGender.Text = mRSVictims.Fields(6).Value
112         If Not mRSVictims.Fields(7).Value = "" Then comVictimsAff.Text = mRSVictims.Fields(7).Value
    
        End If

        '<EhFooter>
        Exit Sub

dxDBGridVictims_OnChangeNode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.dxDBGridVictims_OnChangeNode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     bSaving = False
    
102     Set mCN = New adodb.Connection
104     Set mRSMaster = New adodb.Recordset
106     Set mRSSubjects = New adodb.Recordset
108     Set mRSVictims = New adodb.Recordset
    
        Dim oRSAdminNames As adodb.Recordset
110     Set oRSAdminNames = New adodb.Recordset
    
112     mCN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\DynamicData\WorldVision.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
114     mCN.CursorLocation = g_sGlobalCursorLocation
116     mCN.Open
    
118     mRSMaster.Open "SELECT * FROM [dd_WVISec_mastertable] WHERE [GUID1] = 'dude-undude'", mCN, adOpenDynamic, adLockBatchOptimistic
120     mRSSubjects.Open "SELECT * FROM [dd_WVISec_linkSubjects] WHERE [GUID1] = 'dude-undude'", mCN, adOpenDynamic, adLockBatchOptimistic
122     mRSVictims.Open "SELECT * FROM [dd_WVISec_linkVictims] WHERE [GUID1] = 'dude-undude'", mCN, adOpenDynamic, adLockBatchOptimistic
124     oRSAdminNames.Open "SELECT * FROM [dd_WVISec_ddAdminNames] ORDER BY [option]", mCN, adOpenDynamic, adLockBatchOptimistic
    
126     With oRSAdminNames
    
128         Do Until .EOF
    
130             Select Case .Fields("option").Value
        
                    Case "Admin1"

132                     If .Fields("IsUsed").Value = False Then
134                         lblLocationAdmin1.Visible = False
136                         txtLocationAdmin1.Visible = False
                        Else
138                         lblLocationAdmin1.Visible = True
140                         txtLocationAdmin1.Visible = True
142                         lblLocationAdmin1.caption = .Fields("LayerCaption").Value
144                         lblLocationAdmin1.Tag = .Fields("NameOfField").Value
146                         txtLocationAdmin1.Tag = .Fields("LayerName").Value
                        End If
                    
148                 Case "Admin2"
            
150                     If .Fields("IsUsed").Value = False Then
152                         lblLocationAdmin2.Visible = False
154                         txtLocationAdmin2.Visible = False
                        Else
156                         lblLocationAdmin2.Visible = True
158                         txtLocationAdmin2.Visible = True
160                         lblLocationAdmin2.caption = .Fields("LayerCaption").Value
162                         lblLocationAdmin2.Tag = .Fields("NameOfField").Value
164                         txtLocationAdmin2.Tag = .Fields("LayerName").Value
                        End If

166                 Case "Admin3"
            
168                     If .Fields("IsUsed").Value = False Then
170                         lblLocationAdmin3.Visible = False
172                         txtLocationAdmin3.Visible = False
                        Else
174                         lblLocationAdmin3.Visible = True
176                         txtLocationAdmin3.Visible = True
178                         lblLocationAdmin3.caption = .Fields("LayerCaption").Value
180                         lblLocationAdmin3.Tag = .Fields("NameOfField").Value
182                         txtLocationAdmin3.Tag = .Fields("LayerName").Value
                        End If

184                 Case "Admin4"
            
186                     If .Fields("IsUsed").Value = False Then
188                         lblLocationAdmin4.Visible = False
190                         txtLocationAdmin4.Visible = False
                        Else
192                         lblLocationAdmin4.Visible = True
194                         txtLocationAdmin4.Visible = True
196                         lblLocationAdmin4.caption = .Fields("LayerCaption").Value
198                         lblLocationAdmin4.Tag = .Fields("NameOfField").Value
200                         txtLocationAdmin4.Tag = .Fields("LayerName").Value
                        End If

202                 Case "Admin5"
        
204                     If .Fields("IsUsed").Value = False Then
206                         lblLocationAdmin5.Visible = False
208                         txtLocationAdmin5.Visible = False
                        Else
210                         lblLocationAdmin5.Visible = True
212                         txtLocationAdmin5.Visible = True
214                         lblLocationAdmin5.caption = .Fields("LayerCaption").Value
216                         lblLocationAdmin5.Tag = .Fields("NameOfField").Value
218                         txtLocationAdmin5.Tag = .Fields("LayerName").Value
                        End If

                End Select
    
220             .MoveNext
            Loop
        
222         .Close
    
        End With
    
224     Set oRSAdminNames = Nothing
    
226     mRSMaster.AddNew
228     mRSMaster.Fields(0).Value = GUIDGen
    
230     dxDBGridSubjects.Columns.DestroyColumns
232     dxDBGridSubjects.KeyField = "GUID2"
234     Set dxDBGridSubjects.DataSource = mRSSubjects
236     dxDBGridSubjects.Columns.RetrieveFields
238     dxDBGridSubjects.Columns(0).Visible = False
240     dxDBGridSubjects.Columns(1).Visible = False
242     dxDBGridSubjects.Columns(2).Visible = True
244     dxDBGridSubjects.Columns(3).Visible = True
246     dxDBGridSubjects.Columns(4).Visible = False
248     dxDBGridSubjects.Columns(5).Visible = False
250     dxDBGridSubjects.Columns(6).Visible = False
252     dxDBGridSubjects.Columns(7).Visible = False
    
254     dxDBGridVictims.Columns.DestroyColumns
256     dxDBGridVictims.KeyField = "GUID2"
258     Set dxDBGridVictims.DataSource = mRSVictims
260     dxDBGridVictims.Columns.RetrieveFields
262     dxDBGridVictims.Columns(0).Visible = False
264     dxDBGridVictims.Columns(1).Visible = False
266     dxDBGridVictims.Columns(2).Visible = True
268     dxDBGridVictims.Columns(3).Visible = True
270     dxDBGridVictims.Columns(4).Visible = False
272     dxDBGridVictims.Columns(5).Visible = False
274     dxDBGridVictims.Columns(6).Visible = False
276     dxDBGridVictims.Columns(7).Visible = False

278     PopulateCombo comDetailEnteredBy, "dd_WVISec_mastertable", "DetailEnteredBy"
280     PopulateCombo comDetailReportedBy, "dd_WVISec_mastertable", "DetailReportedBy"
282     PopulateCombo comDetailResponsibleOffice, "dd_WVISec_mastertable", "DetailResponsibleOffice"
284     PopulateCombo comSubjectsNationality, "dd_WVISec_linkSubjects", "SubjectNationality"
286     PopulateCombo comVictimsNationality, "dd_WVISec_linkVictims", "VictimNationality"
288     PopulateCombo comOutcomeImpact, "dd_WVISec_ddOutcomeImp"
290     PopulateCombo comOutcomeMaterial, "dd_WVISec_ddOutcomeMat"
292     PopulateCombo comOutcomeTrend, "dd_WVISec_ddOutcomeTre"
294     PopulateCombo comEventCategory, "dd_WVISec_ddEventCategory"
296     PopulateList listEventCategories, "dd_WVISec_ddEventCategory"
298     PopulateCombo comEventSubject, "dd_WVISec_ddEventSubject"
300     PopulateCombo comSubjectsType, "dd_WVISec_ddEventSubject"
302     PopulateCombo comEventType, "dd_WVISec_ddEventType"
304     PopulateCombo comEventVictim, "dd_WVISec_ddEventVictim"
306     PopulateCombo comVictimsType, "dd_WVISec_ddEventVictim"
308     comDetailEnteredBy.Text = g_sUserName
310     comDetailResponsibleOffice.Text = g_sRemoteTablePrefix
312     txtDetailDateEvent = Format(Now(), "Medium Date")
314     txtDetailDateReport = Format(Now(), "Medium Date")
316     txtDetailTime = Time()
318     comDetailReportedBy.Text = ""
320     txtEventPropCost.Text = ""
    
322     chkCasualties.Value = vbUnchecked
    
324     Call CheckCapturedFields
326     Call CheckDeathFields
328     Call CheckInjuryFields
    
330     PopulateCombo comSubjectsGender, "dd_WVISec_ddGender"
332     PopulateCombo comVictimsGender, "dd_WVISec_ddGender"
334     PopulateCombo comSubjectsAff, "dd_WVISec_ddAffiliation"
336     PopulateCombo comVictimsAff, "dd_WVISec_ddAffiliation"
    
338     C1TTab1Tab.CurrTab = 0

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

100     mCN.Close
102     Set mCN = Nothing

104     If mRSSubjects.State = adStateOpen Then mRSSubjects.Close
106     If mRSMaster.State = adStateOpen Then mRSMaster.Close
108     If mRSVictims.State = adStateOpen Then mRSVictims.Close
110     Set mRSSubjects = Nothing
112     Set mRSMaster = Nothing
114     Set mRSVictims = Nothing

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVIIncidents.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function SynchHistoryAddNew(oCn As adodb.Connection, sNewGUID As String, _
                                    sID As String, _
                                    sTitle As String, _
                                    sDescription As String, _
                                    sBy As String, _
                                    sRFC3339DateTime As String, _
                                    sTableName As String, _
                                    bIsGeoLayer As Boolean, _
                                    supdates As String, _
                                    Optional sSynchHistPrefix As String = "") As Boolean
        '<EhHeader>
        On Error GoTo SynchHistoryAddNew_Err
        '</EhHeader>

100     SynchHistoryAddNew = False

        Dim oRS As New adodb.Recordset
102     oRS.Open "SELECT * FROM [" & sSynchHistPrefix & "SynchHistory] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "'", oCn, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If .RecordCount > 0 Then
                    
110                 m_frmDebug.DebugPrint "(frmWviIncidents.SynchHistoryAddNew) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be 0 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  These records will be deleted.  If this error message persists please contact an OASIS Developer"

112                 Do Until .EOF
114                     .Delete adAffectCurrent
116                     .UpdateBatch adAffectCurrent
118                     .MoveNext
                    Loop
                            
                End If
    
120             If .EOF Then
                    
122                 .AddNew
124                 .Fields("sID").Value = sID
126                 .Fields("sGUID").Value = sNewGUID
128                 .Fields("sTableName").Value = sTableName
130                 .Fields("swhen").Value = sRFC3339DateTime
132                 .Fields("sStatus").Value = "pending"
134                 .Fields("sequence").Value = 1
136                 .Fields("sBy").Value = "[" & sTitle & "] " & sDescription & " (" & sBy & ")"
138                 .Fields("sdelete").Value = "false"
140                 .Fields("updates").Value = supdates
142                 .Fields("noconflict").Value = "false"
144                 .UpdateBatch adAffectCurrent
146                 SynchHistoryAddNew = True
            
                End If
    
            Else
            
148             m_frmDebug.DebugPrint "(frmWviIncidents.SynchHistoryAddNew) Table [" & sSynchHistPrefix & "SynchHistory] failed to open"
            
            End If
        
        End With
        
150     Set oRS = Nothing
        SynchHistoryAddNew = True
        '<EhFooter>
        Exit Function

SynchHistoryAddNew_Err:
        SynchHistoryAddNew = False
        Set oRS = Nothing
        MsgBox "frmWVIIncidents.SynchHistoryAddNew_Err (line " & Erl & "): " & Err.Description
        'Resume Next
        '</EhFooter>
End Function

