VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmIncidentsV2DataEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Security Incidents v2"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9810
   Icon            =   "frmIncidentsV2DataEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic 
      Height          =   5430
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9810
      _cx             =   17304
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "frmIncidentsV2DataEntry.frx":6852
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
      _GridInfo       =   $"frmIncidentsV2DataEntry.frx":7699
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1ESecurityEvent 
         Height          =   1335
         Left            =   90
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   90
         Width           =   6000
         _cx             =   10583
         _cy             =   2355
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
         Caption         =   " SECURITY INCIDENTS DATA ENTRY"
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
         TabIndex        =   24
         Top             =   5010
         Width           =   3570
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   90
         TabIndex        =   23
         Top             =   5010
         Width           =   3570
      End
      Begin C1SizerLibCtl.C1Tab C1TTab1Tab 
         Height          =   3465
         Left            =   90
         TabIndex        =   7
         Top             =   1485
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
         Caption         =   "Detail|Location|Event|Casualities"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   3435
            Left            =   11280
            TabIndex        =   27
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
            _GridInfo       =   $"frmIncidentsV2DataEntry.frx":7773
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.ComboBox comEventSubject 
               Height          =   315
               Left            =   4215
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   450
               Width           =   4125
            End
            Begin VB.CheckBox chkPropertyDamaged 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Property damaged?"
               Height          =   375
               Left            =   4215
               TabIndex        =   16
               Top             =   1575
               Width           =   2070
            End
            Begin VB.ListBox listEventCategories 
               BackColor       =   &H00FFFFFF&
               Columns         =   1
               Enabled         =   0   'False
               Height          =   1635
               ItemData        =   "frmIncidentsV2DataEntry.frx":7848
               Left            =   225
               List            =   "frmIncidentsV2DataEntry.frx":7858
               Style           =   1  'Checkbox
               TabIndex        =   13
               Top             =   1575
               Width           =   3825
            End
            Begin RichTextLib.RichTextBox txtEventDesc 
               Height          =   1260
               Left            =   4215
               TabIndex        =   17
               Top             =   1950
               Width           =   4125
               _ExtentX        =   7276
               _ExtentY        =   2223
               _Version        =   393217
               TextRTF         =   $"frmIncidentsV2DataEntry.frx":7868
            End
            Begin VB.ComboBox comEventVictim 
               Height          =   315
               Left            =   4215
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   1200
               Width           =   4125
            End
            Begin VB.ComboBox comEventType 
               Height          =   315
               Left            =   225
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   450
               Width           =   3825
            End
            Begin VB.ComboBox comEventCategory 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   1830
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   825
               Width           =   2220
            End
            Begin C1SizerLibCtl.C1Elastic C1EEventTarget 
               Height          =   375
               Left            =   4215
               TabIndex        =   43
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
               BackColor       =   16777215
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
               TabIndex        =   42
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
               BackColor       =   16777215
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
               TabIndex        =   41
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
               BackColor       =   16777215
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
               TabIndex        =   49
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
               BackColor       =   16777215
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
               TabIndex        =   50
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
               BackColor       =   16777215
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
            TabIndex        =   26
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
            _GridInfo       =   $"frmIncidentsV2DataEntry.frx":7910
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin XpressEditorsLibCtl.dxTextEdit txtLocationMGRS 
               Height          =   315
               Left            =   6345
               OleObjectBlob   =   "frmIncidentsV2DataEntry.frx":79F2
               TabIndex        =   56
               Top             =   645
               Width           =   1995
            End
            Begin VB.ListBox lstLocations 
               Height          =   840
               Left            =   4230
               Sorted          =   -1  'True
               TabIndex        =   9
               Top             =   1500
               Width           =   4110
            End
            Begin C1SizerLibCtl.C1Elastic C1EDistancesTo 
               Height          =   420
               Left            =   4230
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   1080
               Width           =   4110
               _cx             =   7250
               _cy             =   741
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
               BackColor       =   16777215
               ForeColor       =   64
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
            Begin RichTextLib.RichTextBox txtLocationDesc 
               Height          =   855
               Left            =   4230
               TabIndex        =   10
               Top             =   2355
               Width           =   4110
               _ExtentX        =   7250
               _ExtentY        =   1508
               _Version        =   393217
               TextRTF         =   $"frmIncidentsV2DataEntry.frx":7A4A
            End
            Begin C1SizerLibCtl.C1Elastic txtLocationAdmin2 
               Height          =   435
               Left            =   1875
               TabIndex        =   39
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
               BackColor       =   16777215
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
               TabIndex        =   38
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
               BackColor       =   16777215
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
               TabIndex        =   37
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
               BackColor       =   16777215
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
            Begin C1SizerLibCtl.C1Elastic lblLocationAdmin 
               Height          =   420
               Index           =   2
               Left            =   225
               TabIndex        =   36
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
               BackColor       =   16777215
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
            Begin C1SizerLibCtl.C1Elastic lblLocationAdmin 
               Height          =   435
               Index           =   1
               Left            =   225
               TabIndex        =   35
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
               BackColor       =   16777215
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
            Begin C1SizerLibCtl.C1Elastic lblLocationAdmin 
               Height          =   420
               Index           =   0
               Left            =   225
               TabIndex        =   34
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
               BackColor       =   16777215
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
            Begin VB.CommandButton cmdLocationSet 
               Caption         =   "-- CLICK HERE TO SET LOCATION --"
               Height          =   420
               Left            =   4230
               TabIndex        =   8
               Top             =   225
               Width           =   4110
            End
            Begin C1SizerLibCtl.C1Elastic lblLocationAdmin 
               Height          =   435
               Index           =   3
               Left            =   225
               TabIndex        =   51
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
               BackColor       =   16777215
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
            Begin C1SizerLibCtl.C1Elastic lblLocationAdmin 
               Height          =   420
               Index           =   4
               Left            =   225
               TabIndex        =   52
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
               BackColor       =   16777215
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
               TabIndex        =   53
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
               BackColor       =   16777215
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
               TabIndex        =   54
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
               BackColor       =   16777215
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
            Begin C1SizerLibCtl.C1Elastic C1EMGRS 
               Height          =   435
               Left            =   4230
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   645
               Width           =   2115
               _cx             =   3731
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
               BackColor       =   16777215
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "MGRS:"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   855
               Left            =   225
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   225
               Width           =   4005
               _cx             =   7064
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
               GridRows        =   2
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   0
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmIncidentsV2DataEntry.frx":7B02
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.CommandButton cmdUpdateLocation 
                  Caption         =   "Update location manually"
                  Height          =   765
                  Left            =   2535
                  TabIndex        =   64
                  Top             =   45
                  Width           =   1290
               End
               Begin XpressEditorsLibCtl.dxTextEdit txtLocationX 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0.00000000"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  Height          =   315
                  Left            =   810
                  OleObjectBlob   =   "frmIncidentsV2DataEntry.frx":7B61
                  TabIndex        =   60
                  Top             =   45
                  Width           =   1725
               End
               Begin C1SizerLibCtl.C1Elastic C1EX 
                  Height          =   390
                  Left            =   45
                  TabIndex        =   61
                  TabStop         =   0   'False
                  Top             =   45
                  Width           =   765
                  _cx             =   1349
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
               Begin C1SizerLibCtl.C1Elastic C1EY 
                  Height          =   375
                  Left            =   45
                  TabIndex        =   62
                  TabStop         =   0   'False
                  Top             =   435
                  Width           =   765
                  _cx             =   1349
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
               Begin XpressEditorsLibCtl.dxTextEdit txtLocationY 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "0.00000000"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
                  Height          =   315
                  Left            =   810
                  OleObjectBlob   =   "frmIncidentsV2DataEntry.frx":7BB9
                  TabIndex        =   63
                  Top             =   435
                  Width           =   1725
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   3435
            Left            =   -9180
            TabIndex        =   25
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
            _GridInfo       =   $"frmIncidentsV2DataEntry.frx":7C11
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin XpressEditorsLibCtl.dxTextEdit txtSourceLink 
               Height          =   315
               Left            =   4275
               OleObjectBlob   =   "frmIncidentsV2DataEntry.frx":7CF2
               TabIndex        =   58
               Top             =   2355
               Width           =   4065
            End
            Begin XpressEditorsLibCtl.dxTimeEdit txtDetailTime 
               Height          =   315
               Left            =   4275
               OleObjectBlob   =   "frmIncidentsV2DataEntry.frx":7D4A
               TabIndex        =   3
               Top             =   1050
               Width           =   4065
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic11 
               Height          =   405
               Left            =   225
               TabIndex        =   33
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
               BackColor       =   16777215
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
               Height          =   420
               Left            =   225
               TabIndex        =   32
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
               BackColor       =   16777215
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
               Height          =   480
               Left            =   225
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   1455
               Width           =   4050
               _cx             =   7144
               _cy             =   847
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
               TabIndex        =   30
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
               BackColor       =   16777215
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
               TabIndex        =   29
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
               BackColor       =   16777215
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
               TabIndex        =   28
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
               BackColor       =   16777215
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
               TabIndex        =   6
               Text            =   "Combo3"
               Top             =   2760
               Width           =   4065
            End
            Begin VB.ComboBox comDetailReportedBy 
               Height          =   315
               Left            =   4275
               Sorted          =   -1  'True
               TabIndex        =   5
               Text            =   "Combo2"
               Top             =   1935
               Width           =   4065
            End
            Begin VB.ComboBox comDetailEnteredBy 
               Height          =   315
               Left            =   4275
               TabIndex        =   4
               Text            =   "Combo1"
               Top             =   1455
               Width           =   4065
            End
            Begin XpressEditorsLibCtl.dxDateEdit txtDetailDateEvent 
               Height          =   315
               Left            =   4275
               OleObjectBlob   =   "frmIncidentsV2DataEntry.frx":7E90
               TabIndex        =   2
               Top             =   630
               Width           =   4065
            End
            Begin XpressEditorsLibCtl.dxDateEdit txtDetailDateReport 
               Height          =   315
               Left            =   4275
               OleObjectBlob   =   "frmIncidentsV2DataEntry.frx":7F30
               TabIndex        =   1
               Top             =   225
               Width           =   4065
            End
            Begin C1SizerLibCtl.C1Elastic C1ESourceWeb 
               Height          =   405
               Left            =   225
               TabIndex        =   57
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
               BackColor       =   16777215
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Source web link (if applicable):"
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
            Left            =   11580
            TabIndex        =   45
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
            _GridInfo       =   $"frmIncidentsV2DataEntry.frx":7FD0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.TextBox txtCasInjuries 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   390
               Left            =   3030
               TabIndex        =   19
               Text            =   "txtCasInjuries"
               Top             =   750
               Width           =   1680
            End
            Begin VB.TextBox txtCasDeaths 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   3030
               TabIndex        =   20
               Text            =   "txtCasDeaths"
               Top             =   1140
               Width           =   1680
            End
            Begin VB.CheckBox chkCasualties 
               BackColor       =   &H00FFFFFF&
               Caption         =   "THE NUMBER OF CASUALTIES ARE KNOWN"
               Height          =   465
               Left            =   225
               TabIndex        =   18
               Top             =   285
               Width           =   8115
            End
            Begin VB.TextBox txtCasCaptured 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   390
               Left            =   3030
               TabIndex        =   21
               Text            =   "txtCasCaptured"
               Top             =   1515
               Width           =   1680
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic32 
               Height          =   375
               Left            =   225
               TabIndex        =   46
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
               BackColor       =   16777215
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
               TabIndex        =   47
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
               BackColor       =   16777215
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic16 
               Height          =   390
               Left            =   225
               TabIndex        =   48
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
               BackColor       =   16777215
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
            Begin RichTextLib.RichTextBox txtCasualtyInfo 
               Height          =   765
               Left            =   225
               TabIndex        =   22
               Top             =   2280
               Width           =   8115
               _ExtentX        =   14314
               _ExtentY        =   1349
               _Version        =   393217
               TextRTF         =   $"frmIncidentsV2DataEntry.frx":80A7
            End
         End
      End
   End
End
Attribute VB_Name = "frmIncidentsV2DataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event CloseWindow()
Public Event GetLocationOnMap()
Public Event RefreshMap()
Public Event GetXYFromMGRS(dX As Double, dY As Double, sMGRS As String)
Public Event GetMGRSFromXY(dX As Double, dY As Double, sMGRS As String)
Public Event UpdateNearstPlaces(m_dPTG As XGIS_Point)
Private mRSMaster As New ADODB.Recordset
Private mRSLinkNearbyLoc As New ADODB.Recordset
Private mCN As ADODB.Connection
Private bSaving As Boolean

Private dOldX As Double
Private dOldY As Double
Private sOldMGRS As String

Private Sub SetOldCoordVals()
    dOldX = txtLocationX
    dOldY = txtLocationY
    DOLDMGRS = txtLocationMGRS
End Sub

Private Sub PopulateCombo(Combo1 As ComboBox, _
                          sTableName As String, _
                          Optional sFieldName As String = "option")
        '<EhHeader>
        On Error GoTo PopulateCombo_Err
        '</EhHeader>

        Dim sSQL As String
        Dim oRS As New ADODB.Recordset

100     Combo1.Clear
102     sSQL = "SELECT DISTINCT [" & sFieldName & "] FROM [" & sTableName & "] ORDER BY [" & sTableName & "].[" & sFieldName & "]"
    
104     With oRS
    
106         .Open sSQL, mCN, adOpenDynamic, adLockBatchOptimistic
        
108         Do Until .EOF
        
110             If Not IsNull(.Fields(0).value) Then
112                 Combo1.AddItem .Fields(0).value
                End If

114             .MoveNext
        
            Loop
     
116         .Close
    
        End With
    
118     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

PopulateCombo_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.PopulateCombo " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GuidGUIDForValue(sValue As String, _
                                  sTable As String) As String
        '<EhHeader>
        On Error GoTo GuidGUIDForValue_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset

100     oRS.Open "SELECT [GUID1] FROM [" & sTable & "] WHERE [option] = '" & sValue & "'", mCN, adOpenDynamic, adLockBatchOptimistic

102     If Not oRS.EOF Then
104         GuidGUIDForValue = oRS.Fields(0).value
        Else
106         GuidGUIDForValue = ""
        End If
    
108     oRS.Close
110     Set oRS = Nothing

        '<EhFooter>
        Exit Function

GuidGUIDForValue_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.GuidGUIDForValue " & "at line " & Erl
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
        Dim oRS As New ADODB.Recordset
    
100     List1.Clear
102     sSQL = "SELECT DISTINCT [" & sFieldName & "] FROM [" & sTableName & "] ORDER BY [" & sTableName & "].[" & sFieldName & "]"
    
104     With oRS
    
106         .Open sSQL, mCN, adOpenDynamic, adLockBatchOptimistic
        
108         Do Until .EOF

110             If Not IsNull(.Fields(0).value) Then
112                 List1.AddItem .Fields(0).value
                End If

114             .MoveNext
            Loop
     
116         .Close
    
        End With
    
118     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

PopulateList_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.PopulateList " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckDeathFields()
        '<EhHeader>
        On Error GoTo CheckDeathFields_Err
        '</EhHeader>

100     txtCasDeaths = IIf(chkCasualties.value = vbUnchecked, "", 0)
102     txtCasDeaths.Enabled = IIf(chkCasualties.value = vbUnchecked, False, True)
104     txtCasDeaths.BackColor = IIf(chkCasualties.value = vbUnchecked, chkCasualties.BackColor, vbWhite)

        '<EhFooter>
        Exit Sub

CheckDeathFields_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.CheckDeathFields " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckInjuryFields()
        '<EhHeader>
        On Error GoTo CheckInjuryFields_Err
        '</EhHeader>

100     txtCasInjuries = IIf(chkCasualties.value = vbUnchecked, "", 0)
102     txtCasInjuries.Enabled = IIf(chkCasualties.value = vbUnchecked, False, True)
104     txtCasInjuries.BackColor = IIf(chkCasualties.value = vbUnchecked, chkCasualties.BackColor, vbWhite)
    
        '<EhFooter>
        Exit Sub

CheckInjuryFields_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.CheckInjuryFields " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CheckCapturedFields()
        '<EhHeader>
        On Error GoTo CheckCapturedFields_Err
        '</EhHeader>
    
100     txtCasCaptured = IIf(chkCasualties.value = vbUnchecked, "", 0)
102     txtCasCaptured.Enabled = IIf(chkCasualties.value = vbUnchecked, False, True)
104     txtCasCaptured.BackColor = IIf(chkCasualties.value = vbUnchecked, chkCasualties.BackColor, vbWhite)
    
        '<EhFooter>
        Exit Sub

CheckCapturedFields_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.CheckCapturedFields " & "at line " & Erl
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
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.chkCasualties_Click " & "at line " & Erl
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
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.cmdCancel_Click " & "at line " & Erl
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
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.cmdLocationSet_Click " & "at line " & Erl
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
                             sMGRS As String)
        '<EhHeader>
        On Error GoTo SetLocationParams_Err
        '</EhHeader>

100     txtLocationAdmin1.caption = sAdmin1
102     txtLocationAdmin2.caption = sAdmin2
104     txtLocationAdmin3.caption = sAdmin3
106     txtLocationAdmin4.caption = sAdmin4
108     txtLocationAdmin5.caption = sAdmin5
    
110     txtLocationX = sX
112     txtLocationY = sY
        txtLocationMGRS = sMGRS
        '<EhFooter>
        Exit Sub

SetLocationParams_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.SetLocationParams " & "at line " & Erl
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
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.CheckIfListboxTicked " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>

        Dim i As Long
        Dim sGUID2 As String
100     bSaving = True

102     With mRSMaster
    
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Page 1: Detail
            '
104         .Fields("DetailReportDate").value = txtDetailDateReport
106         .Fields("DetailEventDate").value = txtDetailDateEvent
108         .Fields("DetailTimeOfEvent").value = CStr(Format(txtDetailTime, "hh:mm"))
110         .Fields("DetailEnteredBy").value = comDetailEnteredBy.Text
112         .Fields("DetailReportedBy").value = comDetailReportedBy.Text
114         .Fields("DetailResponsibleOffice").value = comDetailResponsibleOffice.Text
        
            If DoFieldExists(mRSMaster, "DetailReportedByWebLink") Then
                .Fields("DetailReportedByWebLink").value = txtSourceLink
            End If
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Page 2: Location
            '
116         .Fields("Longitude").value = "0" & txtLocationX
118         .Fields("Latitude").value = "0" & txtLocationY
120         .Fields("LocationAdmin1").value = txtLocationAdmin1
122         .Fields("LocationAdmin2").value = txtLocationAdmin2
124         .Fields("LocationAdmin3").value = txtLocationAdmin3
126         .Fields("LocationAdmin4").value = txtLocationAdmin4
128         .Fields("LocationAdmin5").value = txtLocationAdmin5
130         .Fields("LocationDescription").value = txtLocationDesc.Text

            If DoFieldExists(mRSMaster, "WKT") Then
                .Fields("WKT").value = "POINT (" & txtLocationX & " " & txtLocationY & ")"
            End If
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Page 3: Event
            '
132         .Fields("EventRiskCategorySP").value = CheckIfListboxTicked(comEventCategory.List(0), listEventCategories)
134         .Fields("EventRiskCategoryCS").value = CheckIfListboxTicked(comEventCategory.List(1), listEventCategories)
136         .Fields("EventRiskCategoryCN").value = CheckIfListboxTicked(comEventCategory.List(2), listEventCategories)
138         .Fields("EventRiskCategoryTR").value = CheckIfListboxTicked(comEventCategory.List(3), listEventCategories)
140         .Fields("EventRiskCategoryKI").value = CheckIfListboxTicked(comEventCategory.List(4), listEventCategories)
142         .Fields("EventRiskCategoryHU").value = CheckIfListboxTicked(comEventCategory.List(5), listEventCategories)
144         .Fields("EventRiskCategoryIN").value = CheckIfListboxTicked(comEventCategory.List(6), listEventCategories)
        
146         Select Case comEventCategory.Text
        
                Case "Social & Political"
148                 .Fields("EventRiskCategorySP").value = True

150             Case "Crime & Security"
152                 .Fields("EventRiskCategoryCS").value = True

154             Case "Conflict"
156                 .Fields("EventRiskCategoryCN").value = True

158             Case "Terrorism"
160                 .Fields("EventRiskCategoryTR").value = True

162             Case "Kidnapping"
164                 .Fields("EventRiskCategoryKI").value = True

166             Case "Humanitarian Space"
168                 .Fields("EventRiskCategoryHU").value = True

170             Case "Infrastructure"
172                 .Fields("EventRiskCategoryIN").value = True
        
            End Select
        
174         .Fields("EventPropDamaged").value = IIf(chkPropertyDamaged.value = vbChecked, True, False)
176         .Fields("EventEventDesc").value = txtEventDesc.Text
        
178         .Fields("dd_" & g_sIncidentsV2DDName & "_ddEventType").value = GuidGUIDForValue(comEventType.Text, "dd_" & g_sIncidentsV2DDName & "_ddEventType")
180         .Fields("dd_" & g_sIncidentsV2DDName & "_ddEventCategory").value = GuidGUIDForValue(comEventCategory.Text, "dd_" & g_sIncidentsV2DDName & "_ddEventCategory")
182         .Fields("dd_" & g_sIncidentsV2DDName & "_ddEventSubject").value = GuidGUIDForValue(comEventSubject.Text, "dd_" & g_sIncidentsV2DDName & "_ddEventSubject")
184         .Fields("dd_" & g_sIncidentsV2DDName & "_ddEventVictim").value = GuidGUIDForValue(comEventVictim.Text, "dd_" & g_sIncidentsV2DDName & "_ddEventVictim")
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Page 4: Casualities
            '
186         .Fields("Casualties").value = IIf(chkCasualties.value = vbChecked, True, False)
188         .Fields("CasualtiesInjuries").value = "0" & txtCasInjuries
190         .Fields("CasualtiesDeaths").value = "0" & txtCasDeaths
192         .Fields("CasualtiesCaptured").value = "0" & txtCasCaptured
194         .Fields("CasualtiesDesc").value = txtCasualtyInfo.Text
            
196         SynchHistoryAddNew mCN, GUIDGen, .Fields("GUID1").value, "Incidents v2 Module", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, "dd_" & g_sIncidentsV2DDName & "_mastertable", False, "false", "dd_" & g_sIncidentsV2DDName & "_"
198         .UpdateBatch adAffectCurrent
        End With
        
200     i = 0

202     Do Until i = lstLocations.ListCount
            sGUID2 = GUIDGen
204         mRSLinkNearbyLoc.AddNew
206         mRSLinkNearbyLoc.Fields("GUID1").value = mRSMaster.Fields("GUID1").value
208         mRSLinkNearbyLoc.Fields("GUID2").value = sGUID2
210         mRSLinkNearbyLoc.Fields("NearbyLocation").value = lstLocations.List(i)
            mRSLinkNearbyLoc.UpdateBatch adAffectCurrent
            SynchHistoryAddNew mCN, GUIDGen, sGUID2, "Incidents v2 Module", g_sRemoteTablePrefix, g_sUserName, RFC3339DateTime, "dd_" & g_sIncidentsV2DDName & "_linkNearbyLocations", False, "false", "dd_" & g_sIncidentsV2DDName & "_"
212         i = i + 1
        Loop
    
214     MsgBox "Saved", vbInformation, "Incident Added"
216     RaiseEvent RefreshMap
218     Unload Me

        '<EhFooter>
        Exit Sub

cmdSave_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.cmdSave_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GetGUIDsForRS(oRSPassed As ADODB.Recordset, _
                          sFieldName As String)
        '<EhHeader>
        On Error GoTo GetGUIDsForRS_Err
        '</EhHeader>

        Dim sSQL As String
        Dim oRSTemp As New ADODB.Recordset
    
100     sSQL = "SELECT [GUID1], [option] FROM [" & sFieldName & "]"
102     oRSTemp.Open sSQL, mCN, adOpenDynamic, adLockBatchOptimistic

104     If Not oRSPassed.EOF Then oRSPassed.MoveFirst
    
106     Do Until oRSPassed.EOF
    
108         With oRSTemp
    
110             If Not .EOF Then
    
112                 .MoveFirst
114                 .Find "[option] = '" & oRSPassed.Fields(sFieldName).value & "'"

116                 If Not .EOF Then
                    
118                     oRSPassed.Fields(sFieldName).value = .Fields("GUID1").value
                    
                    End If
                
                End If
    
            End With

120         oRSPassed.MoveNext
        Loop

        '<EhFooter>
        Exit Sub

GetGUIDsForRS_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.GetGUIDsForRS " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdUpdateLocation_Click()

    Dim sMGRS As String
    Dim oPoint As New XGIS_Point
    Dim dX As Double
    Dim dY As Double
    
    If dOldX <> txtLocationX Or dOldY <> txtLocationY Then
        
        ' If MsgBox("Do you want to update the MGRS value from these coordinates and refresh the point on the map?", vbYesNo) = vbYes Then
        RaiseEvent GetMGRSFromXY(txtLocationX, txtLocationY, sMGRS)
        txtLocationMGRS = sMGRS
        SetOldCoordVals
        oPoint.Prepare txtLocationX, txtLocationY
        RaiseEvent UpdateNearstPlaces(oPoint)
        ' End If
        
    ElseIf sOldMGRS <> txtLocationMGRS Then
        
        '  If MsgBox("Do you want to update the XY coordinates from this MGRS value and refresh the point on the map?", vbYesNo) = vbYes Then
        RaiseEvent GetXYFromMGRS(dX, dY, txtLocationMGRS)
        txtLocationX = dX
        txtLocationY = dY
        SetOldCoordVals
        oPoint.Prepare txtLocationX, txtLocationY
        RaiseEvent UpdateNearstPlaces(oPoint)
        ' End If
    
    End If
   
End Sub

Private Sub comEventCategory_Click()
        '<EhHeader>
        On Error GoTo comEventCategory_Click_Err
        '</EhHeader>

        Dim i As Long
100     i = listEventCategories.ListCount - 1
102     PopulateList listEventCategories, "dd_" & g_sIncidentsV2DDName & "_ddEventCategory"

104     Do Until i < 0

106         If listEventCategories.List(i) = comEventCategory.Text Then
108             listEventCategories.RemoveItem i
            End If

110         i = i - 1
        Loop

        '<EhFooter>
        Exit Sub

comEventCategory_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.comEventCategory_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetTextRetValue(sSQL As String)
        '<EhHeader>
        On Error GoTo GetTextRetValue_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
    
100     With oRS
    
102         .Open sSQL, mCN, adOpenDynamic, adLockBatchOptimistic

104         If Not .EOF Then GetTextRetValue = .Fields(0).value
106         .Close
    
        End With
    
108     Set oRS = Nothing
    
        '<EhFooter>
        Exit Function

GetTextRetValue_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.GetTextRetValue " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub comEventType_Click()
        '<EhHeader>
        On Error GoTo comEventType_Click_Err
        '</EhHeader>

100     comEventCategory.Text = GetTextRetValue("SELECT dd_" & g_sIncidentsV2DDName & "_ddEventCategory.[option] FROM dd_" & g_sIncidentsV2DDName & "_ddEventType LEFT JOIN dd_" & g_sIncidentsV2DDName & "_ddEventCategory ON dd_" & g_sIncidentsV2DDName & "_ddEventType.dd_" & g_sIncidentsV2DDName & "_ddEventCategory = dd_" & g_sIncidentsV2DDName & "_ddEventCategory.GUID1 WHERE dd_" & g_sIncidentsV2DDName & "_ddEventType.[option] = '" & comEventType & "'")

        'SELECT [dd_CARESec_ddEventCategory] FROM [dd_" & g_sIncidentsV2DDName & "_ddEventType] WHERE [option] = '" & comEventType & "'")
102     Call comEventCategory_Click
    
104     comEventCategory.Enabled = True
106     listEventCategories.Enabled = True
108     comEventCategory.BackColor = vbWhite
110     listEventCategories.BackColor = vbWhite
   
        '<EhFooter>
        Exit Sub

comEventType_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.comEventType_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

100     bSaving = False
        Dim iCount As Integer
        Dim i As Integer
102     Set mCN = New ADODB.Connection
104     Set mRSMaster = New ADODB.Recordset
    
106     mCN.ConnectionString = g_sIncidentsV2ConnectionString
108     mCN.CursorLocation = g_sGlobalCursorLocation
110     mCN.Open
    
112     mRSMaster.Open "SELECT * FROM [dd_" & g_sIncidentsV2DDName & "_mastertable] WHERE [GUID1] = 'dude-undude'", mCN, adOpenDynamic, adLockBatchOptimistic
114     mRSLinkNearbyLoc.Open "SELECT * FROM [dd_" & g_sIncidentsV2DDName & "_linkNearbyLocations] WHERE [GUID1] = 'dude-undude'", mCN, adOpenDynamic, adLockBatchOptimistic
116     iCount = 0
118     i = 0

120     Do Until iCount = 4
122         SafeMoveFirst g_RSAppSettings
124         g_RSAppSettings.Find "SettingName = 'AdminLevel" & iCount & "'"

126         If Not IsNull(g_RSAppSettings.Fields("SettingValue4").value) Then
128             lblLocationAdmin(i).caption = g_RSAppSettings.Fields("SettingValue4").value
130             i = i + 1
            Else
132             lblLocationAdmin(i).caption = ""
            End If

134         iCount = iCount + 1
        Loop

136     SafeMoveFirst g_RSAppSettings
138     g_RSAppSettings.Find "SettingName = 'AdminLocation'"

140     If Not IsNull(g_RSAppSettings.Fields("SettingValue4").value) Then
142         lblLocationAdmin(i).caption = g_RSAppSettings.Fields("SettingValue4").value
144         i = i + 1
        Else
146         lblLocationAdmin(i).caption = ""
        End If
        
148     Do Until i = 5
150         lblLocationAdmin(i).caption = ""
152         i = i + 1
        Loop
            
154     mRSMaster.AddNew
156     mRSMaster.Fields(0).value = GUIDGen
    
158     PopulateCombo comDetailEnteredBy, "dd_" & g_sIncidentsV2DDName & "_mastertable", "DetailEnteredBy"
160     PopulateCombo comDetailReportedBy, "dd_" & g_sIncidentsV2DDName & "_mastertable", "DetailReportedBy"
162     PopulateCombo comDetailResponsibleOffice, "dd_" & g_sIncidentsV2DDName & "_mastertable", "DetailResponsibleOffice"
164     PopulateCombo comEventCategory, "dd_" & g_sIncidentsV2DDName & "_ddEventCategory"
166     PopulateList listEventCategories, "dd_" & g_sIncidentsV2DDName & "_ddEventCategory"
168     PopulateCombo comEventSubject, "dd_" & g_sIncidentsV2DDName & "_ddEventSubject"
170     PopulateCombo comEventType, "dd_" & g_sIncidentsV2DDName & "_ddEventType"
172     PopulateCombo comEventVictim, "dd_" & g_sIncidentsV2DDName & "_ddEventVictim"
174     comDetailEnteredBy.Text = g_sUserName
176     comDetailResponsibleOffice.Text = g_sRemoteTablePrefix
178     txtDetailDateEvent = Format(Now(), "Medium Date")
180     txtDetailDateReport = Format(Now(), "Medium Date")
182     txtDetailTime = Time()
184     comDetailReportedBy.Text = ""
186     chkCasualties.value = vbUnchecked
        txtLocationX = 0
        txtLocationY = 0
        txtLocationMGRS = "not set"
        SetOldCoordVals
188     Call CheckCapturedFields
190     Call CheckDeathFields
192     Call CheckInjuryFields
    
194     C1TTab1Tab.CurrTab = 0

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

100     mCN.Close
102     Set mCN = Nothing

104     If mRSMaster.State = adStateOpen Then mRSMaster.Close
106     Set mRSMaster = Nothing
 
108     If mRSLinkNearbyLoc.State = adStateOpen Then mRSLinkNearbyLoc.Close
110     Set mRSLinkNearbyLoc = Nothing
        RaiseEvent CloseWindow
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.Form_Unload " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function SynchHistoryAddNew(oCn As ADODB.Connection, _
                                    sNewGUID As String, _
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

        Dim oRS As New ADODB.Recordset
102     oRS.Open "SELECT * FROM [" & sSynchHistPrefix & "SynchHistory] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "'", oCn, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If .RecordCount > 0 Then
                    
110                 DebugPrint "(frmIncidentsV2DataEntry.SynchHistoryAddNew) Something dodgy is going on the table: [" & sSynchHistPrefix & "SynchHistory].  There should be 0 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  These records will be deleted.  If this error message persists please contact an OASIS Developer"

112                 Do Until .EOF
114                     .Delete adAffectCurrent
116                     .UpdateBatch adAffectCurrent
118                     .MoveNext
                    Loop
                            
                End If
    
120             If .EOF Then
                    
122                 .AddNew
124                 .Fields("sID").value = sID
126                 .Fields("sGUID").value = sNewGUID
128                 .Fields("sTableName").value = sTableName
130                 .Fields("swhen").value = sRFC3339DateTime
132                 .Fields("sStatus").value = "pending"
134                 .Fields("sequence").value = 1
136                 .Fields("sBy").value = "[" & sTitle & "] " & sDescription & " (" & sBy & ")"
138                 .Fields("sdelete").value = "false"
140                 .Fields("updates").value = supdates
142                 .Fields("noconflict").value = "false"
144                 .UpdateBatch adAffectCurrent
146                 SynchHistoryAddNew = True
            
                End If
    
            Else
            
148             DebugPrint "(frmIncidentsV2DataEntry.SynchHistoryAddNew) Table [" & sSynchHistPrefix & "SynchHistory] failed to open"
            
            End If
        
        End With
        
150     Set oRS = Nothing
152     SynchHistoryAddNew = True
        '<EhFooter>
        Exit Function

SynchHistoryAddNew_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmIncidentsV2DataEntry.SynchHistoryAddNew " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub txtLocationMGRS_KeyPress(KeyAscii As Integer)
  
    Dim dX As Double
    Dim dY As Double
    Dim oPoint As New XGIS_Point
    
    If KeyAscii = 13 Then
    
        If sOldMGRS <> txtLocationMGRS Then
        
            If MsgBox("Do you want to update the XY coordinates from this MGRS value and refresh the point on the map?", vbYesNo) = vbYes Then
                RaiseEvent GetXYFromMGRS(dX, dY, txtLocationMGRS)
                txtLocationX = dX
                txtLocationY = dY
                SetOldCoordVals
                oPoint.Prepare txtLocationX, txtLocationY
                RaiseEvent UpdateNearstPlaces(oPoint)
            End If
        
        End If
    
    End If
    
End Sub

Private Sub txtLocationX_KeyPress(KeyAscii As Integer)
    Dim sMGRS As String
    Dim oPoint As New XGIS_Point
    
    If KeyAscii = 13 Then
    
        If dOldX <> txtLocationMGRS Or dOldY <> txtLocationMGRS Then
        
            If MsgBox("Do you want to update the MGRS value from these coordinates and refresh the point on the map?", vbYesNo) = vbYes Then
                RaiseEvent GetMGRSFromXY(txtLocationX, txtLocationY, sMGRS)
                txtLocationMGRS = sMGRS
                SetOldCoordVals
                oPoint.Prepare txtLocationX, txtLocationY
                RaiseEvent UpdateNearstPlaces(oPoint)
            End If
        
        End If
    
    End If
End Sub

Private Sub txtLocationY_KeyPress(KeyAscii As Integer)

    Dim sMGRS As String
    Dim oPoint As New XGIS_Point
    
    If KeyAscii = 13 Then
    
        If dOldX <> txtLocationMGRS Or dOldY <> txtLocationMGRS Then
        
            If MsgBox("Do you want to update the MGRS value from these coordinates and refresh the point on the map?", vbYesNo) = vbYes Then
                RaiseEvent GetMGRSFromXY(txtLocationX, txtLocationY, sMGRS)
                txtLocationMGRS = sMGRS
                SetOldCoordVals
                oPoint.Prepare txtLocationX, txtLocationY
                RaiseEvent UpdateNearstPlaces(oPoint)
            End If
        
        End If
    
    End If

End Sub

